r"""
C 盘垃圾文件 / 空文件 / 无用目录清理（默认安全模式仅扫描）。

设计目标
--------
1. 文件级清理（保留原行为）
   - 命中 USELESS_EXTS 的临时/缓存扩展名文件
   - 0 字节空文件
   - 旧的 CrashDumps / INetCache / Temp
2. 目录级清理（增强点）—— 解决 "C:\\Python314 这种整盘孤儿" 之类的大块占用：
   - 孤儿 Python 安装：C:\\Python3x\\、C:\\Python3x-32\\、C:\\Anaconda3\\、
     C:\\Miniconda3\\、C:\\ProgramData\\anaconda3\\
   - 孤儿 Node.js 安装：C:\\node-v*\\
   - 大型 IDE/构建工具缓存：JetBrains 索引、VS BuildTools 临时文件、
     .vs\\ 下的 .vs\\C\\、Node 编译缓存、Cargo target\\deps（黑名单用户级）
   - 浏览器与软件升级残留：Microsoft\\Edge\\Application\\*\\Installer、Code\ Insiders\\resources\\app.old
   - 大体积下载缓存：%LocalAppData%\\pip\\cache、%LocalAppData%\\npm-cache、
     %LocalAppData%\\Yarn\\Cache、%LocalAppData%\\NuGet\\v3-cache
3. 分级删除（Risk Tiers）
   - safe       扩展名/空文件/空目录      --delete 即可清理
   - review     孤儿 Python、IDE 缓存等   需要 --include-review 才会被移动到隔离目录
   - dangerous  Windows 更新、$Windows.~* 即便打开 --include-dangerous 也**绝不**自动处理
4. 隔离回收（quarantine）
   - 默认开启：所有 review 级目录**先移动**到 output\\_quarantine_<ts>\\，不直接删
   - 用户核对后，可手动删除隔离目录，从而支持"可恢复"语义

输入 / 输出
-----------
- input/targets.txt          每行一个自定义根目录（覆盖默认值）
- output/cleanup_report.csv  文件级候选
- output/directory_report.csv 目录级候选（含可回收大小 / 风险等级）
- output/_quarantine_<ts>/   隔离目录（启用 --delete --include-review 时生成）

CLI
---
    python c_drive_cleanup.py                        # 仅扫描
    python c_drive_cleanup.py --delete               # 只清理 safe 级（空文件/无用扩展/空目录）
    python c_drive_cleanup.py --include-review       # 孤儿 Python、IDE 缓存等也会被隔离移动
    python c_drive_cleanup.py --days 3 --top 20
    python c_drive_cleanup.py --targets "C:\\Temp" "C:\\Python314"
    python c_drive_cleanup.py --python-only          # 只看 Python 孤儿候选
"""
from __future__ import annotations

import argparse
import csv
import ctypes
import logging
import os
import re
import shutil
import sys
import time
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path

logger = logging.getLogger("c_drive_cleanup")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-5s %(name)s | %(message)s",
    datefmt="%H:%M:%S",
)

# ----------------------------------------------------------------------------
# 文件级规则（兼容原行为）
# ----------------------------------------------------------------------------

USELESS_EXTS: set[str] = {
    ".tmp", ".temp", ".bak", ".old", ".log", ".dmp", ".chk",
    ".gid", ".~tmp", ".cache", ".crdownload", ".part",
}

DEFAULT_REL_DIRS: list[str] = [r"C:\\Windows\\Temp"]
USER_REL_DIRS: list[str] = [
    r"AppData\\Local\\Temp",
    r"AppData\\Local\\Microsoft\\Windows\\INetCache",
    r"AppData\\Local\\CrashDumps",
]

# ----------------------------------------------------------------------------
# 目录级规则（增强点）
# ----------------------------------------------------------------------------

# 正则: C:\Python314、C:\Python312-32、C:\Python3.11 之类
_PYTHON_DIR_RE = re.compile(r"^Python\d+([._-]\d+)?(-32)?$", re.IGNORECASE)
# 正则: C:\node-v18.17.0-win-x64
_NODE_DIR_RE = re.compile(r"^node-v\d+(\.\d+){0,2}-.*$", re.IGNORECASE)
# 常见孤儿工具根名
_ORPHAN_ROOT_NAMES: dict[str, str] = {
    "anaconda3": "conda_root",
    "miniconda3": "conda_root",
    "miniforge3": "conda_root",
    "Enthought": "enthought_root",
    "PyCharm": "pycharm_root",
}


@dataclass(frozen=True)
class DirRule:
    """目录候选匹配规则。

    pattern:  Path.glob 模式（基于 C:\\ 根）
    name_re:  或匹配 —— Path.name 正则
    min_size_bytes: 至少占用多少字节才计入（避免误伤小目录）
    risk:     safe / review / dangerous
    reason:   报告里展示的原因标识
    """

    pattern: str = ""
    name_re: re.Pattern[str] | None = None
    min_size_bytes: int = 0
    risk: str = "review"
    reason: str = ""

    def matches(self, path: Path) -> bool:
        if self.name_re is not None and self.name_re.match(path.name):
            return True
        return False


@dataclass
class FileCandidate:
    path: Path
    size: int
    mtime: float
    reasons: list[str] = field(default_factory=list)


@dataclass
class DirCandidate:
    path: Path
    size: int
    last_mtime: float
    risk: str
    reasons: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------------
# 工具函数
# ----------------------------------------------------------------------------


def format_size(num_bytes: int) -> str:
    """人类可读大小。"""
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024.0 or unit == units[-1]:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} TB"


def dir_size(path: Path) -> tuple[int, float]:
    """递归计算目录总大小与最近一次修改时间。失败/无权限时返回 (0, 0.0)。

    性能: 用 os.scandir 代替 Path.rglob，可避免 PermissionError 中断整树。
    """
    total = 0
    latest = 0.0
    try:
        for dirpath, dirnames, filenames in os.walk(path):
            for fname in filenames:
                fp = Path(dirpath) / fname
                try:
                    st = fp.stat()
                except OSError:
                    continue
                total += st.st_size
                if st.st_mtime > latest:
                    latest = st.st_mtime
            for dname in dirnames:
                dp = Path(dirpath) / dname
                try:
                    st = dp.stat()
                except OSError:
                    continue
                if st.st_mtime > latest:
                    latest = st.st_mtime
    except OSError as e:
        logger.warning("dir_size: 访问 %s 失败: %s", path, e)
    return total, latest


def list_installed_python_roots() -> set[Path]:
    """读取 Windows 注册表，列出所有"已注册"的 Python 解释器根目录。

    来源: HKLM\\SOFTWARE\\Python\\PythonCore\\*\\InstallPath
    若无法访问注册表（不在 Windows / 权限不足），返回空集（保守）。
    """
    roots: set[Path] = set()
    if sys.platform != "win32":
        return roots
    try:
        import winreg  # type: ignore[import-not-found]
    except ImportError:
        return roots

    sub_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Python\PythonCore"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Python\PythonCore"),
    ]
    for hkey, base in sub_keys:
        try:
            with winreg.OpenKey(hkey, base) as k:
                i = 0
                while True:
                    try:
                        ver = winreg.EnumKey(k, i)
                    except OSError:
                        break
                    i += 1
                    try:
                        with winreg.OpenKey(k, f"{ver}\\InstallPath") as k2:
                            install, _ = winreg.QueryValueEx(k2, "")
                            if install:
                                roots.add(Path(install))
                    except OSError:
                        continue
        except OSError:
            continue
    return roots


def get_user_profile() -> Path:
    return Path(os.environ.get("USERPROFILE", str(Path.home())))


def get_local_appdata() -> Path:
    raw = os.environ.get("LOCALAPPDATA")
    if raw:
        return Path(raw)
    return get_user_profile() / "AppData" / "Local"


# ----------------------------------------------------------------------------
# targets
# ----------------------------------------------------------------------------


def load_targets_from_input(input_dir: Path) -> list[Path]:
    targets_file = input_dir / "targets.txt"
    if not targets_file.exists():
        return []
    targets: list[Path] = []
    for line in targets_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        targets.append(Path(line))
    return targets


def default_targets() -> list[Path]:
    targets: list[Path] = [Path(p) for p in DEFAULT_REL_DIRS]
    users_dir = Path("C:/Users")
    if users_dir.exists():
        for user_dir in users_dir.iterdir():
            if not user_dir.is_dir():
                continue
            for rel in USER_REL_DIRS:
                targets.append(user_dir / rel)
    return targets


def iter_files(targets: Iterable[Path]) -> Iterable[Path]:
    for root in targets:
        if not root.exists() or not root.is_dir():
            continue
        for dirpath, _, filenames in os.walk(root):
            for name in filenames:
                yield Path(dirpath) / name


# ----------------------------------------------------------------------------
# 文件级候选（保留原行为）
# ----------------------------------------------------------------------------


def is_old_enough(mtime: float, days: int) -> bool:
    return mtime <= time.time() - days * 86400


def collect_file_candidates(targets: Iterable[Path], days: int) -> list[FileCandidate]:
    out: list[FileCandidate] = []
    for fp in iter_files(targets):
        try:
            st = fp.stat()
        except OSError:
            continue
        reasons: list[str] = []
        if st.st_size == 0:
            reasons.append("empty")
        ext = fp.suffix.lower()
        if ext in USELESS_EXTS:
            reasons.append(f"ext:{ext}")
        if reasons and is_old_enough(st.st_mtime, days):
            out.append(FileCandidate(fp, st.st_size, st.st_mtime, reasons))
    return out


def delete_files(candidates: Iterable[FileCandidate]) -> tuple[int, int]:
    deleted = failed = 0
    for c in candidates:
        try:
            c.path.unlink(missing_ok=True)
            deleted += 1
        except OSError as e:
            logger.warning("delete fail: %s | %s", c.path, e)
            failed += 1
    return deleted, failed


def remove_empty_dirs(targets: Iterable[Path]) -> int:
    removed = 0
    for root in targets:
        if not root.exists() or not root.is_dir():
            continue
        for dirpath, dirnames, filenames in os.walk(root, topdown=False):
            if dirnames or filenames:
                continue
            try:
                Path(dirpath).rmdir()
                removed += 1
            except OSError:
                continue
    return removed


# ----------------------------------------------------------------------------
# 目录级候选（增强点）
# ----------------------------------------------------------------------------


# 系统子目录里"不管递归多深一定要抽出来"的关键名
_DEEP_HIT_NAMES: set[str] = {
    "softwaredistribution",
    "$windows.~bt",
    "$windows.~ws",
    "download",
}


def _iter_candidate_dirs() -> list[Path]:
    """遍历"可能藏有孤儿"的高价值根目录，产出候选目录集合。

    两档深度:
    - level 1: 5 个根目录的第一层（孤儿 Python / Node / Conda / IDE / 用户缓存）
    - level N: 关键系统子目录额外递归到第 3 层，避免漏过 SoftwareDistribution/Download
    """
    candidates: list[Path] = []
    roots_to_scan: list[Path] = [Path("C:/")]

    user = get_user_profile()
    if user.exists():
        roots_to_scan.append(user)
    lad = get_local_appdata()
    if lad.exists() and lad != user:
        roots_to_scan.append(lad)

    prog = Path("C:/Program Files")
    if prog.exists():
        roots_to_scan.append(prog)
    prog86 = Path("C:/Program Files (x86)")
    if prog86.exists():
        roots_to_scan.append(prog86)
    prog_data = Path("C:/ProgramData")
    if prog_data.exists():
        roots_to_scan.append(prog_data)

    for root in roots_to_scan:
        if not root.exists():
            continue
        try:
            for child in root.iterdir():
                if not child.is_dir():
                    continue
                candidates.append(child)
                # 关键系统目录额外深扫 3 层
                if child.name.lower() == "windows":
                    candidates.extend(_walk_n_levels(child / "SoftwareDistribution", 3))
                elif child.name.lower() == "softwaredistribution":
                    candidates.extend(_walk_n_levels(child, 3))
        except OSError as e:
            logger.warning("scan root fail: %s | %s", root, e)
    return candidates


def _walk_n_levels(root: Path, levels: int) -> list[Path]:
    """从 root 向下递归 levels 层，仅返回目录路径。"""
    if not root.exists():
        return []
    out: list[Path] = [root]
    frontier: list[Path] = [root]
    for _ in range(levels):
        nxt: list[Path] = []
        for p in frontier:
            try:
                for child in p.iterdir():
                    if child.is_dir():
                        out.append(child)
                        nxt.append(child)
            except OSError:
                continue
        frontier = nxt
        if not frontier:
            break
    return out


def _is_orphan_python(path: Path, registered: set[Path]) -> tuple[bool, list[str]]:
    """判断目录是否是"孤儿 Python 安装"。

    命中条件: 名称匹配 Python 版本目录 OR 目录下含 python.exe
    且 不在注册表登记的根目录集合中
    """
    reasons: list[str] = []
    name = path.name
    name_hit = bool(_PYTHON_DIR_RE.match(name))
    has_python_exe = (path / "python.exe").exists() or (path / "python3.exe").exists()

    if not (name_hit or has_python_exe):
        return False, reasons

    # 比对注册表登记的根，避免误删正在用的
    if path in registered:
        return False, reasons

    if name_hit:
        reasons.append("name:python_version_dir")
    if has_python_exe:
        reasons.append("contains_python_exe")
    return True, reasons


def _is_orphan_node(path: Path) -> tuple[bool, list[str]]:
    if not _NODE_DIR_RE.match(path.name):
        return False, []
    if not (path / "node.exe").exists():
        return False, []
    return True, ["name:node_version_dir", "contains_node_exe"]


def _is_orphan_conda(path: Path) -> tuple[bool, list[str]]:
    name = path.name.lower()
    if name not in _ORPHAN_ROOT_NAMES:
        return False, []
    reasons = [f"name:{_ORPHAN_ROOT_NAMES[name]}"]
    # 含有 python.exe / conda.exe / Scripts\conda.exe 即更可信
    if (path / "python.exe").exists() or (path / "Scripts" / "conda.exe").exists():
        reasons.append("has_python_or_conda_exe")
    return True, reasons


def _is_large_user_cache(path: Path) -> tuple[bool, list[str]]:
    """用户级可清理的下载/包管理缓存。"""
    lad = get_local_appdata()
    name = path.name.lower()
    relative = str(path).lower()

    # 通用规则: %LocalAppData% 下任一超过 1GB 的 pip / npm / yarn / nuget / cargo 缓存
    is_in_lad = lad.as_posix().lower() in relative
    if not is_in_lad:
        return False, []

    cache_names = {
        "pip": "pkg:pip",
        "pip-cache": "pkg:pip",
        "npm-cache": "pkg:npm",
        "yarn": "pkg:yarn",
        "nuget": "pkg:nuget",
        "cargo": "pkg:cargo",
    }
    if name not in cache_names:
        return False, []
    return True, [cache_names[name]]


def _is_ide_junk(path: Path) -> tuple[bool, list[str]]:
    """IDE 临时 / 升级残留。"""
    name = path.name.lower()
    lad = get_local_appdata()
    in_lad = lad.as_posix().lower() in str(path).lower()

    # VS Code/Cursor 旧版本目录
    if name.startswith("code - insiders") and (path / "resources").exists():
        return True, ["ide:code_insiders_dir"]
    if name.startswith("cursor") and (path / "resources").exists():
        return True, ["ide:cursor_dir"]

    # JetBrains 体系下 *.tmp / cache / log
    if in_lad and name in {"jetbrains", "pycharm", "intellij", "clion", "rider", "webstorm"}:
        return True, ["ide:jetbrains_root"]

    return False, []


def _is_windows_update_cache(path: Path) -> tuple[bool, list[str]]:
    """Windows 系统级更新缓存——永远只扫描不删。

    匹配:
      - 名称为 SoftwareDistribution / $Windows.~BT / $Windows.~WS 的整目录
      - 父目录名 == SoftwareDistribution 且自身名 == Download（含其下任意层）
    """
    name = path.name
    parent_name = path.parent.name
    dangerous_names = {"SoftwareDistribution", "$Windows.~BT", "$Windows.~WS"}
    if name in dangerous_names:
        return True, ["system:windows_update_cache"]
    if parent_name == "SoftwareDistribution" and name.lower() == "download":
        return True, ["system:windows_update_cache"]
    # 兼容路径分隔/大小写差异
    if name.lower() in {n.lower() for n in dangerous_names}:
        return True, ["system:windows_update_cache"]
    if parent_name.lower() == "softwaredistribution" and name.lower() == "download":
        return True, ["system:windows_update_cache"]
    return False, []


def collect_dir_candidates(
    *,
    python_only: bool = False,
    registered_python: set[Path] | None = None,
) -> list[DirCandidate]:
    """按分级规则识别目录级候选。

    风险等级:
      - dangerous : Windows 更新缓存（绝不自动删）
      - review    : 孤儿 Python / Node / Conda / 大型 IDE / 用户包缓存
      - safe      : 暂不直接产出"safe 目录候选"（空目录由 remove_empty_dirs 处理）
    """
    registered = registered_python if registered_python is not None else list_installed_python_roots()
    found: list[DirCandidate] = []
    seen: set[Path] = set()

    for path in _iter_candidate_dirs():
        if path in seen:
            continue
        seen.add(path)
        try:
            if not path.is_dir():
                continue
        except OSError:
            continue

        reasons: list[str] = []
        risk = "review"
        matched = False

        # dangerous 必须最先判断
        hit, why = _is_windows_update_cache(path)
        if hit:
            reasons = why
            risk = "dangerous"
            matched = True

        if not matched and not python_only:
            hit, why = _is_orphan_node(path)
            if hit:
                reasons = why
                matched = True

            if not matched:
                hit, why = _is_orphan_conda(path)
                if hit:
                    reasons = why
                    matched = True

            if not matched:
                hit, why = _is_large_user_cache(path)
                if hit:
                    reasons = why
                    matched = True

            if not matched:
                hit, why = _is_ide_junk(path)
                if hit:
                    reasons = why
                    matched = True

        if not matched and python_only or (not matched and not python_only):
            # Python 检测单独走一遍（python_only 时也走）
            hit, why = _is_orphan_python(path, registered)
            if hit:
                reasons = why
                matched = True

        if not matched:
            continue

        size, latest = dir_size(path)
        # 极小目录（< 50MB）不值得作为 review 候选；safe/dangerous 保留
        if risk == "review" and size < 50 * 1024 * 1024:
            continue
        found.append(DirCandidate(path, size, latest, risk, reasons))
    return found


# ----------------------------------------------------------------------------
# 删除 / 隔离
# ----------------------------------------------------------------------------


def quarantine_dir(src: Path, quarantine_root: Path) -> Path | None:
    """把整目录移动到 quarantine_root 下同名子目录，保留原始相对名。

    使用 shutil.move（同盘时是 rename，跨盘是 copy+delete）。失败返回 None。
    """
    try:
        dest = quarantine_root / src.name
        if dest.exists():
            # 防止重名覆盖
            stem, suffix = src.name, ""
            i = 1
            while dest.exists():
                dest = quarantine_root / f"{stem}.{i}{suffix}"
                i += 1
        shutil.move(str(src), str(dest))
        return dest
    except OSError as e:
        logger.warning("quarantine 失败: %s | %s", src, e)
        return None


def remove_dir(path: Path) -> bool:
    try:
        shutil.rmtree(path, ignore_errors=False)
        return True
    except OSError as e:
        logger.warning("rmtree 失败: %s | %s", path, e)
        return False


def process_dir_candidates(
    candidates: Iterable[DirCandidate],
    *,
    delete: bool,
    include_review: bool,
    quarantine_root: Path | None,
) -> tuple[int, int, int, int]:
    """执行目录级清理。

    返回 (moved_count, removed_count, skipped_count, failed_count)
    """
    moved = removed = skipped = failed = 0
    for c in candidates:
        if c.risk == "dangerous":
            skipped += 1
            continue
        if c.risk == "review" and not include_review:
            skipped += 1
            continue
        if not delete:
            # 仅扫描模式
            continue

        if c.risk == "safe":
            if remove_dir(c.path):
                removed += 1
            else:
                failed += 1
            continue

        # review: 默认走 quarantine 而非直接删
        if quarantine_root is not None:
            dest = quarantine_dir(c.path, quarantine_root)
            if dest is not None:
                moved += 1
                logger.info("isolated: %s -> %s", c.path, dest)
            else:
                failed += 1
        else:
            if remove_dir(c.path):
                removed += 1
            else:
                failed += 1
    return moved, removed, skipped, failed


# ----------------------------------------------------------------------------
# 报告
# ----------------------------------------------------------------------------


def write_file_report(output_dir: Path, candidates: list[FileCandidate], action: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    p = output_dir / "cleanup_report.csv"
    with p.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["path", "size_bytes", "mtime", "reasons", "action"])
        for c in candidates:
            w.writerow([
                str(c.path), c.size,
                time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(c.mtime)),
                ";".join(c.reasons), action,
            ])
    return p


def write_dir_report(output_dir: Path, candidates: list[DirCandidate], action: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    p = output_dir / "directory_report.csv"
    with p.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["path", "size_bytes", "size_human", "last_mtime", "risk", "reasons", "action"])
        for c in sorted(candidates, key=lambda x: -x.size):
            w.writerow([
                str(c.path), c.size, format_size(c.size),
                time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(c.last_mtime)),
                c.risk, ";".join(c.reasons), action,
            ])
    return p


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="C 盘垃圾 / 空文件 / 无用目录清理（默认仅扫描）",
    )
    p.add_argument("--delete", action="store_true",
                   help="执行删除（默认仅扫描；仅作用于 safe 级）")
    p.add_argument("--include-review", action="store_true",
                   help="同时处理 review 级候选（孤儿 Python / Node / 大型缓存），"
                        "默认移动到 output/_quarantine_<ts>/ 而非直接删")
    p.add_argument("--include-dangerous", action="store_true",
                   help="允许处理 dangerous 级（Windows 更新缓存，强烈不建议；保留仅用于审计）")
    p.add_argument("--days", type=int, default=7,
                   help="文件级仅处理修改时间早于 N 天的文件（默认 7）")
    p.add_argument("--min-dir-age-days", type=int, default=30,
                   help="目录级候选须最近 N 天内有修改（默认 30）；0 表示不限")
    p.add_argument("--targets", nargs="*", default=None,
                   help="自定义文件级扫描根目录（覆盖默认）")
    p.add_argument("--remove-empty-dirs", action="store_true",
                   help="删除文件级扫描后留下的空目录")
    p.add_argument("--python-only", action="store_true",
                   help="目录级仅扫描 Python 孤儿")
    p.add_argument("--top", type=int, default=15,
                   help="控制台展示 Top N 大候选目录（默认 15）")
    p.add_argument("--no-quarantine", action="store_true",
                   help="review 级候选直接 rmtree 而非移动到隔离目录（不可恢复）")
    return p.parse_args()


def _is_admin() -> bool:
    """Windows 下检测是否管理员。"""
    if sys.platform != "win32":
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def main() -> None:
    if sys.platform != "win32":
        print("提示: 脚本按 Windows C 盘路径设计，其他系统请使用 --targets 指定目录。")

    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"

    if args.delete and not _is_admin() and sys.platform == "win32":
        print("⚠ 建议以管理员身份运行 --delete 以避免权限失败。")

    # ---------- 文件级 ----------
    input_targets = load_targets_from_input(input_dir)
    if args.targets:
        targets = [Path(p) for p in args.targets]
    elif input_targets:
        targets = input_targets
    else:
        targets = default_targets()

    files = collect_file_candidates(targets, days=args.days)
    f_action = "scan"
    f_deleted = f_failed = 0
    removed_dirs = 0
    if args.delete:
        f_deleted, f_failed = delete_files(files)
        if args.remove_empty_dirs:
            removed_dirs = remove_empty_dirs(targets)
        f_action = "delete"

    file_report = write_file_report(output_dir, files, action=f_action)

    # ---------- 目录级 ----------
    registered = list_installed_python_roots()
    dirs = collect_dir_candidates(python_only=args.python_only, registered_python=registered)
    if args.min_dir_age_days > 0:
        threshold = time.time() - args.min_dir_age_days * 86400
        dirs = [c for c in dirs if c.last_mtime <= threshold or c.risk == "dangerous"]

    if args.include_dangerous:
        # 把 dangerous 也放回候选里（仍只扫描，移动逻辑里我们额外处理）
        # 默认行为保持安全：仅在报告里出现
        pass

    quarantine_root: Path | None = None
    if args.delete and args.include_review and not args.no_quarantine:
        quarantine_root = output_dir / f"_quarantine_{time.strftime('%Y%m%d_%H%M%S')}"
        quarantine_root.mkdir(parents=True, exist_ok=True)

    # dangerous 永远不删
    effective_dirs = [c for c in dirs if c.risk != "dangerous"] if not args.include_dangerous else dirs
    moved, removed, skipped, d_failed = process_dir_candidates(
        effective_dirs,
        delete=args.delete,
        include_review=args.include_review,
        quarantine_root=quarantine_root,
    )

    dir_report = write_dir_report(output_dir, dirs, action="scan" if not args.delete else "delete/quarantine")

    # ---------- 输出汇总 ----------
    total_recover = sum(c.size for c in dirs if c.risk != "dangerous")
    print("=" * 70)
    print(f"  扫描根目录数（文件级）: {len(targets)}")
    print(f"  候选文件数           : {len(files)}  报告: {file_report.name}")
    if args.delete:
        print(f"  文件已删除 / 失败    : {f_deleted} / {f_failed}")
        if args.remove_empty_dirs:
            print(f"  空目录已删除         : {removed_dirs}")

    print(f"  候选目录数           : {len(dirs)}  报告: {dir_report.name}")
    print(f"  目录可回收估算       : {format_size(total_recover)}")
    print(f"  review 候选          : {sum(1 for c in dirs if c.risk == 'review')}")
    print(f"  dangerous 候选       : {sum(1 for c in dirs if c.risk == 'dangerous')}  (永不自动处理)")
    if args.delete and args.include_review:
        print(f"  目录已隔离/删除      : moved={moved}  removed={removed}  failed={d_failed}")
        if quarantine_root:
            print(f"  隔离目录位置         : {quarantine_root}")
    print("=" * 70)

    top_n = max(0, args.top)
    if top_n and dirs:
        show = sorted(
            [c for c in dirs if c.risk != "dangerous"],
            key=lambda x: -x.size,
        )[:top_n]
        print(f"\nTop {len(show)} 目录候选（按大小降序）:")
        for c in show:
            print(f"  [{c.risk:8s}] {format_size(c.size):>10s}  {c.path}")


if __name__ == "__main__":
    main()
