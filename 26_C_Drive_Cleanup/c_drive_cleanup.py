"""
C 盘垃圾文件/空文件/无用文件清理（安全模式默认仅扫描）。

默认扫描目录（仅 C 盘常见临时/缓存位置）：
- C:\\Windows\\Temp
- C:\\Users\\<User>\\AppData\\Local\\Temp
- C:\\Users\\<User>\\AppData\\Local\\Microsoft\\Windows\\INetCache
- C:\\Users\\<User>\\AppData\\Local\\CrashDumps

输入/输出：
- input/ 可放 targets.txt（每行一个自定义目录，存在则优先使用）
- output/ 输出清理报告 cleanup_report.csv

用法：
  python c_drive_cleanup.py
  python c_drive_cleanup.py --delete
  python c_drive_cleanup.py --days 3 --remove-empty-dirs
  python c_drive_cleanup.py --targets "C:\\Temp" "C:\\Users\\me\\AppData\\Local\\Temp"
"""
from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from pathlib import Path
from collections.abc import Iterable


USELESS_EXTS = {
    ".tmp",
    ".temp",
    ".bak",
    ".old",
    ".log",
    ".dmp",
    ".chk",
    ".gid",
    ".~tmp",
    ".cache",
    ".crdownload",
    ".part",
}

DEFAULT_REL_DIRS = [
    r"C:\\Windows\\Temp",
]

USER_REL_DIRS = [
    r"AppData\\Local\\Temp",
    r"AppData\\Local\\Microsoft\\Windows\\INetCache",
    r"AppData\\Local\\CrashDumps",
]


class Candidate:
    def __init__(self, path: Path, size: int, mtime: float, reasons: list[str]):
        self.path = path
        self.size = size
        self.mtime = mtime
        self.reasons = reasons


def load_targets_from_input(input_dir: Path) -> list[Path]:
    targets_file = input_dir / "targets.txt"
    if not targets_file.exists():
        return []
    targets = []
    for line in targets_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        targets.append(Path(line))
    return targets


def default_targets() -> list[Path]:
    targets = [Path(p) for p in DEFAULT_REL_DIRS]
    users_dir = Path("C:/Users")
    if users_dir.exists():
        for user_dir in users_dir.iterdir():
            if not user_dir.is_dir():
                continue
            for rel in USER_REL_DIRS:
                targets.append(user_dir / rel)
    return targets


def is_old_enough(mtime: float, days: int) -> bool:
    threshold = time.time() - days * 86400
    return mtime <= threshold


def iter_files(targets: Iterable[Path]) -> Iterable[Path]:
    for root in targets:
        if not root.exists() or not root.is_dir():
            continue
        for dirpath, _, filenames in os.walk(root):
            for name in filenames:
                yield Path(dirpath) / name


def collect_candidates(targets: Iterable[Path], days: int) -> list[Candidate]:
    candidates: list[Candidate] = []
    for file_path in iter_files(targets):
        try:
            stat = file_path.stat()
        except Exception:
            continue

        reasons: list[str] = []
        ext = file_path.suffix.lower()
        if stat.st_size == 0:
            reasons.append("empty")
        if ext in USELESS_EXTS:
            reasons.append(f"ext:{ext}")
        if reasons and is_old_enough(stat.st_mtime, days):
            candidates.append(Candidate(file_path, stat.st_size, stat.st_mtime, reasons))
    return candidates


def delete_files(candidates: list[Candidate]) -> tuple[int, int]:
    deleted = 0
    failed = 0
    for c in candidates:
        try:
            c.path.unlink(missing_ok=True)
            deleted += 1
        except Exception:
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
            except Exception:
                continue
    return removed


def write_report(output_dir: Path, candidates: list[Candidate], action: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / "cleanup_report.csv"
    with report_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["path", "size_bytes", "mtime", "reasons", "action"])
        for c in candidates:
            writer.writerow([
                str(c.path),
                c.size,
                time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(c.mtime)),
                ";".join(c.reasons),
                action,
            ])
    return report_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="C 盘垃圾/空文件清理（默认仅扫描，安全模式）"
    )
    parser.add_argument(
        "--delete",
        action="store_true",
        help="执行删除（默认仅扫描）",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=7,
        help="仅处理修改时间早于 N 天的文件（默认 7 天）",
    )
    parser.add_argument(
        "--targets",
        nargs="*",
        default=None,
        help="自定义扫描目录列表（覆盖默认目录）",
    )
    parser.add_argument(
        "--remove-empty-dirs",
        action="store_true",
        help="删除清理后产生的空目录",
    )
    return parser.parse_args()


def main() -> None:
    if sys.platform != "win32":
        print("提示：该脚本按 Windows C 盘路径设计，其他系统请使用 --targets 指定目录。")

    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"

    input_targets = load_targets_from_input(input_dir)
    if args.targets is not None and len(args.targets) > 0:
        targets = [Path(p) for p in args.targets]
    elif input_targets:
        targets = input_targets
    else:
        targets = default_targets()

    candidates = collect_candidates(targets, days=args.days)

    action = "scan"
    deleted = 0
    failed = 0
    if args.delete:
        deleted, failed = delete_files(candidates)
        if args.remove_empty_dirs:
            removed_dirs = remove_empty_dirs(targets)
            print(f"已删除空目录：{removed_dirs} 个")
        action = "delete"

    report_path = write_report(output_dir, candidates, action=action)

    print(f"扫描目录数：{len(targets)}")
    print(f"候选文件数：{len(candidates)}")
    if args.delete:
        print(f"已删除：{deleted}，失败：{failed}")
    print(f"报告已生成：{report_path}")


if __name__ == "__main__":
    main()
