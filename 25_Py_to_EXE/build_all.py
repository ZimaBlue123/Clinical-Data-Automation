#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
25_Py_to_EXE/build_all.py — 批量打包核心

被 `py_to_exe.py --batch` 调用，也可直接 `python build_all.py` 运行。

设计原则：
- 单一职责：只负责批量调度 + 报告
- 不修改被打包模块的任何文件
- 默认失败继续 + 末尾汇总；--fail-fast 立即中断
- 报告三件套：JSON + Markdown + Excel
- PyInstaller 调用走 `sys.executable -m PyInstaller`（与 18_PDF_eCTD_Converter 一致）

报告输出：
  25_PY_to_EXE/output/_batch_report/
    ├─ batch_report_<timestamp>.json   # 完整数据
    ├─ batch_report_<timestamp>.md     # 人可读汇总
    └─ batch_report_<timestamp>.xlsx   # Excel（复用 18 的 openpyxl 写法）
"""
from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

# ---------------------------------------------------------------------------
# 依赖检查（pyyaml / pandas / openpyxl）
# ---------------------------------------------------------------------------

try:
    import yaml
except ModuleNotFoundError as exc:
    raise SystemExit(
        f"未找到 pyyaml。请运行: {sys.executable} -m pip install pyyaml"
    ) from exc

try:
    import pandas as pd
except ModuleNotFoundError as exc:
    raise SystemExit(
        f"未找到 pandas。请运行: {sys.executable} -m pip install pandas openpyxl"
    ) from exc

try:
    import openpyxl  # noqa: F401
except ModuleNotFoundError as exc:
    raise SystemExit(
        f"未找到 openpyxl。请运行: {sys.executable} -m pip install openpyxl"
    ) from exc

# ---------------------------------------------------------------------------
# 路径基准（frozen 兼容）
# ---------------------------------------------------------------------------

BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
# 假设 25_Py_to_EXE 位于仓库根的下一级（与现有所有 NN_*/ 同级）
REPO_ROOT = BASE_DIR.parent

logger = logging.getLogger("build_all")
if not logger.handlers:
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(_h)
logger.setLevel(logging.INFO)


# ===========================================================================
# 数据类
# ===========================================================================

@dataclass
class PyInstallerConfig:
    """单次 PyInstaller 调用的参数。"""
    onefile: bool = True
    console: bool = True
    icon: str | None = None
    collect_submodules: list[str] = field(default_factory=list)
    excludes: list[str] = field(default_factory=list)
    datas: list[tuple[str, str]] = field(default_factory=list)
    hiddenimports: list[str] = field(default_factory=list)
    clean_artifacts: bool = True


@dataclass
class ModuleSpec:
    """manifest 中一个模块的打包规格。"""
    module: str
    script: str
    output_name: str | None = None
    config: PyInstallerConfig = field(default_factory=PyInstallerConfig)
    warning: str = ""

    @property
    def module_dir(self) -> Path:
        return REPO_ROOT / self.module

    @property
    def script_path(self) -> Path:
        return self.module_dir / self.script

    @property
    def exe_name(self) -> str:
        base = self.output_name or Path(self.script).stem
        return base + (".exe" if sys.platform == "win32" else "")

    @property
    def out_dir(self) -> Path:
        return BASE_DIR / "output" / self.module


@dataclass
class BatchResult:
    """单次打包的结果。"""
    module: str
    script: str
    status: str                       # success | failed | skipped
    output_exe: str = ""
    size_mb: float = 0.0
    duration_sec: float = 0.0
    exit_code: int | None = None
    warning: str = ""
    error: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))


# ===========================================================================
# Manifest 加载
# ===========================================================================

def _to_pyinstaller_config(d: dict | None) -> PyInstallerConfig:
    d = d or {}
    return PyInstallerConfig(
        onefile=bool(d.get("onefile", True)),
        console=bool(d.get("console", True)),
        icon=d.get("icon"),
        collect_submodules=list(d.get("collect_submodules", []) or []),
        excludes=list(d.get("excludes", []) or []),
        datas=[tuple(x) for x in (d.get("datas", []) or [])],
        hiddenimports=list(d.get("hiddenimports", []) or []),
        clean_artifacts=bool(d.get("clean_artifacts", True)),
    )


def load_manifest(path: Path) -> tuple[PyInstallerConfig, list[ModuleSpec]]:
    """从 YAML 加载 manifest。返回 (default_config, modules)。"""
    if not path.exists():
        raise SystemExit(f"manifest 不存在: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    batch = data.get("batch", {})
    default_cfg = _to_pyinstaller_config(batch.get("default_pyinstaller"))
    modules: list[ModuleSpec] = []
    for entry in batch.get("modules", []) or []:
        # 合并 default + per-module 覆盖
        merged = {**(batch.get("default_pyinstaller", {}) or {}), **(entry.get("pyinstaller", {}) or {})}
        modules.append(ModuleSpec(
            module=entry["module"],
            script=entry["script"],
            output_name=entry.get("output_name"),
            config=_to_pyinstaller_config(merged),
            warning=str(entry.get("warning", "") or ""),
        ))
    return default_cfg, modules


# ===========================================================================
# 自动发现
# ===========================================================================

# 排除的目录前缀（不是模块）
EXCLUDE_DIR_PREFIXES = ("__pycache__", ".git", "tests", "docs", "scripts", "src", ".github")
# 排除的脚本前缀（util_/test_/lib_ 等）
EXCLUDE_SCRIPT_PREFIXES = ("util_", "lib_", "test_", "_")
EXCLUDE_SCRIPT_SUFFIXES = ("_test.py",)


def auto_discover(repo_root: Path, only: Iterable[str] | None = None) -> list[ModuleSpec]:
    """启发式扫描 NN_*/ 目录，取根目录语义化命名的 .py 为主程序。

    规则：
      - 目录名以两个数字 + 下划线开头（NN_xxx）视为模块
      - 取模块根目录顶层 .py，过滤 util_/test_/lib_/_ 前缀和 _test.py 后缀
      - 特殊：模块根目录的 main.py 也视为主入口（兼容 13/21 等历史遗留）
    """
    only_set = set(only or [])
    out: list[ModuleSpec] = []
    for entry in sorted(repo_root.iterdir()):
        if not entry.is_dir():
            continue
        # 名称必须以两个数字 + 下划线开头
        if not (len(entry.name) >= 3 and entry.name[:2].isdigit() and entry.name[2] == "_"):
            continue
        if any(entry.name.startswith(p) for p in EXCLUDE_DIR_PREFIXES):
            continue
        if only_set and entry.name not in only_set:
            continue

        candidates: list[Path] = []
        for p in entry.iterdir():
            if not p.is_file() or p.suffix != ".py":
                continue
            if p.stem == "__init__":
                continue
            if p.name.startswith(EXCLUDE_SCRIPT_PREFIXES):
                continue
            if p.name.endswith(EXCLUDE_SCRIPT_SUFFIXES):
                continue
            candidates.append(p)

        # main.py 优先放最前
        candidates.sort(key=lambda x: (0 if x.name == "main.py" else 1, x.name))

        for script in candidates:
            out.append(ModuleSpec(
                module=entry.name,
                script=script.name,
                config=PyInstallerConfig(),
            ))
    return out


# ===========================================================================
# 打包执行
# ===========================================================================

def _build_pyinstaller_cmd(spec: ModuleSpec) -> list[str]:
    """为单个 ModuleSpec 构造 PyInstaller 命令行。"""
    cfg = spec.config
    spec.out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--distpath", str(spec.out_dir),
        "--workpath", str(spec.out_dir / "_build"),
        "--specpath", str(spec.out_dir / "_spec"),
        "--name", Path(spec.exe_name).stem,
    ]
    if cfg.onefile:
        cmd.append("--onefile")
    if not cfg.console:
        cmd.append("--windowed")
    if cfg.icon and Path(cfg.icon).exists():
        cmd.extend(["--icon", str(Path(cfg.icon).resolve())])
    for m in cfg.collect_submodules:
        cmd.extend(["--collect-submodules", m])
    for m in cfg.excludes:
        cmd.extend(["--exclude-module", m])
    for src, dst in cfg.datas:
        # PyInstaller 6: Windows 用 ; 分隔，Unix 用 :
        sep = ";" if os.name == "nt" else ":"
        cmd.extend(["--add-data", f"{src}{sep}{dst}"])
    for h in cfg.hiddenimports:
        cmd.extend(["--hidden-import", h])
    # 入口脚本（放最后）
    cmd.append(str(spec.script_path.resolve()))
    return cmd


def _truncate(s: str | None, n: int = 2000) -> str:
    if not s:
        return ""
    s = s.strip()
    return s if len(s) <= n else s[:n] + f"\n... (truncated, total {len(s)} chars)"


def build_one(spec: ModuleSpec) -> BatchResult:
    """对单个模块执行打包。失败时返回 status='failed'，不抛异常。"""
    res = BatchResult(module=spec.module, script=spec.script, status="skipped")
    res.warning = spec.warning

    # 前置校验
    if not spec.module_dir.exists():
        res.status = "failed"
        res.error = f"模块目录不存在: {spec.module_dir}"
        return res
    if not spec.script_path.exists():
        res.status = "failed"
        res.error = f"入口脚本不存在: {spec.script_path}"
        return res

    # 构造命令
    cmd = _build_pyinstaller_cmd(spec)
    logger.info("执行: %s", " ".join(cmd))

    # 执行
    t0 = time.monotonic()
    try:
        proc = subprocess.run(cmd, text=True, capture_output=True, timeout=1800)  # 30 分钟硬上限
    except subprocess.TimeoutExpired as exc:
        res.status = "failed"
        res.duration_sec = time.monotonic() - t0
        res.error = f"打包超时（30 分钟上限）\nstdout: {_truncate(exc.stdout)}\nstderr: {_truncate(exc.stderr)}"
        return res
    except Exception as exc:  # noqa: BLE001 — 任何子进程异常都吞掉，写入 result
        res.status = "failed"
        res.duration_sec = time.monotonic() - t0
        res.error = f"执行异常: {type(exc).__name__}: {exc}"
        return res

    res.duration_sec = time.monotonic() - t0
    res.exit_code = proc.returncode

    # 找生成的 EXE
    exe_path = spec.out_dir / spec.exe_name
    if proc.returncode == 0 and exe_path.exists():
        res.status = "success"
        res.output_exe = str(exe_path)
        res.size_mb = round(exe_path.stat().st_size / (1024 * 1024), 2)
    else:
        res.status = "failed"
        # 截取 stderr（PyInstaller 主要错误信息在 stderr）
        res.error = _truncate(proc.stderr) or _truncate(proc.stdout) or f"PyInstaller exit {proc.returncode}, 未找到 EXE: {exe_path}"

    # 浮点尾数处理（time.monotonic 差值会有 ~1e-14 的浮点残差）
    res.duration_sec = round(res.duration_sec, 2)

    # 清理中间产物
    if spec.config.clean_artifacts:
        for sub in ("_build", "_spec"):
            d = spec.out_dir / sub
            if d.exists():
                shutil.rmtree(d, ignore_errors=True)
        # 顺手清掉根目录的 .spec 文件（specpath 已指定到 _spec，但 PyInstaller 偶尔也会落一份在 dist 根）
        stray = spec.out_dir / f"{Path(spec.exe_name).stem}.spec"
        if stray.exists():
            try:
                stray.unlink()
            except OSError:
                pass

    return res


# ===========================================================================
# 报告输出（JSON / Markdown / Excel 三件套）
# ===========================================================================

def write_reports(results: list[BatchResult], report_dir: Path) -> tuple[Path, Path, Path]:
    """写三件套报告。返回 (json_path, md_path, xlsx_path)。"""
    report_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = report_dir / f"batch_report_{ts}.json"
    md_path = report_dir / f"batch_report_{ts}.md"
    xlsx_path = report_dir / f"batch_report_{ts}.xlsx"

    # --- JSON ---
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "base_dir": str(BASE_DIR),
        "repo_root": str(REPO_ROOT),
        "summary": {
            "total": len(results),
            "success": sum(1 for r in results if r.status == "success"),
            "failed": sum(1 for r in results if r.status == "failed"),
            "skipped": sum(1 for r in results if r.status == "skipped"),
        },
        "results": [asdict(r) for r in results],
    }
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # --- Markdown ---
    total = payload["summary"]["total"]
    succ = payload["summary"]["success"]
    fail = payload["summary"]["failed"]
    skip = payload["summary"]["skipped"]
    total_size = sum(r.size_mb for r in results)
    total_time = sum(r.duration_sec for r in results)
    md_lines = [
        f"# 批量打包报告 — {payload['generated_at']}",
        "",
        f"- **基目录**：`{BASE_DIR}`",
        f"- **仓库根**：`{REPO_ROOT}`",
        "",
        "## 汇总",
        "",
        "| 指标 | 值 |",
        "|------|----|",
        f"| 总数 | {total} |",
        f"| 成功 | {succ} |",
        f"| 失败 | {fail} |",
        f"| 跳过 | {skip} |",
        f"| EXE 总大小 | {total_size:.1f} MB |",
        f"| 总耗时 | {total_time:.1f} 秒 |",
        "",
        "## 明细",
        "",
        "| 序号 | 模块 | 入口脚本 | 状态 | EXE 大小 | 耗时(秒) | 退出码 | 警告 |",
        "|------|------|----------|------|----------|----------|--------|------|",
    ]
    for i, r in enumerate(results, 1):
        warn = r.warning.replace("\n", " ")[:50] if r.warning else ""
        md_lines.append(
            f"| {i} | {r.module} | {r.script} | {r.status} | "
            f"{r.size_mb:.1f} MB | {r.duration_sec:.1f} | "
            f"{r.exit_code if r.exit_code is not None else '-'} | {warn} |"
        )
    # 失败详情
    failed_results = [r for r in results if r.status == "failed"]
    if failed_results:
        md_lines += ["", "## 失败详情", ""]
        for r in failed_results:
            md_lines += [
                f"### {r.module} ({r.script})",
                "",
                "```",
                r.error or "(无错误信息)",
                "```",
                "",
            ]
    md_path.write_text("\n".join(md_lines), encoding="utf-8")

    # --- Excel ---
    df_all = pd.DataFrame([asdict(r) for r in results])
    # 列顺序：把常用字段放前面
    preferred_cols = ["module", "script", "status", "size_mb", "duration_sec", "exit_code", "warning", "error", "output_exe", "timestamp"]
    df_all = df_all[[c for c in preferred_cols if c in df_all.columns]]
    df_warn = df_all[df_all["warning"].astype(str).str.len() > 0] if "warning" in df_all.columns else df_all.iloc[0:0]
    df_fail = df_all[df_all["status"] == "failed"] if "status" in df_all.columns else df_all.iloc[0:0]
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="全部", index=False)
        df_warn.to_excel(writer, sheet_name="含警告", index=False)
        df_fail.to_excel(writer, sheet_name="失败", index=False)
        # 汇总 sheet
        summary_df = pd.DataFrame([payload["summary"]])
        summary_df.insert(0, "generated_at", payload["generated_at"])
        summary_df.insert(1, "exe_total_mb", round(total_size, 2))
        summary_df.insert(2, "total_sec", round(total_time, 2))
        summary_df.to_excel(writer, sheet_name="汇总", index=False)

    return json_path, md_path, xlsx_path


# ===========================================================================
# 批量调度器
# ===========================================================================

class BatchRunner:
    """批量打包调度器。串行执行（PyInstaller CPU 密集，并发收益有限）。"""

    def __init__(
        self,
        *,
        base_dir: Path,
        manifest_path: Path | None = None,
        auto_discover: bool = False,
        modules_filter: str | None = None,
        fail_fast: bool = False,
        workers: int = 1,
    ) -> None:
        self.base_dir = base_dir
        self.manifest_path = manifest_path
        self.auto_discover = auto_discover
        self.modules_filter = (
            [m.strip() for m in modules_filter.split(",") if m.strip()]
            if modules_filter else None
        )
        self.fail_fast = fail_fast
        self.workers = max(1, workers)  # TODO: 未来按 workers > 1 走 ProcessPoolExecutor

    def _resolve_modules(self) -> list[ModuleSpec]:
        """根据参数解析要打包的模块清单。优先级：--modules > manifest > --auto-discover。"""
        if self.modules_filter:
            if not (self.manifest_path or self.auto_discover):
                raise SystemExit(
                    "--modules 需要配合 --manifest 或 --auto-discover 提供基础信息"
                )
            if self.manifest_path and self.manifest_path.exists():
                _, all_specs = load_manifest(self.manifest_path)
            else:
                all_specs = auto_discover(REPO_ROOT)
            keep = set(self.modules_filter)
            selected = [s for s in all_specs if s.module in keep]
            missing = keep - {s.module for s in selected}
            if missing:
                logger.warning("--modules 中未匹配的模块: %s", ", ".join(sorted(missing)))
            return selected

        if self.manifest_path:
            if not self.manifest_path.exists():
                raise SystemExit(
                    f"manifest 不存在: {self.manifest_path}\n"
                    f"提示：可从 manifest.example.yaml 复制后修改"
                )
            _, specs = load_manifest(self.manifest_path)
            logger.info("从 manifest 加载: %s (共 %d 个模块)", self.manifest_path, len(specs))
            return specs

        if self.auto_discover:
            specs = auto_discover(REPO_ROOT)
            logger.info("自动发现完成: 共 %d 个模块", len(specs))
            return specs

        raise SystemExit(
            "批量模式需要指定以下任一参数：\n"
            "  --manifest <path>     使用 YAML 清单\n"
            "  --auto-discover       自动扫描仓库\n"
            "  --modules <a,b,c>     与上面任意一个配合，覆盖为指定模块子集"
        )

    def run(self) -> int:
        specs = self._resolve_modules()
        if not specs:
            logger.warning("没有模块需要打包。")
            return 0

        logger.info("=" * 60)
        logger.info("批量打包开始：共 %d 个模块（失败策略：%s）",
                    len(specs), "遇错即停" if self.fail_fast else "失败继续")
        logger.info("=" * 60)

        results: list[BatchResult] = []
        t_start = time.monotonic()
        for i, spec in enumerate(specs, 1):
            logger.info("[%d/%d] %s / %s", i, len(specs), spec.module, spec.script)
            if spec.warning:
                logger.warning("  警告: %s", spec.warning)
            res = build_one(spec)
            results.append(res)
            tag = "OK" if res.status == "success" else f"FAIL ({res.status})"
            logger.info("  -> %s  size=%sMB  time=%.1fs",
                        tag,
                        f"{res.size_mb:.1f}" if res.size_mb else "-",
                        res.duration_sec)
            if res.status == "failed" and res.error:
                # 把错误前几行打到日志
                first_lines = "\n".join(res.error.splitlines()[:5])
                logger.error("  error(head):\n%s", first_lines)

            if self.fail_fast and res.status == "failed":
                logger.error("--fail-fast 触发，停止后续模块。")
                break

        total_time = time.monotonic() - t_start
        succ = sum(1 for r in results if r.status == "success")
        fail = sum(1 for r in results if r.status == "failed")
        skip = sum(1 for r in results if r.status == "skipped")

        # 写报告
        report_dir = self.base_dir / "output" / "_batch_report"
        json_p, md_p, xlsx_p = write_reports(results, report_dir)

        # 终端汇总
        logger.info("=" * 60)
        logger.info("批量打包结束：成功=%d 失败=%d 跳过=%d 总耗时=%.1fs",
                    succ, fail, skip, total_time)
        logger.info("报告：")
        logger.info("  JSON : %s", json_p)
        logger.info("  MD   : %s", md_p)
        logger.info("  XLSX : %s", xlsx_p)
        logger.info("=" * 60)

        # 失败列表（便于重试）
        if fail:
            failed = [r for r in results if r.status == "failed"]
            logger.error("失败模块：")
            for r in failed:
                logger.error("  - %s (%s)", r.module, r.script)
            retry_cmd = ",".join(r.module for r in failed)
            logger.error("重试命令：python py_to_exe.py --batch --auto-discover --modules %s", retry_cmd)
            return 1
        return 0


# ---------------------------------------------------------------------------
# 独立 CLI 入口（也可被 py_to_exe.py --batch 调用）
# ---------------------------------------------------------------------------

def _standalone_main() -> int:
    import argparse
    p = argparse.ArgumentParser(description="批量打包：直接运行 build_all.py")
    p.add_argument("--manifest", help="manifest YAML 路径")
    p.add_argument("--auto-discover", action="store_true", help="自动扫描仓库")
    p.add_argument("--modules", help="逗号分隔的模块名列表")
    p.add_argument("--fail-fast", action="store_true", help="遇错即停")
    p.add_argument("--workers", type=int, default=1, help="并发数（默认 1）")
    args = p.parse_args()
    runner = BatchRunner(
        base_dir=BASE_DIR,
        manifest_path=Path(args.manifest).resolve() if args.manifest else None,
        auto_discover=args.auto_discover,
        modules_filter=args.modules,
        fail_fast=args.fail_fast,
        workers=args.workers,
    )
    return runner.run()


if __name__ == "__main__":
    sys.exit(_standalone_main())
