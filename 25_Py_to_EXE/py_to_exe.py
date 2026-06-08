#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
25_Py_to_EXE — Python 脚本转 EXE（单文件 / 批量 / manifest 驱动）

设计目标：
- 单脚本打包（旧用法，100% 向后兼容；上一版 157 行 `py_to_exe.py` 全部能力保留）
- 批量打包（manifest 驱动 + 自动扫描，见 build_all.py）
- 路径以 `sys.executable` 为基准（frozen 兼容；与 18_PDF_eCTD_Converter 一致）
- 启动时自动创建 input/ output/ 目录
- 批量模式默认失败继续 + 末尾汇总；--fail-fast 立即中断

用法：
  # 旧用法（向后兼容）
  python py_to_exe.py
  python py_to_exe.py --input "input/demo.py" --name "demo" --onefile
  python py_to_exe.py --icon "input/app.ico"

  # 批量新用法
  python py_to_exe.py --batch --manifest ./manifest.yaml
  python py_to_exe.py --batch --auto-discover
  python py_to_exe.py --batch --modules 18_PDF_eCTD_Converter,19_PDF_Merge
  python py_to_exe.py --batch --auto-discover --fail-fast
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

# 路径基准（frozen 兼容：与 18_PDF_eCTD_Converter 同样的判断方式）
BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)


# ---------------------------------------------------------------------------
# 单脚本模式（旧用法，100% 向后兼容）
# ---------------------------------------------------------------------------

def _ensure_dirs() -> tuple[Path, Path]:
    """启动时确保 input/ output/ 存在。"""
    input_dir = BASE_DIR / "input"
    output_dir = BASE_DIR / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return input_dir, output_dir


def _build_pyinstaller_cmd_single(
    script_path: Path,
    out_dir: Path,
    *,
    onefile: bool,
    noconsole: bool,
    name: str | None,
    icon: Path | None,
) -> list[str]:
    """单脚本模式的 PyInstaller 命令构造。"""
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--distpath", str(out_dir),
        "--workpath", str(out_dir / "_build"),
        "--specpath", str(out_dir / "_spec"),
    ]
    if onefile:
        cmd.append("--onefile")
    if noconsole:
        cmd.append("--windowed")  # PyInstaller 原生参数是 --windowed（非 --noconsole）
    if name:
        cmd.extend(["--name", name])
    if icon and icon.exists():
        cmd.extend(["--icon", str(icon.resolve())])
    cmd.append(str(script_path))
    return cmd


def _run_single(args: argparse.Namespace, input_dir: Path, output_dir: Path) -> int:
    """单脚本模式：选输入脚本 → 跑 PyInstaller → 清理。"""
    # 1. 选输入脚本
    if args.input:
        script_path = Path(args.input)
        if not script_path.is_absolute():
            script_path = (BASE_DIR / script_path).resolve()
    else:
        candidates = sorted(input_dir.glob("*.py"))
        if not candidates:
            print(
                f"错误：{input_dir} 目录下未找到 .py 文件，请放入代码后再运行。",
                file=sys.stderr,
            )
            return 1
        script_path = candidates[0]

    if not script_path.exists():
        print(f"错误：找不到输入脚本: {script_path}", file=sys.stderr)
        return 1

    # 2. 选输出目录
    if args.output:
        out_dir = Path(args.output)
        if not out_dir.is_absolute():
            out_dir = (BASE_DIR / out_dir).resolve()
    else:
        out_dir = output_dir

    # 3. 选打包模式
    is_onefile = not args.dir_mode  # --dir 开启时为目录模式（关闭 onefile）

    # 4. 图标
    icon_path = None
    if args.icon:
        icon_path = Path(args.icon)
        if not icon_path.is_absolute():
            icon_path = (BASE_DIR / icon_path).resolve()

    # 5. 缓存目录
    build_dir = out_dir / "_build"
    spec_dir = out_dir / "_spec"
    spec_dir.mkdir(parents=True, exist_ok=True)

    # 6. 构造并执行命令
    cmd = _build_pyinstaller_cmd_single(
        script_path=script_path,
        out_dir=out_dir,
        onefile=is_onefile,
        noconsole=args.noconsole,
        name=args.name,
        icon=icon_path,
    )
    print(f"执行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, text=True)
    rc = result.returncode
    if rc != 0:
        print("PyInstaller 打包失败，请检查上方的控制台报错信息。", file=sys.stderr)

    # 7. 清理（默认开）
    if args.clean_artifacts:
        print("正在清理临时文件...")
        if build_dir.exists():
            shutil.rmtree(build_dir, ignore_errors=True)
        if spec_dir.exists():
            shutil.rmtree(spec_dir, ignore_errors=True)

    if rc == 0:
        print(f"\n打包任务结束！请检查输出目录：{out_dir}")
    return rc


# ---------------------------------------------------------------------------
# 批量模式（新用法，分发到 build_all）
# ---------------------------------------------------------------------------

def _run_batch(args: argparse.Namespace) -> int:
    """批量模式：转发到 build_all.BatchRunner。"""
    # 延迟 import 避免不必要的依赖加载（用户走单脚本模式时不需要 yaml/pandas）
    try:
        from build_all import BatchRunner
    except ModuleNotFoundError as exc:
        print(
            f"批量模式依赖加载失败: {exc}\n"
            f"请运行: {sys.executable} -m pip install pyyaml pandas openpyxl",
            file=sys.stderr,
        )
        return 2
    runner = BatchRunner(
        base_dir=BASE_DIR,
        manifest_path=Path(args.manifest).resolve() if args.manifest else None,
        auto_discover=args.auto_discover,
        modules_filter=args.modules,
        fail_fast=args.fail_fast,
        workers=max(1, args.workers),
    )
    return runner.run()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将 Python 脚本打包为 EXE（支持单脚本与批量模式）"
    )

    # === 旧用法（向后兼容） ===
    parser.add_argument("--input", help="单脚本模式：输入 .py 路径")
    parser.add_argument("--output", help="单脚本模式：输出目录（默认 output/）")
    parser.add_argument("--name", help="EXE 文件名（不含扩展名）")
    parser.add_argument(
        "--onefile", dest="onefile", action="store_true", default=True,
        help="单文件模式（默认开启）",
    )
    parser.add_argument(
        "--dir", dest="dir_mode", action="store_true",
        help="目录模式（关闭 onefile）",
    )
    parser.add_argument("--noconsole", action="store_true", help="隐藏控制台黑窗口")
    parser.add_argument("--icon", help="图标路径（.ico）")
    parser.add_argument(
        "--clean-artifacts", dest="clean_artifacts", action="store_true", default=True,
        help="打包后清理 build/spec 中间产物（默认开启）",
    )
    parser.add_argument(
        "--no-clean-artifacts", dest="clean_artifacts", action="store_false",
        help="不清理中间产物",
    )

    # === 新增：批量模式 ===
    parser.add_argument(
        "--batch", action="store_true",
        help="批量打包模式（需配合 --manifest / --auto-discover / --modules）",
    )
    parser.add_argument(
        "--manifest",
        help="manifest YAML 路径（默认 ./manifest.yaml；传 --batch 时生效）",
    )
    parser.add_argument(
        "--auto-discover", action="store_true",
        help="自动扫描仓库各 NN_*/ 模块的主程序（需配合 --batch）",
    )
    parser.add_argument(
        "--modules",
        help="逗号分隔的模块名列表（如 18_PDF_eCTD_Converter,19_PDF_Merge），覆盖 manifest",
    )
    parser.add_argument(
        "--fail-fast", action="store_true",
        help="批量模式遇错即停（默认失败继续）",
    )
    parser.add_argument(
        "--workers", type=int, default=1,
        help="批量模式并发数（默认 1；PyInstaller CPU 密集，建议 ≤ CPU 核数）",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_dir, output_dir = _ensure_dirs()

    if args.batch:
        return _run_batch(args)
    return _run_single(args, input_dir, output_dir)


if __name__ == "__main__":
    sys.exit(main())
