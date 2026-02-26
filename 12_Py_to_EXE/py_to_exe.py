# -*- coding: utf-8 -*-
"""
将 Python 脚本打包为 EXE（12 模块）

依赖：pyinstaller

默认目录：
- 输入：12_Py_to_EXE/input/
- 输出：12_Py_to_EXE/output/

用法：
  python py_to_exe.py
  python py_to_exe.py --input "input/demo.py" --name "demo" --onefile
  python py_to_exe.py --icon "input/app.ico"
"""
from __future__ import annotations

import argparse
import subprocess
import shutil
from dataclasses import dataclass
from pathlib import Path



def run_pyinstaller(
    script_path: Path,
    output_dir: Path,
    name: str | None = None,
    onefile: bool = True,
    console: bool = True,
    icon: Path | None = None,
    build_dir: Path | None = None,
    spec_dir: Path | None = None,
) -> None:
    cmd = ["pyinstaller", str(script_path), "--distpath", str(output_dir)]

    if build_dir:
        cmd.extend(["--workpath", str(build_dir)])
    if spec_dir:
        cmd.extend(["--specpath", str(spec_dir)])

    if onefile:
        cmd.append("--onefile")
    if not console:
        cmd.append("--noconsole")
    if name:
        cmd.extend(["--name", name])
    if icon:
        cmd.extend(["--icon", str(icon)])

    result = subprocess.run(cmd, text=True)
    if result.returncode != 0:
        raise RuntimeError("PyInstaller 打包失败，请检查日志输出。")



@dataclass
class Args:
    input: str | None = None
    output: str | None = None
    name: str | None = None
    onefile: bool = False
    dir: bool = False
    noconsole: bool = False
    icon: str | None = None
    clean_artifacts: bool = False


def parse_args() -> Args:
    parser = argparse.ArgumentParser(description="将 Python 脚本打包为 EXE")
    _ = parser.add_argument("--input", default=None, help="输入脚本（.py）路径")
    _ = parser.add_argument("--output", default=None, help="输出目录（默认 output/）")
    _ = parser.add_argument("--name", default=None, help="输出 EXE 名称（不含扩展名）")
    _ = parser.add_argument("--onefile", action="store_true", help="单文件模式（默认开启）")
    _ = parser.add_argument("--dir", action="store_true", help="目录模式（关闭 onefile）")
    _ = parser.add_argument("--noconsole", action="store_true", help="隐藏控制台窗口")
    _ = parser.add_argument("--icon", default=None, help="图标路径（.ico）")
    _ = parser.add_argument(
        "--clean-artifacts",
        action="store_true",
        help="打包后删除 build 目录与 .spec 文件",
    )
    return parser.parse_args(namespace=Args())




def main() -> None:
    base_dir = Path(__file__).resolve().parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    args = parse_args()

    if args.input:
        script_path = Path(args.input)
        if not script_path.is_absolute():
            script_path = base_dir / script_path
    else:
        candidates = list(input_dir.glob("*.py"))
        if not candidates:
            raise FileNotFoundError("input 目录下未找到 .py 文件")
        script_path = sorted(candidates)[0]

    if not script_path.exists():
        raise FileNotFoundError(f"找不到输入脚本: {script_path}")

    if args.output:
        out_dir = Path(args.output)
        if not out_dir.is_absolute():
            out_dir = output_dir / out_dir
    else:
        out_dir = output_dir

    onefile = True
    if args.dir:
        onefile = False
    elif args.onefile:
        onefile = True

    icon_path = Path(args.icon) if args.icon else None
    if icon_path and not icon_path.is_absolute():
        icon_path = base_dir / icon_path

    build_dir = out_dir / "_build"
    spec_dir = out_dir / "_spec"

    run_pyinstaller(
        script_path,
        out_dir,
        name=args.name,
        onefile=onefile,
        console=not args.noconsole,
        icon=icon_path,
        build_dir=build_dir,
        spec_dir=spec_dir,
    )

    if args.clean_artifacts:
        if build_dir.exists():
            shutil.rmtree(build_dir, ignore_errors=True)

        if spec_dir.exists():
            shutil.rmtree(spec_dir, ignore_errors=True)


    print(f"打包完成，输出目录：{out_dir}")



if __name__ == "__main__":
    main()
