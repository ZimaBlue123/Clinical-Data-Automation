# -*- coding: utf-8 -*-
"""
将 Python 脚本打包为 EXE

依赖：pyinstaller (需要 pip install pyinstaller)

默认目录：
- 输入：input/
- 输出：output/

用法：
  python py_to_exe.py
  python py_to_exe.py --input "input/demo.py" --name "demo" --onefile
  python py_to_exe.py --icon "input/app.ico"
"""
from __future__ import annotations

import argparse
import subprocess
import shutil
import sys
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
    """构建并执行 PyInstaller 命令"""
    # 基础命令
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
    if icon and icon.exists():
        cmd.extend(["--icon", str(icon)])

    print(f"执行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, text=True)
    
    if result.returncode != 0:
        raise RuntimeError("PyInstaller 打包失败，请检查上方的控制台报错信息。")


def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description="将 Python 脚本打包为 EXE")
    parser.add_argument("--input", default=None, help="输入脚本（.py）路径")
    parser.add_argument("--output", default=None, help="输出目录（默认 output/）")
    parser.add_argument("--name", default=None, help="输出 EXE 名称（不含扩展名）")
    parser.add_argument("--onefile", action="store_true", help="单文件模式（默认开启）")
    parser.add_argument("--dir", action="store_true", help="目录模式（关闭 onefile）")
    parser.add_argument("--noconsole", action="store_true", help="隐藏控制台黑窗口")
    parser.add_argument("--icon", default=None, help="图标路径（.ico）")
    parser.add_argument(
        "--clean-artifacts",
        action="store_true",
        default=True, # 建议默认开启清理，保持工作区整洁
        help="打包后删除 build 目录与 .spec 文件",
    )
    return parser.parse_args()


def main() -> None:
    # 1. 确保基础目录存在
    base_dir = Path(__file__).resolve().parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"
    
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    args = parse_args()

    # 2. 确定输入文件
    if args.input:
        script_path = Path(args.input)
        if not script_path.is_absolute():
            script_path = base_dir / script_path
    else:
        # 自动查找 input 目录下的第一个 .py 文件
        candidates = list(input_dir.glob("*.py"))
        if not candidates:
            print(f"错误：{input_dir} 目录下未找到 .py 文件，请放入代码后再运行。", file=sys.stderr)
            return
        script_path = sorted(candidates)[0]

    if not script_path.exists():
        print(f"错误：找不到输入脚本: {script_path}", file=sys.stderr)
        return

    # 3. 确定输出目录
    if args.output:
        out_dir = Path(args.output)
        if not out_dir.is_absolute():
            out_dir = output_dir / out_dir
    else:
        out_dir = output_dir

    # 4. 确定打包模式 (默认单文件，如果是 --dir 则改为多文件目录模式)
    is_onefile = not args.dir

    # 5. 图标处理
    icon_path = None
    if args.icon:
        icon_path = Path(args.icon)
        if not icon_path.is_absolute():
            icon_path = base_dir / icon_path

    # 设置缓存目录
    build_dir = out_dir / "_build"
    spec_dir = out_dir / "_spec"
    spec_dir.mkdir(parents=True, exist_ok=True) # 确保 spec 存放目录存在

    try:
        # 开始打包
        run_pyinstaller(
            script_path=script_path,
            output_dir=out_dir,
            name=args.name,
            onefile=is_onefile,
            console=not args.noconsole,
            icon=icon_path,
            build_dir=build_dir,
            spec_dir=spec_dir,
        )
    except Exception as e:
        print(f"打包过程中发生错误: {e}")
    finally:
        # 6. 清理生成的临时文件
        if args.clean_artifacts:
            print("正在清理临时文件...")
            if build_dir.exists():
                shutil.rmtree(build_dir, ignore_errors=True)
            if spec_dir.exists():
                shutil.rmtree(spec_dir, ignore_errors=True)

    print(f"\n✅ 打包任务结束！请检查输出目录：{out_dir}")


if __name__ == "__main__":
    main()
