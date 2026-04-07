# -*- coding: utf-8 -*-
"""
批量将 PPT/PPTX 转换为 PDF。

默认目录：
- 输入：05_PPT_to_PDF/input/
- 输出：05_PPT_to_PDF/output/

依赖：Windows + 已安装 PowerPoint + pywin32

用法：
  python ppt_to_pdf.py
  python ppt_to_pdf.py --input "D:\\PPT" --output "D:\\PDF"
  python ppt_to_pdf.py --overwrite
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def ensure_windows() -> None:
    if sys.platform != "win32":
        print("错误：PPT 转 PDF 需要 Windows + PowerPoint。", file=sys.stderr)
        sys.exit(1)


def export_ppt_to_pdf(ppt_path: Path, pdf_path: Path, overwrite: bool = False) -> bool:
    try:
        import win32com.client
    except ImportError:
        print("错误：需要安装 pywin32。请执行: pip install pywin32", file=sys.stderr)
        return False

    if not ppt_path.exists():
        print(f"跳过：文件不存在 {ppt_path}")
        return False

    if pdf_path.exists() and not overwrite:
        print(f"跳过（已存在）：{pdf_path.name}")
        return False

    app = None
    prs = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = 1
        prs = app.Presentations.Open(str(ppt_path), WithWindow=False)
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        prs.SaveAs(str(pdf_path), FileFormat=32)  # 32 = PDF
        print(f"已转换：{ppt_path.name} -> {pdf_path.name}")
        return True
    except Exception as exc:
        print(f"转换失败：{ppt_path.name}，原因：{exc}")
        return False
    finally:
        if prs:
            try:
                prs.Close()
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass


def collect_ppts(input_path: Path) -> list[Path]:
    if input_path.is_file() and input_path.suffix.lower() in {".ppt", ".pptx"}:
        return [input_path]
    if input_path.is_dir():
        files = list(input_path.glob("*.ppt")) + list(input_path.glob("*.pptx"))
        return sorted(files)
    return []


def main() -> None:
    ensure_windows()

    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(description="批量将 PPT/PPTX 转换为 PDF")
    parser.add_argument(
        "--input",
        default=str(default_input),
        help="输入目录或单个 PPT/PPTX 文件",
    )
    parser.add_argument(
        "--output",
        default=str(default_output),
        help="输出目录（默认 05_PPT_to_PDF/output）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖已存在的输出文件",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    ppt_files = collect_ppts(input_path)
    if not ppt_files:
        print(f"未找到 PPT/PPTX 文件：{input_path}")
        return

    total = 0
    success = 0
    for ppt_path in ppt_files:
        total += 1
        pdf_name = f"{ppt_path.stem}.pdf"
        pdf_path = output_dir / pdf_name
        if export_ppt_to_pdf(ppt_path, pdf_path, overwrite=args.overwrite):
            success += 1

    print(f"✅ 处理完成：{success}/{total} 个文件已转换")


if __name__ == "__main__":
    main()
