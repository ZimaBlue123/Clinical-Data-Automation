# -*- coding: utf-8 -*-
"""
批量将 Word 文档转换为 PDF。

默认目录：
- 输入：09_Word_to_PDF/input/
- 输出：09_Word_to_PDF/output/

依赖：Windows + 已安装 Microsoft Word + pywin32

用法：
  python word_to_pdf.py
  python word_to_pdf.py --input "D:\\DOCS" --output "D:\\PDF"
  python word_to_pdf.py --overwrite
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def ensure_windows() -> None:
    if sys.platform != "win32":
        print("错误：Word 转 PDF 需要 Windows + Microsoft Word。", file=sys.stderr)
        sys.exit(1)


def export_word_to_pdf(doc_path: Path, pdf_path: Path, overwrite: bool = False) -> bool:
    try:
        import win32com.client
    except ImportError:
        print("错误：需要安装 pywin32。请执行: pip install pywin32", file=sys.stderr)
        return False

    if not doc_path.exists():
        print(f"跳过：文件不存在 {doc_path}")
        return False

    if pdf_path.exists() and not overwrite:
        print(f"跳过（已存在）：{pdf_path.name}")
        return False

    app = None
    doc = None
    try:
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        doc = app.Documents.Open(str(doc_path))
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = PDF
        print(f"已转换：{doc_path.name} -> {pdf_path.name}")
        return True
    except Exception as exc:
        print(f"转换失败：{doc_path.name}，原因：{exc}")
        return False
    finally:
        if doc:
            try:
                doc.Close(False)
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass


def collect_docs(input_path: Path) -> list[Path]:
    if input_path.is_file() and input_path.suffix.lower() in {".doc", ".docx"}:
        return [input_path]
    if input_path.is_dir():
        files = list(input_path.glob("*.doc")) + list(input_path.glob("*.docx"))
        return sorted(files)
    return []


def main() -> None:
    ensure_windows()

    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(description="批量将 Word 文档转换为 PDF")
    parser.add_argument(
        "--input",
        default=str(default_input),
        help="输入目录或单个 Word 文件",
    )
    parser.add_argument(
        "--output",
        default=str(default_output),
        help="输出目录（默认 09_Word_to_PDF/output）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖已存在的输出文件",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    doc_files = collect_docs(input_path)
    if not doc_files:
        print(f"未找到 Word 文件：{input_path}")
        return

    total = 0
    success = 0
    for doc_path in doc_files:
        total += 1
        pdf_name = f"{doc_path.stem}.pdf"
        pdf_path = output_dir / pdf_name
        if export_word_to_pdf(doc_path, pdf_path, overwrite=args.overwrite):
            success += 1

    print(f"✅ 处理完成：{success}/{total} 个文件已转换")


if __name__ == "__main__":
    main()
