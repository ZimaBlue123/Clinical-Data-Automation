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
  python word_to_pdf.py --no-recursive
  python word_to_pdf.py --no-keep-structure
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def ensure_windows() -> None:
    if sys.platform != "win32":
        print("错误：Word 转 PDF 需要 Windows + Microsoft Word。", file=sys.stderr)
        sys.exit(1)


class WordPdfConverter:
    def __init__(self) -> None:
        self._win32 = None
        self._app = None

    def __enter__(self) -> "WordPdfConverter":
        try:
            import win32com.client  # type: ignore
        except ImportError:
            print("错误：需要安装 pywin32。请执行: pip install pywin32", file=sys.stderr)
            raise

        self._win32 = win32com.client
        self._app = win32com.client.Dispatch("Word.Application")
        self._app.Visible = False
        try:
            self._app.DisplayAlerts = 0
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._app:
            try:
                self._app.Quit()
            except Exception:
                pass
        self._app = None
        self._win32 = None

    def export(self, doc_path: Path, pdf_path: Path, overwrite: bool = False) -> bool:
        if not doc_path.exists():
            print(f"跳过：文件不存在 {doc_path}")
            return False

        if pdf_path.exists() and not overwrite:
            print(f"跳过（已存在）：{pdf_path}")
            return False

        doc = None
        try:
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            doc = self._app.Documents.Open(str(doc_path))  # type: ignore[union-attr]
            doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = PDF
            print(f"已转换：{doc_path} -> {pdf_path}")
            return True
        except Exception as exc:
            print(f"转换失败：{doc_path}，原因：{exc}")
            return False
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass


def collect_docs(input_path: Path, recursive: bool = True) -> list[Path]:
    allowed = {".doc", ".docx", ".docm"}
    if input_path.is_file() and input_path.suffix.lower() in allowed:
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*" if recursive else "*"
        files: list[Path] = []
        for p in input_path.glob(pattern):
            if not p.is_file():
                continue
            if p.suffix.lower() not in allowed:
                continue
            # 跳过 Office 生成的临时文件，如 "~$xxx.docx"
            if p.name.startswith("~$"):
                continue
            files.append(p.resolve())
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
    parser.add_argument(
        "--recursive",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="是否递归遍历子文件夹（默认开启）",
    )
    parser.add_argument(
        "--keep-structure",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="当输入为目录时，是否在输出目录中保留相对目录结构（默认开启）",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    doc_files = collect_docs(input_path, recursive=args.recursive)
    if not doc_files:
        print(f"未找到 Word 文件：{input_path}")
        return

    total = 0
    success = 0
    base_input_dir = input_path.resolve() if input_path.is_dir() else None
    try:
        with WordPdfConverter() as converter:
            for doc_path in doc_files:
                total += 1
                if base_input_dir and args.keep_structure:
                    rel = doc_path.relative_to(base_input_dir)
                    pdf_path = (output_dir / rel).with_suffix(".pdf")
                else:
                    pdf_path = output_dir / f"{doc_path.stem}.pdf"
                if converter.export(doc_path, pdf_path, overwrite=args.overwrite):
                    success += 1
    except ImportError:
        return

    print(f"✅ 处理完成：{success}/{total} 个文件已转换")


if __name__ == "__main__":
    main()
