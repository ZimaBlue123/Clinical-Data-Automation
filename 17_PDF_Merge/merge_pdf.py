# -*- coding: utf-8 -*-
"""
合并多个 PDF 为一个 PDF（按自然排序从低到高）。

默认目录：
- 输入：17_PDF_Merge/input/
- 输出：17_PDF_Merge/output/

排序规则（自然排序）：
- 数字按数值比较：1, 2, 10（而不是 1, 10, 2）
- 字母不区分大小写：a, b, c...
- 当 input 为目录时，按“相对路径”排序（可包含子文件夹），例如：
  input/1.pdf
  input/2.pdf
  input/a/b.pdf

用法：
  python merge_pdf.py
  python merge_pdf.py --output-name merged.pdf
  python merge_pdf.py --input "D:\\PDFS" --output "D:\\OUT" --output-name all.pdf
  python merge_pdf.py --overwrite
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path


_SPLIT_RE = re.compile(r"(\\d+)")
logger = logging.getLogger(__name__)


def natural_key(text: str) -> list[object]:
    parts: list[object] = []
    for p in _SPLIT_RE.split(text):
        if not p:
            continue
        if p.isdigit():
            parts.append(int(p))
        else:
            parts.append(p.casefold())
    return parts


def path_natural_key(rel_path: Path) -> list[object]:
    key: list[object] = []
    for part in rel_path.parts:
        key.extend(natural_key(part))
        key.append("\0")  # 分隔符，避免 "ab" + "c" 与 "a" + "bc" 混淆
    return key


def collect_pdfs(input_path: Path) -> tuple[list[Path], Path | None]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            return ([], None)
        return ([input_path.resolve()], None)

    if input_path.is_dir():
        pdfs = [p.resolve() for p in input_path.rglob("*.pdf") if p.is_file()]
        return (sorted(pdfs), input_path.resolve())

    return ([], None)


def merge_pdfs(
    pdf_files: list[Path],
    base_dir: Path | None,
    output_pdf: Path,
    overwrite: bool,
) -> int:
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore
    except ImportError:
        logger.error("action=dependency_missing name=pypdf")
        return 2

    if output_pdf.exists() and not overwrite:
        logger.error("action=output_exists output=%s hint=use_overwrite", output_pdf)
        return 1

    if base_dir:
        pdf_files = sorted(pdf_files, key=lambda p: path_natural_key(p.relative_to(base_dir)))
    else:
        pdf_files = sorted(pdf_files, key=lambda p: path_natural_key(Path(p.name)))

    writer = PdfWriter()
    added_files = 0
    added_pages = 0

    for pdf_path in pdf_files:
        try:
            reader = PdfReader(str(pdf_path))
            if getattr(reader, "is_encrypted", False):
                try:
                    reader.decrypt("")  # type: ignore[attr-defined]
                except Exception:
                    logger.warning("action=skip_encrypted file=%s reason=decrypt_failed", pdf_path)
                    continue

            pages = list(reader.pages)
            for page in pages:
                writer.add_page(page)
            added_files += 1
            added_pages += len(pages)
            logger.info("已加入: file=%s pages=%s", pdf_path, len(pages))
        except Exception as exc:
            logger.warning("跳过读取失败: file=%s reason=%s", pdf_path, exc)

    if added_pages == 0:
        logger.error("action=merge_empty reason=no_pages_added")
        return 1

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open("wb") as f:
        writer.write(f)

    logger.info("合并完成: files=%s/%s pages=%s output=%s", added_files, len(pdf_files), added_pages, output_pdf)
    return 0


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(description="合并多个 PDF 为一个 PDF（自然排序）")
    parser.add_argument("--input", "-i", default=str(default_input), help="输入目录或单个 PDF 文件")
    parser.add_argument("--output", "-o", default=str(default_output), help="输出目录（默认 17_PDF_Merge/output）")
    parser.add_argument("--output-name", default="merged.pdf", help="输出文件名（默认 merged.pdf）")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()
    output_pdf = output_dir / args.output_name

    pdf_files, base_input_dir = collect_pdfs(input_path)
    if not pdf_files:
        logger.error("action=input_not_found input=%s", input_path)
        raise SystemExit(1)

    raise SystemExit(merge_pdfs(pdf_files, base_input_dir, output_pdf, overwrite=args.overwrite))


if __name__ == "__main__":
    main()

