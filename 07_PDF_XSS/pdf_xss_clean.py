# -*- coding: utf-8 -*-
"""
批量清理 PDF 中潜在 XSS / 恶意脚本内容（保留外部 URL，仅移除 JS / 恶意协议链接）。

默认目录：
- 输入：07_PDF_XSS/input/
- 输出：07_PDF_XSS/output/

用法：
  python pdf_xss_clean.py
  python pdf_xss_clean.py --input "D:\\pdfs" --output "D:\\pdfs_clean"
  python pdf_xss_clean.py --overwrite
  python pdf_xss_clean.py --no-recursive
  python pdf_xss_clean.py --no-keep-structure
"""
from __future__ import annotations

import argparse
from pathlib import Path

import fitz  # PyMuPDF

BAD_PROTOCOLS = ("javascript:", "data:", "vbscript:", "file:")


def is_js_or_malicious_url(link: dict) -> bool:
    """判断链接是否为 JavaScript 或恶意协议。"""
    uri = (link.get("uri") or "").lower()

    # JS 动作
    if "javascript" in uri:
        return True

    # 常见恶意协议
    if any(uri.startswith(proto) for proto in BAD_PROTOCOLS):
        return True

    # JS 内嵌动作，如 /A << /S /JavaScript /JS (...) >>
    if link.get("kind") == 2:
        action = str(link)
        if "/JavaScript" in action or "/JS" in action:
            return True

    return False


def clean_pdf(pdf_path: Path, output_path: Path, overwrite: bool = False) -> bool:
    """
    清理单个 PDF。

    Args:
        pdf_path: 输入 PDF 文件路径
        output_path: 输出 PDF 文件路径
        overwrite: 是否覆盖已存在输出

    Returns:
        是否成功处理
    """
    if not pdf_path.exists():
        print(f"跳过：文件不存在 {pdf_path}")
        return False

    if output_path.exists() and not overwrite:
        print(f"跳过（已存在）：{output_path.name}")
        return False

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"打开失败：{pdf_path.name}，原因：{exc}")
        return False

    try:
        for page_index in range(len(doc)):
            page = doc[page_index]

            # ---- 删除注释（常包含 JS）----
            annot = page.first_annot
            while annot:
                next_annot = annot.next
                page.delete_annot(annot)
                annot = next_annot

            # ---- 检查每个链接 ----
            for link in page.get_links():
                if is_js_or_malicious_url(link):
                    page.delete_link(link)

        # ---- 删除嵌入文件 ----
        emb_count = doc.embfile_count()
        if emb_count > 0:
            for i in range(emb_count - 1, -1, -1):
                doc.embfile_del(i)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path, garbage=4, deflate=True)
        print(f"已清理：{pdf_path.name} -> {output_path.name}")
        return True
    except Exception as exc:
        print(f"处理失败：{pdf_path.name}，原因：{exc}")
        return False
    finally:
        doc.close()


def collect_pdfs(input_path: Path, recursive: bool = True) -> list[Path]:
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return sorted([p.resolve() for p in input_path.glob(pattern) if p.is_file()])
    return []


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(
        description="批量清理 PDF 中潜在 XSS/恶意脚本内容"
    )
    parser.add_argument(
        "--input",
        default=str(default_input),
        help="输入 PDF 文件夹或单个 PDF 文件",
    )
    parser.add_argument(
        "--output",
        default=str(default_output),
        help="输出文件夹（默认 07_PDF_XSS/output）",
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

    pdf_files = collect_pdfs(input_path, recursive=args.recursive)
    if not pdf_files:
        print(f"未找到 PDF 文件：{input_path}")
        return

    total = 0
    success = 0
    base_input_dir = input_path.resolve() if input_path.is_dir() else None
    for pdf_path in pdf_files:
        total += 1
        if base_input_dir and args.keep_structure:
            rel = pdf_path.relative_to(base_input_dir)
            output_path = (output_dir / rel).with_name(f"{rel.stem}_cleaned.pdf")
        else:
            output_path = output_dir / f"{pdf_path.stem}_cleaned.pdf"
        if clean_pdf(pdf_path, output_path, overwrite=args.overwrite):
            success += 1

    print(f"✅ 处理完成：{success}/{total} 个 PDF 已清理")


if __name__ == "__main__":
    main()
