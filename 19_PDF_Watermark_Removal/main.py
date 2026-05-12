from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    # Allow `import src.*` when running from module directory.
    sys.path.insert(0, str(ROOT_DIR))

from steps.step01_triage import triage_document
from steps.step02_vector_detect import detect_vector_boxes
from steps.step03_ocr_detect import detect_ocr_boxes
from steps.step04_merge_boxes import merge_boxes_by_page
from steps.step05_audit_render import render_audit_pdf
from steps.step06_safe_extract import extract_clean_text_by_page, save_text_map_json
from steps.utils import Box, configure_tesseract, save_boxes_json_v2


logger = logging.getLogger("watermark_exclusion")


def iter_pdfs(input_path: Path, *, recursive: bool) -> List[Path]:
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        pdfs = [p.resolve() for p in input_path.glob(pattern) if p.is_file()]
        return sorted(pdfs)
    return []


def resolve_io_paths(
    base_input_dir: Optional[Path],
    pdf_path: Path,
    output_root: Path,
    *,
    keep_structure: bool,
) -> Path:
    if base_input_dir and keep_structure:
        rel_parent = pdf_path.resolve().relative_to(base_input_dir).parent
        out_dir = (output_root / rel_parent).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        return out_dir
    output_root.mkdir(parents=True, exist_ok=True)
    return output_root


def parse_keywords(raw: str | None) -> List[str]:
    if not raw:
        return [r"CONFIDENTIAL", r"DRAFT", r"仅供审查", r"内部使用"]
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    return parts or [r"CONFIDENTIAL", r"DRAFT", r"仅供审查", r"内部使用"]


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    parser = argparse.ArgumentParser(description="19_PDF_Watermark_Removal: detect interference zones + audit masks + clean extracted text.")
    parser.add_argument("--input", "-i", default="input", help="输入 PDF（文件或目录，默认相对当前模块目录）")
    parser.add_argument("--output", "-o", default="output", help="输出目录（默认相对当前模块目录）")
    parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=True, help="递归遍历子文件夹（默认开启）")
    parser.add_argument("--keep-structure", action=argparse.BooleanOptionalAction, default=True, help="保留相对目录结构（默认开启）")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在输出")

    parser.add_argument("--keywords", default=None, help="逗号分隔的关键词列表（用于 vector/OCR 命中；用于定位疑似水印/页眉页脚）")
    parser.add_argument("--vector-min-hit-pages", type=int, default=1, help="vector 命中页面数达到该阈值则优先使用矢量定位（默认 1）")

    parser.add_argument("--ocr-dpi", type=int, default=200, help="OCR 渲染分辨率（dpi，默认 200）")
    parser.add_argument("--ocr-conf-thresh", type=float, default=50.0, help="OCR 词置信度阈值（默认 50）")
    parser.add_argument("--ocr-lang", default="chi_sim+eng", help="tesseract 语言（默认 chi_sim+eng）")
    parser.add_argument("--ocr-repeated-heuristic", action=argparse.BooleanOptionalAction, default=True, help="若关键词命中失败，启用重复词启发式（默认开启）")
    parser.add_argument("--ocr-repeated-min-pages", type=int, default=3, help="重复词启发式最少跨页数（默认 3）")

    args = parser.parse_args()

    module_dir = Path(__file__).resolve().parent
    input_path = Path(args.input)
    if not input_path.is_absolute():
        input_path = module_dir / input_path
    output_root = Path(args.output)
    if not output_root.is_absolute():
        output_root = module_dir / output_root
    output_root.mkdir(parents=True, exist_ok=True)

    pdfs = iter_pdfs(input_path, recursive=args.recursive)
    if not pdfs:
        logger.error("action=input_not_found input=%s", input_path)
        raise SystemExit(1)

    # Configure OCR engine once
    ocr_ok = configure_tesseract()
    if not ocr_ok:
        logger.warning("action=ocr_config_warning backend=tesseract status=unavailable")

    base_input_dir = input_path.resolve() if input_path.is_dir() else None

    success_count = 0
    failed_count = 0

    for pdf_path in pdfs:
        out_dir = resolve_io_paths(
            base_input_dir,
            pdf_path,
            output_root,
            keep_structure=args.keep_structure,
        )

        stem = pdf_path.stem
        boxes_json_path = out_dir / f"{stem}_boxes.json"
        audit_pdf_path = out_dir / f"{stem}_audit_masked.pdf"
        clean_text_json_path = out_dir / f"{stem}_clean_text_by_page.json"
        report_json_path = out_dir / f"{stem}_watermark_report.json"

        if not args.overwrite:
            if boxes_json_path.exists() and audit_pdf_path.exists() and clean_text_json_path.exists():
                logger.info("action=skip_existing stem=%s", stem)
                continue

        keywords = parse_keywords(args.keywords)
        logger.info("action=process_start file=%s keywords=%s", pdf_path.name, ",".join(keywords))

        mode_used = "unknown"
        boxes_by_page: Dict[str, List[Box]] = {}
        page_meta_by_page: Dict[str, Dict[str, object]] = {}

        doc = fitz.open(pdf_path)
        try:
            # 1) Triage route
            route, _ = triage_document(
                doc,
                vector_keywords=keywords,
                min_vector_hit_pages=args.vector_min_hit_pages,
            )
            mode_used = route

            # 2) Detect boxes
            if mode_used == "vector":
                boxes_by_page = detect_vector_boxes(doc, vector_keywords=keywords, pad=2.0)
            else:
                boxes_by_page = detect_ocr_boxes(
                    doc,
                    vector_keywords=keywords,
                    ocr_dpi=args.ocr_dpi,
                    conf_thresh=args.ocr_conf_thresh,
                    lang=args.ocr_lang,
                    repeated_heuristic=bool(args.ocr_repeated_heuristic),
                    repeated_min_pages=args.ocr_repeated_min_pages,
                )

            # 3) Merge / refine
            boxes_by_page = merge_boxes_by_page(doc, boxes_by_page, pad=2.0)

            # Collect page meta for rotation/cropbox/mediabox auditing.
            for pno in range(len(doc)):
                page = doc[pno]
                pk = str(pno + 1)
                page_meta_by_page[pk] = {
                    "rotation": getattr(page, "rotation", 0),
                    "mediabox": [float(v) for v in page.mediabox],
                    "cropbox": [float(v) for v in page.cropbox],
                    "page_width": float(page.rect.width),
                    "page_height": float(page.rect.height),
                }

            # 4) Save audit + audit boxes
            save_boxes_json_v2(
                boxes_json_path,
                boxes_by_page,
                page_meta_by_page=page_meta_by_page,
            )

            # 5) Audit render (overlay only)
            if args.overwrite and audit_pdf_path.exists():
                audit_pdf_path.unlink(missing_ok=True)
            render_audit_pdf(doc, boxes_by_page, output_path=str(audit_pdf_path))

            # 6) Safe extraction for downstream text/keyword matching
            clean_text_map = extract_clean_text_by_page(
                pdf_path,
                boxes_by_page,
                page_meta_by_page=page_meta_by_page,
            )
            save_text_map_json(clean_text_json_path, clean_text_map)

            # 7) Minimal report for traceability
            page_counts = {str(p): len(b) for p, b in boxes_by_page.items()}
            report = {
                "input_pdf": str(pdf_path),
                "mode_used": mode_used,
                "keywords": keywords,
                "pages_detected": sorted([int(p) for p in page_counts.keys()]) if page_counts else [],
                "boxes_count_by_page": page_counts,
            }
            if report_json_path.exists() and args.overwrite:
                report_json_path.unlink(missing_ok=True)

            with report_json_path.open("w", encoding="utf-8") as f:
                json.dump(report, f, indent=2, ensure_ascii=False)

            logger.info("action=process_success file=%s output_dir=%s", pdf_path.name, out_dir)
            success_count += 1
        except Exception as exc:
            failed_count += 1
            logger.exception("处理失败: file=%s", pdf_path.name)
        finally:
            doc.close()

    logger.info("批处理完成: success=%s failed=%s total=%s", success_count, failed_count, len(pdfs))
    raise SystemExit(0 if failed_count == 0 else 1)


if __name__ == "__main__":
    main()

