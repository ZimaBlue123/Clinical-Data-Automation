from __future__ import annotations

import re
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from pytesseract import Output

from .utils import Box, clamp_boxes_to_page


_DEFAULT_STOPWORDS = {
    # Common English words
    "the",
    "and",
    "for",
    "with",
    "from",
    "this",
    "that",
    "are",
    "was",
    "were",
    "you",
    "your",
    "not",
    # Common doc meta tokens (very frequent)
    "page",
    "pdf",
}


def _normalize_word(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    # Keep alnum/underscore; keep CJK as-is
    s = re.sub(r"[^\w\u4e00-\u9fa5]+", "", s)
    return s.lower()


def detect_ocr_boxes(
    doc: fitz.Document,
    *,
    vector_keywords: List[str],
    ocr_dpi: int = 200,
    conf_thresh: float = 50.0,
    lang: str = "chi_sim+eng",
    repeated_heuristic: bool = True,
    repeated_min_pages: int = 3,
) -> Dict[str, List[Box]]:
    """
    OCR-based exclusion box detection for scanned PDFs.

    Strategy:
    1) Prefer keyword matching (vector_keywords) from OCR word-level boxes.
    2) If no keyword boxes found, optionally use a repeated-word heuristic:
       frequent OCR words across pages become candidates.
    """
    kw_lowers = [k.strip().lower() for k in vector_keywords if k and k.strip()]
    boxes_by_page: Dict[str, List[Box]] = {}

    # For repeated heuristic
    word_page_boxes: Dict[str, List[Tuple[int, Box]]] = {}

    for pno in range(len(doc)):
        page = doc[pno]
        page_rect = page.rect
        page_w = float(page_rect.width)
        page_h = float(page_rect.height)

        # Render to image
        pix = page.get_pixmap(dpi=ocr_dpi, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        data = pytesseract.image_to_data(img, output_type=Output.DICT, lang=lang)
        n = len(data.get("text", []))

        page_boxes: List[Box] = []

        for i in range(n):
            text = (data["text"][i] or "").strip()
            if not text:
                continue

            try:
                conf = float(data["conf"][i])
            except Exception:
                conf = -1.0
            if conf < conf_thresh:
                continue

            left = int(data["left"][i])
            top = int(data["top"][i])
            width = int(data["width"][i])
            height = int(data["height"][i])

            scale_x = float(pix.width) / page_w if page_w else 1.0
            scale_y = float(pix.height) / page_h if page_h else 1.0
            if scale_x == 0 or scale_y == 0:
                continue

            x0 = left / scale_x
            top_pdf = top / scale_y
            x1 = (left + width) / scale_x
            bottom_pdf = (top + height) / scale_y
            box = (float(x0), float(top_pdf), float(x1), float(bottom_pdf))

            text_lower = text.lower()
            matched_kw = False
            for kw in kw_lowers:
                if kw and kw in text_lower:
                    matched_kw = True
                    break

            if matched_kw and kw_lowers:
                page_boxes.append(box)

            if repeated_heuristic:
                norm = _normalize_word(text)
                if not norm:
                    continue
                if norm in _DEFAULT_STOPWORDS:
                    continue
                # Very short tokens are rarely useful for watermarking
                if len(norm) < 2:
                    continue

                word_page_boxes.setdefault(norm, []).append((pno + 1, box))

        if page_boxes:
            page_boxes = clamp_boxes_to_page(page_boxes, page_width=page_w, page_height=page_h)
            boxes_by_page[str(pno + 1)] = page_boxes

    # If keyword matching succeeded, use it.
    if boxes_by_page:
        return boxes_by_page

    if not repeated_heuristic:
        return boxes_by_page

    # Choose frequent words across pages
    word_to_pages: Dict[str, set[int]] = {}
    for word, occurrences in word_page_boxes.items():
        pages = {p for (p, _b) in occurrences}
        word_to_pages[word] = pages

    candidates = [w for w, pages in word_to_pages.items() if len(pages) >= repeated_min_pages]
    if not candidates:
        return boxes_by_page

    # Build boxes from candidates
    # Note: we do not try to perfectly distinguish "watermark" vs "real text" here;
    # exclusions are an *audit-driven* best-effort.
    for word in candidates:
        for p, box in word_page_boxes.get(word, []):
            boxes_by_page.setdefault(str(p), []).append(box)

    # Final clamp
    for pno_str, page_boxes in list(boxes_by_page.items()):
        # Convert page number
        pno = int(pno_str)
        page = doc[pno - 1]
        page_rect = page.rect
        page_boxes = clamp_boxes_to_page(
            page_boxes, page_width=float(page_rect.width), page_height=float(page_rect.height)
        )
        if page_boxes:
            boxes_by_page[pno_str] = page_boxes
        else:
            boxes_by_page.pop(pno_str, None)

    return boxes_by_page

