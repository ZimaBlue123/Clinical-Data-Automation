from __future__ import annotations


import fitz  # PyMuPDF

from .utils import Box, clamp_boxes_to_page, merge_boxes


def merge_boxes_by_page(
    doc: fitz.Document,
    boxes_by_page: dict[str, list[Box]],
    *,
    pad: float = 2.0,
) -> dict[str, list[Box]]:
    """
    Merge overlapping boxes to reduce fragmentation and improve stability.
    """
    out: dict[str, list[Box]] = {}
    for page_key, page_boxes in boxes_by_page.items():
        try:
            pno = int(page_key)
        except Exception:
            continue
        if pno < 1 or pno > len(doc):
            continue
        page = doc[pno - 1]
        page_rect = page.rect
        page_boxes_merged = merge_boxes(page_boxes, pad=pad)
        page_boxes_merged = clamp_boxes_to_page(
            page_boxes_merged,
            page_width=float(page_rect.width),
            page_height=float(page_rect.height),
        )
        if page_boxes_merged:
            out[str(pno)] = page_boxes_merged
    return out

