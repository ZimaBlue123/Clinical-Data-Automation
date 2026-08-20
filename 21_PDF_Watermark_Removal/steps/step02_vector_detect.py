from __future__ import annotations


import fitz  # PyMuPDF

from .utils import Box, clamp_boxes_to_page


def detect_vector_boxes(
    doc: fitz.Document,
    *,
    vector_keywords: list[str],
    pad: float = 2.0,
) -> dict[str, list[Box]]:
    """
    Detect exclusion boxes from vector text using `page.search_for`.

    Returns:
        boxes_by_page where page key is 1-based string, value is list of boxes.
    """
    boxes_by_page: dict[str, list[Box]] = {}

    for pno in range(len(doc)):
        page = doc[pno]
        page_rect = page.rect
        page_w = float(page_rect.width)
        page_h = float(page_rect.height)

        page_boxes: list[Box] = []
        for kw in vector_keywords:
            if not kw:
                continue
            try:
                rects = page.search_for(kw)
            except Exception:
                rects = []
            for r in rects:
                page_boxes.append(
                    (
                        float(r.x0) - pad,
                        float(r.y0) - pad,
                        float(r.x1) + pad,
                        float(r.y1) + pad,
                    )
                )

        # Clamp boxes into page coordinate space
        page_boxes = clamp_boxes_to_page(page_boxes, page_width=page_w, page_height=page_h)
        if page_boxes:
            boxes_by_page[str(pno + 1)] = page_boxes

    return boxes_by_page
