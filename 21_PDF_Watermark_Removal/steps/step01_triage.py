from __future__ import annotations


import fitz  # PyMuPDF

from .utils import Box


def triage_document(
    doc: fitz.Document,
    *,
    vector_keywords: list[str],
    min_vector_hit_pages: int = 1,
    max_pages_to_probe: int = 3,
) -> tuple[str, dict[int, list[Box]]]:
    """
    Document triage:
    - Try fast vector keyword search using `page.search_for`.
    - If we find hits on at least `min_vector_hit_pages`, route to "vector".
    - Otherwise, route to OCR-based localization ("ocr").
    """
    per_page_hits: dict[int, list[Box]] = {}
    hit_pages = 0

    probe_pages = min(len(doc), max_pages_to_probe)
    for pno in range(probe_pages):
        page = doc[pno]
        bboxes: list[Box] = []
        for kw in vector_keywords:
            if not kw:
                continue
            try:
                rects = page.search_for(kw)
            except Exception:
                rects = []
            for r in rects:
                # PyMuPDF Rect: (x0,y0,x1,y1). We treat y as "top" coordinate.
                bboxes.append((float(r.x0), float(r.y0), float(r.x1), float(r.y1)))
        if bboxes:
            hit_pages += 1
            per_page_hits[pno + 1] = bboxes

    if hit_pages >= min_vector_hit_pages:
        return "vector", per_page_hits
    return "ocr", {}
