from __future__ import annotations

from typing import Dict, List

import fitz  # PyMuPDF

from .utils import Box


def render_audit_pdf(
    doc: fitz.Document,
    boxes_by_page: Dict[str, List[Box]],
    *,
    output_path: str,
    stroke_rgb: tuple[float, float, float] = (1.0, 0.0, 0.0),
    fill_rgb: tuple[float, float, float] = (1.0, 0.8, 0.8),
    opacity: float = 0.35,
) -> None:
    """
    Render audit overlay rectangles onto a copy of the PDF.

    Note: this does NOT remove any watermark; it only adds an audit mask overlay
    so downstream reviewers can see detected exclusion zones.
    """
    for page_key, boxes in boxes_by_page.items():
        try:
            pno = int(page_key)
        except Exception:
            continue
        if pno < 1 or pno > len(doc):
            continue
        page = doc[pno - 1]
        for box in boxes:
            rect = fitz.Rect(box)
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=stroke_rgb, fill=fill_rgb)
            try:
                annot.set_opacity(opacity)
            except Exception:
                # Some PyMuPDF versions may not support opacity on annotations
                pass
            annot.update()

    doc.save(output_path)

