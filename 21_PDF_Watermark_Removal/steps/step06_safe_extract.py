from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Optional

import logging

from src.pdf_reader import extract_text_from_pdf

from .utils import Box

logger = logging.getLogger(__name__)


def extract_clean_text_by_page(
    pdf_path: str | Path,
    boxes_by_page: Dict[str, List[Box]],
    *,
    page_meta_by_page: Optional[Dict[str, Dict[str, object]]] = None,
) -> Dict[str, str]:
    """
    Safe extraction: extract page text after excluding detected interference zones.
    """
    # Convert page keys to int for src.pdf_reader
    exclusion_boxes_by_page: Dict[int, object] = {}
    for page_key, boxes in boxes_by_page.items():
        try:
            pn = int(page_key)
        except Exception:
            continue
        if not boxes:
            continue

        if page_meta_by_page and page_key in page_meta_by_page:
            meta = page_meta_by_page[page_key]
            exclusion_boxes_by_page[pn] = {
                "rotation": meta.get("rotation", 0),
                "page_width": meta.get("page_width", None),
                "page_height": meta.get("page_height", None),
                "mediabox": meta.get("mediabox", None),
                "cropbox": meta.get("cropbox", None),
                "boxes": boxes,
            }
        else:
            exclusion_boxes_by_page[pn] = boxes

    text_map = extract_text_from_pdf(
        pdf_path,
        page_numbers=None,
        exclusion_boxes_by_page=exclusion_boxes_by_page,
    )

    # Ensure string keys
    out: Dict[str, str] = {str(pn): text for pn, text in text_map.items()}
    return out


def save_text_map_json(output_path: str | Path, text_map: Dict[str, str]) -> None:
    p = Path(output_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(text_map, f, indent=2, ensure_ascii=False)

