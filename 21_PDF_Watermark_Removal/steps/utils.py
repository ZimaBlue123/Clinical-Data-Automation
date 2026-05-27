from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pytesseract
from PIL import Image


Box = Tuple[float, float, float, float]  # (x0, top, x1, bottom)


def tesseract_cmd_from_env_and_default(
    env_var: str = "TESSERACT_CMD_PATH",
    default_cmd: str = r"D:\Tesseract-OCR\tesseract.exe",
) -> Optional[str]:
    """Resolve tesseract executable path (env var -> default)."""
    env_path = os.getenv(env_var)
    if env_path and Path(env_path).exists():
        return env_path
    if Path(default_cmd).exists():
        return default_cmd
    return None


def configure_tesseract(
    env_var: str = "TESSERACT_CMD_PATH",
    default_cmd: str = r"D:\Tesseract-OCR\tesseract.exe",
) -> bool:
    """Configure pytesseract.pytesseract.tesseract_cmd. Returns whether it succeeded."""
    cmd = tesseract_cmd_from_env_and_default(env_var=env_var, default_cmd=default_cmd)
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
        return True
    return False


def box_intersects(a: Box, b: Box) -> bool:
    ax0, atop, ax1, abot = a
    bx0, btop, bx1, bbot = b
    # Non-overlap cases
    if ax1 < bx0 or ax0 > bx1:
        return False
    if abot < btop or atop > bbot:
        return False
    return True


def box_union(a: Box, b: Box) -> Box:
    ax0, atop, ax1, abot = a
    bx0, btop, bx1, bbot = b
    return (min(ax0, bx0), min(atop, btop), max(ax1, bx1), max(abot, bbot))


def merge_boxes(boxes: Sequence[Box], *, pad: float = 2.0, max_boxes: int = 200) -> List[Box]:
    """
    Merge overlapping boxes with a small padding.
    Note: This is heuristic; tune pad if boxes are too tight/too loose.
    """
    if not boxes:
        return []

    normalized: List[Box] = []
    for (x0, top, x1, bottom) in boxes:
        nx0 = x0 - pad
        ntop = top - pad
        nx1 = x1 + pad
        nbottom = bottom + pad
        normalized.append((float(nx0), float(ntop), float(nx1), float(nbottom)))

    # Greedy merge
    merged: List[Box] = []
    for box in sorted(normalized, key=lambda x: (x[1], x[0])):  # sort by top, then x0
        cur = box
        did_merge = True
        while did_merge:
            did_merge = False
            for i in range(len(merged)):
                if box_intersects(cur, merged[i]):
                    cur = box_union(cur, merged[i])
                    merged.pop(i)
                    did_merge = True
                    break
        merged.append(cur)
        if len(merged) >= max_boxes:
            break

    # Final pass: merge any remaining overlaps
    final: List[Box] = []
    for box in merged:
        merged_into_existing = False
        for i in range(len(final)):
            if box_intersects(box, final[i]):
                final[i] = box_union(box, final[i])
                merged_into_existing = True
                break
        if not merged_into_existing:
            final.append(box)
    return final


def clamp_boxes_to_page(
    boxes: Sequence[Box],
    *,
    page_width: float,
    page_height: float,
) -> List[Box]:
    """Clamp boxes into [0, page_width] x [0, page_height] coordinate space."""
    clamped: List[Box] = []
    for x0, top, x1, bottom in boxes:
        nx0 = max(0.0, min(x0, page_width))
        nx1 = max(0.0, min(x1, page_width))
        ntop = max(0.0, min(top, page_height))
        nbottom = max(0.0, min(bottom, page_height))
        if nx1 <= nx0 or nbottom <= ntop:
            continue
        clamped.append((nx0, ntop, nx1, nbottom))
    return clamped


def load_boxes_json(path: str | Path) -> Dict[str, List[Box]]:
    """Load boxes.json in { "1": [[x0,top,x1,bottom], ...], ... } format."""
    import json

    p = Path(path)
    with p.open("r", encoding="utf-8") as f:
        raw = json.load(f)

    out: Dict[str, List[Box]] = {}
    for page_key, boxes in raw.items():
        if not isinstance(boxes, list):
            continue
        page_boxes: List[Box] = []
        for b in boxes:
            if not isinstance(b, list) or len(b) != 4:
                continue
            page_boxes.append((float(b[0]), float(b[1]), float(b[2]), float(b[3])))
        out[str(page_key)] = page_boxes
    return out


def save_boxes_json(path: str | Path, boxes_by_page: Dict[str, List[Box]]) -> None:
    """
    Save legacy boxes.json format:
      { "1": [[x0,top,x1,bottom], ...], "2": [...] }

    Prefer `save_boxes_json_v2()` for rotation/mediabox/cropbox metadata.
    """
    import json

    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    serializable: Dict[str, List[List[float]]] = {}
    for page_key, boxes in boxes_by_page.items():
        serializable[str(page_key)] = [[b[0], b[1], b[2], b[3]] for b in boxes]
    with p.open("w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=2, ensure_ascii=False)


def save_boxes_json_v2(
    path: str | Path,
    boxes_by_page: Dict[str, List[Box]],
    *,
    page_meta_by_page: Optional[Dict[str, Dict[str, Any]]] = None,
) -> None:
    """
    Save v2 boxes.json:
    {
      "0": {"rotation":0,"mediabox":[...],"cropbox":[...],"boxes":[[...],...]},
      "1": {...}
    }
    """
    import json

    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    serializable: Dict[str, Dict[str, Any]] = {}

    page_meta_by_page = page_meta_by_page or {}

    for page_key, boxes in boxes_by_page.items():
        pk = str(page_key)
        meta = page_meta_by_page.get(pk, {})
        serializable[pk] = {
            "rotation": meta.get("rotation", None),
            "mediabox": meta.get("mediabox", None),
            "cropbox": meta.get("cropbox", None),
            "page_width": meta.get("page_width", None),
            "page_height": meta.get("page_height", None),
            "boxes": [[b[0], b[1], b[2], b[3]] for b in boxes],
        }

    with p.open("w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=2, ensure_ascii=False)

