# -*- coding: utf-8 -*-
"""
从 PDF 中按关键词/页码检索文本或表格，供写入 Excel 使用。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pdfplumber

logger = logging.getLogger(__name__)

Box = Tuple[float, float, float, float]  # (x0, top, x1, bottom)
ExclusionPageData = Union[List[Box], Dict[str, Any]]


def _intersects(a: Box, b: Box) -> bool:
    ax0, atop, ax1, abot = a
    bx0, btop, bx1, bbot = b
    if ax1 < bx0 or ax0 > bx1:
        return False
    if abot < btop or atop > bbot:
        return False
    return True


def _rotation_to_int(rotation: Any) -> int:
    """Normalize rotation degrees into {0,90,180,270} clockwise."""
    try:
        r = int(rotation) % 360
    except Exception:
        r = 0
    if r == 0:
        return 0
    if r == 90:
        return 90
    if r == 180:
        return 180
    if r == 270:
        return 270
    # If the PDF uses unexpected angles, fail closed to no-transform.
    return 0


def _unrotate_point_to_base(x: float, y: float, r_fitz: int, base_w: float, base_h: float) -> Tuple[float, float]:
    """Convert (x,y) from a rotated coordinate system into rotation-0 base system."""
    if r_fitz == 0:
        return (x, y)
    if r_fitz == 90:
        xb = y
        yb = base_h - x
        return (xb, yb)
    if r_fitz == 180:
        xb = base_w - x
        yb = base_h - y
        return (xb, yb)
    if r_fitz == 270:
        xb = base_w - y
        yb = x
        return (xb, yb)
    return (x, y)


def _rotate_base_to_target(xb: float, yb: float, r_target: int, base_w: float, base_h: float) -> Tuple[float, float]:
    """Convert base (rotation-0) point into target rotated coordinate system."""
    if r_target == 0:
        return (xb, yb)
    if r_target == 90:
        xr = base_h - yb
        yr = xb
        return (xr, yr)
    if r_target == 180:
        xr = base_w - xb
        yr = base_h - yb
        return (xr, yr)
    if r_target == 270:
        xr = yb
        yr = base_w - xb
        return (xr, yr)
    return (xb, yb)


def _map_box_between_rotations(
    box: Box,
    *,
    r_fitz: int,
    fitz_page_width: float,
    fitz_page_height: float,
    r_pdf: int,
) -> Box:
    """
    Map a rectangle box from fitz coordinate system into pdfplumber coordinate system,
    accounting for rotation (0/90/180/270) via point rotation around rotation-0 base frame.
    """
    # Base dims: when rotating 90/270, width/height swap.
    if r_fitz in (0, 180):
        base_w = fitz_page_width
        base_h = fitz_page_height
    else:
        base_w = fitz_page_height
        base_h = fitz_page_width

    x0, top, x1, bottom = box

    corners = [(x0, top), (x1, top), (x0, bottom), (x1, bottom)]
    mapped: List[Tuple[float, float]] = []
    for x, y in corners:
        xb, yb = _unrotate_point_to_base(x, y, r_fitz=r_fitz, base_w=base_w, base_h=base_h)
        xr, yr = _rotate_base_to_target(xb, yb, r_target=r_pdf, base_w=base_w, base_h=base_h)
        mapped.append((xr, yr))

    xs = [p[0] for p in mapped]
    ys = [p[1] for p in mapped]
    return (min(xs), min(ys), max(xs), max(ys))


def _clamp_box_to_page(box: Box, *, page_width: float, page_height: float) -> Optional[Box]:
    x0, top, x1, bottom = box
    nx0 = max(0.0, min(float(x0), page_width))
    nx1 = max(0.0, min(float(x1), page_width))
    ntop = max(0.0, min(float(top), page_height))
    nbottom = max(0.0, min(float(bottom), page_height))
    if nx1 <= nx0 or nbottom <= ntop:
        return None
    return (nx0, ntop, nx1, nbottom)


def _box_area(box: Box) -> float:
    x0, top, x1, bottom = box
    w = float(x1) - float(x0)
    h = float(bottom) - float(top)
    if w <= 0 or h <= 0:
        return 0.0
    return w * h


# 正交旋转下面积应守恒；超出此区间视为可疑（浮点 + 非标准页）
_AREA_RATIO_OK_MIN = 0.98
_AREA_RATIO_OK_MAX = 1.02


def prepare_exclusion_boxes_with_audit(
    page: "pdfplumber.page.Page",
    exclusion_page: ExclusionPageData,
) -> Tuple[List[Box], Dict[str, Any]]:
    """
    将排除框映射到当前 pdfplumber 页坐标系，并输出单页审计明细。

    Returns:
        (用于 filter 的最终 boxes, page_detail 字典)
    """
    detail: Dict[str, Any] = {
        "original_rotation_fitz": 0,
        "pdfplumber_rotation": _rotation_to_int(getattr(page, "rotation", 0)),
        "fitz_page_width": None,
        "fitz_page_height": None,
        "pdfplumber_page_width": float(getattr(page, "width", 0) or 0),
        "pdfplumber_page_height": float(getattr(page, "height", 0) or 0),
        "input_boxes_count": 0,
        "mapped_boxes_count": 0,
        "clamped_count": 0,
        "dropped_count": 0,
        "severe_area_distortion": False,
        "box_details": [],
    }

    if not exclusion_page:
        return [], detail

    if isinstance(exclusion_page, list):
        exclusion_boxes = exclusion_page
        fitz_page_width = None
        fitz_page_height = None
        r_fitz = 0
    else:
        exclusion_boxes = exclusion_page.get("boxes") or []
        r_fitz = _rotation_to_int(exclusion_page.get("rotation"))
        fitz_page_width = exclusion_page.get("page_width")
        fitz_page_height = exclusion_page.get("page_height")

    detail["original_rotation_fitz"] = r_fitz
    detail["fitz_page_width"] = fitz_page_width
    detail["fitz_page_height"] = fitz_page_height

    exclusion_boxes_norm = [
        (float(x0), float(top), float(x1), float(bottom)) for (x0, top, x1, bottom) in exclusion_boxes
    ]
    detail["input_boxes_count"] = len(exclusion_boxes_norm)

    page_w = detail["pdfplumber_page_width"]
    page_h = detail["pdfplumber_page_height"]
    r_pdf = detail["pdfplumber_rotation"]

    final_boxes: List[Box] = []

    if not exclusion_boxes_norm:
        detail["mapped_boxes_count"] = 0
        return [], detail

    # 无旋转元数据或旋转一致：不做 fitz→pdfplumber 映射，直接使用原框（仍可做 clamp 防越界）
    need_map = (
        (fitz_page_width is not None)
        and (fitz_page_height is not None)
        and (r_fitz != r_pdf)
    )

    for idx, b in enumerate(exclusion_boxes_norm):
        area_orig = _box_area(b)
        entry: Dict[str, Any] = {
            "index": idx,
            "original": [b[0], b[1], b[2], b[3]],
            "area_original": round(area_orig, 4),
        }

        if need_map:
            try:
                mb = _map_box_between_rotations(
                    b,
                    r_fitz=r_fitz,
                    fitz_page_width=float(fitz_page_width),
                    fitz_page_height=float(fitz_page_height),
                    r_pdf=r_pdf,
                )
                area_mapped = _box_area(mb)
                entry["after_rotation_map"] = [mb[0], mb[1], mb[2], mb[3]]
                entry["area_after_map"] = round(area_mapped, 4)
                if area_orig > 1e-9:
                    ratio = area_mapped / area_orig
                    entry["area_retention_ratio"] = round(ratio, 6)
                    if ratio < _AREA_RATIO_OK_MIN or ratio > _AREA_RATIO_OK_MAX:
                        detail["severe_area_distortion"] = True
                else:
                    entry["area_retention_ratio"] = None
                pre_clamp = mb
            except Exception as e:
                logger.warning("box rotation map failed, using raw box: %s", e)
                pre_clamp = b
                entry["after_rotation_map"] = None
                entry["map_error"] = str(e)
        else:
            pre_clamp = b
            entry["after_rotation_map"] = None
            entry["area_retention_ratio"] = 1.0 if area_orig > 1e-9 else None

        if page_w and page_h:
            cb = _clamp_box_to_page(pre_clamp, page_width=page_w, page_height=page_h)
            if cb is None:
                detail["dropped_count"] += 1
                entry["dropped"] = True
                entry["after_clamp"] = None
                detail["box_details"].append(entry)
                continue

            entry["dropped"] = False
            entry["after_clamp"] = [cb[0], cb[1], cb[2], cb[3]]
            entry["area_after_clamp"] = round(_box_area(cb), 4)

            if (
                abs(cb[0] - pre_clamp[0]) > 1e-6
                or abs(cb[1] - pre_clamp[1]) > 1e-6
                or abs(cb[2] - pre_clamp[2]) > 1e-6
                or abs(cb[3] - pre_clamp[3]) > 1e-6
            ):
                detail["clamped_count"] += 1
                entry["clamped"] = True
                entry["clamp_delta"] = {
                    "x0": round(cb[0] - pre_clamp[0], 4),
                    "top": round(cb[1] - pre_clamp[1], 4),
                    "x1": round(cb[2] - pre_clamp[2], 4),
                    "bottom": round(cb[3] - pre_clamp[3], 4),
                }
            else:
                entry["clamped"] = False
                entry["clamp_delta"] = None

            final_boxes.append(cb)
        else:
            final_boxes.append(pre_clamp)
            entry["dropped"] = False
            entry["after_clamp"] = [pre_clamp[0], pre_clamp[1], pre_clamp[2], pre_clamp[3]]

        detail["box_details"].append(entry)

    detail["mapped_boxes_count"] = len(final_boxes)
    return final_boxes, detail


def build_mapping_audit_for_pdf(
    pdf_path: str | Path,
    exclusion_boxes_by_page: Dict[int, ExclusionPageData],
) -> Dict[str, Any]:
    """
    对整份 PDF 逐页计算排除框坐标映射审计（面积守恒、钳制、丢弃）。

    在 `03_PDF_to_Excel` 使用 `--exclusion-json` 时调用，结果可合并入 `*_watermark_report.json`
    的 `mapping_audit` 节点，满足可追溯审计要求。
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")
    if not pdf_path.is_file():
        raise ValueError(f"路径不是文件: {pdf_path}")

    mapping_audit: Dict[str, Any] = {
        "total_pages_processed": 0,
        "pages_with_exclusion": [],
        "pages_with_rotation": [],
        "anomalies": {
            "clamped_boxes_count": 0,
            "dropped_boxes_count": 0,
            "severe_area_distortion_pages": [],
        },
        "page_level_details": {},
    }

    with pdfplumber.open(pdf_path) as pdf:
        n = len(pdf.pages)
        mapping_audit["total_pages_processed"] = n

        for pno in range(1, n + 1):
            page_excl = exclusion_boxes_by_page.get(pno) if exclusion_boxes_by_page else None
            if not page_excl:
                continue

            page = pdf.pages[pno - 1]
            _, detail = prepare_exclusion_boxes_with_audit(page, page_excl)

            mapping_audit["pages_with_exclusion"].append(pno)

            r_fitz = int(detail.get("original_rotation_fitz") or 0)
            r_pdf = int(detail.get("pdfplumber_rotation") or 0)
            if (r_fitz % 360) != 0 or (r_pdf % 360) != 0:
                mapping_audit["pages_with_rotation"].append(pno)

            mapping_audit["anomalies"]["clamped_boxes_count"] += int(detail.get("clamped_count") or 0)
            mapping_audit["anomalies"]["dropped_boxes_count"] += int(detail.get("dropped_count") or 0)
            if detail.get("severe_area_distortion"):
                mapping_audit["anomalies"]["severe_area_distortion_pages"].append(pno)

            mapping_audit["page_level_details"][str(pno)] = detail

    return mapping_audit


def _filter_page_by_exclusion(
    page: "pdfplumber.page.Page",
    exclusion_page: ExclusionPageData,
):
    """
    Filter out *text/char-like* objects inside exclusion boxes before calling `extract_text` / `extract_tables`.

    Important:
    - Do NOT filter geometry objects (lines/rect/curves). Only exclude text-like objects,
      otherwise table gridlines may be broken and cell alignment may shift.
    """
    if not exclusion_page:
        return page

    exclusion_boxes_norm, _detail = prepare_exclusion_boxes_with_audit(page, exclusion_page)  # noqa: F841

    def predicate(obj) -> bool:
        # pdfplumber 'filter' runs over page layout objects (chars/lines/words/rectangles depending on extraction).
        # We only exclude objects that look like text chars; geometry objects should always be preserved.
        obj_text = obj.get("text")
        is_text_obj = isinstance(obj_text, str) and bool(obj_text.strip())
        if not is_text_obj:
            return True

        try:
            x0 = float(obj.get("x0"))
            top = float(obj.get("top"))
            x1 = float(obj.get("x1"))
            bottom = float(obj.get("bottom"))
        except Exception:
            # Keep objects that don't have bbox info
            return True

        obj_box: Box = (x0, top, x1, bottom)
        for ex_box in exclusion_boxes_norm:
            if _intersects(obj_box, ex_box):
                return False
        return True

    try:
        return page.filter(predicate)
    except Exception as e:
        logger.warning("page.filter failed; fallback to unfiltered extraction: %s", e)
        return page


def extract_text_from_pdf(
    pdf_path: str | Path,
    page_numbers: Optional[list[int]] = None,
    exclusion_boxes_by_page: Optional[Dict[int, ExclusionPageData]] = None,
) -> dict[int, str]:
    """
    提取 PDF 指定页的文本。不指定页码则提取全部页。
    
    Args:
        pdf_path: PDF 文件路径
        page_numbers: 要提取的页码列表（从1开始），None 表示提取全部页
        
    Returns:
        { 页码: 该页全文 } 字典
        
    Raises:
        FileNotFoundError: PDF 文件不存在
        ValueError: PDF 文件无法打开或读取
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")
    
    if not pdf_path.is_file():
        raise ValueError(f"路径不是文件: {pdf_path}")

    result = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                logger.warning(f"PDF 文件没有页面: {pdf_path}")
                return result
                
            pages = page_numbers if page_numbers is not None else range(1, len(pdf.pages) + 1)
            for pno in pages:
                if pno < 1 or pno > len(pdf.pages):
                    logger.warning(f"页码 {pno} 超出范围 (1-{len(pdf.pages)})，跳过")
                    continue
                try:
                    page = pdf.pages[pno - 1]
                    if exclusion_boxes_by_page:
                        page_excl = exclusion_boxes_by_page.get(pno)
                    else:
                        page_excl = None

                    if page_excl:
                        page = _filter_page_by_exclusion(page, page_excl)

                    text = page.extract_text() or ""
                    result[pno] = text
                except Exception as e:
                    logger.error(f"提取第 {pno} 页文本时出错: {e}")
                    result[pno] = ""
    except Exception as e:
        raise ValueError(f"无法打开或读取 PDF 文件 {pdf_path}: {e}") from e
        
    return result


def search_by_keyword(
    pdf_path: str | Path,
    keyword: str,
    page_number: Optional[int] = None,
    context_chars: int = 500,
    exclusion_boxes_by_page: Optional[Dict[int, ExclusionPageData]] = None,
) -> list[tuple[int, str]]:
    """
    在 PDF 中按关键词检索，返回 (页码, 匹配到的段落/上下文) 列表。
    
    Args:
        pdf_path: PDF 文件路径
        keyword: 要搜索的关键词
        page_number: 指定页码（从1开始），None 表示搜索全部页
        context_chars: 上下文字符数
        
    Returns:
        (页码, 匹配文本) 列表
    """
    if not keyword or not keyword.strip():
        logger.warning("关键词为空，返回空结果")
        return []
    
    if context_chars < 0:
        context_chars = 500
        logger.warning(f"context_chars 为负数，重置为 {context_chars}")
    
    pdf_path = Path(pdf_path)
    try:
        pages_text = extract_text_from_pdf(
            pdf_path,
            [page_number] if page_number is not None else None,
            exclusion_boxes_by_page=exclusion_boxes_by_page,
        )
    except Exception as e:
        logger.error(f"提取 PDF 文本失败: {e}")
        return []
    
    hits = []
    keyword = keyword.strip()
    for pno, text in pages_text.items():
        if not text or keyword not in text:
            continue
        # 按行或块取包含关键词的一段
        parts = text.replace("\n", " ").split(keyword)
        for i, part in enumerate(parts):
            if i == 0:
                continue
            start = max(0, len(part) - context_chars)
            end = min(len(part) + context_chars, len(part) + len(keyword) + context_chars)
            next_part = parts[i + 1][:context_chars] if i + 1 < len(parts) else ""
            segment = (part[start:] + keyword + next_part)[: context_chars * 2]
            hits.append((pno, segment.strip()))
    
    if not hits:
        # 整页作为后备
        for pno, text in pages_text.items():
            if keyword in text:
                hits.append((pno, text[: 2000]))
                break
    
    return hits


def extract_tables_from_pdf(
    pdf_path: str | Path,
    page_number: Optional[int] = None,
    near_keyword: Optional[str] = None,
    exclusion_boxes_by_page: Optional[Dict[int, ExclusionPageData]] = None,
) -> list[list[list[str]]]:
    """
    从 PDF 中提取表格。
    
    Args:
        pdf_path: PDF 文件路径
        page_number: 只在该页取表；不指定则所有页
        near_keyword: 若指定，只保留在含该关键词的页上的表格
        
    Returns:
        [ 页1表格列表, 页2表格列表, ... ]，每页表格为 list[list[str]]（行→列）
        
    Raises:
        FileNotFoundError: PDF 文件不存在
        ValueError: PDF 文件无法打开或读取
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")
    
    if not pdf_path.is_file():
        raise ValueError(f"路径不是文件: {pdf_path}")

    all_tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                logger.warning(f"PDF 文件没有页面: {pdf_path}")
                return all_tables
                
            pages = [page_number] if page_number is not None else range(1, len(pdf.pages) + 1)
            for pno in pages:
                if pno < 1 or pno > len(pdf.pages):
                    logger.warning(f"页码 {pno} 超出范围 (1-{len(pdf.pages)})，跳过")
                    continue
                try:
                    page = pdf.pages[pno - 1]
                    if exclusion_boxes_by_page:
                        page_excl = exclusion_boxes_by_page.get(pno)
                    else:
                        page_excl = None
                    if page_excl:
                        page = _filter_page_by_exclusion(page, page_excl)

                    if near_keyword:
                        text = page.extract_text() or ""
                        if near_keyword not in text:
                            continue
                    tables = page.extract_tables()
                    if tables:
                        all_tables.append(tables)
                except Exception as e:
                    logger.error(f"提取第 {pno} 页表格时出错: {e}")
                    continue
    except Exception as e:
        raise ValueError(f"无法打开或读取 PDF 文件 {pdf_path}: {e}") from e
        
    return all_tables


def get_first_table_near_keyword(
    pdf_path: str | Path,
    keyword: str,
    page_number: Optional[int] = None,
    exclusion_boxes_by_page: Optional[Dict[int, ExclusionPageData]] = None,
) -> Optional[list[list[str]]]:
    """
    获取在关键词所在页的第一个表格；若某表内包含关键词则优先返回该表。
    
    Args:
        pdf_path: PDF 文件路径
        keyword: 关键词
        page_number: 指定页码（从1开始），None 表示搜索全部页
        
    Returns:
        第一个匹配的表格（list[list[str]]），未找到则返回 None
    """
    if not keyword or not keyword.strip():
        logger.warning("关键词为空，返回 None")
        return None
    
    try:
        tables_by_page = extract_tables_from_pdf(
            pdf_path,
            page_number=page_number,
            near_keyword=keyword,
            exclusion_boxes_by_page=exclusion_boxes_by_page,
        )
    except Exception as e:
        logger.error(f"提取表格失败: {e}")
        return None
    
    keyword = keyword.strip()
    for page_tables in tables_by_page:
        # 优先返回包含关键词的表格
        for table in page_tables:
            if not table:
                continue
            if any(any(cell and keyword in str(cell) for cell in row) for row in table):
                return table
        # 否则返回第一个非空表格
        for table in page_tables:
            if table:
                return table
    return None
