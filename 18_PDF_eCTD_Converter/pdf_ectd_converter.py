#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
eCTD 合规装甲 & XSS 深度清理器（18_PDF_eCTD_Converter）

该模块融合防止恶意脚本/XSS 的能力，并强制执行《eCTD验证标准V1.1》附件6常见要求（按可实现程度落地）：
- 可读性校验（6.1）：可打开且页数 > 0
- 剥离密码/安全设置（6.19 / 6.21）：需要密码才能打开的 PDF 直接拒收；输出统一保存为未加密
- 移除所有附件（6.17）
- 移除除超文本链接外的所有注释（6.18）
- 清理/拦截外部链接与非法协议（6.3 / 6.10 / 6.11）
- 强制设置初始视图（6.20）：UseOutlines + OneColumn（PyMuPDF 可稳定设置）
- 大于 5 页必须有书签（6.23）；可选用 ``--add-auto-bookmarks`` 自动补写书签
- 字体规范化（6.26）：将 Times-Roman / Helvetica 等 PDF 内置名映射为 eCTD 常见认可名，并 subset 嵌入所用字体
- 书签修复（6.5 / 6.6 / 6.8）：为无动作书签补全 GoTo、修正越界目标、统一承前缩放（zoom=0）
- 启用快速 Web 查看（Fast Web View / Linearization）（6.22；若当前 MuPDF 不支持线性化则自动降级保存）
- 导出合规审计 Excel 报告（便于审计与回溯）

用法：
  cd 18_PDF_eCTD_Converter
  python pdf_ectd_converter.py --input "./input" --output "./output" --report "./ectd_report.xlsx" --overwrite

  # 默认已开启 outline 自动书签；若需禁用：--no-add-auto-bookmarks

  # Windows 上若默认 python 不是安装依赖的环境，请用该解释器的完整路径运行（与 pip 安装 pymupdf 的解释器一致）。

  # 指定输入目录/单文件
  python pdf_ectd_converter.py --input "D:\\pdfs"
  python pdf_ectd_converter.py --input "D:\\pdfs\\a.pdf"

  # 仅校验（也会写审计报告）
  python pdf_ectd_converter.py --validate-only --report "output/ectd_report.xlsx"

  # 超过 5 页且无书签时自动补全书签（省略样式时默认为 outline）
  python pdf_ectd_converter.py --input "./input" --add-auto-bookmarks
  python pdf_ectd_converter.py --input "./input" --add-auto-bookmarks pages
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ModuleNotFoundError as exc:  # pragma: no cover
    raise SystemExit(
        "未找到 PyMuPDF（import fitz）。请使用与 pip 对应的解释器安装依赖，例如：\n"
        f"  {sys.executable} -m pip install pymupdf pandas openpyxl\n"
        "若在 Windows 上存在多个 Python，请用 `where python` 核对当前使用的可执行文件。"
    ) from exc

try:
    import pandas as pd
except ModuleNotFoundError as exc:  # pragma: no cover
    raise SystemExit(
        f"未找到 pandas。请运行: {sys.executable} -m pip install pandas openpyxl"
    ) from exc

try:
    import openpyxl  # noqa: F401 — pandas Excel 导出引擎
except ModuleNotFoundError as exc:  # pragma: no cover
    raise SystemExit(
        f"未找到 openpyxl（Excel 报告依赖）。请运行: {sys.executable} -m pip install openpyxl"
    ) from exc


# eCTD 与 XSS 共同封杀的恶意协议与外部前缀（保守策略：外部一律视为风险）
BAD_PROTOCOLS = ("javascript:", "data:", "vbscript:", "file:", "http://", "https://", "mailto:")
logger = logging.getLogger(__name__)


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _reset_mupdf_warnings() -> None:
    fitz.TOOLS.reset_mupdf_warnings()


def _drain_mupdf_warnings() -> str:
    return (fitz.TOOLS.mupdf_warnings() or "").strip()


def _summarize_mupdf_warnings(warnings: str, *, max_len: int = 500) -> str:
    if not warnings:
        return ""
    lines = [ln.strip() for ln in warnings.splitlines() if ln.strip()]
    xref_hits = sum(1 for ln in lines if "xref" in ln.lower())
    parts = [f"共 {len(lines)} 条"]
    if xref_hits:
        parts.append(f"xref 相关 {xref_hits} 条")
    summary = "；".join(parts)
    sample = lines[0][:120]
    if len(lines) > 1:
        sample += f" …（另 {len(lines) - 1} 条）"
    text = f"{summary}：{sample}"
    return text[:max_len]


def _open_pdf(pdf_path: Path) -> tuple[fitz.Document, str]:
    """打开 PDF 并收集 MuPDF 结构警告；必要时尝试 repair()。"""
    _reset_mupdf_warnings()
    doc = fitz.open(pdf_path)
    warnings = _drain_mupdf_warnings()
    if getattr(doc, "is_repaired", False) or ("xref" in warnings.lower()):
        try:
            doc.repair()
            warnings = _drain_mupdf_warnings() or warnings
        except Exception:
            logger.debug("PDF repair() 失败: file=%s", pdf_path.name, exc_info=True)
    return doc, warnings


def _inspect_open_pdf(doc: fitz.Document, warnings: str) -> tuple[bool, str, dict]:
    """对已打开的 PDF 做基础校验并返回 meta。"""
    meta: dict = {}
    if doc.needs_pass:
        return False, "错误 (6.21): 存在密码保护，无法打开", meta
    if doc.page_count <= 0:
        return False, "错误 (6.1): 页数为0或内容无效", meta
    toc = doc.get_toc()
    meta["page_count"] = int(doc.page_count)
    meta["has_toc"] = bool(toc)
    if warnings:
        meta["mupdf_warnings"] = warnings
    if getattr(doc, "is_repaired", False):
        meta["mupdf_repaired"] = True
    return True, "OK", meta


def _save_pdf_document(doc: fitz.Document, tmp_path: Path, *, linearize: bool) -> tuple[bool, str]:
    """保存 PDF；返回 (是否线性化, 附加提示)。"""
    save_opts = dict(
        incremental=False,
        garbage=4,
        deflate=True,
        # 进一步压缩：对图片/字体单独 deflate（对体积通常更敏感）
        deflate_images=True,
        deflate_fonts=True,
        clean=True,
        encryption=fitz.PDF_ENCRYPT_NONE,
        # 使用对象流（object streams）与更高压缩强度，通常可显著降低“重写后变大”的概率
        use_objstms=True,
        compression_effort=100,
    )
    _reset_mupdf_warnings()
    note = ""
    try:
        if linearize:
            doc.save(str(tmp_path), linear=True, **save_opts)
            return True, note
        doc.save(str(tmp_path), **save_opts)
        return False, note
    except TypeError:
        # 当前 PyMuPDF 不支持 linear 参数（或不支持某些 deflate_* 参数），降级保存
        # PyMuPDF 的 save 参数在不同版本/平台可能略有差异，这里做一次“收敛参数”降级
        fallback = dict(save_opts)
        fallback.pop("compression_effort", None)
        fallback.pop("use_objstms", None)
        doc.save(str(tmp_path), **fallback)
        if linearize:
            return False, "提示: 当前 PyMuPDF 不支持 linear 参数，已按非线性方式保存"
        return False, note
    except Exception as exc:
        if "linear" in str(exc).lower():
            doc.save(str(tmp_path), **save_opts)
            return False, "提示: 当前运行库不支持 PDF 线性化，已按非线性方式保存（6.22 不可用）"
        # clean=True 在个别损坏 PDF 上可能失败，降级重试一次
        try:
            save_opts["clean"] = False
            doc.save(str(tmp_path), **save_opts)
            return False, "提示: 标准清洗保存失败，已降级为无 clean 模式保存"
        except Exception:
            raise exc


def _check_disk_space(target_dir: Path, needed_bytes: int) -> tuple[bool, str]:
    """检查目标目录所在磁盘是否有足够可用空间（needed_bytes 为预估写入量）。"""
    try:
        usage = shutil.disk_usage(target_dir)
    except OSError as exc:
        return False, f"无法检查磁盘空间: {exc}"
    if usage.free < needed_bytes:
        free_mb = usage.free / (1024 * 1024)
        need_mb = needed_bytes / (1024 * 1024)
        return False, f"磁盘空间不足（需要约 {need_mb:.0f} MB，可用 {free_mb:.0f} MB）"
    return True, ""


def _verify_output_pdf(
    output_path: Path,
    *,
    expected_pages: int,
    require_toc: bool,
) -> tuple[bool, str]:
    if not output_path.is_file() or output_path.stat().st_size < 128:
        return False, "输出文件不存在或过小"
    try:
        doc = fitz.open(output_path)
    except Exception as exc:
        return False, f"输出 PDF 无法打开: {exc}"
    try:
        if doc.needs_pass:
            return False, "输出 PDF 仍受密码保护"
        if doc.page_count != expected_pages:
            return False, f"输出页数不一致（期望 {expected_pages}，实际 {doc.page_count}）"
        if require_toc and doc.page_count > 5 and not doc.get_toc():
            return False, "输出 PDF 仍缺少书签（规则 6.23）"
        toc_errors = _validate_toc_resolvable(doc)
        if toc_errors:
            return False, "输出 PDF 书签仍无效: " + "; ".join(toc_errors[:3])
        risky = _ectd_risky_font_names(doc)
        if risky:
            sample = ", ".join(sorted(risky)[:5])
            return False, f"输出仍含 eCTD 高风险字体名: {sample}"
        return True, "输出校验通过"
    finally:
        doc.close()


def _bookmark_title_sanitize(text: str, max_len: int = 180) -> str:
    t = " ".join((text or "").split())
    if not t:
        return "文档"
    return t[:max_len]


def _bookmark_root_title(doc: fitz.Document, pdf_path: Path) -> str:
    m = doc.metadata or {}
    raw = (m.get("title") or "").strip()
    if raw:
        return _bookmark_title_sanitize(raw)
    stem = (pdf_path.stem or "").strip()
    if stem:
        return _bookmark_title_sanitize(stem)
    return "文档"


def _toc_from_outline_heuristic(
    doc: fitz.Document,
    *,
    max_scan_pages: int = 24,
    min_font: float = 12.5,
) -> list[list]:
    """
    从正文前几页中，按「较大字号的一行文字」猜测章节标题，生成一级书签。
    仅为启发式，复杂排版可能不准。
    """
    entries: list[list] = []
    last = ""
    scan = min(max_scan_pages, doc.page_count)
    for pno in range(scan):
        page = doc[pno]
        textdict = page.get_text("dict") or {}
        for block in textdict.get("blocks") or []:
            if block.get("type") != 0:
                continue
            for line in block.get("lines") or []:
                spans = line.get("spans") or []
                if not spans:
                    continue
                text = "".join((s.get("text") or "") for s in spans).strip()
                if len(text) < 2 or len(text) > 150:
                    continue
                compact = text.replace(" ", "")
                if compact.isdigit() or set(compact) <= set("0123456789./-"):
                    continue
                sizes = [float(s.get("size") or 0) for s in spans]
                max_sz = max(sizes) if sizes else 0.0
                if max_sz < min_font:
                    continue
                title = _bookmark_title_sanitize(text)
                if title == last:
                    continue
                last = title
                entries.append([1, title, pno + 1])
    return entries


def _build_auto_toc(doc: fitz.Document, pdf_path: Path, style: str) -> list[list]:
    """在「无现有书签且页数 > 5」前提下生成待写入的 TOC（PyMuPDF: [level, title, page]，页码从 1 起）。"""
    if doc.page_count <= 5 or doc.get_toc():
        return []
    if style == "minimal":
        return [[1, _bookmark_root_title(doc, pdf_path), 1]]
    if style == "pages":
        return [[1, f"第 {i} 页", i] for i in range(1, doc.page_count + 1)]
    if style == "outline":
        guessed = _toc_from_outline_heuristic(doc)
        if len(guessed) >= 3:
            return guessed
        return [[1, f"第 {i} 页", i] for i in range(1, doc.page_count + 1)]
    raise ValueError(f"未知的自动书签样式: {style}")


# eCTD 6.26：PDF 内置字体名 -> 常见校验器认可名（仅改对象字典中的名称，不重绘正文）
FONT_XREF_RENAMES: dict[str, str] = {
    "/Times-Roman": "/TimesNewRomanPSMT",
    "/Times-Bold": "/TimesNewRomanPS-BoldMT",
    "/Times-Italic": "/TimesNewRomanPS-ItalicMT",
    "/Times-BoldItalic": "/TimesNewRomanPS-BoldItalicMT",
    "/Helvetica": "/ArialMT",
    "/Helvetica-Bold": "/Arial-BoldMT",
    "/Helvetica-Oblique": "/Arial-ItalicMT",
    "/Helvetica-BoldOblique": "/Arial-BoldItalicMT",
    "/Courier": "/CourierNewPSMT",
    "/Courier-Bold": "/CourierNewPS-BoldMT",
    "/Courier-Oblique": "/CourierNewPS-ItalicMT",
    "/Courier-BoldOblique": "/CourierNewPS-BoldItalicMT",
}

FONT_XREF_KEYS = ("BaseFont", "FontName", "Name")


def _pdf_name_token(value: str) -> str | None:
    raw = (value or "").strip()
    if not raw:
        return None
    if raw.startswith("/"):
        return raw
    if raw.startswith("(") and raw.endswith(")"):
        return f"/{raw[1:-1]}"
    return f"/{raw}"


def _normalize_pdf_font_names(doc: fitz.Document) -> tuple[int, list[str]]:
    """将 PDF 内置 Base-14 名称映射为 eCTD 常见白名单名称。"""
    renamed = 0
    notes: list[str] = []
    seen: set[str] = set()
    for xref in range(1, doc.xref_length()):
        try:
            keys = doc.xref_get_keys(xref)
        except Exception:
            continue
        for key in FONT_XREF_KEYS:
            if key not in keys:
                continue
            try:
                _typ, value = doc.xref_get_key(xref, key)
            except Exception:
                continue
            if not isinstance(value, str):
                continue
            token = _pdf_name_token(value)
            if not token:
                continue
            new_token = FONT_XREF_RENAMES.get(token)
            if not new_token:
                continue
            try:
                doc.xref_set_key(xref, key, new_token)
                renamed += 1
                pair = f"{token} -> {new_token}"
                if pair not in seen:
                    seen.add(pair)
                    notes.append(pair)
            except Exception:
                logger.debug("字体名映射失败: xref=%s key=%s", xref, key, exc_info=True)
    return renamed, notes


def _embed_fonts_subset(doc: fitz.Document) -> tuple[bool, str]:
    """嵌入文档所用字形子集（需 fonttools；失败时不中断主流程）。"""
    try:
        doc.subset_fonts()
        return True, ""
    except AttributeError:
        return False, "提示: 当前 PyMuPDF 无 subset_fonts，跳过字体嵌入"
    except Exception as exc:
        hint = "（可安装 fonttools: pip install fonttools）" if "fonttools" in str(exc).lower() else ""
        return False, f"警告: subset_fonts 未成功{hint}: {exc}"


def _ectd_risky_font_names(doc: fitz.Document) -> set[str]:
    risky_exact = {
        "timesroman",
        "timesbold",
        "timesitalic",
        "timesbolditalic",
        "helvetica",
        "helveticabold",
        "helveticaoblique",
        "helveticaboldoblique",
        "helv",
        "courier",
        "courierbold",
        "courieroblique",
        "courierboldoblique",
    }
    risky: set[str] = set()
    for pno in range(doc.page_count):
        try:
            fonts = doc[pno].get_fonts(full=True)
        except Exception:
            continue
        for item in fonts:
            if len(item) < 4:
                continue
            base = str(item[3] or "")
            norm = base.lower().replace(" ", "").replace("-", "")
            if "+" in norm:
                norm = norm.split("+", 1)[1]
            if norm in risky_exact:
                risky.add(base)
    return risky


def _make_goto_dest(page_1based: int, page_count: int) -> dict:
    """eCTD 6.5 / 6.8：显式 GoTo + 承前缩放（zoom=0）。"""
    p0 = max(0, min(page_count - 1, page_1based - 1))
    return {
        "kind": fitz.LINK_GOTO,
        "page": p0,
        "to": fitz.Point(72, 72),
        "zoom": 0.0,
    }


def _dest_has_valid_action(dest: object, page_count: int) -> bool:
    """eCTD 6.5：书签须具备 GoTo / GoToR / Launch 类动作（不接受无 dest 的容器项）。"""
    if dest is None or dest == 0:
        return False
    if not isinstance(dest, dict):
        return False
    if dest.get("collapse"):
        return False
    kind = dest.get("kind")
    if kind == fitz.LINK_GOTO:
        tp = dest.get("page")
        return isinstance(tp, int) and 0 <= tp < page_count
    if kind in (fitz.LINK_GOTOR, fitz.LINK_LAUNCH):
        return True
    return False


def _effective_toc_pages(toc: list, page_count: int) -> list[int]:
    """为 page=0 或越界的目录项推断可用页码（用于父级书签补动作）。"""
    pages: list[int] = []
    for item in toc:
        if len(item) < 3 or not isinstance(item[2], int):
            pages.append(0)
        else:
            pages.append(item[2])

    last_good = 1
    for i, p in enumerate(pages):
        if p < 1:
            pages[i] = last_good
        else:
            pages[i] = min(p, page_count)
            last_good = pages[i]

    next_good = last_good
    for i in range(len(pages) - 1, -1, -1):
        if pages[i] < 1:
            pages[i] = next_good
        else:
            next_good = pages[i]

    return [max(1, min(page_count, p)) for p in pages]


def _normalize_toc_dest(
    dest: object,
    page_1based: int,
    page_count: int,
) -> dict:
    """统一为带承前缩放的 GoTo；禁止 URI/Named 等 eCTD 6.4 不允许的类型。"""
    if _dest_has_valid_action(dest, page_count) and isinstance(dest, dict):
        kind = dest.get("kind")
        if kind == fitz.LINK_GOTO:
            d = dict(dest)
            d.pop("collapse", None)
            d.pop("xref", None)
            d["zoom"] = 0.0
            tp = d.get("page")
            if not isinstance(tp, int) or tp < 0 or tp >= page_count:
                d["page"] = max(0, min(page_count - 1, page_1based - 1))
            if d.get("to") is None:
                d["to"] = fitz.Point(72, 72)
            return d
        if kind in (fitz.LINK_GOTOR, fitz.LINK_LAUNCH):
            d = dict(dest)
            d.pop("collapse", None)
            d.pop("xref", None)
            d["zoom"] = 0.0
            return d
    return _make_goto_dest(page_1based, page_count)


def _toc_needs_flatten(toc: list, page_count: int) -> bool:
    """多级目录在 PyMuPDF 保存时易生成带 collapse 的无动作父书签（触发 6.5）。"""
    max_level = 1
    for item in toc:
        if len(item) < 1:
            continue
        max_level = max(max_level, int(item[0]))
        dest = item[3] if len(item) > 3 else None
        if isinstance(dest, dict) and dest.get("collapse"):
            return True
        if not _dest_has_valid_action(dest, page_count):
            return True
    return max_level > 1


def _flatten_toc_for_ectd(
    toc: list,
    effective_pages: list[int],
    page_count: int,
) -> list[list]:
    """扁平为一级书签，标题缩进保留层级感；每条均带 GoTo + 承前缩放。"""
    flat: list[list] = []
    for i, item in enumerate(toc):
        if len(item) < 3:
            continue
        lvl = int(item[0])
        title = str(item[1])
        page = effective_pages[i]
        indent = "  " * max(0, lvl - 1)
        flat.append(
            [
                1,
                indent + title,
                page,
                _make_goto_dest(page, page_count),
            ]
        )
    return flat


def _repair_toc(doc: fitz.Document) -> tuple[int, list[str]]:
    """重写目录：补全无动作书签(6.5)、修正越界(6.6)、承前缩放(6.8)。"""
    try:
        toc = doc.get_toc(simple=False)
    except Exception as exc:
        return 0, [f"无法读取书签: {exc}"]
    if not toc:
        return 0, []

    page_count = doc.page_count
    effective_pages = _effective_toc_pages(toc, page_count)
    repaired = 0
    notes: list[str] = []
    new_toc: list = []

    if _toc_needs_flatten(toc, page_count):
        new_toc = _flatten_toc_for_ectd(toc, effective_pages, page_count)
        repaired = len(new_toc)
        notes.append(
            f"目录已扁平化为一级书签（共 {len(new_toc)} 条），避免无动作父节点（6.5）"
        )
    else:
        for i, item in enumerate(toc):
            if len(item) < 3:
                continue
            lvl, title = item[0], str(item[1])
            page = effective_pages[i]
            dest = item[3] if len(item) > 3 else None
            title_short = (title[:48] + "…") if len(title) > 48 else title
            new_dest = _normalize_toc_dest(dest, page, page_count)

            if not _dest_has_valid_action(dest, page_count):
                repaired += 1
                notes.append(f"书签「{title_short}」无有效动作，已补全 GoTo（第 {page} 页）")
            elif isinstance(dest, dict):
                if dest.get("zoom") not in (0, 0.0):
                    repaired += 1
                    notes.append(f"书签「{title_short}」缩放已设为承前缩放")
                elif item[2] != page:
                    repaired += 1
                    notes.append(f"书签「{title_short}」页码 {item[2]} -> {page}")

            new_toc.append([lvl, title, page, new_dest])

    doc.set_toc(new_toc)

    remaining = _validate_toc_resolvable(doc)
    if remaining:
        notes.append("警告: 书签修复后仍有未通过项: " + "; ".join(remaining[:3]))

    return repaired, notes


def _validate_toc_resolvable(doc: fitz.Document) -> list[str]:
    errors: list[str] = []
    try:
        toc = doc.get_toc(simple=False)
    except Exception as exc:
        return [f"无法读取书签: {exc}"]
    page_count = doc.page_count
    for item in toc:
        if len(item) < 3:
            continue
        title = str(item[1])
        page = item[2]
        dest = item[3] if len(item) > 3 else None
        title_short = (title[:48] + "…") if len(title) > 48 else title
        if isinstance(page, int) and (page < 1 or page > page_count):
            errors.append(f"书签「{title_short}」页码 {page} 仍越界")
        if not _dest_has_valid_action(dest, page_count):
            errors.append(f"书签「{title_short}」仍无有效动作（6.5）")
            continue
        if isinstance(dest, dict) and dest.get("kind") == fitz.LINK_NAMED:
            try:
                doc.resolve_link(dest)
            except Exception as exc:
                errors.append(f"书签「{title_short}」无法解析: {exc}")
        if isinstance(dest, dict) and dest.get("zoom") not in (0, 0.0):
            errors.append(f"书签「{title_short}」未使用承前缩放（6.8）")
    return errors


def _collect_pdfs(input_path: Path, recursive: bool = True) -> list[Path]:
    input_path = input_path.expanduser()
    if not input_path.exists():
        return []
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return sorted([p.resolve() for p in input_path.glob(pattern) if p.is_file()])
    return []


def _validate_pdf_basic(pdf_path: Path) -> tuple[bool, str, dict]:
    """独立校验入口（validate-only 模式使用）。"""
    meta: dict = {}
    try:
        doc, warnings = _open_pdf(pdf_path)
    except fitz.FileDataError:
        return False, "错误 (6.1): 文件被破坏或不可读", meta
    except Exception as exc:
        logger.exception("PDF 基础校验异常: file=%s", pdf_path)
        return False, f"未知异常: {exc}", meta

    try:
        return _inspect_open_pdf(doc, warnings)
    finally:
        doc.close()


class ECTDComplianceCleaner:
    def __init__(
        self,
        input_dir: Path,
        output_dir: Path,
        report_path: Path,
        overwrite: bool,
        *,
        auto_bookmarks: str | None = None,
        subset_fonts: bool = False,
        linearize: bool = True,
    ):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.report_path = report_path
        self.overwrite = overwrite
        self.auto_bookmarks = auto_bookmarks
        self.subset_fonts = subset_fonts
        self.linearize = linearize
        self.report_rows: list[dict] = []

    def _is_illegal_link(self, link: dict) -> tuple[bool, str]:
        """
        违规链接判定（保守）：
        - 任何 URI 链接（LINK_URI）视为外部链接风险（6.3 / 6.10）
        - 协议黑名单（javascript/data/vbscript/file/http/https/mailto）
        - Named action 视为风险（常见为 JS 或内部命令）
        """
        uri = (link.get("uri") or "").strip()
        uri_lower = uri.lower()
        kind = link.get("kind")

        # 1) 协议黑名单与外部链接
        for proto in BAD_PROTOCOLS:
            if uri_lower.startswith(proto):
                return True, f"包含非法协议或外部链接前缀（{proto}）"

        # 2) URI 链接默认拒绝（即便是相对路径/片段也可能被解释为外部跳转）
        if kind == getattr(fitz, "LINK_URI", 1):
            return True, "包含外部 URI 链接（LINK_URI）"

        # 3) Named action（如 JavaScript / Named destinations 的变种）
        if kind == getattr(fitz, "LINK_NAMED", 5):
            return True, "包含 Named Action（LINK_NAMED）"

        # 4) Launch 动作风险较高（可能启动外部程序/文件）
        if kind == getattr(fitz, "LINK_LAUNCH", 4):
            return True, "包含 Launch 动作（LINK_LAUNCH）"

        # 其余：允许 LINK_GOTO（文档内跳转）与 LINK_GOTOR（跨文档跳转，是否合规取决于提交结构；这里不强杀）
        return False, ""

    def _append_report(
        self,
        filename: str,
        status: str,
        detail: str,
        *,
        input_size_bytes: int | None = None,
        output_size_bytes: int | None = None,
        page_count: int | None = None,
        has_toc: bool | None = None,
        removed_embedded: int | None = None,
        removed_annots: int | None = None,
        removed_links: int | None = None,
        has_searchable_text: bool | None = None,
        pagemode_set: bool | None = None,
        pagelayout_set: bool | None = None,
        linearized: bool | None = None,
        mupdf_warnings: str | None = None,
        output_verified: bool | None = None,
        font_renames: int | None = None,
        toc_repairs: int | None = None,
        fonts_subset: bool | None = None,
    ) -> None:
        self.report_rows.append(
            {
                "文件名": filename,
                "处理时间": _now_str(),
                "状态": status,
                "输入大小(字节)": input_size_bytes,
                "输出大小(字节)": output_size_bytes,
                "页数": page_count,
                "是否有书签": has_toc,
                "删除附件数": removed_embedded,
                "删除注释数": removed_annots,
                "删除违规链接数": removed_links,
                "是否有可搜索文本": has_searchable_text,
                "初始视图UseOutlines": pagemode_set,
                "页面布局OneColumn": pagelayout_set,
                "FastWebView(Linear)": linearized,
                "字体名映射数": font_renames,
                "书签修复数": toc_repairs,
                "字体子集嵌入": fonts_subset,
                "源PDF结构警告": mupdf_warnings or None,
                "输出校验通过": output_verified,
                "详细信息": detail,
            }
        )

    def process_pdf(self, pdf_path: Path, output_path: Path, *, validate_only: bool) -> bool:
        status = "FAILED"
        details: list[str] = []
        msg = ""
        page_count: int | None = None
        report_has_toc = False
        removed_embedded = 0
        removed_annots = 0
        removed_links = 0
        has_searchable_text = False
        pagemode_set = False
        pagelayout_set = False
        linearized = False
        output_verified: bool | None = None
        font_renames = 0
        toc_repairs = 0
        fonts_subset = False
        all_warnings = ""
        input_size_bytes: int | None = None
        output_size_bytes: int | None = None

        if not pdf_path.exists():
            logger.error("输入文件不存在: %s", pdf_path)
            self._append_report(pdf_path.name, status, "文件不存在")
            return False

        if not pdf_path.is_file() or pdf_path.stat().st_size < 64:
            logger.error("输入文件无效或过小: %s", pdf_path)
            self._append_report(pdf_path.name, status, "输入文件无效或过小")
            return False

        if output_path.exists() and not self.overwrite and not validate_only:
            logger.warning("跳过已存在输出: input=%s output=%s", pdf_path, output_path)
            self._append_report(pdf_path.name, "SKIPPED", "文件已存在且未开启覆盖")
            return False

        doc = None
        reported = False
        try:
            input_size_bytes = int(pdf_path.stat().st_size)
            try:
                doc, open_warnings = _open_pdf(pdf_path)
            except fitz.FileDataError:
                status = "FAILED"
                details.append("错误 (6.1): 文件被破坏或不可读")
                return False
            except Exception as exc:
                status = "FAILED"
                details.append(f"无法打开 PDF: {exc}")
                logger.exception("打开 PDF 失败: file=%s", pdf_path)
                return False

            ok, msg, meta = _inspect_open_pdf(doc, open_warnings)
            page_count = meta.get("page_count")
            has_toc = meta.get("has_toc")
            mupdf_warnings = meta.get("mupdf_warnings", "")
            all_warnings = mupdf_warnings
            if meta.get("mupdf_repaired"):
                logger.warning("源 PDF 存在结构问题，MuPDF 已尝试修复: %s", pdf_path.name)
            if mupdf_warnings:
                logger.warning(
                    "源 PDF 结构警告: %s => %s",
                    pdf_path.name,
                    _summarize_mupdf_warnings(mupdf_warnings),
                )
            if not ok:
                status = "FAILED"
                details.append(msg)
                return False

            allow_autofill = self.auto_bookmarks is not None
            if isinstance(page_count, int) and page_count > 5 and not has_toc and not allow_autofill:
                status = "FAILED"
                details.append("错误: 大于5页的文件缺少书签（规则 6.23）")
                return False

            if validate_only:
                status = "SUCCESS"
                detail = "校验通过（validate-only）"
                if mupdf_warnings:
                    detail += f"；源 PDF 结构警告: {_summarize_mupdf_warnings(mupdf_warnings)}"
                if allow_autofill and isinstance(page_count, int) and page_count > 5 and not has_toc:
                    detail += f"；已启用 --add-auto-bookmarks={self.auto_bookmarks}，转换时将写入书签"
                details.append(detail)
                reported = True
                self._append_report(
                    pdf_path.name,
                    status,
                    detail,
                    page_count=page_count,
                    has_toc=has_toc,
                    mupdf_warnings=_summarize_mupdf_warnings(mupdf_warnings) or None,
                )
                return True

            report_has_toc = bool(has_toc)

            if doc.get_toc():
                repaired, repair_notes = _repair_toc(doc)
                toc_repairs += repaired
                for note in repair_notes[:8]:
                    details.append(f"书签修复: {note}")
                if len(repair_notes) > 8:
                    details.append(f"书签修复: …另有 {len(repair_notes) - 8} 条")
                report_has_toc = bool(doc.get_toc())

            if self.auto_bookmarks and doc.page_count > 5 and not doc.get_toc():
                try:
                    new_toc = _build_auto_toc(doc, pdf_path, self.auto_bookmarks)
                    if new_toc:
                        doc.set_toc(new_toc)
                        report_has_toc = True
                        details.append(
                            f"已自动补全书签（模式 {self.auto_bookmarks}，共 {len(new_toc)} 条，规则 6.23）"
                        )
                except Exception:
                    logger.exception("自动补全书签失败: file=%s", pdf_path)
                if not doc.get_toc():
                    raise RuntimeError("已启用自动书签但仍无书签，不满足规则 6.23")

            # 6.17: 删除嵌入文件（附件）
            emb_count = int(doc.embfile_count())
            if emb_count > 0:
                for i in range(emb_count - 1, -1, -1):
                    doc.embfile_del(i)
                removed_embedded = emb_count
                details.append(f"清理了 {emb_count} 个附件（6.17）")

            # 6.18 / 6.10 / 6.11: 删除非链接注释 + 删除违规链接
            for page in doc:
                # 6.18: 移除除超文本链接外的所有注释
                annot = page.first_annot
                while annot:
                    next_annot = annot.next
                    try:
                        if annot.type[0] != fitz.PDF_ANNOT_LINK:
                            page.delete_annot(annot)
                            removed_annots += 1
                    except Exception:
                        # 注释损坏时，尽量删除
                        try:
                            page.delete_annot(annot)
                            removed_annots += 1
                        except Exception:
                            logger.debug("删除注释失败: file=%s page=%s", pdf_path.name, page.number + 1, exc_info=True)
                    annot = next_annot

                # 删除违规链接（先收集再删）
                bad_links: list[dict] = []
                for link in page.get_links():
                    illegal, reason = self._is_illegal_link(link)
                    if illegal:
                        bad_links.append(link)
                        if reason and reason not in details:
                            details.append(f"链接清理: {reason}")
                for link in bad_links:
                    try:
                        page.delete_link(link)
                        removed_links += 1
                    except Exception:
                        logger.debug("删除链接失败: file=%s page=%s", pdf_path.name, page.number + 1, exc_info=True)

            # 6.25: 可搜索文本检查（预警）
            has_searchable_text = any((p.get_text() or "").strip() for p in doc)
            if not has_searchable_text:
                details.append("警告: 未检测到可搜索文本，建议执行 OCR（6.25）")

            # 6.20: 初始视图（稳定可用）
            try:
                doc.set_pagemode("UseOutlines")
                pagemode_set = True
            except Exception:
                details.append("警告: 无法设置初始视图 UseOutlines（6.20）")

            try:
                doc.set_pagelayout("OneColumn")
                pagelayout_set = True
            except Exception:
                details.append("警告: 无法设置页面布局 OneColumn（6.20）")

            rename_count, rename_notes = _normalize_pdf_font_names(doc)
            font_renames += rename_count
            if rename_notes:
                details.append(
                    f"字体名映射（6.26）: {rename_count} 处；"
                    + "；".join(rename_notes[:4])
                    + (f" …共 {len(rename_notes)} 种" if len(rename_notes) > 4 else "")
                )

            if self.subset_fonts:
                subset_ok, subset_note = _embed_fonts_subset(doc)
                fonts_subset = subset_ok
                if subset_note:
                    details.append(subset_note)
            else:
                # 默认不做字体子集嵌入：该步骤在部分 PDF 上会显著增大体积
                fonts_subset = False

            risky_before_save = _ectd_risky_font_names(doc)
            if risky_before_save:
                sample = ", ".join(sorted(risky_before_save)[:5])
                details.append(
                    f"警告: 仍检测到 eCTD 高风险字体名（可能需人工换字体或重排版）: {sample}"
                )

            if doc.get_toc():
                toc_repairs += _repair_toc(doc)[0]

            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 预估输出约为源文件 1.2 倍（压缩后通常更小，留余量防磁盘满）
            src_bytes = pdf_path.stat().st_size
            ok_space, space_msg = _check_disk_space(output_path.parent, int(src_bytes * 1.2) + 50 * 1024 * 1024)
            if not ok_space:
                raise RuntimeError(space_msg)

            # 原子写入：先写入同目录临时文件再替换，避免保存中断留下半截 PDF
            fd, tmp_name = tempfile.mkstemp(
                suffix=".pdf",
                dir=str(output_path.parent),
                prefix=".ectd_tmp_",
            )
            os.close(fd)
            tmp_path = Path(tmp_name)
            saved_ok = False
            try:
                linearized, save_note = _save_pdf_document(doc, tmp_path, linearize=self.linearize)
                if save_note:
                    details.append(save_note)
                save_warnings = _drain_mupdf_warnings()
                if save_warnings:
                    all_warnings = f"{all_warnings}\n{save_warnings}".strip() if all_warnings else save_warnings
                os.replace(tmp_path, output_path)
                saved_ok = True
            except OSError as ose:
                raise RuntimeError(f"无法写入输出文件（磁盘空间/权限/路径问题）: {ose}") from ose
            finally:
                if not saved_ok and tmp_path.exists():
                    try:
                        tmp_path.unlink()
                    except OSError:
                        pass

            try:
                output_size_bytes = int(output_path.stat().st_size)
            except OSError:
                output_size_bytes = None

            expected_pages = int(doc.page_count)
            verified, verify_msg = _verify_output_pdf(
                output_path,
                expected_pages=expected_pages,
                require_toc=expected_pages > 5,
            )
            output_verified = verified
            if not verified:
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                raise RuntimeError(verify_msg)
            if all_warnings:
                details.append(
                    f"警告: 源/保存阶段 PDF 结构告警（已重写输出）: {_summarize_mupdf_warnings(all_warnings)}"
                )
            status = "SUCCESS"
            details.insert(0, "清洗合规通过")
            return True
        except Exception as exc:
            logger.exception("eCTD 处理失败: file=%s output=%s", pdf_path, output_path)
            details.append(f"处理终止: {exc}")
            return False
        finally:
            if doc is not None:
                doc.close()
            if not reported:
                self._append_report(
                    pdf_path.name,
                    status,
                    " | ".join(details) if details else msg,
                    input_size_bytes=input_size_bytes,
                    output_size_bytes=output_size_bytes,
                    page_count=page_count if isinstance(page_count, int) else None,
                    has_toc=report_has_toc if isinstance(report_has_toc, bool) else None,
                    removed_embedded=removed_embedded,
                    removed_annots=removed_annots,
                    removed_links=removed_links,
                    has_searchable_text=has_searchable_text,
                    pagemode_set=pagemode_set,
                    pagelayout_set=pagelayout_set,
                    linearized=linearized,
                    mupdf_warnings=_summarize_mupdf_warnings(all_warnings) or None,
                    output_verified=output_verified,
                    font_renames=font_renames or None,
                    toc_repairs=toc_repairs or None,
                    fonts_subset=fonts_subset if fonts_subset else None,
                )

    def export_report(self) -> None:
        if not self.report_rows:
            logger.warning("无审计数据可导出。")
            return
        df = pd.DataFrame(self.report_rows)
        self.report_path.parent.mkdir(parents=True, exist_ok=True)
        warn_df = df[df["源PDF结构警告"].notna()][["文件名", "状态", "源PDF结构警告", "输出校验通过"]]
        try:
            with pd.ExcelWriter(self.report_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="全部", index=False)
                if not warn_df.empty:
                    warn_df.to_excel(writer, sheet_name="结构警告", index=False)
        except (PermissionError, OSError):
            alt = self.report_path.with_name(f"{self.report_path.stem}_{_now_str().replace(':', '-')}{self.report_path.suffix}")
            logger.warning("报告目标无法写入，改用备用路径: %s", alt)
            with pd.ExcelWriter(alt, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="全部", index=False)
                if not warn_df.empty:
                    warn_df.to_excel(writer, sheet_name="结构警告", index=False)
            self.report_path = alt
        logger.info("eCTD 审计报告已生成: %s", self.report_path)


def main() -> None:
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
            except (AttributeError, OSError, ValueError):
                pass
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    # 结构警告改由 mupdf_warnings() 采集并写入 Excel，避免 C 层 stderr 刷屏
    fitz.TOOLS.mupdf_display_errors(False)
    fitz.TOOLS.mupdf_display_warnings(False)
    # 兼容源码运行与 PyInstaller 单文件/单目录运行：
    # - 源码运行：以脚本所在目录为根目录
    # - 冻结运行：以 exe 所在目录为根目录（便于复制到任意路径/电脑直接用）
    base_dir = (
        Path(sys.executable).resolve().parent
        if getattr(sys, "frozen", False)
        else Path(__file__).resolve().parent
    )
    default_input = base_dir / "input"
    default_output = base_dir / "output"
    default_report = base_dir / "ectd_report.xlsx"

    parser = argparse.ArgumentParser(description="eCTD 合规装甲 & 批量 PDF XSS 深度清理器")
    parser.add_argument("--input", "-i", default=str(default_input), help="输入 PDF 文件夹或单个 PDF 文件")
    parser.add_argument("--output", "-o", default=str(default_output), help="输出文件夹（默认 output/）")
    parser.add_argument("--report", default=str(default_report), help="Excel 审计报告输出路径")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    parser.add_argument("--keep-name", action="store_true", help="输出文件名与源文件一致（默认追加 _ectd）")
    parser.add_argument("--validate-only", action="store_true", help="仅做校验，不做输出（仍生成审计报告）")
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
    parser.add_argument(
        "--add-auto-bookmarks",
        nargs="?",
        const="outline",
        default="outline",
        choices=["minimal", "pages", "outline"],
        metavar="STYLE",
        help="超过 5 页且无书签时自动写入书签（默认 outline）：minimal=单条根书签；pages=每页一条；"
        "outline=尝试按大号字体提取标题（不足 3 条则退化为每页一条）",
    )
    parser.add_argument(
        "--no-add-auto-bookmarks",
        action="store_const",
        const=None,
        dest="add_auto_bookmarks",
        help="禁用自动补全书签（严格按 6.23 拒收无书签文件）",
    )
    parser.add_argument(
        "--subset-fonts",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="是否尝试嵌入所用字体子集（可能增大文件体积；默认关闭）",
    )
    parser.add_argument(
        "--linearize",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="是否尝试线性化（Fast Web View / Linearization；默认开启）",
    )
    args = parser.parse_args()

    # 自动创建输入/输出目录，保证 exe 复制到新位置后可即开即用
    default_input.mkdir(parents=True, exist_ok=True)
    default_output.mkdir(parents=True, exist_ok=True)

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output).expanduser()
    if not output_dir.is_absolute():
        output_dir = (base_dir / output_dir).resolve()
    else:
        output_dir = output_dir.resolve()
    report_path = Path(args.report).expanduser()
    if not report_path.is_absolute():
        report_path = (base_dir / report_path).resolve()
    else:
        report_path = report_path.resolve()

    pdf_files = _collect_pdfs(input_path, recursive=args.recursive)
    if not pdf_files:
        if input_path.is_dir() and input_path == default_input:
            logger.error("未找到 PDF: %s（目录为空，请先放入待处理 PDF 后重试）", input_path)
            raise SystemExit(1)
        hint = "路径不存在" if not input_path.exists() else "目录下无 .pdf 文件"
        logger.error("未找到 PDF: %s（%s）", input_path, hint)
        raise SystemExit(1)

    logger.info("启动 eCTD 处理: files=%s", len(pdf_files))

    cleaner = ECTDComplianceCleaner(
        input_dir=input_path,
        output_dir=output_dir,
        report_path=report_path,
        overwrite=args.overwrite,
        auto_bookmarks=args.add_auto_bookmarks,
        subset_fonts=bool(args.subset_fonts),
        linearize=bool(args.linearize),
    )

    total = 0
    success = 0
    base_input_dir = input_path.resolve() if input_path.is_dir() else None
    file_total = len(pdf_files)

    for pdf_path in pdf_files:
        total += 1
        size_mb = pdf_path.stat().st_size / (1024 * 1024)
        if size_mb >= 50:
            logger.info("[%d/%d] 大文件: %s (%.1f MB)", total, file_total, pdf_path.name, size_mb)
        else:
            logger.info("[%d/%d] 处理: %s", total, file_total, pdf_path.name)
        if base_input_dir and args.keep_structure:
            try:
                rel = pdf_path.relative_to(base_input_dir)
            except ValueError:
                logger.warning(
                    "无法将文件路径相对于输入根目录解析，将扁平输出: file=%s root=%s",
                    pdf_path,
                    base_input_dir,
                )
                rel = Path(pdf_path.name)
            if args.keep_name:
                out_path = output_dir / rel
            else:
                out_path = (output_dir / rel).with_name(f"{rel.stem}_ectd.pdf")
        else:
            if args.keep_name:
                out_path = output_dir / pdf_path.name
            else:
                out_path = output_dir / f"{pdf_path.stem}_ectd.pdf"

        if cleaner.process_pdf(pdf_path, out_path, validate_only=args.validate_only):
            success += 1
            if args.validate_only:
                logger.info("校验通过: %s", pdf_path.name)
            else:
                logger.info("处理成功: %s => %s", pdf_path.name, out_path.name)
        else:
            logger.error("处理失败: %s", pdf_path.name)

    cleaner.export_report()
    warn_count = sum(1 for r in cleaner.report_rows if r.get("源PDF结构警告"))
    logger.info(
        "执行完毕: success=%s total=%s structure_warnings=%s",
        success,
        total,
        warn_count,
    )
    raise SystemExit(0 if success == total else 1)


if __name__ == "__main__":
    main()

