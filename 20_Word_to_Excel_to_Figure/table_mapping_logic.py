# -*- coding: utf-8 -*-
"""
固化的“Excel 子表 -> Word 表格”映射逻辑。

用途：
- 先根据 Template 骨架与 input Word/RTF，生成候选映射 plan（供人工确认）
- 再读取确认后的映射 JSON，在正式生成 output 时限定抓取来源
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional


@dataclass(frozen=True)
class CellRange:
    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int

    @property
    def nrows(self) -> int:
        return self.max_row - self.min_row + 1


@dataclass(frozen=True)
class ChartSeriesRefs:
    cat: CellRange
    val: CellRange


@dataclass(frozen=True)
class WordTableRef:
    """指定 Word 文档中的某一个表（doc.Content.Tables 的 1-based 索引）。"""

    word_path: str
    table_index: int


def strip_word_cell_text(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    s = s.replace("\x07", "").replace("\r", "").replace("\x0b", "")
    s = s.strip()
    s = re.sub(r"\s+", "", s)
    if s.startswith("'"):
        s = s.lstrip("'")
    return s


def label_tokens(s: str) -> list[str]:
    s_norm = strip_word_cell_text(s).upper()
    return re.findall(r"[A-Z0-9]+", s_norm)


def labels_match(cell_text: Any, target_text: Any) -> bool:
    tgt = strip_word_cell_text(target_text)
    if not tgt:
        return False
    cell = strip_word_cell_text(cell_text)
    if cell == tgt:
        return True
    if tgt in cell or cell in tgt:
        return True
    cell_up = cell.upper()
    toks = label_tokens(tgt)
    return bool(toks) and all(t in cell_up for t in toks)


def subtable_id_from_cell_range(r: CellRange) -> str:
    return f"{r.sheet}|{r.min_row}|{r.max_row}|{r.min_col}|{r.max_col}"


def load_table_mapping_json(table_map_json_path: str) -> dict[str, list[WordTableRef]]:
    """
    从 JSON 读取你确认后的 table 映射。
    规则：
    - 每个 subtable 列表里，如果存在 `selected: true`，则只选这些
    - 否则默认选列表第一个元素
    """
    p = Path(table_map_json_path).expanduser().resolve()
    if not p.is_file():
        raise FileNotFoundError(f"table_map_json 不存在: {p}")
    data = json.loads(p.read_text(encoding="utf-8"))
    subtables = data.get("subtables", {}) or {}
    mapping: dict[str, list[WordTableRef]] = {}
    for sid, cand_list in subtables.items():
        if not cand_list:
            continue
        selected = [c for c in cand_list if bool(c.get("selected"))]
        if not selected:
            selected = cand_list[:1]
        mapping[sid] = [
            WordTableRef(word_path=str(c.get("word_file")), table_index=int(c.get("table_index")))
            for c in selected
        ]
    return mapping


def build_table_mapping_plan(
    *,
    template_path: str,
    word_parts: list[str],
    series_refs: list[ChartSeriesRefs],
    word_open_runner: Any,
    top_k_per_subtable: int = 3,
) -> dict[str, Any]:
    """
    生成候选映射 plan。

    说明：
    - word_open_runner 由调用端提供（用于打开 Word 并枚举 tables）
    - 具体读 Excel/解析 chart series 的逻辑在主程序中完成，这里只负责“打分与组织 plan JSON”
    """
    # 以 cat_range 几何聚类为 subtable
    subtables: dict[str, list[ChartSeriesRefs]] = {}
    for sr in series_refs:
        sid = subtable_id_from_cell_range(sr.cat)
        subtables.setdefault(sid, []).append(sr)

    # 从 template 侧抽取每个 subtable 的 cat label 列表（调用端若需要更强控制，可改为传入）
    # 这里保持为“空”，调用端通常会在 plan JSON 中补充 cat_samples。
    subtable_cat_samples: dict[str, list[str]] = {sid: [] for sid in subtables.keys()}

    # 扫描每个 Word part 的所有 tables 文本
    table_infos: dict[str, list[dict[str, Any]]] = {}
    for wp in word_parts:
        doc = word_open_runner(wp)
        infos: list[dict[str, Any]] = []
        tables = doc.Content.Tables
        for ti in range(1, tables.Count + 1):
            t = tables.Item(ti)
            try:
                raw = str(t.Range.Text)
            except Exception:
                raw = ""
            snippet = raw.replace("\r", " ").replace("\n", " ")
            snippet = re.sub(r"\s+", " ", snippet).strip()
            snippet = snippet[:140] if len(snippet) > 140 else snippet
            infos.append({"table_index": ti, "snippet": snippet, "table_text": strip_word_cell_text(raw)})
        table_infos[str(wp)] = infos

    # 候选打分：这里仅提供“占位实现”，实际建议用主程序的 cat label 参与打分
    subtables_out: dict[str, list[dict[str, Any]]] = {}
    for sid in subtables.keys():
        candidates: list[tuple[int, str, int, str]] = []
        for wp_str, infos in table_infos.items():
            for info in infos:
                # 占位：无 cat_samples 时 score=0，便于你后续按实际需求替换打分逻辑
                candidates.append((0, wp_str, int(info["table_index"]), str(info.get("snippet", ""))))
        candidates.sort(key=lambda x: (-x[0], x[2]))
        top = candidates[:top_k_per_subtable]
        subtables_out[sid] = [
            {"word_file": wp_str, "table_index": ti, "score": score, "snippet": snippet}
            for score, wp_str, ti, snippet in top
        ]

    return {
        "template_path": str(Path(template_path)),
        "subtables": subtables_out,
        "subtable_cat_samples": subtable_cat_samples,
        "top_k_per_subtable": top_k_per_subtable,
        "note": "请在每个 subtable 候选中标记 selected=true，或保留 top1 作为默认。",
    }

