"""GraphPad Prism pzfx 文件解析（只读，纯标准库）。

pzfx 格式
---------
- 早期版本：单一 XML 文件（Prism 4/5）
- 较新版本：ZIP 容器，内含 ``DataTables/*.xml`` 等
- 命名空间：``http://graphpad.com/prism/Prism.htm``（Prism 6+ 用 ``Prism.xsd``）

本模块覆盖这两种格式，统一对外返回 ``PzfxFile`` 数据类。
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

# 两种历史命名空间
NS_OLD = "{http://graphpad.com/prism/Prism.htm}"
NS_NEW = "{http://www.graphpad.com/prism/Prism.xsd}"
NS = (NS_OLD, NS_NEW)

# 标题 → (年龄段, 指标) 解析
_TITLE_RE = re.compile(r"^(40-49岁|50-59岁|60岁以上|50岁以上)(GMC|GMI|阳转率)$")


@dataclass(frozen=True)
class PzfxTable:
    """pzfx 单表。

    Attributes:
        table_id: 内部 ID（如 ``Table2``）
        title:    ``<Title>`` 内容（如 ``"40-49岁GMC"``）
        age_band: ``"40-49岁"`` 等
        metric:   ``"GMC"`` / ``"GMI"`` / ``"阳转率"``
        row_titles: 3 个时间点标签（``["免前", "一免后2个月", "全免后1个月"]``）
        columns:  列表，每项 = ``{title, subcolumns, decimals}``
                  其中 subcolumns 是 ``list[list[str]]``：3 个亚列 × 3 个时间点的 d 文本
    """

    table_id: str
    title: str
    age_band: str
    metric: str
    row_titles: list[str]
    columns: list[dict]  # type: ignore[type-arg]

    def get(self, ycol_idx: int, sub_idx: int) -> list[str]:
        """便捷：取第 ycol_idx 列第 sub_idx 个亚列的 3 个 d。"""
        return self.columns[ycol_idx]["subcolumns"][sub_idx]


@dataclass(frozen=True)
class PzfxFile:
    """pzfx 解析结果。"""

    path: Path
    is_zip: bool
    tables: list[PzfxTable] = field(default_factory=list)
    raw_xml: str = ""

    def get_table(self, age_band: str, metric: str) -> PzfxTable | None:
        for t in self.tables:
            if t.age_band == age_band and t.metric == metric:
                return t
        return None


def _parse_text_xml(xml_bytes: bytes) -> tuple[ET.Element, str]:
    """解析单文件 XML 格式的 pzfx。"""
    try:
        text = xml_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = xml_bytes.decode("utf-8", errors="replace")
    root = ET.fromstring(text)
    return root, text


def _read_zip_xml_entries(pzfx_bytes: bytes) -> list[tuple[str, bytes]]:
    entries: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(io.BytesIO(pzfx_bytes), "r") as zf:
        for n in zf.namelist():
            if n.endswith(".xml"):
                entries.append((n, zf.read(n)))
    return entries


def _extract_tables_from_root(root: ET.Element) -> list[ET.Element]:
    """兼容新旧命名空间，取所有顶层 <Table>。"""
    for ns in NS:
        tbls = root.findall(f"{ns}Table")
        if tbls:
            return tbls
    return []


def _text(el: ET.Element | None, tag: str) -> str:
    if el is None:
        return ""
    return (el.text or "").strip()


def _parse_one_table(tbl: ET.Element) -> PzfxTable | None:
    """解析单个 <Table> 节点为 PzfxTable。"""
    # 找 ID / Title（兼容命名空间）
    tid = tbl.get("ID", "")
    title = ""
    for ns in NS:
        t = tbl.find(f"{ns}Title")
        if t is not None and t.text:
            title = t.text.strip()
            break

    m = _TITLE_RE.match(title)
    if not m:
        return None
    age_band, metric = m.group(1), m.group(2)

    # 行标题
    row_titles: list[str] = []
    for ns in NS:
        rtc = tbl.find(f"{ns}RowTitlesColumn")
        if rtc is not None:
            subs = rtc.findall(f"{ns}Subcolumn")
            for sc in subs:
                ds = sc.findall(f"{ns}d")
                row_titles = [(d.text or "").strip() for d in ds]
            break

    # YColumn
    columns: list[dict] = []  # type: ignore[type-arg]
    for ns in NS:
        ycs = tbl.findall(f"{ns}YColumn")
        if ycs:
            for yc in ycs:
                t_el = yc.find(f"{ns}Title")
                col_title = _text(t_el, "Title")
                sub_list: list[list[str]] = []
                for sc in yc.findall(f"{ns}Subcolumn"):
                    ds = [(d.text or "").strip() for d in sc.findall(f"{ns}d")]
                    sub_list.append(ds)
                columns.append(
                    {
                        "title": col_title,
                        "subcolumns": sub_list,
                        "decimals": yc.get("Decimals"),
                    }
                )
            break

    return PzfxTable(
        table_id=tid,
        title=title,
        age_band=age_band,
        metric=metric,
        row_titles=row_titles,
        columns=columns,
    )


def parse_pzfx(pzfx_path: Path | str) -> PzfxFile:
    """读取 pzfx 文件，返回 PzfxFile。"""
    path = Path(pzfx_path)
    if not path.exists():
        raise FileNotFoundError(f"pzfx 不存在: {path}")
    raw = path.read_bytes()

    # 是否 ZIP 容器
    is_zip = raw[:4] == b"PK\x03\x04"

    if is_zip:
        entries = _read_zip_xml_entries(raw)
        all_tables: list[ET.Element] = []
        for _n, data in entries:
            try:
                sub_root = ET.fromstring(data)
            except ET.ParseError:
                continue
            all_tables.extend(_extract_tables_from_root(sub_root))
        return PzfxFile(
            path=path,
            is_zip=True,
            tables=[t for t in (_parse_one_table(x) for x in all_tables) if t],
        )

    # 单文件 XML
    try:
        root, text = _parse_text_xml(raw)
    except ET.ParseError as e:
        raise ValueError(f"{path.name} 不是合法 XML: {e}") from e

    all_tables = _extract_tables_from_root(root)
    tables = [t for t in (_parse_one_table(x) for x in all_tables) if t]
    return PzfxFile(path=path, is_zip=False, tables=tables, raw_xml=text)


def is_pzfx(pzfx_path: Path | str) -> bool:
    """快速嗅探：是否像 pzfx。"""
    p = Path(pzfx_path)
    if not p.exists() or p.suffix.lower() != ".pzfx":
        return False
    raw = p.read_bytes()[:4]
    return raw[:4] == b"PK\x03\x04" or raw[:5] == b"<?xml"
