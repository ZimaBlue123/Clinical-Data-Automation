"""docx 段落 + 表格抽取（纯标准库 zipfile + xml.etree，不依赖 python-docx）。

设计点
-------
- docx 本质是 ZIP，word/document.xml 是正文段落 + 表格
- 提供三个公共入口：
  - extract_paragraphs(docx_path)        -> list[str]
  - extract_tables(docx_path)            -> list[list[list[str]]]
  - write_paragraphs_to_txt(docx_path)   -> Path
- 段落内 w:t 拼接；表格内 w:tc/w:tr/w:tbl 严格按出现顺序解析
- 鲁棒性：容错 zipfile.BadZipFile、xml.etree.ParseError、KeyError(missing document.xml)
"""

from __future__ import annotations

import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

if sys.version_info >= (3, 11):
    from typing import Self  # noqa: F401
else:
    from typing_extensions import Self  # type: ignore[assignment]  # noqa: F401

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass(frozen=True)
class DocxContent:
    """docx 解析结果。

    Attributes:
        paragraphs: 段落纯文本列表（含空段；与 docx w:p 1:1 对应）
        tables: 顶层表格的纯文本矩阵列表（不含嵌套表）
        table_titles: 与 tables 一一对应的"表号"标题
    """

    paragraphs: list[str] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)
    table_titles: list[str] = field(default_factory=list)

    def find_paragraph(self, keyword: str) -> str:
        """返回包含 keyword 的第一个段落（跳过空段）。找不到抛 KeyError。"""
        for p in self.paragraphs:
            if keyword in p:
                return p
        raise KeyError(f"未找到含 {keyword!r} 的段落")

    def find_paragraphs(self, keyword: str) -> list[str]:
        """返回所有包含 keyword 的段落（跳过空段）。"""
        return [p for p in self.paragraphs if keyword in p]

    def non_empty_count(self) -> int:
        """返回非空段落数（与 _docx_paragraphs.txt 里的 P 编号语义接近，但 P 是 1-based 全编号）"""
        return sum(1 for p in self.paragraphs if p.strip())


def _cell_text(tc: ET.Element) -> str:
    parts: list[str] = []
    for p in tc.iter(W + "p"):
        ts = [t.text or "" for t in p.iter(W + "t")]
        parts.append("".join(ts))
    return "\n".join(parts).strip()


def _table_to_rows(tbl: ET.Element) -> list[list[str]]:
    return [[_cell_text(tc) for tc in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]


def parse_docx(docx_path: Path | str) -> DocxContent:
    """解析 docx 文件，返回段落（含空段占位）+ 顶层表。

    段落列表与 docx 中 w:p 严格 1:1 对应（空段保留为空字符串 ""）。
    这样 antibody_mapping 等下游模块可以用 paragraphs[i] + paragraphs[i-1] 的
    索引关系做"上一行判定"，不会因为空段被压缩而错位。

    嵌套表会被识别但不会出现在结果中（按用户场景：临床阶段小结无嵌套表需求）。
    """
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"docx 不存在: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"仅支持 .docx，不支持 {path.suffix}: {path}")
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                raise KeyError(f"{path.name} 不含 word/document.xml，不是合法 docx")
            doc_xml = zf.read("word/document.xml")
    except zipfile.BadZipFile as e:
        raise ValueError(f"{path.name} 不是合法 zip/docx: {e}") from e

    try:
        root = ET.fromstring(doc_xml)
    except ET.ParseError as e:
        raise ValueError(f"{path.name} 的 document.xml 解析失败: {e}") from e

    body = root.find(W + "body")
    if body is None:
        return DocxContent()

    paragraphs: list[str] = []
    tables: list[list[list[str]]] = []
    table_titles: list[str] = []
    last_title: str = ""

    for child in body:
        tag = child.tag
        if tag == W + "p":
            texts = [t.text or "" for t in child.iter(W + "t")]
            txt = "".join(texts).strip()
            # 注意：保留空段占位（用空字符串），保证索引与 docx 一致
            paragraphs.append(txt)
            if txt:
                # 标题判定：表号型行（"表 2-1 受试者分布" / "表 31 ..."）
                if txt.startswith("表 ") and len(txt) < 60:
                    last_title = txt
        elif tag == W + "tbl":
            tables.append(_table_to_rows(child))
            table_titles.append(last_title)
            last_title = ""

    return DocxContent(paragraphs=paragraphs, tables=tables, table_titles=table_titles)


def extract_paragraphs(docx_path: Path | str) -> list[str]:
    """返回非空段落列表（剔除空段）。"""
    return [p for p in parse_docx(docx_path).paragraphs if p.strip()]


def extract_tables(docx_path: Path | str) -> list[list[list[str]]]:
    return parse_docx(docx_path).tables


def write_paragraphs_to_txt(
    docx_path: Path | str,
    out_path: Path | str,
    *,
    encoding: str = "utf-8",
) -> Path:
    """把段落写入 txt，每行格式 ``[P{idx}] {text}``（P 为 1-based，与 docx w:p 序号对齐）。

    空段也会写出（行首为空），方便外部脚本做行号对照。
    """
    p = Path(docx_path)
    o = Path(out_path)
    content = parse_docx(p)
    with o.open("w", encoding=encoding) as fh:
        for i, txt in enumerate(content.paragraphs, start=1):
            fh.write(f"[P{i}] {txt}\n")
    return o
