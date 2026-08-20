"""
OOXML 文本替换核心（lib/）：跨 w:t / w:delText、表格单元格、页眉页脚等。
由 replace_docx.py 调用，勿直接运行。
"""

from __future__ import annotations

import glob
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from collections.abc import Sequence

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
WT = f"{{{W_NS}}}t"
WDEL = f"{{{W_NS}}}delText"
WTC = f"{{{W_NS}}}tc"
WP = f"{{{W_NS}}}p"

TEXT_TAGS = {WT, WDEL}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


@dataclass
class ReplaceStats:
    """单次 docx 处理统计。"""

    before: int = 0
    after: int = 0
    detail: dict = field(default_factory=dict)

    @property
    def replaced(self) -> int:
        return self.before - self.after


@dataclass
class ReplaceRuleSet:
    """替换规则集合。"""

    literals: list[tuple[str, str]] = field(default_factory=list)
    regexes: list[tuple[re.Pattern[str], str]] = field(default_factory=list)

    def count_in_text(self, text: str) -> int:
        n = 0
        for old, _ in self.literals:
            n += text.count(old)
        for rx, _ in self.regexes:
            n += sum(1 for _ in rx.finditer(text))
        return n

    def apply(self, text: str) -> tuple[str, int]:
        total = 0
        out = text
        for old, new in self.literals:
            c = out.count(old)
            if c:
                out = out.replace(old, new)
                total += c
        for rx, new in self.regexes:
            out, c = rx.subn(new, out)
            total += c
        return out, total


def build_date_rules() -> ReplaceRuleSet:
    x1 = r"[XxＸｘ]"
    x_pair = rf"{x1}\s*{x1}"
    year = r"(?:2026|２０２６)"
    re_slash = re.compile(rf"{year}\s*(?:/|／)?\s*{x_pair}\s*(?:/|／)?\s*{x_pair}\s*日?")
    re_cn = re.compile(rf"{year}\s*年\s*{x_pair}\s*月\s*{x_pair}\s*日")
    return ReplaceRuleSet(
        literals=[],
        regexes=[
            (re_cn, "2026年05月27日"),
            (re_slash, "2026/05/27"),
        ],
    )


def build_study_id_rules() -> ReplaceRuleSet:
    new_full = "YDSWX（TVAX-009）-004（III）"
    # 长模式优先；含「Ⅳ) / IV)」前缀的一并去掉，避免出现 I) 残留
    literals = [
        ("Ⅳ) YDSWX（TVAX-009）-004（Ⅳ）", new_full),
        ("IV) YDSWX（TVAX-009）-004（IV）", new_full),
        ("Ⅳ)YDSWX（TVAX-009）-004（Ⅳ）", new_full),
        ("IV)YDSWX（TVAX-009）-004(IV)", new_full),
        ("YDSWX（TVAX-009）-004（Ⅳ）", new_full),
        ("YDSWX(TVAX-009)-004(Ⅳ)", "YDSWX(TVAX-009)-004(III)"),
        ("YDSWX（TVAX-009）-004(Ⅳ)", "YDSWX（TVAX-009）-004(III)"),
        ("YDSWX(TVAX-009)-004(IV)", "YDSWX(TVAX-009)-004(III)"),
        # 修复历史错误替换产生的 I) 前缀
        ("I) YDSWX（TVAX-009）-004（III）", new_full),
        ("I)YDSWX（TVAX-009）-004（III）", new_full),
    ]
    literals = sorted(dict.fromkeys(literals), key=lambda x: -len(x[0]))
    return ReplaceRuleSet(literals=literals, regexes=[])


def build_default_rules() -> ReplaceRuleSet:
    d = build_date_rules()
    s = build_study_id_rules()
    literals = sorted(
        list(dict.fromkeys(d.literals + s.literals)),
        key=lambda x: -len(x[0]),
    )
    return ReplaceRuleSet(literals=literals, regexes=d.regexes + s.regexes)


def iter_text_nodes(container) -> list[etree._Element]:
    """按文档顺序收集 w:t / w:delText（含修订删除线文本）。"""
    return [el for el in container.iter() if el.tag in TEXT_TAGS]


def _write_back(nodes: Sequence[etree._Element], s_new: str, lengths: Sequence[int]) -> None:
    """写回 w:t / w:delText；替换后长度变化时合并到首个节点，避免 run 错位（如 I) 残留）。"""
    if not nodes:
        return
    if len(nodes) == 1:
        nodes[0].text = s_new
        return
    old_total = sum(lengths)
    if len(s_new) != old_total:
        nodes[0].text = s_new
        for el in nodes[1:]:
            el.text = ""
        return
    idx = 0
    for i, el in enumerate(nodes):
        if i < len(nodes) - 1:
            el.text = s_new[idx : idx + lengths[i]]
            idx += lengths[i]
        else:
            el.text = s_new[idx:]


def replace_in_nodes(nodes: Sequence[etree._Element], rules: ReplaceRuleSet, do_replace: bool) -> int:
    if not nodes:
        return 0
    texts = [el.text or "" for el in nodes]
    joined = "".join(texts)
    n = rules.count_in_text(joined)
    if n == 0 or not do_replace:
        return n
    s_new, _ = rules.apply(joined)
    _write_back(nodes, s_new, [len(x) for x in texts])
    return n


def replace_in_xml_tree(root, rules: ReplaceRuleSet, do_replace: bool) -> int:
    """按段落 w:p 处理，避免整格拼接多段文字；单段内仍跨 run 拼接。"""
    total = 0
    for p in root.xpath(".//w:p", namespaces=NS):
        total += replace_in_nodes(iter_text_nodes(p), rules, do_replace)
    return total


def iter_package_xml_files(extract_dir: str) -> list[str]:
    """处理 word/ 下全部 XML（正文、页眉页脚、脚注等）。"""
    word_dir = os.path.join(extract_dir, "word")
    if not os.path.isdir(word_dir):
        return []
    return sorted(glob.glob(os.path.join(word_dir, "*.xml")))


def process_docx(in_path: str, output_path: str, rules: ReplaceRuleSet) -> ReplaceStats:
    work_dir = tempfile.mkdtemp(prefix="docx_ooxml_")
    stats = ReplaceStats()
    try:
        with zipfile.ZipFile(in_path, "r") as z:
            z.extractall(work_dir)

        xml_files = iter_package_xml_files(work_dir)
        stats.before = sum(replace_in_xml_tree(etree.parse(x).getroot(), rules, False) for x in xml_files)

        for x in xml_files:
            tree = etree.parse(x)
            replace_in_xml_tree(tree.getroot(), rules, True)
            tree.write(x, encoding="UTF-8", xml_declaration=True)

        stats.after = sum(replace_in_xml_tree(etree.parse(x).getroot(), rules, False) for x in xml_files)

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for folder, _, filenames in os.walk(work_dir):
                for fn in filenames:
                    abs_fp = os.path.join(folder, fn)
                    z.write(abs_fp, os.path.relpath(abs_fp, work_dir))
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    return stats


def unique_output_path(in_path: str, output_dir: str, suffix: str = "_updated") -> str:
    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(in_path))[0]
    ext = os.path.splitext(in_path)[1] or ".docx"
    out = os.path.join(output_dir, base + suffix + ext)
    if not os.path.exists(out):
        return out
    i = 2
    while True:
        cand = os.path.join(output_dir, base + suffix + f"_{i}" + ext)
        if not os.path.exists(cand):
            return cand
        i += 1
