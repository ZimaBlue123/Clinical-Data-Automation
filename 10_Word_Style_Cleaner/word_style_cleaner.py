import argparse
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree as ET

WD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def q(tag: str) -> str:
    return f"{{{WD_NS}}}{tag}"


STYLE_XML = "word/styles.xml"


@dataclass(frozen=True)
class StyleInfo:
    style_id: str
    style_type: str  # paragraph / character / table / numbering
    name: str | None
    custom: bool
    outline_level: int | None
    linked_style_id: str | None  # for character styles


def _iter_word_xml_parts(z: zipfile.ZipFile):
    for name in z.namelist():
        if not name.startswith("word/"):
            continue
        if not name.endswith(".xml"):
            continue
        if name == STYLE_XML:
            continue
        yield name


def collect_style_dependency_map(styles_root: ET.Element) -> dict[str, set[str]]:
    """
    styleId -> other styleIds it depends on (basedOn / link / next).
    Keeping these parents avoids broken style inheritance after pruning.
    """

    deps: dict[str, set[str]] = {}
    ref_tags = (q("basedOn"), q("link"), q("next"))
    for style in styles_root.findall(f".//{q('style')}"):
        sid = style.get(q("styleId"))
        if not sid:
            continue
        refs: set[str] = set()
        for tag in ref_tags:
            el = style.find(tag)
            if el is None:
                continue
            val = el.get(q("val"))
            if val:
                refs.add(val)
        if refs:
            deps[sid] = refs
    return deps


def expand_used_style_ids(used_ids: set[str], deps: dict[str, set[str]]) -> set[str]:
    expanded = set(used_ids)
    queue = list(used_ids)
    while queue:
        sid = queue.pop()
        for parent in deps.get(sid, ()):
            if parent not in expanded:
                expanded.add(parent)
                queue.append(parent)
    return expanded


def detect_used_style_ids(z: zipfile.ZipFile) -> set[str]:
    """
    Scan all parts of the docx (except styles.xml) and collect styleIds referenced by:
    w:pStyle, w:rStyle, w:tblStyle, w:tcStyle, w:trStyle
    """

    style_prop_tags = {q("pStyle"), q("rStyle"), q("tblStyle"), q("tcStyle"), q("trStyle")}
    used: set[str] = set()

    for xml_name in _iter_word_xml_parts(z):
        data = z.read(xml_name)
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue

        for el in root.iter():
            if el.tag in style_prop_tags:
                val = el.get(q("val"))
                if val:
                    used.add(val)
    return used


def parse_styles(styles_xml: bytes) -> tuple[ET.Element, list[StyleInfo]]:
    root = ET.fromstring(styles_xml)
    styles: list[StyleInfo] = []

    for style in root.findall(f".//{q('style')}"):
        style_id = style.get(q("styleId")) or ""
        style_type = style.get(q("type")) or ""
        custom = style.get(q("customStyle")) == "1"

        name_el = style.find(q("name"))
        name = name_el.get(q("val")) if name_el is not None else None

        outline_level = None
        ppr = style.find(q("pPr"))
        if ppr is not None:
            outline = ppr.find(q("outlineLvl"))
            if outline is not None:
                val = outline.get(q("val"))
                if val is not None and str(val).isdigit():
                    outline_level = int(val)

        linked_style_id = None
        link_el = style.find(q("link"))
        if link_el is not None:
            linked_style_id = link_el.get(q("val"))

        styles.append(
            StyleInfo(
                style_id=style_id,
                style_type=style_type,
                name=name,
                custom=custom,
                outline_level=outline_level,
                linked_style_id=linked_style_id,
            )
        )

    return root, styles


def extract_root_start_tag(xml_bytes: bytes) -> str | None:
    """
    Extract the full root start-tag (e.g. `<w:styles ...>`).
    We reuse it in output because ElementTree serialization may drop namespace declarations
    that Word expects (especially when `mc:Ignorable` references prefixes).
    """

    # Find the root start tag (skip XML declaration / doctype).
    # Prefer the canonical Word root tag name if present.
    start = xml_bytes.find(b"<w:styles")
    if start < 0:
        start = xml_bytes.find(b"<styles")
    if start < 0:
        # fallback: first '<' that is not '<?' or '<!'
        m = re.search(rb"<(?!\?)(?!\!)[^>]+>", xml_bytes[:20000])
        if not m:
            return None
        return m.group(0).decode("utf-8", errors="ignore")

    end = xml_bytes.find(b">", start)
    if end < 0:
        return None
    return xml_bytes[start : end + 1].decode("utf-8", errors="ignore")


def extract_xmlns_decls(xml_bytes: bytes) -> dict[str, str]:
    tag = extract_root_start_tag(xml_bytes)
    if not tag:
        return {}
    decls: dict[str, str] = {}
    for pfx, uri in re.findall(r'xmlns:([\w.-]+)=["\']([^"\']+)["\']', tag):
        decls[pfx] = uri
    return decls


def replace_root_start_tag(serialized_xml: bytes, original_start_tag: str) -> bytes:
    """
    Replace the first root start-tag in serialized_xml with original_start_tag.
    This keeps Word's expected namespace declarations intact.
    """

    s = serialized_xml.decode("utf-8", errors="ignore")

    # find root tag start
    i = s.find("<w:styles")
    if i < 0:
        i = s.find("<styles")
    if i < 0:
        # e.g. <ns0:styles ...>
        m = re.search(r"<[A-Za-z_][\w\-\.]*:styles\b", s)
        if m:
            i = m.start()
    if i < 0:
        return serialized_xml

    j = s.find(">", i)
    if j < 0:
        return serialized_xml

    return (s[:i] + original_start_tag + s[j + 1 :]).encode("utf-8")


def remove_unused_custom_styles(styles_root: ET.Element, styles: list[StyleInfo], used_ids: set[str]):
    custom_ids = {s.style_id for s in styles if s.custom and s.style_id}
    removed_ids = sorted(custom_ids.difference(used_ids))

    removed_details: list[tuple[str, str | None]] = []
    removed_set = set(removed_ids)

    for parent in list(styles_root.iter()):
        for child in list(parent):
            if child.tag != q("style"):
                continue
            if child.get(q("customStyle")) != "1":
                continue
            sid = child.get(q("styleId"))
            if not sid or sid not in removed_set:
                continue
            name_el = child.find(q("name"))
            style_name = name_el.get(q("val")) if name_el is not None else None
            removed_details.append((sid, style_name))
            parent.remove(child)

    kept_custom_ids = sorted(custom_ids.intersection(used_ids))
    return kept_custom_ids, removed_details


def strip_orphan_style_references(zin: zipfile.ZipFile, valid_style_ids: set[str]) -> dict[str, bytes]:
    """
    Remove w:pStyle/w:rStyle/... that point to deleted styleIds.
    Safety net when pruning leaves dangling references in document parts.
    """

    prop_tags = {q("pStyle"), q("rStyle"), q("tblStyle"), q("tcStyle"), q("trStyle")}
    patched: dict[str, bytes] = {}

    for xml_name in _iter_word_xml_parts(zin):
        data = zin.read(xml_name)
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue

        changed = False
        for parent in root.iter():
            for child in list(parent):
                if child.tag not in prop_tags:
                    continue
                val = child.get(q("val"))
                if val and val not in valid_style_ids:
                    parent.remove(child)
                    changed = True

        if changed:
            patched[xml_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    return patched


def _safe_set_style_name(styles_root: ET.Element, style_id: str, new_name: str) -> bool:
    for style in styles_root.findall(f".//{q('style')}"):
        if style.get(q("styleId")) != style_id:
            continue
        name_el = style.find(q("name"))
        if name_el is None:
            name_el = ET.SubElement(style, q("name"))
        name_el.set(q("val"), new_name)
        return True
    return False


def normalize_style_names(  # noqa: PLR0915 - TODO: 下个迭代重构 # noqa: PLR0912 - TODO: 下个迭代重构
    styles_root: ET.Element,
    styles: list[StyleInfo],
    kept_custom_ids: list[str],
) -> list[tuple[str, str | None, str]]:
    """
    Rename *kept* custom styles by updating w:name/@w:val (do NOT change styleId).

    Strategy (stricter naming system):
    - Chapter headings:
      - Paragraph styles with outlineLvl 0..8 (excluding caption-like keywords) => 标题{lvl+1}
      - Character styles linked to those headings => 标题{lvl} 字符
    - Body paragraphs => 正文{n} (always with Arabic digit)
    - Table styles => 表格{n} (always with Arabic digit)
    - Caption/title-like paragraph styles (keyword-based, always with Arabic digit):
      - 图标题{n}, 表标题{n}, 图表题{n}, 题注{n}
      - Character styles linked to those captions => 同名 + " 字符"
    - Anything else: keep original name if present, otherwise keep styleId.
    """

    by_id = {s.style_id: s for s in styles if s.style_id}
    kept = [by_id[sid] for sid in kept_custom_ids if sid in by_id]

    def _kw(s: StyleInfo) -> str:
        return (s.name or "").strip()

    def _is_caption_like(name: str) -> bool:
        n = name.lower()
        return any(k in n for k in ("图标题", "表标题", "图表题", "题注", "caption"))

    def _caption_category(name: str) -> str | None:
        # Order matters
        if "图表题" in name:
            return "图表题"
        if "图标题" in name or "图题" in name:
            return "图标题"
        if "表标题" in name or "表题" in name:
            return "表标题"
        if "题注" in name:
            return "题注"
        if "caption" in name.lower():
            return "题注"
        return None

    heading_para: dict[str, int] = {}
    for s in kept:
        if s.style_type == "paragraph" and s.outline_level is not None and 0 <= s.outline_level <= 8:
            # Exclude caption-like names even if outlineLvl is set.
            if _is_caption_like(_kw(s)):
                continue
            heading_para[s.style_id] = s.outline_level + 1

    # Assign body paragraph styles
    body_candidates = [
        s
        for s in kept
        if s.style_type == "paragraph"
        and s.style_id not in heading_para
        and (s.name and ("正文" in s.name or "Body" in s.name or "body" in s.name))
        and (_caption_category(_kw(s)) is None)
    ]
    body_candidates_sorted = sorted(body_candidates, key=lambda x: (x.name or "", x.style_id))
    body_assign: dict[str, int] = {}
    for idx, s in enumerate(body_candidates_sorted, start=1):
        body_assign[s.style_id] = idx

    # Assign table styles
    table_candidates = [s for s in kept if s.style_type == "table"]
    table_candidates_sorted = sorted(table_candidates, key=lambda x: (x.name or "", x.style_id))
    table_assign: dict[str, int] = {}
    for idx, s in enumerate(table_candidates_sorted, start=1):
        table_assign[s.style_id] = idx

    # Assign caption-like paragraph styles
    caption_candidates = [s for s in kept if s.style_type == "paragraph" and (_caption_category(_kw(s)) is not None)]
    caption_by_cat: dict[str, list[StyleInfo]] = {}
    for s in caption_candidates:
        cat = _caption_category(_kw(s))
        if cat is None:
            continue
        caption_by_cat.setdefault(cat, []).append(s)

    caption_assign: dict[str, tuple[str, int]] = {}  # styleId -> (cat, idx)
    for cat, lst in caption_by_cat.items():
        for idx, s in enumerate(sorted(lst, key=lambda x: (x.name or "", x.style_id)), start=1):
            caption_assign[s.style_id] = (cat, idx)

    renames: list[tuple[str, str | None, str]] = []
    used_target_names: set[str] = set()

    for s in kept:
        old = s.name
        new_name = None

        # Caption-like names (strict numbering)
        if s.style_type == "paragraph" and s.style_id in caption_assign:
            cat, idx = caption_assign[s.style_id]
            new_name = f"{cat}{idx}"
        elif s.style_type == "character" and s.linked_style_id and s.linked_style_id in caption_assign:
            cat, idx = caption_assign[s.linked_style_id]
            new_name = f"{cat}{idx} 字符"
        elif s.style_type == "paragraph" and s.style_id in heading_para:
            new_name = f"标题{heading_para[s.style_id]}"
        elif s.style_type == "character" and s.linked_style_id and s.linked_style_id in heading_para:
            new_name = f"标题{heading_para[s.linked_style_id]} 字符"
        elif s.style_type == "paragraph" and s.style_id in body_assign:
            n = body_assign[s.style_id]
            new_name = f"正文{n}"
        elif s.style_type == "table" and s.style_id in table_assign:
            n = table_assign[s.style_id]
            new_name = f"表格{n}"

        if new_name is None:
            # keep old name, but try to normalize obvious "X级标题" to "标题N"
            if old:
                m = re.search(r"([一二三四五六七八九十])级标题", old)
                if m:
                    cn = m.group(1)
                    cn_map = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
                    if cn in cn_map:
                        new_name = f"标题{cn_map[cn]}"
                else:
                    new_name = old
            else:
                new_name = s.style_id

        # Avoid duplicate target names among renamed custom styles.
        if new_name:
            base = new_name
            if base in used_target_names:
                i = 2
                while f"{base}（{i}）" in used_target_names:
                    i += 1
                new_name = f"{base}（{i}）"
            used_target_names.add(new_name)

        if new_name != old and _safe_set_style_name(styles_root, s.style_id, new_name):
            renames.append((s.style_id, old, new_name))

    return renames


def _write_success_report(
    report_path: str,
    in_docx: str,
    out_docx: str,
    removed_details: list[tuple[str, str | None]],
    renames: list[tuple[str, str | None, str]],
    orphan_patches: int,
) -> None:
    removed_details_sorted = sorted(removed_details, key=lambda x: ((x[1] or ""), x[0]))
    renames_sorted = sorted(renames, key=lambda x: (x[2], x[0]))

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("Word 自定义样式清理报告\n")
        f.write("====================\n")
        f.write(f"输入: {in_docx}\n")
        f.write(f"输出: {out_docx}\n")
        f.write("\n")
        f.write(f"删除未使用自定义样式数: {len(removed_details_sorted)}\n")
        f.write(f"重命名自定义样式数: {len(renames_sorted)}\n")
        f.write(f"修复孤儿样式引用（XML 部件数）: {orphan_patches}\n")
        f.write("\n")

        f.write("移除的自定义样式（styleId -> name）\n")
        for sid, sname in removed_details_sorted:
            f.write(f"- {sid} -> {sname}\n")

        f.write("\n重命名的自定义样式（styleId: old -> new）\n")
        for sid, old, new in renames_sorted:
            f.write(f"- {sid}: {old} -> {new}\n")


def process_one_docx(  # noqa: PLR0915 - TODO: 下个迭代重构 # noqa: PLR0912 - TODO: 下个迭代重构
    in_docx: str,
    out_docx: str,
    report_path: str,
    rename_styles: bool,
) -> bool:
    try:
        with zipfile.ZipFile(in_docx, "r") as zin:
            if STYLE_XML not in zin.namelist():
                raise RuntimeError(f"docx 中未找到 {STYLE_XML}")

            styles_xml = zin.read(STYLE_XML)
            original_root_tag = extract_root_start_tag(styles_xml)
            xmlns_decls = extract_xmlns_decls(styles_xml)
            for pfx, uri in xmlns_decls.items():
                try:
                    ET.register_namespace(pfx, uri)
                except ValueError:
                    pass

            styles_root, styles = parse_styles(styles_xml)
            used_ids = detect_used_style_ids(zin)
            used_ids = expand_used_style_ids(used_ids, collect_style_dependency_map(styles_root))

            kept_custom_ids, removed_details = remove_unused_custom_styles(styles_root, styles, used_ids)

            renames: list[tuple[str, str | None, str]] = []
            if rename_styles:
                new_styles_xml_bytes = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
                styles_root2, styles2 = parse_styles(new_styles_xml_bytes)
                renames = normalize_style_names(styles_root2, styles2, kept_custom_ids)
                styles_root = styles_root2

            final_styles_xml_bytes = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
            if original_root_tag:
                final_styles_xml_bytes = replace_root_start_tag(final_styles_xml_bytes, original_root_tag)

            valid_style_ids = {s.style_id for s in parse_styles(final_styles_xml_bytes)[1] if s.style_id}
            orphan_patches = strip_orphan_style_references(zin, valid_style_ids)

            out_dir = os.path.dirname(os.path.abspath(out_docx)) or "."
            os.makedirs(out_dir, exist_ok=True)
            fd, tmp_path = tempfile.mkstemp(suffix=".docx", dir=out_dir)
            os.close(fd)
            try:
                with zipfile.ZipFile(tmp_path, "w") as zout:
                    for item in zin.infolist():
                        if item.filename == STYLE_XML:
                            zout.writestr(item, final_styles_xml_bytes)
                        elif item.filename in orphan_patches:
                            zout.writestr(item, orphan_patches[item.filename])
                        else:
                            zout.writestr(item, zin.read(item.filename))
                os.replace(tmp_path, out_docx)
            finally:
                if os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass

        _write_success_report(
            report_path,
            in_docx,
            out_docx,
            removed_details,
            renames,
            len(orphan_patches),
        )
        return True
    except (OSError, zipfile.BadZipFile, ET.ParseError, RuntimeError, KeyError) as e:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("Word 自定义样式清理报告（失败）\n")
            f.write("========================\n")
            f.write(f"输入: {in_docx}\n")
            f.write(f"输出: {out_docx}\n")
            f.write("\n")
            f.write(f"错误: {type(e).__name__}: {e}\n")
        return False


def iter_docx_files(input_dir: str, recursive: bool) -> list[str]:
    docx_files: list[str] = []
    if recursive:
        for root, _dirs, files in os.walk(input_dir):
            for fn in files:
                if fn.lower().endswith(".docx") and not fn.startswith("~$"):
                    docx_files.append(os.path.join(root, fn))
    else:
        for fn in os.listdir(input_dir):
            if fn.lower().endswith(".docx") and not fn.startswith("~$"):
                docx_files.append(os.path.join(input_dir, fn))
    return sorted(docx_files)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default="./input", help="输入文件夹")
    ap.add_argument("--output", default="./output", help="输出文件夹")
    ap.add_argument("--report-dir", default=None, help="报告输出目录（默认与输出文件同目录）")
    ap.add_argument("--suffix", default="_styles_cleaned", help="输出文件名后缀（不含扩展名）")
    ap.add_argument("--recursive", dest="recursive", action="store_true", help="递归处理子目录（默认）")
    ap.add_argument("--no-recursive", dest="recursive", action="store_false", help="仅处理输入目录顶层")
    ap.set_defaults(recursive=True)
    ap.add_argument("--overwrite", action="store_true", help="输出文件已存在则覆盖")
    ap.add_argument("--no-rename-styles", action="store_true", help="不做样式命名规范化，仅清理未用自定义样式")
    args = ap.parse_args()

    input_dir = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output)
    report_dir = os.path.abspath(args.report_dir) if args.report_dir else output_dir

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(report_dir, exist_ok=True)

    docx_files = iter_docx_files(input_dir, recursive=args.recursive)
    if not docx_files:
        raise SystemExit(f"未找到 docx：{input_dir}")

    ok = 0
    failed = 0
    skipped = 0

    for in_docx in docx_files:
        base = os.path.splitext(os.path.basename(in_docx))[0]
        out_docx = os.path.join(output_dir, f"{base}{args.suffix}.docx")
        report_path = os.path.join(report_dir, f"{base}{args.suffix}_report.txt")

        if (not args.overwrite) and (os.path.exists(out_docx) or os.path.exists(report_path)):
            skipped += 1
            print(f"[跳过] {os.path.basename(in_docx)}")
            continue

        print(f"[处理] {os.path.basename(in_docx)}")
        if process_one_docx(
            in_docx=in_docx,
            out_docx=out_docx,
            report_path=report_path,
            rename_styles=not args.no_rename_styles,
        ):
            ok += 1
            print(f"  -> {os.path.basename(out_docx)}")
        else:
            failed += 1
            print(f"  -> 失败，见报告: {os.path.basename(report_path)}", file=sys.stderr)

    print("完成。")
    print(f"成功 {ok}，失败 {failed}，跳过 {skipped}")
    print("输出目录:", output_dir)
    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
