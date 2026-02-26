# -*- coding: utf-8 -*-
"""
Apply Clinical Publication Color Scheme to ADR charts（仅改配色）.

策略：复制 input 到 output，只修改工作簿内图表的系列颜色，其余内容与格式原样保留。
实现：在 zip 内对 xl/charts/chart*.xml 做纯正则替换（不解析 XML），保证 xlsx 结构不被破坏。
"""

from __future__ import annotations

import argparse
import logging
import re
import shutil
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = BASE / "input"
DEFAULT_OUTPUT_DIR = BASE / "output"
DEFAULT_INPUT = DEFAULT_INPUT_DIR / "不同剂量组ADR分析 (TFL).xlsx"
DEFAULT_OUTPUT = DEFAULT_OUTPUT_DIR / "不同剂量组ADR分析_clinical_colors.xlsx"

logger = logging.getLogger(__name__)

# 期刊调色板库（原始 HEX 可含 #，注入前会统一 strip；slug 用于输出文件名后缀）
JOURNAL_PALETTES = {
    "1": {
        "name": "NPG (Nature)",
        "slug": "NPG",
        "hex": ["#E64B35", "#4DBBD5", "#00A087", "#3C5488", "#F39B7F"],
    },
    "2": {
        "name": "Lancet",
        "slug": "Lancet",
        "hex": ["#00468B", "#ED0000", "#42B540", "#0099B4", "#925E9F"],
    },
    "3": {
        "name": "NEJM",
        "slug": "NEJM",
        "hex": ["#BC3C29", "#0072B5", "#E18727", "#20854E"],
    },
}

# 默认调色板（无 #，Excel srgbClr 用）
DEFAULT_PALETTE = ["5B9BD5", "254061", "ED7D31", "C00000", "7F7F7F"]

# 匹配 <c:ser ...> ... </c:ser> 整块（含换行），非贪婪
SERIES_BLOCK_RE = re.compile(r"<c:ser[^>]*>.*?</c:ser>", re.DOTALL)
# 匹配系列内的 srgbClr 标签，允许 val 前后有空格
SRGB_CLR_RE = re.compile(r'<a:srgbClr\s+val="[0-9A-Fa-f]{6}"\s*/>', re.IGNORECASE)


def _strip_hex(hex_list: list[str]) -> list[str]:
    """去掉 #、转大写，供 Excel srgbClr 使用（禁止 #）。"""
    out = []
    for s in hex_list:
        h = (s or "").strip().lstrip("#").upper()
        if len(h) == 6:
            out.append(h)
    return out if out else DEFAULT_PALETTE


def _interactive_color_menu() -> tuple[list[str], str, str]:
    """
    显示 ASCII 菜单，读取 1/2/3，返回 (cleaned_hex_list, scheme_name, slug)。
    无效输入默认 1 (NPG)。
    """
    print()
    print("[?] Select Color Scheme:")
    print("    1. NPG (Nature) - Red/Blue/Green match")
    print("    2. Lancet - High Contrast")
    print("    3. NEJM - Professional/Steady")
    print()
    choice = (input("Choice [1]: ").strip() or "1")
    if choice not in JOURNAL_PALETTES:
        choice = "1"
    entry = JOURNAL_PALETTES[choice]
    palette = _strip_hex(entry["hex"])
    slug = entry.get("slug", entry["name"].split()[0])
    return (palette, entry["name"], slug)


def _patch_chart_xml_by_regex(xml_bytes: bytes, palette: list[str]) -> bytes:
    """
    仅通过正则替换修改图表系列颜色，不解析 XML。
    对每个 <c:ser>...</c:ser> 块内的所有 <a:srgbClr val="..."/> 替换为调色板中对应颜色；
    按系列顺序循环使用 palette（6 位 HEX 无 #），保证结构 100% 不变。
    """
    try:
        content = xml_bytes.decode("utf-8")
    except Exception:
        return xml_bytes

    if not palette:
        palette = DEFAULT_PALETTE

    parts: list[str] = []
    last_end = 0
    series_index = 0

    for m in SERIES_BLOCK_RE.finditer(content):
        parts.append(content[last_end : m.start()])
        block = m.group(0)
        color = palette[series_index % len(palette)]
        new_block = SRGB_CLR_RE.sub(f'<a:srgbClr val="{color}"/>', block)
        parts.append(new_block)
        last_end = m.end()
        series_index += 1

    parts.append(content[last_end:])
    return "".join(parts).encode("utf-8")


def build_with_clinical_colors(tfl_path: Path, out_path: Path, palette: list[str]) -> None:
    """
    复制 input 到 output，仅对 xl/charts/chart*.xml 做正则配色替换（使用给定 palette），其余 zip 成员原样写回。
    """
    if not tfl_path.exists():
        raise FileNotFoundError(f"源 Excel 文件不存在: {tfl_path}")

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink(missing_ok=True)
    shutil.copy2(tfl_path, out_path)

    chart_prefix = "xl/charts/chart"
    chart_suffix = ".xml"
    to_patch = set()

    with zipfile.ZipFile(out_path, "r") as zread:
        for n in zread.namelist():
            if n.startswith(chart_prefix) and n.endswith(chart_suffix) and "/_rels/" not in n:
                to_patch.add(n)
        new_entries: list[tuple[str, bytes]] = []
        for n in zread.namelist():
            data = zread.read(n)
            if n in to_patch:
                try:
                    data = _patch_chart_xml_by_regex(data, palette)
                except Exception as e:
                    logger.warning("跳过图表配色 %s: %s", n, e)
            new_entries.append((n, data))

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in new_entries:
            zout.writestr(name, data)

    print("Clinical color charts applied (colors only, regex):", out_path)


def _select_palette(choice: str | None) -> tuple[list[str], str, str]:
    if choice:
        key = choice.strip()
        key_map = {"npg": "1", "lancet": "2", "nejm": "3"}
        key = key_map.get(key.lower(), key)
        if key in JOURNAL_PALETTES:
            entry = JOURNAL_PALETTES[key]
            palette = _strip_hex(entry["hex"])
            slug = entry.get("slug", entry["name"].split()[0])
            return palette, entry["name"], slug
    return _interactive_color_menu()


def main() -> None:
    parser = argparse.ArgumentParser(description="仅对 ADR 图表应用期刊配色（正则替换），保留 input 其余内容与格式")
    parser.add_argument("--input", "-i", default=str(DEFAULT_INPUT), help="输入 Excel 或输入文件夹（批量）")
    parser.add_argument("--output", "-o", default=str(DEFAULT_OUTPUT), help="输出 Excel 或输出文件夹（批量）")
    parser.add_argument("--batch", "-b", action="store_true", help="批量：对 input 下所有 .xlsx 仅改配色")
    parser.add_argument(
        "--palette",
        default=None,
        help="指定配色方案：1(NPG)/2(Lancet)/3(NEJM) 或 NPG/Lancet/NEJM",
    )
    args = parser.parse_args()

    # 非交互环境默认使用 NPG
    palette_choice = args.palette
    if not palette_choice and not sys.stdin.isatty():
        palette_choice = "1"

    palette, scheme_name, slug = _select_palette(palette_choice)

    try:
        DEFAULT_INPUT_DIR.mkdir(parents=True, exist_ok=True)
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        logger.warning("创建默认目录失败: %s", e)

    if args.batch:
        input_dir = Path(args.input)
        output_dir = Path(args.output)
        if not input_dir.exists() or not input_dir.is_dir():
            raise FileNotFoundError(f"输入文件夹不存在或不是目录: {input_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
        files = [p for p in input_dir.glob("*.xlsx") if not p.name.startswith("~$")]
        if not files:
            print(f"批量模式：在 {input_dir} 下未找到 .xlsx 文件。")
            return
        print(f"批量模式：共 {len(files)} 个 Excel，仅改图表配色（正则）...")
        for idx, src in enumerate(files, start=1):
            dst = output_dir / f"{src.stem}_clinical_colors_{slug}.xlsx"
            print(f"[{idx}/{len(files)}] {src.name} -> {dst.name}")
            try:
                build_with_clinical_colors(src, dst, palette)
            except Exception as e:
                logger.error("处理 %s 失败: %s", src, e)
        print("批量完成，输出目录:", output_dir)
    else:
        tfl_path = Path(args.input)
        out_path = Path(args.output)
        if not tfl_path.exists():
            raise FileNotFoundError(f"输入文件不存在: {tfl_path}")
        # 若为默认输出路径，则文件名后缀体现配色方案
        if out_path == Path(DEFAULT_OUTPUT):
            out_path = DEFAULT_OUTPUT_DIR / f"{tfl_path.stem}_clinical_colors_{slug}.xlsx"
        build_with_clinical_colors(tfl_path, out_path, palette)

    print(f"Done. Applied {scheme_name} style.")


if __name__ == "__main__":
    main()
