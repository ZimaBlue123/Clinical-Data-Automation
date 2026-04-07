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
# 匹配系列内的 srgbClr / schemeClr（支持自闭合与非自闭合）
SRGB_VAL_RE = re.compile(r'(<a:srgbClr\b[^>]*\bval=")[0-9A-Fa-f]{6}(")', re.IGNORECASE)
SCHEME_SELF_RE = re.compile(r'<a:schemeClr\b[^>]*\bval="[^"]+"\s*/>', re.IGNORECASE)
SCHEME_BLOCK_RE = re.compile(r'<a:schemeClr\b[^>]*\bval="[^"]+"\b[^>]*>.*?</a:schemeClr>', re.IGNORECASE | re.DOTALL)
# solidFill 常用于线/填充：直接替换为指定 srgbClr
SOLIDFILL_RE = re.compile(r"<a:solidFill>.*?</a:solidFill>", re.DOTALL | re.IGNORECASE)

# 从 <c:ser> 中提取系列标题（优先 a:t，其次 c:v）
SERIES_TITLE_RE = re.compile(
    r"<c:tx\b[^>]*>.*?(?:<a:t>(?P<at>.*?)</a:t>|<c:v>(?P<cv>.*?)</c:v>).*?</c:tx>",
    re.DOTALL | re.IGNORECASE,
)


def _normalize_series_key(title: str) -> str:
    """
    将系列标题归一化为“组键”，用于让同一组的柱/线共享颜色。
    规则尽量保守：仅移除常见的度量后缀/括号内容，避免误合并。
    """
    t = (title or "").strip()
    if not t:
        return ""
    # 去掉括号内容（中英文括号）
    t = re.sub(r"[\(（].*?[\)）]", "", t).strip()
    # 去掉常见度量词（仅末尾）
    t = re.sub(r"(发生率|例数|人数|频数|n|N|%|％)\s*$", "", t).strip()
    # 压缩空白
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _extract_series_title(block: str) -> str:
    m = SERIES_TITLE_RE.search(block)
    if not m:
        return ""
    return (m.group("at") or m.group("cv") or "").strip()


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


def _apply_color_in_series_block(block: str, color: str) -> tuple[str, dict[str, int]]:
    """
    在单个 <c:ser>...</c:ser> 块内强制应用同一颜色。
    覆盖路径：
    - <a:solidFill>...</a:solidFill>（最常见的线/填充定义）
    - <a:srgbClr val="......">（自闭合或非自闭合）
    - <a:schemeClr val="accentX">（主题色引用）
    """
    stats = {"solidFill": 0, "srgbVal": 0, "scheme": 0}

    def _solidfill_sub(_: re.Match) -> str:
        stats["solidFill"] += 1
        return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

    block2 = SOLIDFILL_RE.sub(_solidfill_sub, block)

    # srgbClr val 改写为目标色（不改变标签其余属性/结构）
    def _srgb_val_sub(m: re.Match) -> str:
        return f"{m.group(1)}{color}{m.group(2)}"

    block3, n_srgb = SRGB_VAL_RE.subn(_srgb_val_sub, block2)
    stats["srgbVal"] += n_srgb

    # schemeClr（自闭合或带子节点）统一替换为 srgbClr
    def _scheme_to_srgb(_: re.Match) -> str:
        stats["scheme"] += 1
        return f'<a:srgbClr val="{color}"/>'

    block4 = SCHEME_SELF_RE.sub(_scheme_to_srgb, block3)
    block5 = SCHEME_BLOCK_RE.sub(_scheme_to_srgb, block4)

    return block5, stats


def _patch_chart_xml_by_regex(xml_bytes: bytes, palette: list[str]) -> tuple[bytes, dict[str, int]]:
    """
    仅通过正则替换修改图表系列颜色，不解析 XML。
    对每个 <c:ser>...</c:ser> 块内的颜色定义替换为调色板中对应颜色；
    按系列顺序循环使用 palette（6 位 HEX 无 #），保证结构 100% 不变。
    """
    try:
        content = xml_bytes.decode("utf-8")
    except Exception:
        return xml_bytes, {"series": 0, "solidFill": 0, "srgbVal": 0, "scheme": 0}

    if not palette:
        palette = DEFAULT_PALETTE

    parts: list[str] = []
    last_end = 0
    # 同一“组键”强制同色；色号只在组首次出现时递增
    group_color: dict[str, str] = {}
    group_index = 0
    total = {"series": 0, "solidFill": 0, "srgbVal": 0, "scheme": 0}

    matches = list(SERIES_BLOCK_RE.finditer(content))
    for m in matches:
        parts.append(content[last_end : m.start()])
        block = m.group(0)
        title = _extract_series_title(block)
        key = _normalize_series_key(title) or title or f"__series_{total['series']}"
        if key not in group_color:
            group_color[key] = palette[group_index % len(palette)]
            group_index += 1
        color = group_color[key]
        new_block, stats = _apply_color_in_series_block(block, color)
        parts.append(new_block)
        last_end = m.end()
        total["series"] += 1
        total["solidFill"] += stats["solidFill"]
        total["srgbVal"] += stats["srgbVal"]
        total["scheme"] += stats["scheme"]

    parts.append(content[last_end:])
    patched = "".join(parts)

    # 兜底：若没有 <c:ser>（少数图表/结构），则全局替换主题色引用为首个颜色，至少保证“看得见”变化
    if not matches:
        color = palette[0]
        patched2 = SCHEME_SELF_RE.sub(f'<a:srgbClr val="{color}"/>', patched)
        patched2 = SCHEME_BLOCK_RE.sub(f'<a:srgbClr val="{color}"/>', patched2)
        patched2, n_srgb = SRGB_VAL_RE.subn(lambda m: f"{m.group(1)}{color}{m.group(2)}", patched2)
        total["scheme"] += 0  # 不精确统计兜底次数，避免误导
        total["srgbVal"] += n_srgb
        patched = patched2

    return patched.encode("utf-8"), total


def build_with_clinical_colors(tfl_path: Path, out_path: Path, palette: list[str], *, verbose: bool = False) -> None:
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
    agg = {"charts": 0, "series": 0, "solidFill": 0, "srgbVal": 0, "scheme": 0}

    with zipfile.ZipFile(out_path, "r") as zread:
        for n in zread.namelist():
            if n.startswith(chart_prefix) and n.endswith(chart_suffix) and "/_rels/" not in n:
                to_patch.add(n)
        new_entries: list[tuple[str, bytes]] = []
        for n in zread.namelist():
            data = zread.read(n)
            if n in to_patch:
                try:
                    data, stats = _patch_chart_xml_by_regex(data, palette)
                    agg["charts"] += 1
                    agg["series"] += stats.get("series", 0)
                    agg["solidFill"] += stats.get("solidFill", 0)
                    agg["srgbVal"] += stats.get("srgbVal", 0)
                    agg["scheme"] += stats.get("scheme", 0)
                except Exception as e:
                    logger.warning("跳过图表配色 %s: %s", n, e)
            new_entries.append((n, data))

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in new_entries:
            zout.writestr(name, data)

    print("Clinical color charts applied (colors only, regex):", out_path)
    if verbose:
        print(
            f"Patched charts={agg['charts']} series={agg['series']} solidFill={agg['solidFill']} "
            f"srgbVal={agg['srgbVal']} schemeClr={agg['scheme']}"
        )


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
    parser.add_argument("--verbose", "-v", action="store_true", help="输出自检统计（被替换的颜色节点数量）")
    parser.add_argument(
        "--n-colors",
        type=int,
        default=3,
        help="每个图表最多使用的颜色数量（默认 3；同一组柱/线共享颜色，超过则循环）",
    )
    parser.add_argument(
        "--palette",
        default=None,
        help="指定配色方案：1(NPG)/2(Lancet)/3(NEJM) 或 NPG/Lancet/NEJM",
    )
    args = parser.parse_args()

    import sys
    # 非交互环境默认使用 NPG
    palette_choice = args.palette
    if not palette_choice and not sys.stdin.isatty():
        palette_choice = "1"

    palette, scheme_name, slug = _select_palette(palette_choice)
    if args.n_colors and args.n_colors > 0:
        palette = palette[: args.n_colors] if len(palette) >= args.n_colors else palette

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
                build_with_clinical_colors(src, dst, palette, verbose=args.verbose)
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
        build_with_clinical_colors(tfl_path, out_path, palette, verbose=args.verbose)

    print(f"Done. Applied {scheme_name} style.")


if __name__ == "__main__":
    main()
