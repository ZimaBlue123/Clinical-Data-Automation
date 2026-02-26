from __future__ import annotations

"""
统一科研/医学风格的配色方案。

说明：
- 所有图表（柱状图、折线图、散点/标记等）应通过本模块获取颜色，避免在各处硬编码。
- 颜色均使用 HEX（不带 # 的 6 位大写）以便直接给 openpyxl 等库使用。
"""

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class ColorTheme:
    """全局配色主题定义。"""

    # 方案 A：多组对比（主系列）
    group_experimental: str = "2E5B88"  # Steel Blue
    group_control: str = "C0504D"  # Muted Brick Red
    group_placebo: str = "788496"  # Slate Gray
    group_other: str = "698B69"  # Sage Green

    # 方案 B：趋势/强调
    gridline: str = "D9D9D9"  # 辅助线/网格
    text_major: str = "333333"  # 主要文字（坐标轴、标题等）


# 单例主题
COLOR_THEME = ColorTheme()


def _hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    """HEX（不带 #）→ RGB 三元组。"""
    hex_str = hex_str.strip().lstrip("#")
    return int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)


def _rgb_to_hex(r: int, g: int, b: int) -> str:
    return f"{max(0, min(255, r)):02X}{max(0, min(255, g)):02X}{max(0, min(255, b)):02X}"


def _lighten(hex_color: str, factor: float) -> str:
    """
    适度提亮颜色，避免过暗。
    factor > 1 时越大越亮，这里建议 1.0–1.4 区间。
    """
    r, g, b = _hex_to_rgb(hex_color)
    # 转到 HLS，只调亮度
    import colorsys

    h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
    l = min(1.0, l * factor)
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return _rgb_to_hex(int(r2 * 255), int(g2 * 255), int(b2 * 255))


# 将业务中的组别名称映射到语义分组（英文 key），再映射到具体颜色。
_CATEGORY_CANONICAL_MAP: Dict[str, str] = {
    # 英文名 / 代码中的类别名
    "Experimental": "experimental",
    "Control": "control",
    "Placebo": "placebo",
    "Other": "other",
    # 中文业务名称映射
    "低剂量试验组": "experimental",
    "高剂量试验组": "experimental",
    "低剂量佐剂组": "control",
    "高剂量佐剂组": "control",
    "安慰剂组": "placebo",
}

_CANONICAL_TO_COLOR: Dict[str, str] = {
    "experimental": COLOR_THEME.group_experimental,
    "control": COLOR_THEME.group_control,
    "placebo": COLOR_THEME.group_placebo,
    "other": COLOR_THEME.group_other,
}


# 针对本项目中 5 个具体组别，提供“同色系不同明度”的精细配色：
# - 低/高剂量试验组：以 Steel Blue 为基色，高剂量略深，低剂量略浅；
# - 低/高剂量佐剂组：以 Muted Brick Red 为基色，同样做明度区分；
# - 安慰剂组：中性色 Slate Gray。
_GROUP_NAME_OVERRIDE: Dict[str, str] = {
    "低剂量试验组": _lighten(COLOR_THEME.group_experimental, 1.25),
    "高剂量试验组": _lighten(COLOR_THEME.group_experimental, 1.05),
    "低剂量佐剂组": _lighten(COLOR_THEME.group_control, 1.25),
    "高剂量佐剂组": _lighten(COLOR_THEME.group_control, 1.05),
    "安慰剂组": _lighten(COLOR_THEME.group_placebo, 1.15),
}


def _canonical_category(name: str) -> str | None:
    """将任意类别名称映射到规范 key，若无法识别则返回 None。"""
    if not name:
        return None
    key = name.strip()
    if key in _CATEGORY_CANONICAL_MAP:
        return _CATEGORY_CANONICAL_MAP[key]
    # 简单的英文/大小写兼容
    lowered = key.lower()
    if lowered in ("exp", "experimental", "treatment", "trial"):
        return "experimental"
    if lowered in ("ctl", "control", "positive control"):
        return "control"
    if lowered in ("placebo", "neg", "negative"):
        return "placebo"
    if lowered in ("other", "misc"):
        return "other"
    return None


def _hsl_to_rgb_hex(h: float, s: float, l: float) -> str:
    """简易 HSL → HEX，用于生成低饱和度回退色。"""
    import colorsys

    r, g, b = colorsys.hls_to_rgb(h, l, s)  # 注意：colorsys 使用 HLS 顺序
    return f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"


def _fallback_color_for_name(name: str) -> str:
    """
    对于未在 COLOR_THEME 中显式定义的类别，
    通过“低饱和度 + 中高亮度”的 HSL 生成稳定、舒适的颜色，而不是随机 RGB。
    """
    if not name:
        # 非法名称时，回退为中性灰
        return COLOR_THEME.group_placebo

    # 使用 hash 保证：同一名称在不同图表中颜色保持一致，但不同名称有一定区分度。
    h_seed = hash(name) & 0xFFFFFFFF
    # 均匀分布到 [0,1) 的色相
    h = (h_seed % 360) / 360.0
    # 低饱和度 + 偏亮，保证“高级感”
    s = 0.25
    l = 0.70
    return _hsl_to_rgb_hex(h, s, l)


def get_series_color(category_name: str) -> str:
    """
    根据数据类别名称返回合适的系列颜色（HEX，6 位大写，不带 #）。
    1. 若是本项目的 5 个剂量组名称，优先使用为其定制的“同色系不同明度”配色；
    2. 否则按 COLOR_THEME 的语义映射；
    3. 若仍未命中，则生成一个低饱和度的稳定回退色。
    """
    if category_name in _GROUP_NAME_OVERRIDE:
        return _GROUP_NAME_OVERRIDE[category_name]

    canonical = _canonical_category(category_name)
    if canonical and canonical in _CANONICAL_TO_COLOR:
        return _CANONICAL_TO_COLOR[canonical]
    return _fallback_color_for_name(category_name or "")


__all__ = [
    "COLOR_THEME",
    "ColorTheme",
    "get_series_color",
]

