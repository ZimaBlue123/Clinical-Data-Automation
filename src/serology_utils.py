from __future__ import annotations

import re
from typing import Final

# 五项指标在 Excel 中的固定顺序
OUTPUT_MARKERS: Final[tuple[str, ...]] = (
    "Anti-HBs",
    "HBsAg",
    "Anti-HBc",
    "Anti-HBe",
    "HBeAg",
)

# 单位（与各模块输出表头一致）
MARKER_UNITS: Final[dict[str, str]] = {
    "Anti-HBs": "mIU/ml",
    "HBsAg": "IU/ml",
    "Anti-HBc": "S/CO",
    "Anti-HBe": "S/CO",
    "HBeAg": "S/CO",
}


def canonical_sample_id(raw: str) -> str:
    """
    样品 ID 跨模块归一化（PDF/Word/OCR 可能造成的字符误差）。

    目标形态：`001-D0-a` / `725-M1-a`
    """
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""

    # 常见分隔符/连接符统一
    s = s.replace("—", "-").replace("–", "-").replace("_", "-")
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")

    parts = [p for p in s.split("-") if p]
    if len(parts) < 3:
        m = re.search(r"([A-Za-z0-9]+)-([A-Za-z0-9]+)-([A-Za-z0-9]+)", s)
        if m:
            parts = [m.group(1), m.group(2), m.group(3)]
        else:
            return s

    p0, p1, p2 = parts[0], parts[1], parts[2]

    # OCR 常见误识别修正
    p0 = p0.replace("O", "0").replace("o", "0").replace("I", "1").replace("l", "1")
    if p0.isdigit():
        p0 = p0.zfill(3)
    p1 = p1.upper().replace("O", "0").replace("I", "1").replace("L", "1")
    p2 = p2.lower()

    return f"{p0}-{p1}-{p2}"

