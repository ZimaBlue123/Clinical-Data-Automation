"""Pytest 路径：仓库根目录与 11_Word_Text_Replace/lib。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORD_REPLACE_LIB = ROOT / "11_Word_Text_Replace" / "lib"

for path in (ROOT, WORD_REPLACE_LIB):
    s = str(path)
    if s not in sys.path:
        sys.path.insert(0, s)
