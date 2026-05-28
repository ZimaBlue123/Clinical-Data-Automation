"""src.serology_utils 单元测试。"""

from __future__ import annotations

import pytest

from src.serology_utils import OUTPUT_MARKERS, canonical_sample_id


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("001-D0-a", "001-D0-a"),
        ("1-D0-a", "001-D0-a"),
        ("725-M1-a", "725-M1-a"),
        ("", ""),
        (None, ""),
        ("001—D0—a", "001-D0-a"),
    ],
)
def test_canonical_sample_id(raw, expected):
    assert canonical_sample_id(raw) == expected


def test_output_markers():
    assert len(OUTPUT_MARKERS) == 5
    assert "Anti-HBs" in OUTPUT_MARKERS
