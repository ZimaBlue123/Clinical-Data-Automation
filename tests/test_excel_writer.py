"""src.excel_writer 单元测试。"""

from __future__ import annotations

import pytest

from src.excel_writer import cell_to_row_col


@pytest.mark.parametrize(
    ("cell", "row", "col"),
    [
        ("B3", 3, 2),
        ("A1", 1, 1),
        ("AA10", 10, 27),
    ],
)
def test_cell_to_row_col(cell, row, col):
    assert cell_to_row_col(cell) == (row, col)


def test_cell_to_row_col_invalid():
    with pytest.raises(ValueError):
        cell_to_row_col("")
