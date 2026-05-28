"""共享库导入冒烟测试（CI 无 Office / Paddle 环境）。"""

from __future__ import annotations


def test_import_src_packages():
    import src.color_theme  # noqa: F401
    import src.excel_writer  # noqa: F401
    import src.pdf_reader  # noqa: F401
    import src.serology_utils  # noqa: F401


def test_color_theme_series_color():
    from src.color_theme import get_series_color

    color = get_series_color("试验组")
    assert isinstance(color, str) and len(color) == 6 and color.isalnum()
