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


def test_ectd_converter_basic_processing(tmp_path):
    import sys
    from pathlib import Path
    import fitz

    module_dir = Path(__file__).resolve().parent.parent / "18_PDF_eCTD_Converter"
    if str(module_dir) not in sys.path:
        sys.path.insert(0, str(module_dir))

    from pdf_ectd_converter import ECTDComplianceCleaner

    # 创建测试 PDF
    input_pdf = tmp_path / "sample.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((50, 50), "Hello eCTD Test Document")
    doc.save(str(input_pdf))
    doc.close()

    output_pdf = tmp_path / "sample_ectd.pdf"
    report_xlsx = tmp_path / "report.xlsx"

    cleaner = ECTDComplianceCleaner(
        input_dir=tmp_path,
        output_dir=tmp_path,
        report_path=report_xlsx,
        overwrite=True,
    )

    # 1. 仅校验模式：不生成输出，源文件保留
    val_res = cleaner.process_pdf(input_pdf, output_pdf, validate_only=True)
    assert val_res is True
    assert input_pdf.exists()

    # 2. 正常转换模式：输出生成成功
    proc_res = cleaner.process_pdf(input_pdf, output_pdf, validate_only=False)
    assert proc_res is True
    assert output_pdf.exists()

    # 3. 模拟成功后安全清理输入文件
    if input_pdf.resolve() != output_pdf.resolve():
        input_pdf.unlink()
    assert not input_pdf.exists()

