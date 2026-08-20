"""单元测试：17_PDF_Title_Renamer 标题提取与重命名规范。"""

from __future__ import annotations

import sys
from pathlib import Path
import pytest

# 导入目标模块
MODULE_DIR = Path(__file__).resolve().parent.parent / "17_PDF_Title_Renamer"
if str(MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(MODULE_DIR))

from pdf_sanitizer import PDFSanitizer  # noqa: E402


def test_smart_title_case_scientific_tokens() -> None:
    """测试科学缩写与大小写格式化。"""
    raw = "optimal approaches for cpg and mrna vaccines in phase IIa trials"
    simplified = PDFSanitizer._simplify_filename(raw)
    assert "Optimal_Approaches_for_CpG_and_mRNA_Vaccines_in_Phase_IIa_Trials" in simplified


def test_smart_colon_and_clinical_phase() -> None:
    """测试冒号处理：副标题包含 clinical/phase 时保留，否则截断。"""
    # 包含 phase/trial 保留
    title_with_phase = "A Novel Adjuvant: A Phase I/IIa Randomized Trial"
    simplified1 = PDFSanitizer._simplify_filename(title_with_phase)
    assert "Phase_I_IIa" in simplified1
    assert "Randomized_Trial" in simplified1

    # 普通副标题截断
    title_normal = "Vaccine Development: An Overview of the Last Century"
    simplified2 = PDFSanitizer._simplify_filename(title_normal)
    assert simplified2 == "Vaccine_Development"

    # 中文冒号硬截断
    title_cn = "重组新型冠状病毒疫苗：临床安全性评估与观察"
    simplified_cn = PDFSanitizer._simplify_filename(title_cn)
    assert simplified_cn == "重组新型冠状病毒疫苗"


def test_smart_bracket_removal() -> None:
    """测试括号处理：保留专业术语内容，删除无意义纯符号。"""
    raw = "Expression of Antigen in (Pichia pastoris) System [10.1016/j.vaccine]"
    simplified = PDFSanitizer._simplify_filename(raw)
    assert "Pichia_Pastoris" in simplified
    assert "[" not in simplified
    assert "]" not in simplified


def test_masthead_and_publisher_noise_filtering() -> None:
    """测试期刊名、待刊横幅、出版商元数据黑名单过滤。"""
    assert PDFSanitizer._is_journal_masthead_only("Vaccine")
    assert PDFSanitizer._is_journal_masthead_only("Nature Communications")
    assert PDFSanitizer._is_journal_masthead_only("PLOS ONE")
    assert PDFSanitizer._is_journal_masthead_only("Vaccines")
    assert PDFSanitizer._is_journal_masthead_only("Cells")

    # 普通正文标题不可被误判为 masthead
    assert not PDFSanitizer._is_journal_masthead_only("Optimal approaches to data collection")
    assert not PDFSanitizer._is_journal_masthead_only("Safety and immunogenicity of novel adjuvants")

    # 待刊横幅与出版商角色
    assert PDFSanitizer._is_publisher_status_line("ARTICLE IN PRESS")
    assert PDFSanitizer._is_publisher_status_line("Accepted Manuscript")
    assert PDFSanitizer._is_publisher_metadata_only("Academic Editor")
    assert PDFSanitizer._is_publisher_metadata_only("Author Contributions")


def test_volume_and_boilerplate_filtering() -> None:
    """测试卷期页码与出版商链接过滤。"""
    assert PDFSanitizer._is_volume_header_line("Vaccine 31 (2013) 1870–1876")
    assert PDFSanitizer._is_volume_header_line("31 (2013) 1870– 1876")
    assert PDFSanitizer._is_boilerplate_line("Contents lists available at SciVerse ScienceDirect")
    assert PDFSanitizer._is_boilerplate_line("journal homepage: www.elsevier.com/locate/vaccine")
    assert PDFSanitizer._is_boilerplate_line("j ourna l ho me pag e: www.elsevier.com/locate/vaccine")
    assert PDFSanitizer._is_boilerplate_line("a r t i c l e i n f o")


def test_target_pdf_extraction_mock(tmp_path: Path) -> None:
    """使用动态构建的 PDF 文档测试端到端标题提取与年份验证。"""
    try:
        import fitz
    except ImportError:
        pytest.skip("PyMuPDF (fitz) is not installed")

    doc = fitz.open()
    doc.set_metadata(
        {
            "title": (
                "Optimal approaches to data collection and analysis of potential "
                "immune mediated disorders in clinical trials of new vaccines"
            ),
            "creationDate": "D:20130214015733Z",
        }
    )
    page = doc.new_page()
    page.insert_text((50, 50), "Vaccine 31 (2013) 1870-1876", fontsize=6.4)
    page.insert_text((50, 100), "Contents lists available at SciVerse ScienceDirect", fontsize=8.0)
    page.insert_text(
        (50, 150),
        "Optimal approaches to data collection and analysis of potential "
        "immune mediated disorders in clinical trials of new vaccines",
        fontsize=13.5,
    )
    page.insert_text((50, 200), "Fernanda Tavares Da Silva, Filip De Keyser", fontsize=10.5)

    pdf_file = tmp_path / "sample.pdf"
    doc.save(pdf_file)
    doc.close()

    sanitizer = PDFSanitizer(output_dir=str(tmp_path / "out"))
    raw_title, year = sanitizer._scan_payload(pdf_file)

    assert "Optimal approaches to data collection and analysis" in raw_title
    assert "immune mediated disorders" in raw_title
    assert year == "2013"

    simplified = sanitizer._simplify_filename(raw_title)
    assert simplified.startswith("Optimal_Approaches_to_Data_Collection_and_Analysis_of_Potential_Immune")


def test_dedupe_filename(tmp_path: Path) -> None:
    """测试文件名防重名递增编号。"""
    sanitizer = PDFSanitizer(output_dir=str(tmp_path))
    (tmp_path / "Document_Name-2023.pdf").touch()
    name1 = sanitizer._dedupe_filename(tmp_path, "Document_Name-2023")
    assert name1 == "Document_Name-2023_1"
    (tmp_path / "Document_Name-2023_1.pdf").touch()
    name2 = sanitizer._dedupe_filename(tmp_path, "Document_Name-2023")
    assert name2 == "Document_Name-2023_2"
