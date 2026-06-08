import logging
import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd
import pdfplumber
import pytesseract
import docx
from pdf2image import convert_from_path

from settings import Settings


def configure_ocr(settings: Settings) -> None:
    if settings.tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = settings.tesseract_cmd


def extract_pdf_text(pdf_path: str, settings: Settings, ocr_lang: str = "chi_sim+eng") -> str:
    extracted_text: list[str] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text or len(text.strip()) < 50:
                    logging.info(f"PDF页面 {page_num + 1} 文本层缺失/极少，触发 OCR 补偿...")
                    text = _perform_ocr(pdf_path, page_num, settings=settings, lang=ocr_lang)
                if text:
                    extracted_text.append(text)
    except Exception as e:
        logging.error("PDF 打开或读取失败 [%s]: %s", os.path.basename(pdf_path), e)
        return ""
    return "\n".join(extracted_text)


def _perform_ocr(pdf_path: str, page_index: int, settings: Settings, lang: str) -> str:
    try:
        kwargs = {
            "first_page": page_index + 1,
            "last_page": page_index + 1,
        }
        if settings.poppler_path:
            kwargs["poppler_path"] = settings.poppler_path

        images = convert_from_path(pdf_path, **kwargs)
        if images:
            return pytesseract.image_to_string(images[0], lang=lang)
        return ""
    except Exception as e:
        logging.error(f"OCR 引擎异常: {e}")
        return ""


def parse_txt(file_path: str) -> str:
    for encoding in ["utf-8", "utf-8-sig", "gbk", "utf-16"]:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
        except OSError as e:
            raise ValueError(f"读取文本失败: {e}") from e
    raise ValueError("无法识别的文本编码格式")


def parse_docx_fast(file_path: str) -> str:
    text_blocks: list[str] = []
    try:
        with zipfile.ZipFile(file_path) as docx_zip:
            xml_content = docx_zip.read("word/document.xml")

        tree = ET.fromstring(xml_content)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        for paragraph in tree.findall(".//w:p", namespaces=ns):
            texts = [node.text for node in paragraph.findall(".//w:t", namespaces=ns) if node.text]
            if texts:
                text_blocks.append("".join(texts))

        logging.info(f"DOCX 极速模式解析完成: {os.path.basename(file_path)}")
        return "\n".join(text_blocks)
    except Exception as e:
        logging.warning(f"极速模式解析失败，回退至标准模式 [{os.path.basename(file_path)}]: {e}")
        return parse_docx_standard(file_path)


def parse_docx_standard(file_path: str) -> str:
    doc = docx.Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])


def parse_excel(file_path: str) -> str:
    try:
        sheets_dict = pd.read_excel(file_path, sheet_name=None)
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {e}") from e

    text_blocks: list[str] = []
    for sheet_name, df in sheets_dict.items():
        df = df.copy()
        df.dropna(how="all", inplace=True)
        df.dropna(axis=1, how="all", inplace=True)
        if df.empty:
            continue
        text_blocks.append(f"--- 表格: {sheet_name} ---")
        text_blocks.append(df.to_string(index=False))
    if not text_blocks:
        return ""
    return "\n\n".join(text_blocks)


def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
