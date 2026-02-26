# -*- coding: utf-8 -*-
"""
从 PDF 中按关键词/页码检索文本或表格，供写入 Excel 使用。
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

import pdfplumber

logger = logging.getLogger(__name__)


def extract_text_from_pdf(
    pdf_path: str | Path,
    page_numbers: Optional[list[int]] = None,
) -> dict[int, str]:
    """
    提取 PDF 指定页的文本。不指定页码则提取全部页。
    
    Args:
        pdf_path: PDF 文件路径
        page_numbers: 要提取的页码列表（从1开始），None 表示提取全部页
        
    Returns:
        { 页码: 该页全文 } 字典
        
    Raises:
        FileNotFoundError: PDF 文件不存在
        ValueError: PDF 文件无法打开或读取
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")
    
    if not pdf_path.is_file():
        raise ValueError(f"路径不是文件: {pdf_path}")

    result = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                logger.warning(f"PDF 文件没有页面: {pdf_path}")
                return result
                
            pages = page_numbers if page_numbers is not None else range(1, len(pdf.pages) + 1)
            for pno in pages:
                if pno < 1 or pno > len(pdf.pages):
                    logger.warning(f"页码 {pno} 超出范围 (1-{len(pdf.pages)})，跳过")
                    continue
                try:
                    page = pdf.pages[pno - 1]
                    text = page.extract_text() or ""
                    result[pno] = text
                except Exception as e:
                    logger.error(f"提取第 {pno} 页文本时出错: {e}")
                    result[pno] = ""
    except Exception as e:
        raise ValueError(f"无法打开或读取 PDF 文件 {pdf_path}: {e}") from e
        
    return result


def search_by_keyword(
    pdf_path: str | Path,
    keyword: str,
    page_number: Optional[int] = None,
    context_chars: int = 500,
) -> list[tuple[int, str]]:
    """
    在 PDF 中按关键词检索，返回 (页码, 匹配到的段落/上下文) 列表。
    
    Args:
        pdf_path: PDF 文件路径
        keyword: 要搜索的关键词
        page_number: 指定页码（从1开始），None 表示搜索全部页
        context_chars: 上下文字符数
        
    Returns:
        (页码, 匹配文本) 列表
    """
    if not keyword or not keyword.strip():
        logger.warning("关键词为空，返回空结果")
        return []
    
    if context_chars < 0:
        context_chars = 500
        logger.warning(f"context_chars 为负数，重置为 {context_chars}")
    
    pdf_path = Path(pdf_path)
    try:
        pages_text = extract_text_from_pdf(
            pdf_path,
            [page_number] if page_number is not None else None,
        )
    except Exception as e:
        logger.error(f"提取 PDF 文本失败: {e}")
        return []
    
    hits = []
    keyword = keyword.strip()
    for pno, text in pages_text.items():
        if not text or keyword not in text:
            continue
        # 按行或块取包含关键词的一段
        parts = text.replace("\n", " ").split(keyword)
        for i, part in enumerate(parts):
            if i == 0:
                continue
            start = max(0, len(part) - context_chars)
            end = min(len(part) + context_chars, len(part) + len(keyword) + context_chars)
            next_part = parts[i + 1][:context_chars] if i + 1 < len(parts) else ""
            segment = (part[start:] + keyword + next_part)[: context_chars * 2]
            hits.append((pno, segment.strip()))
    
    if not hits:
        # 整页作为后备
        for pno, text in pages_text.items():
            if keyword in text:
                hits.append((pno, text[: 2000]))
                break
    
    return hits


def extract_tables_from_pdf(
    pdf_path: str | Path,
    page_number: Optional[int] = None,
    near_keyword: Optional[str] = None,
) -> list[list[list[str]]]:
    """
    从 PDF 中提取表格。
    
    Args:
        pdf_path: PDF 文件路径
        page_number: 只在该页取表；不指定则所有页
        near_keyword: 若指定，只保留在含该关键词的页上的表格
        
    Returns:
        [ 页1表格列表, 页2表格列表, ... ]，每页表格为 list[list[str]]（行→列）
        
    Raises:
        FileNotFoundError: PDF 文件不存在
        ValueError: PDF 文件无法打开或读取
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")
    
    if not pdf_path.is_file():
        raise ValueError(f"路径不是文件: {pdf_path}")

    all_tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                logger.warning(f"PDF 文件没有页面: {pdf_path}")
                return all_tables
                
            pages = [page_number] if page_number is not None else range(1, len(pdf.pages) + 1)
            for pno in pages:
                if pno < 1 or pno > len(pdf.pages):
                    logger.warning(f"页码 {pno} 超出范围 (1-{len(pdf.pages)})，跳过")
                    continue
                try:
                    page = pdf.pages[pno - 1]
                    if near_keyword:
                        text = page.extract_text() or ""
                        if near_keyword not in text:
                            continue
                    tables = page.extract_tables()
                    if tables:
                        all_tables.append(tables)
                except Exception as e:
                    logger.error(f"提取第 {pno} 页表格时出错: {e}")
                    continue
    except Exception as e:
        raise ValueError(f"无法打开或读取 PDF 文件 {pdf_path}: {e}") from e
        
    return all_tables


def get_first_table_near_keyword(
    pdf_path: str | Path,
    keyword: str,
    page_number: Optional[int] = None,
) -> Optional[list[list[str]]]:
    """
    获取在关键词所在页的第一个表格；若某表内包含关键词则优先返回该表。
    
    Args:
        pdf_path: PDF 文件路径
        keyword: 关键词
        page_number: 指定页码（从1开始），None 表示搜索全部页
        
    Returns:
        第一个匹配的表格（list[list[str]]），未找到则返回 None
    """
    if not keyword or not keyword.strip():
        logger.warning("关键词为空，返回 None")
        return None
    
    try:
        tables_by_page = extract_tables_from_pdf(pdf_path, page_number=page_number, near_keyword=keyword)
    except Exception as e:
        logger.error(f"提取表格失败: {e}")
        return None
    
    keyword = keyword.strip()
    for page_tables in tables_by_page:
        # 优先返回包含关键词的表格
        for table in page_tables:
            if not table:
                continue
            if any(any(cell and keyword in str(cell) for cell in row) for row in table):
                return table
        # 否则返回第一个非空表格
        for table in page_tables:
            if table:
                return table
    return None
