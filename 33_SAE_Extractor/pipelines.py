import logging
import os
import re
from dataclasses import dataclass
from pathlib import Path
from collections.abc import Iterable

import pandas as pd

from parsers import (
    extract_pdf_text,
    parse_docx_fast,
    parse_docx_standard,
    parse_excel,
    parse_txt,
    ensure_parent_dir,
)
from sae_extractor import ClinicalDataGuard
from settings import Settings


DEFAULT_KEYWORDS: list[str] = [
    r"不良事件",
    r"SAE",
    r"严重不良",
    r"CTCAE",
    r"因果关系",
    r"转归",
    r"停药",
    r"死亡",
    r"住院",
    r"致残",
    r"后遗症",
    r"相关性",
    r"异常",
]


@dataclass(frozen=True)
class PipelineIO:
    input_dir: Path
    output_file: Path


def truncate_by_keywords(
    raw_text: str,
    *,
    keywords: Iterable[str] = DEFAULT_KEYWORDS,
    min_len: int = 2000,
    window: int = 900,
    hard_cap: int = 6000,
) -> str:  # noqa: E501
    total_len = len(raw_text)
    if total_len < min_len:
        return raw_text

    pattern = re.compile("|".join(list(keywords)), re.IGNORECASE)
    matches = [m.span() for m in pattern.finditer(raw_text)]
    if not matches:
        logging.info("未命中 SAE 关键词锚点，执行保守截断（前 %s 字符）。", min_len)
        return raw_text[:min_len]

    intervals: list[list[int]] = [[max(0, s - window), min(total_len, e + window)] for s, e in matches]
    intervals.sort(key=lambda x: x[0])

    merged: list[list[int]] = [intervals[0]]
    for cur in intervals[1:]:
        prev = merged[-1]
        if cur[0] <= prev[1]:
            prev[1] = max(prev[1], cur[1])
        else:
            merged.append(cur)

    snippets = [raw_text[s:e] for s, e in merged]
    final_text = "\n\n...[已截断]...\n\n".join(snippets)
    if len(final_text) > hard_cap:
        final_text = final_text[:hard_cap]

    logging.info("降噪压缩: %s 字符 -> %s 字符", total_len, len(final_text))
    return final_text


def export_records_to_excel(records: list[dict], output_file: Path) -> None:
    if not records:
        logging.warning("数据池为空，中止导出。")
        return

    ensure_parent_dir(output_file)
    df = pd.DataFrame(records)
    cols = df.columns.tolist()
    if "source_file" in cols:
        cols.insert(0, cols.pop(cols.index("source_file")))
        df = df[cols]

    df.to_excel(str(output_file), index=False, engine="openpyxl")
    logging.info("批处理完成，共生成 %s 条记录。输出至: %s", len(df), output_file)


def _list_files(input_dir: Path, exts: tuple[str, ...]) -> list[Path]:
    if not input_dir.exists():
        input_dir.mkdir(parents=True, exist_ok=True)
        logging.warning("目录 %s 已创建，请存入文档后重试。", input_dir)
        return []
    files = [
        input_dir / name
        for name in os.listdir(input_dir)
        if name.lower().endswith(exts) and not name.startswith("~$") and (input_dir / name).is_file()
    ]
    files.sort(key=lambda p: p.name.lower())
    return files


def _extract_text_for_file(path: Path, *, settings: Settings, fast_docx: bool) -> str:
    ext = path.suffix.lower()
    if ext == ".txt":
        return parse_txt(str(path))
    if ext == ".pdf":
        return extract_pdf_text(str(path), settings=settings)
    if ext in [".docx", ".doc"]:
        return parse_docx_fast(str(path)) if fast_docx else parse_docx_standard(str(path))
    if ext in [".xlsx", ".xls"]:
        return parse_excel(str(path))
    return ""


def run_omni_batch(*, engine: ClinicalDataGuard, io: PipelineIO, settings: Settings, fast_docx: bool = False) -> None:
    files = _list_files(io.input_dir, (".txt", ".pdf", ".docx", ".xlsx", ".xls"))
    if not files:
        logging.warning("未发现支持的文档格式，流水线终止。")
        return

    logging.info("共发现 %s 份多源文件，启动全能提取引擎。", len(files))
    records: list[dict] = []

    for idx, path in enumerate(files, start=1):
        logging.info("正在处理 [%s/%s]: %s", idx, len(files), path.name)
        try:
            raw = _extract_text_for_file(path, settings=settings, fast_docx=fast_docx)
        except Exception as e:
            logging.error("物理层解析失败 [%s]: %s", path.name, e)
            continue

        if not raw.strip():
            logging.error("[%s] 数据剥离为空，已跳过。", path.name)
            continue

        cleaned = truncate_by_keywords(raw)
        extracted = engine.extract(cleaned)
        if extracted:
            extracted["source_file"] = path.name
            records.append(extracted)
        else:
            logging.warning("[%s] 网关解析未返回有效结构。", path.name)

    export_records_to_excel(records, io.output_file)


def run_pdf_batch(*, engine: ClinicalDataGuard, io: PipelineIO, settings: Settings) -> None:
    files = _list_files(io.input_dir, (".pdf",))
    if not files:
        logging.warning("目录 %s 中未发现 .pdf 文件，流水线终止。", io.input_dir)
        return

    logging.info("检测到 %s 份 PDF 记录，流水线启动。", len(files))
    records: list[dict] = []

    for idx, path in enumerate(files, start=1):
        logging.info("正在处理 [%s/%s]: %s", idx, len(files), path.name)
        try:
            raw = extract_pdf_text(str(path), settings=settings)
        except Exception as e:
            logging.error("PDF 解析失败 [%s]: %s", path.name, e)
            continue

        if not raw.strip():
            logging.error("[%s] 文本提取为空，跳过结构化抽取。", path.name)
            continue

        cleaned = truncate_by_keywords(raw)
        extracted = engine.extract(cleaned)
        if extracted:
            extracted["source_file"] = path.name
            records.append(extracted)
        else:
            logging.warning("[%s] 提取失败或网关返回无效结构。", path.name)

    export_records_to_excel(records, io.output_file)
