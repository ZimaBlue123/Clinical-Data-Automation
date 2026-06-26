# -*- coding: utf-8 -*-
"""
Word（.doc/.docx/.rtf）-> Excel（.xlsx）

与 08_Word_Tables_to_Excel 复用同一套“COM 读取表格 + openpyxl 写出（带样式）”逻辑，
本模块做批处理：把 input/ 下所有 Word 文件的“全部顶层表格”导出为多个 sheet。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import re
import sys
import traceback
from pathlib import Path
from typing import Optional


MODULE_DIR = Path(__file__).resolve().parent
# 本模块（09_Word_Tables_to_Excel）合并了原 08 的导出逻辑，直接从同目录导入
from word_tables_to_excel import (  # noqa: E402
    export_word_tables_to_excel,
)


def _natural_key(s: str):
    parts = re.split(r"(\d+)", s)
    out = []
    for p in parts:
        if not p:
            continue
        if p.isdigit():
            out.append(int(p))
        else:
            out.append(p.lower())
    return out


def _parse_int_list(s: str) -> list[int]:
    s = (s or "").strip()
    if not s:
        return []
    parts = re.split(r"[,\s]+", s)
    out: list[int] = []
    for p in parts:
        if not p:
            continue
        out.append(int(p))
    return out


def _split_keywords(s: str) -> list[str]:
    s = (s or "").strip()
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _iter_word_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在：{input_dir}")

    word_suffixes = {".doc", ".docx", ".rtf"}
    out: list[Path] = []
    for p in input_dir.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in word_suffixes:
            continue
        # Word 临时锁文件：~$xxxx.docx
        if p.name.startswith("~$"):
            continue
        out.append(p)

    if not out:
        raise FileNotFoundError(f"输入目录未找到 Word/RTF：{input_dir}")

    out.sort(key=lambda x: _natural_key(x.name))
    return out


def _maybe_backup_output(output_path: Path, *, backup_existing: bool) -> None:
    if not output_path.exists():
        return
    if not backup_existing:
        output_path.unlink(missing_ok=True)
        return

    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = output_path.with_name(
        f"{output_path.stem}.bak_{ts}{output_path.suffix}"
    )
    try:
        output_path.rename(backup_path)
    except Exception:
        # 备份失败时仍尝试删除原文件（避免阻塞后续写入）
        output_path.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="批量：Word 全部顶层表格 -> Excel（多 sheet）"
    )
    parser.add_argument(
        "--input-dir",
        "-i",
        default=str(MODULE_DIR / "input"),
        help="输入目录（默认本模块 input/）",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        default=str(MODULE_DIR / "output"),
        help="输出目录（默认本模块 output/）",
    )
    parser.add_argument(
        "--header-rows", type=int, default=1, help="表头占用行数（默认 1）"
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="只统计并打印行列，不写 xlsx"
    )
    parser.add_argument(
        "--quiet", action="store_true", help="减少导出过程输出到 stderr"
    )

    parser.add_argument(
        "--skip-existing", action="store_true", help="output 已存在则跳过该文件"
    )
    parser.add_argument(
        "--no-backup-existing",
        action="store_true",
        help="覆盖前不备份旧输出（默认启用备份）",
    )
    parser.add_argument(
        "--overwrite", action="store_true", help="直接覆盖输出（不备份）"
    )
    parser.add_argument("--fail-fast", action="store_true", help="遇到错误立即中止")
    parser.add_argument(
        "--max-files", type=int, default=None, help="最多处理多少个文件（调试用）"
    )

    # 可选筛表（默认导出全部顶层表格）
    parser.add_argument(
        "--table-indices", default=None, help="导出指定表序号（1-based），如 1,3,5"
    )
    parser.add_argument(
        "--table-index", type=int, default=None, help="导出单个表序号（1-based）"
    )
    parser.add_argument(
        "--header-keywords", default=None, help="按表头关键字筛选（逗号分隔）"
    )
    parser.add_argument(
        "--merge-tables-from",
        type=int,
        default=None,
        help="多段顶层表纵向合并：从第 N 个开始（启用后会仅导出合并区间结果）",
    )
    parser.add_argument(
        "--merge-tables-to",
        type=int,
        default=None,
        help="多段顶层表纵向合并：合并到第 M 个",
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # 覆盖/备份策略
    backup_existing = not bool(args.no_backup_existing) and not bool(args.overwrite)

    table_indices: Optional[list[int]] = None
    if args.table_indices:
        table_indices = _parse_int_list(args.table_indices)
    if args.table_index is not None:
        table_indices = [int(args.table_index)]

    header_keywords: Optional[list[str]] = None
    if args.header_keywords:
        ks = _split_keywords(args.header_keywords)
        header_keywords = ks if ks else None

    merge_from = args.merge_tables_from
    merge_to = args.merge_tables_to

    word_files = _iter_word_files(input_dir)
    if args.max_files is not None:
        word_files = word_files[: max(0, int(args.max_files))]

    ok_cnt = 0
    fail_cnt = 0

    for idx, wp in enumerate(word_files, 1):
        out_name = f"{wp.stem}_all_tables.xlsx"
        out_path = output_dir / out_name

        if out_path.exists():
            if args.skip_existing:
                print(f"[{idx}/{len(word_files)}] 跳过已存在：{out_path.name}")
                continue
            if not args.overwrite and not args.dry_run:
                _maybe_backup_output(out_path, backup_existing=backup_existing)

        try:
            export_word_tables_to_excel(
                input_path=wp,
                output_path=out_path,
                table_indices=table_indices,
                header_keywords=header_keywords,
                header_rows=int(args.header_rows),
                table_title=None,
                merge_tables_from=merge_from,
                merge_tables_to=merge_to,
                quiet=bool(args.quiet),
                dry_run=bool(args.dry_run),
            )
            ok_cnt += 1
            if args.dry_run:
                print(f"[{idx}/{len(word_files)}] dry-run：{wp.name}（不写盘）")
            else:
                print(f"[{idx}/{len(word_files)}] 完成：{wp.name} -> {out_path.name}")
        except Exception as e:
            fail_cnt += 1
            print(f"[{idx}/{len(word_files)}] 失败：{wp.name}：{e}", file=sys.stderr)
            if args.fail_fast:
                raise
            if not args.quiet:
                traceback.print_exc()

    print(f"批处理结束：成功 {ok_cnt} 个，失败 {fail_cnt} 个。输出目录：{output_dir}")


if __name__ == "__main__":
    main()
