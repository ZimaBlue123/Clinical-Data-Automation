"""
replace_docx.py — Word docx 批量文本替换（11_Word_Text_Replace）

内置规则示例：日期占位符、研究编号等；规则定义见 lib/ooxml_replace.py。
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
import traceback
from typing import List

from lib.ooxml_replace import (
    build_date_rules,
    build_default_rules,
    build_study_id_rules,
    process_docx,
    unique_output_path,
)

MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT = os.path.join(MODULE_DIR, "input")
DEFAULT_OUTPUT = os.path.join(MODULE_DIR, "output")
OUT_SUFFIX = "_updated"


def _tk_messagebox():
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        return messagebox
    except Exception:
        return None


def confirm(question: str, auto_yes: bool) -> bool:
    if auto_yes:
        return True
    mb = _tk_messagebox()
    if mb is not None:
        return bool(mb.askyesno("确认", question))
    ans = input(question + " (y/n): ").strip().lower()
    return ans in {"y", "yes"}


def info(message: str, auto_yes: bool = False) -> None:
    if auto_yes:
        print(message)
        return
    mb = _tk_messagebox()
    if mb is not None:
        mb.showinfo("提示", message)
    else:
        print(message)


def collect_docx(input_dir: str, recursive: bool) -> List[str]:
    pattern = (
        os.path.join(input_dir, "**", "*.docx")
        if recursive
        else os.path.join(input_dir, "*.docx")
    )
    files = glob.glob(pattern, recursive=recursive)
    out = []
    for f in files:
        name = os.path.basename(f)
        if name.startswith("~$"):
            continue
        if OUT_SUFFIX in os.path.splitext(name)[0]:
            continue
        out.append(f)
    return sorted(out)


def rules_description(only_dates: bool, only_study: bool) -> str:
    lines = []
    if only_study and not only_dates:
        lines.append("仅研究编号：YDSWX（TVAX-009）-004（Ⅳ）→ …004（III）")
    elif only_dates and not only_study:
        lines.append("仅日期占位符 → 2026/05/27 或 2026年05月27日")
    else:
        lines.append("日期占位符 + 研究编号（Ⅳ→III）")
    return "\n".join(lines)


def pick_rules(only_dates: bool, only_study: bool):
    if only_dates and only_study:
        raise ValueError("不能同时指定 --only-dates 与 --only-study-id")
    if only_dates:
        return build_date_rules()
    if only_study:
        return build_study_id_rules()
    return build_default_rules()


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Word docx 批量文本替换")
    p.add_argument("--input", default=DEFAULT_INPUT)
    p.add_argument("--output", default=DEFAULT_OUTPUT)
    p.add_argument("--recursive", action="store_true")
    p.add_argument("--yes", action="store_true")
    p.add_argument("--only-dates", action="store_true", help="仅替换日期占位符")
    p.add_argument("--only-study-id", action="store_true", help="仅替换研究编号 Ⅳ→III")
    return p.parse_args(argv)


def main(argv=None) -> None:
    args = parse_args(argv)
    auto_yes = args.yes or ("--yes" in sys.argv)

    try:
        rules = pick_rules(args.only_dates, args.only_study_id)
    except ValueError as e:
        info(str(e), auto_yes)
        return

    files = collect_docx(args.input, args.recursive)
    if not files:
        info(f"未找到 Word 文档：{args.input}", auto_yes)
        return

    if not confirm(
        f"替换规则：\n{rules_description(args.only_dates, args.only_study_id)}\n\n"
        f"输入：{args.input}\n输出：{args.output}\n\n继续吗？",
        auto_yes,
    ):
        return

    summary = []
    for path in files:
        if not confirm(f"处理：\n{os.path.basename(path)}\n\n开始吗？", auto_yes):
            continue
        try:
            out_path = unique_output_path(path, args.output, OUT_SUFFIX)
            stats = process_docx(path, out_path, rules)
            summary.append((path, out_path, stats))
            msg = (
                f"完成：{os.path.basename(path)}\n"
                f"输出：{os.path.basename(out_path)}\n"
                f"匹配（替换前）：{stats.before} 处\n"
                f"已替换：{stats.replaced} 处\n"
                f"剩余匹配：{stats.after} 处（期望 0）"
            )
            if not confirm(msg + "\n\n继续下一个？", auto_yes):
                break
        except Exception as e:
            info(f"失败：{os.path.basename(path)}\n{e}\n\n{traceback.format_exc()}", auto_yes)

    if summary:
        total_rep = sum(s.replaced for _, _, s in summary)
        total_left = sum(s.after for _, _, s in summary)
        info(
            f"全部完成。\n总替换：{total_rep} 处\n总剩余匹配：{total_left} 处",
            auto_yes,
        )
    else:
        info("未处理任何文件。", auto_yes)


if __name__ == "__main__":
    main()
