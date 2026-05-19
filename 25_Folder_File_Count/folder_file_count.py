# -*- coding: utf-8 -*-
"""
文件夹文件总数统计（25 模块）

功能：递归统计指定目录下所有文件数量，并输出 TXT 与 Excel。

默认目录：
- 输出：25_Folder_File_Count/output/

用法：
  python folder_file_count.py --path "D:\\data"
  python folder_file_count.py --path "D:\\data" --output "output"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd


def normalize_path(value: str) -> Path:
    cleaned = value.strip().strip('"').strip("'")
    return Path(cleaned)


def count_files(root: Path) -> int:

    return sum(1 for p in root.rglob("*") if p.is_file())


def count_files_by_subdir(root: Path) -> list[dict[str, int | str]]:
    rows: list[dict[str, int | str]] = []
    for subdir in sorted([p for p in root.iterdir() if p.is_dir()]):
        count = sum(1 for p in subdir.rglob("*") if p.is_file())
        rows.append({"path": str(subdir), "file_count": count})
    return rows


def build_tree_summary(root: Path) -> tuple[list[dict[str, int | str]], list[str]]:
    rows: list[dict[str, int | str]] = []
    lines: list[str] = []

    def walk(dir_path: Path, depth: int) -> None:
        try:
            entries = sorted(dir_path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
        except Exception:
            entries = []

        indent = "  " * depth
        lines.append(f"{indent}- {dir_path.name}/")
        rows.append({"path": str(dir_path), "level": depth, "type": "dir"})

        for entry in entries:
            if entry.is_dir():
                walk(entry, depth + 1)
            elif entry.is_file():
                lines.append(f"{indent}  - {entry.name}")
                rows.append({"path": str(entry), "level": depth + 1, "type": "file"})

    walk(root, 0)
    return rows, lines






def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计目录内所有文件数量")
    parser.add_argument("--path", default=None, help="待统计的文件夹路径")
    parser.add_argument("--output", default=None, help="输出目录（默认 output/）")

    return parser.parse_args()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    args = parse_args()
    if args.path:
        target = normalize_path(args.path)
        if not target.is_absolute():
            target = Path.cwd() / target
    else:
        if not sys.stdin.isatty():
            print("未提供 --path，且当前为非交互环境。请使用 --path 指定目录。", file=sys.stderr)
            sys.exit(1)
        user_input = input("请输入要统计的文件夹路径：").strip()
        if not user_input:
            raise ValueError("未输入文件夹路径")
        target = normalize_path(user_input)
        if not target.is_absolute():
            target = Path.cwd() / target


    if not target.exists() or not target.is_dir():
        if sys.stdin.isatty():
            for _ in range(3):
                retry = input("路径不存在或不是文件夹，请重新输入：").strip()
                if not retry:
                    continue
                target = normalize_path(retry)
                if not target.is_absolute():
                    target = Path.cwd() / target
                if target.exists() and target.is_dir():
                    break
        if not target.exists() or not target.is_dir():
            print(f"路径不存在或不是文件夹：{target}")
            print("请确认盘符已挂载、路径无多余空格，并使用完整路径。", file=sys.stderr)
            sys.exit(1)




    out_dir = Path(args.output) if args.output else output_dir
    if not out_dir.is_absolute():
        out_dir = output_dir / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    total = count_files(target)
    subdir_rows = count_files_by_subdir(target)
    tree_rows, tree_lines = build_tree_summary(target)
    file_rows = [row for row in tree_rows if row.get("type") == "file"]
    tree_display_rows = [
        {"level": row["level"], "tree": line}
        for row, line in zip(tree_rows, tree_lines)
    ]





    txt_path = out_dir / "folder_file_count.txt"
    with txt_path.open("w", encoding="utf-8") as f:
        f.write(f"path: {target}\n")
        f.write(f"file_count: {total}\n")
        if subdir_rows:
            f.write("subdir_counts:\n")
            for row in subdir_rows:
                f.write(f"- {row['path']}: {row['file_count']}\n")
        if tree_lines:
            f.write("tree_summary:\n")
            for line in tree_lines:
                f.write(f"{line}\n")
        if file_rows:
            f.write("files:\n")
            for row in file_rows:
                f.write(f"- {Path(str(row['path'])).name}\n")

    df_total = pd.DataFrame([{"path": str(target), "file_count": total}])
    df_subdirs = pd.DataFrame(subdir_rows)
    df_tree = pd.DataFrame(tree_display_rows)


    xlsx_path = out_dir / "folder_file_count.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_total.to_excel(writer, index=False, sheet_name="total")
        df_subdirs.to_excel(writer, index=False, sheet_name="subdirs")
        df_tree.to_excel(writer, index=False, sheet_name="tree")



    print(f"已输出：{txt_path}")
    print(f"已输出：{xlsx_path}")


if __name__ == "__main__":
    main()
