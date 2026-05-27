# -*- coding: utf-8 -*-
"""
PDF 跨文件夹重复分析（20 模块）

在同一根目录下的多个子文件夹之间，检测 PDF 重复：
- 文件名相同（不区分大小写）
- 或第一页提取文本相同

无 input/ 目录：通过 --root 指定待扫描的外部路径；报告写入 output/。

用法（单次任务）：
  python pdf_duplicate_analyzer.py --root "D:\\References" --folders "A,B,C" --label "批次1"

用法（配置文件，多任务）：
  python pdf_duplicate_analyzer.py --config jobs.example.json
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import fitz

MODULE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = MODULE_DIR / "output"


def normalize_path(value: str) -> Path:
    cleaned = value.strip().strip('"').strip("'")
    return Path(cleaned)


def normalize_text(text: str) -> str:
    return " ".join(text.split())


def extract_first_page_text(pdf_path: Path) -> tuple[str, str | None]:
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count == 0:
            doc.close()
            return "", "empty PDF"
        text = normalize_text(doc[0].get_text("text"))
        doc.close()
        return text, None
    except Exception as exc:  # noqa: BLE001
        return "", str(exc)


def resolve_folders(root: Path, folders: list[str] | None) -> list[str]:
    if folders:
        return folders
    if not root.is_dir():
        return []
    return sorted(p.name for p in root.iterdir() if p.is_dir())


def collect_pdfs(root: Path, folder_names: list[str]) -> list[dict]:
    records: list[dict] = []
    for folder in folder_names:
        folder_path = root / folder
        if not folder_path.is_dir():
            print(f"WARNING: missing folder {folder_path}", file=sys.stderr)
            continue
        for pdf in sorted(folder_path.rglob("*.pdf")):
            if not pdf.is_file():
                continue
            text, err = extract_first_page_text(pdf)
            records.append(
                {
                    "folder": folder,
                    "name": pdf.name,
                    "path": pdf,
                    "page1_text": text,
                    "page1_hash": hashlib.md5(text.encode("utf-8")).hexdigest()
                    if text
                    else "",
                    "error": err,
                }
            )
    return records


def build_duplicate_groups(records: list[dict]) -> tuple[dict, dict]:
    by_name: dict[str, list[dict]] = defaultdict(list)
    by_content: dict[str, list[dict]] = defaultdict(list)
    for rec in records:
        by_name[rec["name"].lower()].append(rec)
        if rec["page1_text"] and not rec["error"]:
            by_content[rec["page1_hash"]].append(rec)
    name_dup = {k: v for k, v in by_name.items() if len(v) > 1}
    content_dup = {k: v for k, v in by_content.items() if len(v) > 1}
    return name_dup, content_dup


def directed_pair_key(source: dict, other: dict, reason: str) -> tuple:
    return (source["folder"], source["name"], other["folder"], other["name"], reason)


def find_cross_folder_dupes(records: list[dict]) -> list[dict]:
    """All duplicate relationships (filename or page1), excluding same-file."""
    name_dup, content_dup = build_duplicate_groups(records)
    seen_pairs: set[tuple] = set()
    results: list[dict] = []

    def add_pair(source: dict, other: dict, reason: str) -> None:
        if source["path"] == other["path"]:
            return
        pk = directed_pair_key(source, other, reason)
        if pk in seen_pairs:
            return
        seen_pairs.add(pk)
        results.append(
            {
                "folder": source["folder"],
                "name": source["name"],
                "dup_name": other["name"],
                "dup_folder": other["folder"],
                "reason": reason,
            }
        )

    for group in name_dup.values():
        for i, a in enumerate(group):
            for b in group[i + 1 :]:
                add_pair(a, b, "文件名重复")
                add_pair(b, a, "文件名重复")

    for group in content_dup.values():
        for i, a in enumerate(group):
            for b in group[i + 1 :]:
                if a["name"].lower() == b["name"].lower():
                    continue
                add_pair(a, b, "首页内容重复")
                add_pair(b, a, "首页内容重复")

    results.sort(key=lambda r: (r["folder"], r["name"], r["dup_folder"], r["dup_name"]))
    return results


def format_report(path_label: str, root: Path, folders: list[str], dupes: list[dict]) -> str:
    lines: list[str] = []
    lines.append(f"{'=' * 72}")
    lines.append(path_label)
    lines.append(f"根路径: {root}")
    lines.append(f"扫描文件夹: {', '.join(folders)}")
    lines.append(f"{'=' * 72}")
    lines.append("")

    if not dupes:
        lines.append("（未发现重复 PDF）")
        return "\n".join(lines)

    by_folder: dict[str, dict[str, list[dict]]] = defaultdict(lambda: defaultdict(list))
    for d in dupes:
        by_folder[d["folder"]][d["name"]].append(d)

    for folder in folders:
        folder_dupes = by_folder.get(folder)
        if not folder_dupes:
            continue
        lines.append(f"## 文件夹: {folder}")
        lines.append("")
        for name in sorted(folder_dupes):
            entries = folder_dupes[name]
            reasons = sorted({e["reason"] for e in entries})
            lines.append(f"  - 文件: {name}")
            lines.append(f"    重复类型: {', '.join(reasons)}")
            for e in sorted(entries, key=lambda x: (x["dup_folder"], x["dup_name"])):
                lines.append(
                    f"    → 与 [{e['dup_folder']}] 中的「{e['dup_name']}」重复"
                )
            lines.append("")
        lines.append("")

    lines.append(f"--- 汇总: 共 {len(dupes)} 条重复关系记录 ---")
    unique_files = len({(d["folder"], d["name"]) for d in dupes})
    lines.append(f"涉及文件数（去重）: {unique_files}")
    return "\n".join(lines)


def safe_filename(label: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", label.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:80] or "report"


def run_job(
    root: Path,
    folders: list[str] | None,
    label: str,
    output_dir: Path,
) -> Path:
    folder_list = resolve_folders(root, folders)
    if not folder_list:
        raise ValueError(f"未找到可扫描的子文件夹: {root}")

    print(f"扫描: {label} ({root}) ...", flush=True)
    records = collect_pdfs(root, folder_list)
    dupes = find_cross_folder_dupes(records)
    report = format_report(label, root, folder_list, dupes)

    output_dir.mkdir(parents=True, exist_ok=True)
    out_name = f"duplicate_report_{safe_filename(label)}.txt"
    out_path = output_dir / out_name
    out_path.write_text(report, encoding="utf-8")
    print(report)
    print(f"报告已保存: {out_path}", flush=True)
    return out_path


def load_jobs_from_config(config_path: Path) -> list[dict]:
    data = json.loads(config_path.read_text(encoding="utf-8"))
    jobs = data.get("jobs", data if isinstance(data, list) else None)
    if not isinstance(jobs, list) or not jobs:
        raise ValueError("配置文件需包含非空 jobs 数组")
    return jobs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="在同一根目录的多个子文件夹之间检测 PDF 重复（文件名或首页文本）"
    )
    parser.add_argument(
        "--root",
        default=None,
        help="待扫描根目录（与 --config 二选一；可多次指定需配合 --config）",
    )
    parser.add_argument(
        "--folders",
        default=None,
        help="逗号分隔的子文件夹名；省略则扫描 root 下全部一级子目录",
    )
    parser.add_argument("--label", default=None, help="报告标题（默认使用 root 目录名）")
    parser.add_argument(
        "--config",
        default=None,
        help="JSON 配置文件，格式见 jobs.example.json",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="输出目录（默认 21_PDF_Duplicate_Analyzer/output/）",
    )
    return parser.parse_args()


def parse_folder_arg(value: str | None) -> list[str] | None:
    if not value:
        return None
    return [part.strip() for part in value.split(",") if part.strip()]


def main() -> int:
    args = parse_args()
    output_dir = normalize_path(args.output) if args.output else DEFAULT_OUTPUT_DIR

    jobs: list[dict] = []
    if args.config:
        config_path = normalize_path(args.config)
        if not config_path.is_file():
            print(f"ERROR: 配置文件不存在: {config_path}", file=sys.stderr)
            return 1
        jobs = load_jobs_from_config(config_path)
    elif args.root:
        root = normalize_path(args.root)
        label = args.label or root.name
        jobs = [
            {
                "root": str(root),
                "folders": parse_folder_arg(args.folders),
                "label": label,
            }
        ]
    else:
        print("ERROR: 请指定 --root 或 --config", file=sys.stderr)
        return 1

    written: list[Path] = []
    for idx, job in enumerate(jobs, start=1):
        root_raw = job.get("root")
        if not root_raw:
            print(f"ERROR: jobs[{idx}] 缺少 root", file=sys.stderr)
            return 1
        root = normalize_path(str(root_raw))
        folders_raw = job.get("folders")
        if isinstance(folders_raw, str):
            folders = parse_folder_arg(folders_raw)
        elif isinstance(folders_raw, list):
            folders = [str(f).strip() for f in folders_raw if str(f).strip()]
        else:
            folders = None
        label = str(job.get("label") or root.name)
        try:
            written.append(run_job(root, folders, label, output_dir))
        except ValueError as exc:
            print(f"ERROR: {exc}", file=sys.stderr)
            return 1

    summary_path = output_dir / f"run_summary_{datetime.now():%Y%m%d_%H%M%S}.txt"
    summary_lines = [f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}", f"报告数量: {len(written)}", ""]
    summary_lines.extend(str(p) for p in written)
    summary_path.write_text("\n".join(summary_lines), encoding="utf-8")
    print(f"运行摘要: {summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
