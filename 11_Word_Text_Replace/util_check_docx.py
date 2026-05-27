"""校验 output/ 中 docx 是否仍有待替换内容。"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import tempfile
import zipfile

from lxml import etree

from lib.ooxml_replace import (
    NS,
    build_date_rules,
    build_default_rules,
    build_study_id_rules,
    iter_package_xml_files,
    iter_text_nodes,
)

MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUTPUT = os.path.join(MODULE_DIR, "output")

STUDY_OLD = "YDSWX（TVAX-009）-004（Ⅳ）"
STUDY_NEW = "YDSWX（TVAX-009）-004（III）"


def _joined_text_in_xml(xml_path: str) -> str:
    root = etree.parse(xml_path).getroot()
    parts = []
    for p in root.xpath(".//w:p", namespaces=NS):
        parts.append("".join((n.text or "") for n in iter_text_nodes(p)))
    return "\n".join(parts)


def check_docx(docx_path: str, mode: str) -> None:
    if mode == "dates":
        rules = build_date_rules()
    elif mode == "study":
        rules = build_study_id_rules()
    else:
        rules = build_default_rules()

    tmp = tempfile.mkdtemp(prefix="check_docx_")
    try:
        zipfile.ZipFile(docx_path).extractall(tmp)
        blob = ""
        for xf in iter_package_xml_files(tmp):
            blob += _joined_text_in_xml(xf) + "\n"

        remaining = rules.count_in_text(blob)
        print("file:", os.path.basename(docx_path))
        print("mode:", mode)
        print("rules_remaining:", remaining)
        print("study_old_literal:", blob.count(STUDY_OLD))
        print("study_new_literal:", blob.count(STUDY_NEW))
        print("bad_prefix_I):", blob.count("I) YDSWX"))
        print("bad_prefix_I)no_space:", blob.count("I)YDSWX"))
        print("004_IV_unicode:", blob.count("004（Ⅳ）"))
        print("004_III:", blob.count("004（III）"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="校验 docx 批量替换结果")
    ap.add_argument("--latest", action="store_true")
    ap.add_argument("--file", default=None)
    ap.add_argument("--output-dir", default=DEFAULT_OUTPUT)
    ap.add_argument("--mode", choices=("all", "dates", "study"), default="all")
    args = ap.parse_args()

    if args.latest:
        files = glob.glob(os.path.join(args.output_dir, "*_updated*.docx"))
        if not files:
            raise SystemExit(f"未找到输出：{args.output_dir}")
        docx_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    elif args.file:
        docx_path = args.file
    else:
        raise SystemExit("需要 --latest 或 --file")

    check_docx(docx_path, args.mode)


if __name__ == "__main__":
    main()
