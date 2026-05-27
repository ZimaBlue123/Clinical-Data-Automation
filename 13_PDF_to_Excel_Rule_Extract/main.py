# -*- coding: utf-8 -*-
"""
按 config 从 PDF 检索内容并写入 Excel 对应位置。
用法: python main.py [--config config.yaml]
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
import json

import yaml

from src.excel_writer import (
    ensure_sheet,
    load_or_create_workbook,
    save_workbook,
    write_cell,
    write_table,
    write_text_block,
)
from src.pdf_reader import (
    build_mapping_audit_for_pdf,
    extract_text_from_pdf,
    get_first_table_near_keyword,
    search_by_keyword,
)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


def load_config(config_path: str | Path) -> dict:
    """
    加载 YAML 配置文件。
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        配置字典
        
    Raises:
        FileNotFoundError: 配置文件不存在
        ValueError: 配置文件格式错误
    """
    config_path = Path(config_path)
    if not config_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            if not isinstance(config, dict):
                raise ValueError("配置文件格式错误：根节点必须是字典")
            return config
    except yaml.YAMLError as e:
        raise ValueError(f"YAML 解析错误: {e}") from e
    except Exception as e:
        raise ValueError(f"读取配置文件失败: {e}") from e


def resolve_path(base_dir: Path, path_str: str) -> Path:
    p = Path(path_str)
    if not p.is_absolute():
        p = base_dir / p
    return p.resolve()


def run_rule(
    rule: dict,
    pdf_path: Path,
    wb,
    base_dir: Path,
    exclusion_boxes_by_page: dict[int, object] | None = None,
) -> None:
    """
    执行单个提取规则。
    
    Args:
        rule: 规则配置字典
        pdf_path: PDF 文件路径
        wb: 工作簿对象
        base_dir: 基础目录（用于解析相对路径）
    """
    if not isinstance(rule, dict):
        logger.error("规则配置必须是字典")
        return
    
    name = rule.get("name", "未命名规则")
    search_cfg = rule.get("search", {})
    excel_cfg = rule.get("excel", {})

    if not isinstance(search_cfg, dict):
        logger.error(f"规则 '{name}': search 配置必须是字典")
        return
    
    if not isinstance(excel_cfg, dict):
        logger.error(f"规则 '{name}': excel 配置必须是字典")
        return

    keyword = search_cfg.get("keyword")
    page = search_cfg.get("page")
    page_num = None
    if page is not None:
        try:
            page_num = int(page)
            if page_num < 1:
                logger.warning(f"规则 '{name}': 页码必须 >= 1，忽略")
                page_num = None
        except (ValueError, TypeError):
            logger.warning(f"规则 '{name}': 无效的页码 '{page}'，忽略")

    sheet_name = excel_cfg.get("sheet", "Sheet1")
    cell = excel_cfg.get("cell", "A1")
    extract_mode = excel_cfg.get("extract", "text")  # "text" | "table"

    try:
        ws = ensure_sheet(wb, sheet_name)
    except Exception as e:
        logger.error(f"规则 '{name}': 无法创建或获取工作表 '{sheet_name}': {e}")
        return

    if extract_mode == "table":
        try:
            table = get_first_table_near_keyword(
                pdf_path,
                keyword or "",
                page_num,
                exclusion_boxes_by_page=exclusion_boxes_by_page,
            )
            if table:
                write_table(ws, cell, table)
                logger.info(f"规则 '{name}': 已写入表格到 {sheet_name}!{cell}，共 {len(table)} 行")
                print(f"  [OK] {name}: 已写入表格到 {sheet_name}!{cell}，共 {len(table)} 行")
            else:
                # 回退为关键词附近文本
                hits = search_by_keyword(
                    pdf_path,
                    keyword or "",
                    page_num,
                    exclusion_boxes_by_page=exclusion_boxes_by_page,
                )
                if hits:
                    write_text_block(ws, cell, hits[0][1])
                    logger.info(f"规则 '{name}': 未找到表格，已写入文本到 {sheet_name}!{cell}")
                    print(f"  [OK] {name}: 未找到表格，已写入文本到 {sheet_name}!{cell}")
                else:
                    logger.warning(f"规则 '{name}': 未找到关键词或表格")
                    print(f"  [跳过] {name}: 未找到关键词或表格")
        except Exception as e:
            logger.exception("规则提取表格失败: rule=%s pdf=%s", name, pdf_path)
            print(f"  [错误] {name}: {e}")
        return

    # 文本模式
    if keyword:
        try:
            hits = search_by_keyword(
                pdf_path,
                keyword,
                page_num,
                exclusion_boxes_by_page=exclusion_boxes_by_page,
            )
            if hits:
                write_text_block(ws, cell, hits[0][1])
                logger.info(f"规则 '{name}': 已写入到 {sheet_name}!{cell}")
                print(f"  [OK] {name}: 已写入到 {sheet_name}!{cell}")
            else:
                logger.warning(f"规则 '{name}': 未找到关键词 «{keyword}»")
                print(f"  [跳过] {name}: 未找到关键词 «{keyword}»")
        except Exception as e:
            logger.exception("规则关键词搜索失败: rule=%s keyword=%s pdf=%s", name, keyword, pdf_path)
            print(f"  [错误] {name}: {e}")
        return

    if page_num is not None:
        try:
            pages_text = extract_text_from_pdf(
                pdf_path,
                [page_num],
                exclusion_boxes_by_page=exclusion_boxes_by_page,
            )
            if page_num in pages_text and pages_text[page_num].strip():
                write_text_block(ws, cell, pages_text[page_num].strip())
                logger.info(f"规则 '{name}': 已写入第 {page_num} 页文本到 {sheet_name}!{cell}")
                print(f"  [OK] {name}: 已写入第 {page_num} 页文本到 {sheet_name}!{cell}")
            else:
                logger.warning(f"规则 '{name}': 第 {page_num} 页无文本")
                print(f"  [跳过] {name}: 第 {page_num} 页无文本")
        except Exception as e:
            logger.exception("规则页面提取失败: rule=%s page=%s pdf=%s", name, page_num, pdf_path)
            print(f"  [错误] {name}: {e}")
        return

    logger.warning(f"规则 '{name}': 请配置 keyword 或 page")
    print(f"  [跳过] {name}: 请配置 keyword 或 page")


def main():
    """主函数。"""
    parser = argparse.ArgumentParser(description="从 PDF 按规则检索并写入 Excel")
    parser.add_argument("--config", "-c", default="config.yaml", help="配置文件路径")
    parser.add_argument("--input", "-i", default=None, help="可选：覆盖 config 中的 pdf_path")
    parser.add_argument("--output", "-o", default=None, help="可选：覆盖 config 中的 excel_path")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在输出 Excel")
    parser.add_argument("--verbose", "-v", action="store_true", help="显示详细日志")
    parser.add_argument("--exclusion-json", default=None, help="可选：水印排除框 boxes.json 路径（用于跳过疑似干扰区域）")
    parser.add_argument(
        "--no-mapping-audit",
        action="store_true",
        help="禁用坐标映射审计（默认在提供 exclusion-json 时生成 mapping_audit）",
    )
    parser.add_argument(
        "--mapping-audit-output",
        default=None,
        help="可选：将 mapping_audit 单独写入该 JSON 路径；未指定时合并到同目录 *_watermark_report.json 或写入 *_mapping_audit.json",
    )
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    try:
        # 规则引擎从模块目录运行时，也以“项目根目录”为基准解析 config 中的相对路径
        base_dir = Path(__file__).resolve().parent.parent
        config_path = resolve_path(base_dir, args.config)

        if not config_path.exists():
            print(f"配置文件不存在: {config_path}")
            print("请复制 config.example.yaml 为 config.yaml 并修改。")
            sys.exit(1)

        config = load_config(config_path)
        pdf_path_str = config.get("pdf_path", "")
        if not pdf_path_str:
            print("配置文件中缺少 pdf_path")
            sys.exit(1)
        
        pdf_path = resolve_path(base_dir, args.input if args.input else pdf_path_str)
        excel_path = resolve_path(
            base_dir,
            args.output if args.output else config.get("excel_path", "output.xlsx"),
        )
        rules = config.get("rules", [])

        if excel_path.exists() and not args.overwrite:
            print(f"输出已存在，请使用 --overwrite: {excel_path}")
            sys.exit(1)

        exclusion_boxes_by_page = None
        if args.exclusion_json:
            excl_path = resolve_path(base_dir, args.exclusion_json)
            if not excl_path.exists():
                print(f"exclusion-json 不存在: {excl_path}")
                sys.exit(1)
            with excl_path.open("r", encoding="utf-8") as f:
                raw = json.load(f)
            exclusion_boxes_by_page = {}
            for page_key, boxes in raw.items():
                try:
                    pn = int(page_key)
                except Exception:
                    continue
                # v1: { "1": [[x0,top,x1,bottom], ...], ... }
                if isinstance(boxes, list):
                    page_boxes = []
                    for b in boxes:
                        if isinstance(b, list) and len(b) == 4:
                            page_boxes.append((float(b[0]), float(b[1]), float(b[2]), float(b[3])))
                    if page_boxes:
                        exclusion_boxes_by_page[pn] = page_boxes
                    continue

                # v2: { "0": {rotation, mediabox, cropbox, page_width, page_height, boxes:[...]}, ... }
                if isinstance(boxes, dict):
                    b_list = boxes.get("boxes") or []
                    page_boxes = []
                    for b in b_list:
                        if isinstance(b, list) and len(b) == 4:
                            page_boxes.append((float(b[0]), float(b[1]), float(b[2]), float(b[3])))
                    if page_boxes:
                        exclusion_boxes_by_page[pn] = {
                            "rotation": boxes.get("rotation", 0),
                            "mediabox": boxes.get("mediabox", None),
                            "cropbox": boxes.get("cropbox", None),
                            "page_width": boxes.get("page_width", None),
                            "page_height": boxes.get("page_height", None),
                            "boxes": page_boxes,
                        }
                    continue

        # 坐标映射审计：合并至 *_watermark_report.json 或单独文件
        if exclusion_boxes_by_page and not args.no_mapping_audit:
            excl_path = resolve_path(base_dir, args.exclusion_json) if args.exclusion_json else None
            if excl_path and excl_path.exists():
                try:
                    mapping_audit = build_mapping_audit_for_pdf(pdf_path, exclusion_boxes_by_page)
                    stem = excl_path.stem
                    pdf_stem = stem[: -len("_boxes")] if stem.endswith("_boxes") else stem
                    out_dir = excl_path.parent

                    if args.mapping_audit_output:
                        audit_out = resolve_path(base_dir, args.mapping_audit_output)
                    else:
                        audit_out = None

                    wm_path = out_dir / f"{pdf_stem}_watermark_report.json"
                    standalone_path = out_dir / f"{pdf_stem}_mapping_audit.json"

                    if audit_out:
                        payload = {
                            "mapping_audit": mapping_audit,
                            "source": "12_PDF_to_Excel_Rule_Extract",
                            "pdf": str(pdf_path),
                            "exclusion_json": str(excl_path),
                        }
                        audit_out.parent.mkdir(parents=True, exist_ok=True)
                        with audit_out.open("w", encoding="utf-8") as f:
                            json.dump(payload, f, indent=2, ensure_ascii=False)
                        logger.info("mapping_audit 已写入: %s", audit_out)
                    elif wm_path.exists():
                        with wm_path.open("r", encoding="utf-8") as f:
                            wr = json.load(f)
                        if not isinstance(wr, dict):
                            wr = {}
                        wr["mapping_audit"] = mapping_audit
                        wr["mapping_audit_source"] = "12_PDF_to_Excel_Rule_Extract"
                        with wm_path.open("w", encoding="utf-8") as f:
                            json.dump(wr, f, indent=2, ensure_ascii=False)
                        logger.info("mapping_audit 已合并至: %s", wm_path)
                    else:
                        payload = {
                            "mapping_audit": mapping_audit,
                            "source": "12_PDF_to_Excel_Rule_Extract",
                            "pdf": str(pdf_path),
                            "exclusion_json": str(excl_path),
                            "note": f"未找到 {wm_path.name}，已写入独立审计文件",
                        }
                        with standalone_path.open("w", encoding="utf-8") as f:
                            json.dump(payload, f, indent=2, ensure_ascii=False)
                        logger.info("mapping_audit 已写入: %s", standalone_path)

                    anom = mapping_audit.get("anomalies") or {}
                    logger.info(
                        "mapping_audit 汇总: clamped=%s dropped=%s severe_distortion_pages=%s",
                        anom.get("clamped_boxes_count"),
                        anom.get("dropped_boxes_count"),
                        anom.get("severe_area_distortion_pages"),
                    )
                except Exception:
                    logger.exception("mapping_audit 生成失败（已忽略）: pdf=%s exclusion=%s", pdf_path, excl_path)

        if not isinstance(rules, list):
            logger.error("配置文件中 rules 必须是列表")
            print("配置文件中 rules 必须是列表")
            sys.exit(1)

        if not pdf_path.exists():
            print(f"PDF 不存在: {pdf_path}")
            sys.exit(1)

        print(f"PDF: {pdf_path}")
        print(f"Excel: {excel_path}")
        print(f"规则数: {len(rules)}")
        print()

        wb = load_or_create_workbook(excel_path)
        success_count = 0
        for i, rule in enumerate(rules, 1):
            print(f"处理规则 {i}/{len(rules)}: {rule.get('name', '未命名规则')}")
            try:
                run_rule(rule, pdf_path, wb, base_dir, exclusion_boxes_by_page=exclusion_boxes_by_page)
                success_count += 1
            except Exception as e:
                logger.exception("规则执行失败: index=%s rule=%s pdf=%s", i, rule.get("name", "未命名规则"), pdf_path)
                print(f"  [错误] 规则执行失败: {e}")
            print()
        
        save_workbook(wb, excel_path)
        print(f"完成。成功处理 {success_count}/{len(rules)} 个规则。")
        
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
        sys.exit(130)
    except Exception as e:
        logger.error(f"程序执行失败: {e}", exc_info=True)
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
