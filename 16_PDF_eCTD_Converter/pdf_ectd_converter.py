#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
eCTD 合规装甲 & XSS 深度清理器（16_PDF_eCTD_Converter）

该模块融合防止恶意脚本/XSS 的能力，并强制执行《eCTD验证标准V1.1》附件6常见要求（按可实现程度落地）：
- 可读性校验（6.1）：可打开且页数 > 0
- 剥离密码/安全设置（6.19 / 6.21）：需要密码才能打开的 PDF 直接拒收；输出统一保存为未加密
- 移除所有附件（6.17）
- 移除除超文本链接外的所有注释（6.18）
- 清理/拦截外部链接与非法协议（6.3 / 6.10 / 6.11）
- 强制设置初始视图（6.20）：UseOutlines + OneColumn（PyMuPDF 可稳定设置）
- 大于 5 页必须有书签（6.23）
- 启用快速 Web 查看（Fast Web View / Linearization）（6.22）
- 导出合规审计 Excel 报告（便于审计与回溯）

用法：
  cd 16_PDF_eCTD_Converter
  python pdf_ectd_converter.py --input "./input" --output "./output" --report "./ectd_report.xlsx" --overwrite

  # 指定输入目录/单文件
  python pdf_ectd_converter.py --input "D:\\pdfs"
  python pdf_ectd_converter.py --input "D:\\pdfs\\a.pdf"

  # 仅校验（也会写审计报告）
  python pdf_ectd_converter.py --validate-only --report "output/ectd_report.xlsx"
"""

from __future__ import annotations

import argparse
import logging
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd


# eCTD 与 XSS 共同封杀的恶意协议与外部前缀（保守策略：外部一律视为风险）
BAD_PROTOCOLS = ("javascript:", "data:", "vbscript:", "file:", "http://", "https://", "mailto:")
logger = logging.getLogger(__name__)


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _collect_pdfs(input_path: Path, recursive: bool = True) -> list[Path]:
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return sorted([p.resolve() for p in input_path.glob(pattern) if p.is_file()])
    return []


def _validate_pdf_basic(pdf_path: Path) -> tuple[bool, str, dict]:
    """
    核心入门校验：
    - 6.21: 不允许需要密码才能打开
    - 6.1: 可打开且页数>0（未损坏/可读）
    """
    meta: dict = {}
    try:
        doc = fitz.open(pdf_path)
    except fitz.FileDataError:
        return False, "错误 (6.1): 文件被破坏或不可读", meta
    except Exception as exc:
        logger.exception("PDF 基础校验异常: file=%s", pdf_path)
        return False, f"未知异常: {exc}", meta

    try:
        if doc.needs_pass:
            return False, "错误 (6.21): 存在密码保护，无法打开", meta
        if doc.page_count <= 0:
            return False, "错误 (6.1): 页数为0或内容无效", meta

        toc = doc.get_toc()
        meta["page_count"] = int(doc.page_count)
        meta["has_toc"] = bool(toc)
        return True, "OK", meta
    finally:
        doc.close()


class ECTDComplianceCleaner:
    def __init__(self, input_dir: Path, output_dir: Path, report_path: Path, overwrite: bool):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.report_path = report_path
        self.overwrite = overwrite
        self.report_rows: list[dict] = []

    def _is_illegal_link(self, link: dict) -> tuple[bool, str]:
        """
        违规链接判定（保守）：
        - 任何 URI 链接（LINK_URI）视为外部链接风险（6.3 / 6.10）
        - 协议黑名单（javascript/data/vbscript/file/http/https/mailto）
        - Named action 视为风险（常见为 JS 或内部命令）
        """
        uri = (link.get("uri") or "").strip()
        uri_lower = uri.lower()
        kind = link.get("kind")

        # 1) 协议黑名单与外部链接
        for proto in BAD_PROTOCOLS:
            if uri_lower.startswith(proto):
                return True, f"包含非法协议或外部链接前缀（{proto}）"

        # 2) URI 链接默认拒绝（即便是相对路径/片段也可能被解释为外部跳转）
        if kind == getattr(fitz, "LINK_URI", 1):
            return True, "包含外部 URI 链接（LINK_URI）"

        # 3) Named action（如 JavaScript / Named destinations 的变种）
        if kind == getattr(fitz, "LINK_NAMED", 5):
            return True, "包含 Named Action（LINK_NAMED）"

        # 4) Launch 动作风险较高（可能启动外部程序/文件）
        if kind == getattr(fitz, "LINK_LAUNCH", 4):
            return True, "包含 Launch 动作（LINK_LAUNCH）"

        # 其余：允许 LINK_GOTO（文档内跳转）与 LINK_GOTOR（跨文档跳转，是否合规取决于提交结构；这里不强杀）
        return False, ""

    def _append_report(
        self,
        filename: str,
        status: str,
        detail: str,
        *,
        page_count: int | None = None,
        has_toc: bool | None = None,
        removed_embedded: int | None = None,
        removed_annots: int | None = None,
        removed_links: int | None = None,
        has_searchable_text: bool | None = None,
        pagemode_set: bool | None = None,
        pagelayout_set: bool | None = None,
        linearized: bool | None = None,
    ) -> None:
        self.report_rows.append(
            {
                "文件名": filename,
                "处理时间": _now_str(),
                "状态": status,
                "页数": page_count,
                "是否有书签": has_toc,
                "删除附件数": removed_embedded,
                "删除注释数": removed_annots,
                "删除违规链接数": removed_links,
                "是否有可搜索文本": has_searchable_text,
                "初始视图UseOutlines": pagemode_set,
                "页面布局OneColumn": pagelayout_set,
                "FastWebView(Linear)": linearized,
                "详细信息": detail,
            }
        )

    def process_pdf(self, pdf_path: Path, output_path: Path, *, validate_only: bool) -> bool:
        status = "FAILED"
        details: list[str] = []

        if not pdf_path.exists():
            logger.error("输入文件不存在: %s", pdf_path)
            self._append_report(pdf_path.name, status, "文件不存在")
            return False

        if output_path.exists() and not self.overwrite and not validate_only:
            logger.warning("跳过已存在输出: input=%s output=%s", pdf_path, output_path)
            self._append_report(pdf_path.name, "SKIPPED", "文件已存在且未开启覆盖")
            return False

        ok, msg, meta = _validate_pdf_basic(pdf_path)
        page_count = meta.get("page_count")
        has_toc = meta.get("has_toc")
        if not ok:
            self._append_report(pdf_path.name, "FAILED", msg)
            return False

        # 6.23: 大于5页必须有书签（严格）
        if isinstance(page_count, int) and page_count > 5 and not has_toc:
            self._append_report(
                pdf_path.name,
                "FAILED",
                "错误: 大于5页的文件缺少书签（规则 6.23）",
                page_count=page_count,
                has_toc=has_toc,
            )
            return False

        if validate_only:
            self._append_report(
                pdf_path.name,
                "SUCCESS",
                "校验通过（validate-only）",
                page_count=page_count,
                has_toc=has_toc,
            )
            return True

        removed_embedded = 0
        removed_annots = 0
        removed_links = 0
        has_searchable_text = False
        pagemode_set = False
        pagelayout_set = False
        linearized = False

        doc = None
        try:
            doc = fitz.open(pdf_path)
            if doc.needs_pass:
                raise ValueError("文件存在密码保护，无法解析（规则 6.21）")

            # 6.17: 删除嵌入文件（附件）
            emb_count = int(doc.embfile_count())
            if emb_count > 0:
                for i in range(emb_count - 1, -1, -1):
                    doc.embfile_del(i)
                removed_embedded = emb_count
                details.append(f"清理了 {emb_count} 个附件（6.17）")

            # 6.18 / 6.10 / 6.11: 删除非链接注释 + 删除违规链接
            for page in doc:
                # 6.18: 移除除超文本链接外的所有注释
                annot = page.first_annot
                while annot:
                    next_annot = annot.next
                    try:
                        if annot.type[0] != fitz.PDF_ANNOT_LINK:
                            page.delete_annot(annot)
                            removed_annots += 1
                    except Exception:
                        # 注释损坏时，尽量删除
                        try:
                            page.delete_annot(annot)
                            removed_annots += 1
                        except Exception:
                            logger.debug("删除注释失败: file=%s page=%s", pdf_path.name, page.number + 1, exc_info=True)
                    annot = next_annot

                # 删除违规链接（先收集再删）
                bad_links: list[dict] = []
                for link in page.get_links():
                    illegal, reason = self._is_illegal_link(link)
                    if illegal:
                        bad_links.append(link)
                        if reason and reason not in details:
                            details.append(f"链接清理: {reason}")
                for link in bad_links:
                    try:
                        page.delete_link(link)
                        removed_links += 1
                    except Exception:
                        logger.debug("删除链接失败: file=%s page=%s", pdf_path.name, page.number + 1, exc_info=True)

            # 6.25: 可搜索文本检查（预警）
            has_searchable_text = any((p.get_text() or "").strip() for p in doc)
            if not has_searchable_text:
                details.append("警告: 未检测到可搜索文本，建议执行 OCR（6.25）")

            # 6.20: 初始视图（稳定可用）
            try:
                doc.set_pagemode("UseOutlines")
                pagemode_set = True
            except Exception:
                details.append("警告: 无法设置初始视图 UseOutlines（6.20）")

            try:
                doc.set_pagelayout("OneColumn")
                pagelayout_set = True
            except Exception:
                details.append("警告: 无法设置页面布局 OneColumn（6.20）")

            output_path.parent.mkdir(parents=True, exist_ok=True)

            # 6.22: Fast Web View（linearization）
            try:
                doc.save(
                    output_path,
                    incremental=False,
                    garbage=4,
                    deflate=True,
                    clean=True,
                    linear=True,
                    encryption=fitz.PDF_ENCRYPT_NONE,
                )
                linearized = True
            except TypeError:
                doc.save(
                    output_path,
                    incremental=False,
                    garbage=4,
                    deflate=True,
                    clean=True,
                    encryption=fitz.PDF_ENCRYPT_NONE,
                )
                details.append("提示: 当前 PyMuPDF 不支持 linear 参数，已按非线性方式保存")

            status = "SUCCESS"
            details.insert(0, "清洗合规通过")
            return True
        except Exception as exc:
            logger.exception("eCTD 处理失败: file=%s output=%s", pdf_path, output_path)
            details.append(f"处理终止: {exc}")
            return False
        finally:
            if doc is not None:
                doc.close()
            self._append_report(
                pdf_path.name,
                status,
                " | ".join(details) if details else msg,
                page_count=page_count if isinstance(page_count, int) else None,
                has_toc=has_toc if isinstance(has_toc, bool) else None,
                removed_embedded=removed_embedded,
                removed_annots=removed_annots,
                removed_links=removed_links,
                has_searchable_text=has_searchable_text,
                pagemode_set=pagemode_set,
                pagelayout_set=pagelayout_set,
                linearized=linearized,
            )

    def export_report(self) -> None:
        if not self.report_rows:
            logger.warning("无审计数据可导出。")
            return
        df = pd.DataFrame(self.report_rows)
        self.report_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(self.report_path, index=False)
        logger.info("eCTD 审计报告已生成: %s", self.report_path)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"
    default_report = base_dir / "ectd_report.xlsx"

    parser = argparse.ArgumentParser(description="eCTD 合规装甲 & 批量 PDF XSS 深度清理器")
    parser.add_argument("--input", "-i", default=str(default_input), help="输入 PDF 文件夹或单个 PDF 文件")
    parser.add_argument("--output", "-o", default=str(default_output), help="输出文件夹（默认 output/）")
    parser.add_argument("--report", default=str(default_report), help="Excel 审计报告输出路径")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    parser.add_argument("--keep-name", action="store_true", help="输出文件名与源文件一致（默认追加 _ectd）")
    parser.add_argument("--validate-only", action="store_true", help="仅做校验，不做输出（仍生成审计报告）")
    parser.add_argument(
        "--recursive",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="是否递归遍历子文件夹（默认开启）",
    )
    parser.add_argument(
        "--keep-structure",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="当输入为目录时，是否在输出目录中保留相对目录结构（默认开启）",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()
    report_path = Path(args.report).expanduser()

    pdf_files = _collect_pdfs(input_path, recursive=args.recursive)
    if not pdf_files:
        logger.error("未找到 PDF: %s", input_path)
        raise SystemExit(1)

    logger.info("启动 eCTD 处理: files=%s", len(pdf_files))

    cleaner = ECTDComplianceCleaner(
        input_dir=input_path,
        output_dir=output_dir,
        report_path=report_path,
        overwrite=args.overwrite,
    )

    total = 0
    success = 0
    base_input_dir = input_path.resolve() if input_path.is_dir() else None

    for pdf_path in pdf_files:
        total += 1
        if base_input_dir and args.keep_structure:
            rel = pdf_path.relative_to(base_input_dir)
            if args.keep_name:
                out_path = output_dir / rel
            else:
                out_path = (output_dir / rel).with_name(f"{rel.stem}_ectd.pdf")
        else:
            if args.keep_name:
                out_path = output_dir / pdf_path.name
            else:
                out_path = output_dir / f"{pdf_path.stem}_ectd.pdf"

        if cleaner.process_pdf(pdf_path, out_path, validate_only=args.validate_only):
            success += 1
            if args.validate_only:
                logger.info("校验通过: %s", pdf_path.name)
            else:
                logger.info("处理成功: %s => %s", pdf_path.name, out_path.name)
        else:
            logger.error("处理失败: %s", pdf_path.name)

    cleaner.export_report()
    logger.info("执行完毕: success=%s total=%s", success, total)
    raise SystemExit(0 if success == total else 1)


if __name__ == "__main__":
    main()

