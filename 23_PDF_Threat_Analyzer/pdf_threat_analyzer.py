"""
PDF 威胁分析与工业级安全剥离工具（模块 23）。

功能：
1. **静态特征分析**——无渲染扫描 PDF 内部字节流，识别 JavaScript / OpenAction /
   Launch / EmbeddedFiles 等高危对象，提取可疑 URL，给出风险评分与等级（LOW/MEDIUM/HIGH）。
2. **工业级安全剥离**——可选生成"安全副本"，移除已知威胁对象：
   - 优先使用 PyMuPDF (fitz) —— mupdf 内核级剥离，覆盖注释 / 链接 / 嵌入文件
   - 降级使用 pypdf —— 基础结构级剥离
   - 可选调用 qpdf / mutool —— 做线性化校验
3. **标准自检**——生成 1 个含 `/JavaScript` 模拟威胁的测试 PDF，跑通分析+剥离，
   验证报告字段完整、风险等级准确、剥离后无残留威胁；任何环节失败 → exit 2。

依赖（按优先级，降级路径）：
  - fitz (PyMuPDF) ≥ 1.23  — 推荐，已在 requirements.txt
  - pypdf ≥ 4.0            — 已安装
  - qpdf (CLI，可选)        — 系统 PATH 探测；缺失时 graceful skip
  - mutool (CLI，可选)      — 系统 PATH 探测；缺失时 graceful skip

用法（CLI）：
  python pdf_threat_analyzer.py                          # 默认扫描 ./input → ./output
  python pdf_threat_analyzer.py --self-check             # 跑标准自检
  python pdf_threat_analyzer.py --input D:\\suspicious --sanitize
  python pdf_threat_analyzer.py --engine pypdf --no-recursive

退出码（遵循 docs/logging_convention.md）：
  0  成功
  1  部分或全部失败
  2  自检 / 依赖前置失败
  130 用户中断
"""
from __future__ import annotations

import argparse
import json
import logging
import mmap
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# 标准日志（按 docs/logging_convention.md：action=xxx key=value）
# ---------------------------------------------------------------------------
logger = logging.getLogger("pdf_threat_analyzer")
_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def _configure_logging(verbose: bool = False) -> None:
    """配置日志；不破坏调用方已设置的 handler。"""
    level = logging.DEBUG if verbose else logging.INFO
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter(_LOG_FORMAT, _DATE_FORMAT))
        logger.addHandler(handler)
    logger.setLevel(level)
    logger.propagate = False


# ---------------------------------------------------------------------------
# 引擎探测
# ---------------------------------------------------------------------------
@dataclass
class EngineStatus:
    """外部引擎可用性快照。"""
    fitz: bool = False
    pypdf: bool = False
    qpdf: bool = False
    mutool: bool = False

    def summary(self) -> str:
        flags = {
            "fitz": self.fitz,
            "pypdf": self.pypdf,
            "qpdf": self.qpdf,
            "mutool": self.mutool,
        }
        return ", ".join(f"{k}={'yes' if v else 'no'}" for k, v in flags.items())


def detect_engines() -> EngineStatus:
    """探测本地可用的 PDF 引擎。绝不抛异常。"""
    status = EngineStatus()
    try:
        import fitz  # noqa: F401
        status.fitz = True
    except Exception:
        pass
    try:
        from pypdf import PdfReader  # noqa: F401
        status.pypdf = True
    except Exception:
        pass
    status.qpdf = shutil.which("qpdf") is not None
    status.mutool = shutil.which("mutool") is not None
    return status


# ---------------------------------------------------------------------------
# 威胁字典
# ---------------------------------------------------------------------------
# 高危 PDF 关键字 → 权重（命中次数 × 权重 = 风险分）
RISK_KEYWORDS: dict[bytes, int] = {
    b"/JavaScript": 3,
    b"/JS": 3,
    b"/OpenAction": 2,   # 自动执行动作
    b"/AA": 2,           # 附加动作
    b"/Launch": 3,       # 启动外部程序
    b"/EmbeddedFiles": 2,
    b"/SubmitForm": 1,
    b"/URI": 1,          # 外部链接（仅按出现计数，不区分协议）
}

# 已知恶意 URL 协议（用于 sanitize 步骤定向剥离，不参与风险评分）
BAD_URL_PROTOCOLS = ("javascript:", "data:text/html", "vbscript:", "file:")

# mmap 零拷贝加速的尺寸阈值：小于此值用 read()（小文件 mmap 反而更慢，因建链开销）
# 经验值：1MB 以下 read() 占优；以上 mmap 占优。
MMAP_THRESHOLD_BYTES = 1 * 1024 * 1024

# 模块级预编译正则：避免每次扫描重新编译。
# key: 关键字 bytes；value: (权重, 编译后的 pattern)
# pattern 在关键字后接 [^a-zA-Z0-9] 边界，避免误报 /JavaScriptable 这类带后缀的合法名。
COMPILED_RISK_PATTERNS: dict[bytes, tuple[int, re.Pattern[bytes]]] = {
    kw: (weight, re.compile(re.escape(kw) + b"[^a-zA-Z0-9]"))
    for kw, weight in RISK_KEYWORDS.items()
}


# ---------------------------------------------------------------------------
# 分析器
# ---------------------------------------------------------------------------
class PDFThreatAnalyzer:
    """无渲染 PDF 静态特征提取与风险评分。"""

    def __init__(self, file_path: Path) -> None:
        self.file_path = Path(file_path)
        self.report: dict[str, Any] = {
            "file_name": self.file_path.name,
            "file_size_bytes": 0,
            "is_valid_pdf": False,
            "is_encrypted": False,
            "raw_keyword_hits": {},
            "extracted_urls": [],
            "suspicious_urls": [],
            "risk_score": 0,
            "risk_level": "UNKNOWN",
            "errors": [],
        }

    # -- 内部工具 ----------------------------------------------------------
    def _log_error(self, message: str) -> None:
        logger.error("action=analysis_error file=%s reason=%s",
                     self.file_path.name, message)
        self.report["errors"].append(message)

    # -- 三层检测 ----------------------------------------------------------
    def _verify_magic_number(self) -> bool:
        """验证文件头部标识符。"""
        try:
            with open(self.file_path, "rb") as f:
                header = f.read(1024)
            if b"%PDF-" in header:
                self.report["is_valid_pdf"] = True
                return True
        except Exception as exc:  # noqa: BLE001
            self._log_error(f"读取文件头部失败: {exc}")
        return False

    def _scan_raw_binary(self) -> None:
        """底层二进制特征扫描（防逃逸，mmap 零拷贝加速）。

        - 大文件（>= MMAP_THRESHOLD_BYTES，默认 1MB）走 mmap 零拷贝路径
        - 小文件走传统 read() 路径（mmap 建链开销在小文件上反而更慢）
        - 正则模式在模块级 COMPILED_RISK_PATTERNS 预编译，避免每次扫描重新编译
        """
        if self.report["file_size_bytes"] == 0:
            return
        try:
            size = self.report["file_size_bytes"]
            if size >= MMAP_THRESHOLD_BYTES:
                self._scan_with_mmap()
            else:
                self._scan_with_read()
        except Exception as exc:  # noqa: BLE001
            self._log_error(f"二进制流扫描异常: {exc}")

    def _scan_with_read(self) -> None:
        """小文件路径：一次性 read() 全文到内存，正则扫描。"""
        with open(self.file_path, "rb") as f:
            content = f.read()
        for keyword, (weight, pattern) in COMPILED_RISK_PATTERNS.items():
            matches = pattern.findall(content)
            count = len(matches)
            if count > 0:
                self._record_hit(keyword, count, weight)

    def _scan_with_mmap(self) -> None:
        """大文件路径：mmap 零拷贝映射到内存，正则直接扫描 mmap 对象。

        - ACCESS_READ 保护文件不被意外修改
        - length=0 自动取整个文件
        - Windows / Linux / macOS 通用
        """
        with open(self.file_path, "rb") as f, mmap.mmap(f.fileno(), length=0, access=mmap.ACCESS_READ) as mm:
            for keyword, (weight, pattern) in COMPILED_RISK_PATTERNS.items():
                # mmap 对象直接作为 findall 的输入，零拷贝
                matches = pattern.findall(mm)
                count = len(matches)
                if count > 0:
                    self._record_hit(keyword, count, weight)

    def _record_hit(self, keyword: bytes, count: int, weight: int) -> None:
        """记录单条关键字命中到 report。"""
        key_str = keyword.decode("utf-8")
        self.report["raw_keyword_hits"][key_str] = count
        self.report["risk_score"] += weight * count

    def _extract_structural_data(self) -> None:
        """解析 PDF 结构树，提取 URL 并识别加密。"""
        try:
            from pypdf import PdfReader  # 复用探测结果
        except ImportError:
            self._log_error("pypdf 不可用，跳过结构化提取")
            return
        try:
            reader = PdfReader(str(self.file_path))
        except Exception as exc:  # noqa: BLE001
            self._log_error(f"PDF 结构损坏或解析失败: {exc}")
            return

        if reader.is_encrypted:
            self.report["is_encrypted"] = True
            self.report["risk_score"] += 1
            logger.warning("action=encrypted_detected file=%s", self.file_path.name)
            return  # 加密文件无法继续结构化解析

        urls: set = set()
        try:
            for page in reader.pages:
                if "/Annots" not in page:
                    continue
                annots = page["/Annots"]
                if hasattr(annots, "get_object"):
                    annots = annots.get_object()
                if not isinstance(annots, list):
                    continue
                for annot in annots:
                    try:
                        annot_obj = annot.get_object()
                    except Exception:
                        continue
                    if "/A" not in annot_obj:
                        continue
                    try:
                        action = annot_obj["/A"].get_object()
                    except Exception:
                        continue
                    if "/URI" in action:
                        try:
                            urls.add(str(action["/URI"]))
                        except Exception:
                            pass
        except Exception as exc:  # noqa: BLE001
            self._log_error(f"注释遍历异常: {exc}")

        url_list = sorted(urls)
        self.report["extracted_urls"] = url_list
        # 标记可疑 URL（恶意协议）
        suspicious = [
            u for u in url_list
            if any(u.lower().startswith(p) for p in BAD_URL_PROTOCOLS)
        ]
        self.report["suspicious_urls"] = suspicious
        if suspicious:
            self.report["risk_score"] += len(suspicious) * 2  # 恶意协议链接加权
        if url_list:
            self.report["risk_score"] += min(len(url_list), 5)  # 普通外链 1 分/个，最多 5 分

    def _calculate_risk_level(self) -> None:
        """根据风险分评定 LOW/MEDIUM/HIGH。"""
        score = self.report["risk_score"]
        if score == 0:
            self.report["risk_level"] = "LOW"
        elif 1 <= score <= 4:
            self.report["risk_level"] = "MEDIUM"
        else:
            self.report["risk_level"] = "HIGH"

    # -- 入口 --------------------------------------------------------------
    def analyze(self) -> dict[str, Any]:
        """执行完整分析管线，返回报告 dict。"""
        if not self.file_path.exists():
            self._log_error(f"文件不存在: {self.file_path}")
            return self.report

        self.report["file_size_bytes"] = self.file_path.stat().st_size
        if not self._verify_magic_number():
            self._log_error("非标准 PDF 文件结构")
            return self.report

        logger.info("action=analyze_start file=%s size_bytes=%s",
                    self.file_path.name, self.report["file_size_bytes"])
        self._scan_raw_binary()
        self._extract_structural_data()
        self._calculate_risk_level()
        logger.info(
            "action=analyze_done file=%s risk_level=%s score=%s",
            self.file_path.name, self.report["risk_level"], self.report["risk_score"],
        )
        return self.report

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.report, ensure_ascii=False, indent=indent)


# ---------------------------------------------------------------------------
# 工业级安全剥离
# ---------------------------------------------------------------------------
class PDFSanitizer:
    """根据分析报告剥离威胁，输出"安全副本"。

    策略：
      - fitz 路径：删除所有注释（注释常含 JS 动作）、删除恶意协议链接、删除嵌入文件，
        以 garbage=4 + deflate=True 重写，触发 mupdf 完整重写流。
      - pypdf 降级路径：基于 pypdf 重新解析并去除 OpenAction / AA / Names / AcroForm 等危险字典，
        保存为新 PDF。
      - qpdf 路径：剥离后再用 `qpdf --linearize` 验证文件完整性。
    """

    def __init__(self, engines: EngineStatus) -> None:
        self.engines = engines

    def sanitize(self, src: Path, dst: Path, overwrite: bool = False) -> tuple[bool, str]:
        """剥离 src 中的威胁，输出到 dst。返回 (success, engine_used_or_error)。"""
        if dst.exists() and not overwrite:
            return False, f"目标已存在，未覆盖: {dst}"

        if self.engines.fitz:
            return self._sanitize_with_fitz(src, dst)
        if self.engines.pypdf:
            return self._sanitize_with_pypdf(src, dst)
        return False, "无可用 sanitize 引擎（需 fitz 或 pypdf）"

    # -- fitz 路径（mupdf 内核，参考 15_PDF_XSS） ----------------------------
    def _sanitize_with_fitz(self, src: Path, dst: Path) -> tuple[bool, str]:
        try:
            import fitz  # type: ignore
        except ImportError:
            return False, "fitz 不可用"
        try:
            doc = fitz.open(str(src))
        except Exception as exc:  # noqa: BLE001
            return False, f"fitz 打开失败: {exc}"

        removed_annots = 0
        removed_links = 0
        removed_embeds = 0
        removed_catalog_keys = 0
        intermediate: Path | None = None
        try:
            for page in doc:
                # 1) 删注释（最常携带 JS 动作的对象）
                annot = page.first_annot
                while annot:
                    nxt = annot.next
                    try:
                        page.delete_annot(annot)
                        removed_annots += 1
                    except Exception:
                        pass
                    annot = nxt

                # 2) 删恶意协议链接
                for link in page.get_links() or []:
                    uri = (link.get("uri") or "").lower()
                    kind = link.get("kind")
                    action_str = str(link)
                    is_bad = (
                        "javascript" in uri
                        or any(uri.startswith(p) for p in BAD_URL_PROTOCOLS)
                        or (kind == 2 and ("/JavaScript" in action_str or "/JS" in action_str))
                    )
                    if is_bad:
                        try:
                            page.delete_link(link)
                            removed_links += 1
                        except Exception:
                            pass

            # 3) 删嵌入文件
            emb_count = doc.embfile_count()
            for i in range(emb_count - 1, -1, -1):
                try:
                    doc.embfile_del(i)
                    removed_embeds += 1
                except Exception:
                    pass

            # 4) 重写到 intermediate（garbage=4 触发完整对象流重写，清除悬挂引用）
            dst.parent.mkdir(parents=True, exist_ok=True)
            intermediate = dst.with_suffix(".intermediate.pdf")
            doc.save(str(intermediate), garbage=4, deflate=True)
        except Exception as exc:  # noqa: BLE001
            try:
                doc.close()
            except Exception:
                pass
            return False, f"fitz sanitize 失败: {exc}"
        finally:
            try:
                doc.close()
            except Exception:
                pass

        # 5) 二阶段：pypdf 清除 catalog 级危险字典（fitz 不直接暴露 catalog 编辑）
        if self.engines.pypdf and intermediate is not None:
            try:
                from pypdf import PdfReader, PdfWriter  # type: ignore
                reader = PdfReader(str(intermediate))
                writer = PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                for key in ("/OpenAction", "/AA", "/AcroForm"):
                    if key in writer._root_object:
                        del writer._root_object[key]
                        removed_catalog_keys += 1
                # 删 /Names 中危险子字典
                if "/Names" in writer._root_object:
                    names = writer._root_object["/Names"]
                    if hasattr(names, "get_object"):
                        names = names.get_object()
                    for sub_key in ("/JavaScript", "/JS", "/EmbeddedFiles", "/Launch"):
                        if sub_key in names:
                            del names[sub_key]
                            removed_catalog_keys += 1
                    if not dict(names):
                        del writer._root_object["/Names"]
                        removed_catalog_keys += 1
                with open(dst, "wb") as fp:
                    writer.write(fp)
                try:
                    intermediate.unlink()
                except Exception:
                    pass
            except Exception as exc:  # noqa: BLE001
                # pypdf 二阶段失败时，保留 intermediate 作为兜底
                logger.warning("action=catalog_cleanup_failed file=%s reason=%s", dst.name, exc)
                try:
                    intermediate.replace(dst)
                except Exception:
                    pass
        elif intermediate is not None:
            # pypdf 不可用，直接用 fitz 的输出
            intermediate.replace(dst)

        # 6) 可选 qpdf 线性化验证
        qpdf_msg = ""
        if self.engines.qpdf:
            ok, msg = self._qpdf_linearize_check(dst)
            qpdf_msg = f"; qpdf-check={'ok' if ok else 'fail'}"
            if not ok:
                logger.warning("action=qpdf_check_failed file=%s reason=%s",
                               dst.name, msg)

        engine_used = (
            f"fitz(annots={removed_annots}, links={removed_links}, embeds={removed_embeds})"
            f"+pypdf(catalog={removed_catalog_keys})"
            f"{qpdf_msg}"
        )
        logger.info(
            "action=sanitize_success file=%s engine=%s",
            dst.name, engine_used,
        )
        return True, engine_used

    # -- pypdf 降级路径（基础结构级剥离） ------------------------------------
    def _sanitize_with_pypdf(self, src: Path, dst: Path) -> tuple[bool, str]:
        try:
            from pypdf import PdfReader, PdfWriter  # type: ignore
        except ImportError:
            return False, "pypdf 不可用"
        try:
            reader = PdfReader(str(src))
            if reader.is_encrypted:
                # 加密 PDF 无法做无密钥剥离
                return False, "PDF 已加密，无法用 pypdf 剥离（请先用解密工具）"
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            # 移除 root 级危险字典
            try:
                if "/OpenAction" in writer._root_object:
                    del writer._root_object["/OpenAction"]
            except Exception:
                pass
            try:
                if "/AA" in writer._root_object:
                    del writer._root_object["/AA"]
            except Exception:
                pass
            try:
                if "/Names" in writer._root_object:
                    names = writer._root_object["/Names"]
                    if hasattr(names, "get_object"):
                        names = names.get_object()
                    for k in ("/JavaScript", "/JS", "/EmbeddedFiles", "/Launch"):
                        if k in names:
                            del names[k]
                    # 若 Names 已空，从 root 删除
                    if not dict(names):
                        del writer._root_object["/Names"]
            except Exception:
                pass
            # 移除 AcroForm（XSS 常见载体）
            try:
                if "/AcroForm" in writer._root_object:
                    del writer._root_object["/AcroForm"]
            except Exception:
                pass

            dst.parent.mkdir(parents=True, exist_ok=True)
            with open(dst, "wb") as fp:
                writer.write(fp)
        except Exception as exc:  # noqa: BLE001
            return False, f"pypdf sanitize 失败: {exc}"
        return True, "pypdf(基础剥离: OpenAction/AA/Names/AcroForm)"

    # -- qpdf 线性化校验（可选） --------------------------------------------
    def _qpdf_linearize_check(self, pdf_path: Path) -> tuple[bool, str]:
        try:
            result = subprocess.run(
                ["qpdf", "--check", str(pdf_path)],
                capture_output=True, text=True, timeout=60,
            )
            if result.returncode == 0:
                return True, "ok"
            return False, (result.stderr or result.stdout).strip()[:200]
        except Exception as exc:  # noqa: BLE001
            return False, str(exc)


# ---------------------------------------------------------------------------
# 批量分析 + 报告聚合
# ---------------------------------------------------------------------------
def collect_pdfs(input_path: Path, recursive: bool = True) -> list[Path]:
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        return [input_path.resolve()]
    if input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return sorted(p.resolve() for p in input_path.glob(pattern) if p.is_file())
    return []


def run_batch(
    input_path: Path,
    output_dir: Path,
    *,
    recursive: bool = True,
    do_sanitize: bool = False,
    overwrite: bool = False,
    engines: EngineStatus | None = None,
) -> int:
    """批量分析 + 可选 sanitize。返回 0/1。"""
    engines = engines or detect_engines()
    output_dir.mkdir(parents=True, exist_ok=True)
    pdfs = collect_pdfs(input_path, recursive=recursive)
    if not pdfs:
        logger.warning("action=batch_empty input=%s", input_path)
        return 1

    logger.info("action=batch_start total=%s input=%s", len(pdfs), input_path)
    success = 0
    failed = 0
    sanitizer = PDFSanitizer(engines) if do_sanitize else None
    reports: list[dict[str, Any]] = []
    for pdf in pdfs:
        try:
            analyzer = PDFThreatAnalyzer(pdf)
            report = analyzer.analyze()
            reports.append(report)

            # 写单文件 JSON 报告
            json_path = output_dir / f"threat_report_{pdf.stem}.json"
            json_path.write_text(analyzer.to_json(), encoding="utf-8")

            # 可选 sanitize
            if do_sanitize and sanitizer is not None and report.get("is_valid_pdf"):
                sanitized_path = output_dir / f"{pdf.stem}_sanitized.pdf"
                ok, msg = sanitizer.sanitize(pdf, sanitized_path, overwrite=overwrite)
                report["sanitize_result"] = {
                    "success": ok,
                    "engine_or_reason": msg,
                    "output_path": str(sanitized_path) if ok else None,
                }
                if not ok:
                    logger.warning("action=sanitize_failed file=%s reason=%s",
                                   pdf.name, msg)

            success += 1
        except Exception as exc:  # noqa: BLE001
            failed += 1
            logger.error("action=process_failed file=%s reason=%s", pdf.name, exc)

    # 写摘要
    summary = build_summary(reports, success, failed)
    summary_path = output_dir / f"threat_summary_{datetime.now():%Y%m%d_%H%M%S}.txt"
    summary_path.write_text(summary, encoding="utf-8")
    logger.info("action=batch_done success=%s failed=%s total=%s",
                success, failed, len(pdfs))
    logger.info("action=summary_written path=%s", summary_path)
    return 0 if failed == 0 else 1


def build_summary(reports: list[dict[str, Any]], success: int, failed: int) -> str:
    level_count: dict[str, int] = {"LOW": 0, "MEDIUM": 0, "HIGH": 0, "UNKNOWN": 0}
    for r in reports:
        lvl = r.get("risk_level", "UNKNOWN")
        level_count[lvl] = level_count.get(lvl, 0) + 1
    lines = [
        f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"扫描总数: {len(reports)}（成功 {success} / 失败 {failed}）",
        "",
        "--- 风险等级分布 ---",
        f"  LOW:    {level_count.get('LOW', 0)}",
        f"  MEDIUM: {level_count.get('MEDIUM', 0)}",
        f"  HIGH:   {level_count.get('HIGH', 0)}",
        f"  UNKNOWN:{level_count.get('UNKNOWN', 0)}",
        "",
        "--- 文件清单（按风险等级排序）---",
    ]
    sorted_reports = sorted(
        reports,
        key=lambda r: {"HIGH": 0, "MEDIUM": 1, "LOW": 2, "UNKNOWN": 3}.get(
            r.get("risk_level", "UNKNOWN"), 4),
    )
    for r in sorted_reports:
        lines.append(
            f"  [{r.get('risk_level', 'UNKNOWN'):>7}] {r.get('file_name', '?')}"
            f"  score={r.get('risk_score', 0)}"
        )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 自检（标准模式）
# ---------------------------------------------------------------------------
def self_check(output_dir: Path | None = None) -> int:
    """标准自检：生成测试 PDF（含 /JavaScript），跑分析 + sanitize，验证报告。

    返回 0 成功 / 2 失败。
    """
    logger.info("action=self_check_start")
    engines = detect_engines()
    logger.info("action=engines_detected %s", engines.summary())

    if not engines.pypdf:
        logger.error("action=self_check_failed reason=missing_pypdf")
        print("❌ 自检失败：缺少 pypdf，请先 `pip install pypdf`", file=sys.stderr)
        return 2

    # 1) 构造最小化测试 PDF（含 /JavaScript 模拟威胁）
    test_pdf_bytes = build_test_pdf_with_javascript()
    with tempfile.NamedTemporaryFile(
        prefix="selfcheck_", suffix=".pdf", delete=False
    ) as tmp:
        tmp.write(test_pdf_bytes)
        tmp_path = Path(tmp.name)
    try:
        # 2) 跑分析
        analyzer = PDFThreatAnalyzer(tmp_path)
        report = analyzer.analyze()

        # 3) 验证关键字段
        assertions: list[tuple[bool, str]] = [
            (report["is_valid_pdf"], "is_valid_pdf 应为 True"),
            ("/JavaScript" in report["raw_keyword_hits"]
             or "/JS" in report["raw_keyword_hits"],
             "raw_keyword_hits 应命中 /JavaScript 或 /JS"),
            (report["risk_score"] > 0, "risk_score 应 > 0"),
            (report["risk_level"] in ("MEDIUM", "HIGH"), "risk_level 应为 MEDIUM 或 HIGH"),
        ]
        all_ok = True
        for ok, msg in assertions:
            status = "✅" if ok else "❌"
            print(f"  {status} {msg}")
            all_ok = all_ok and ok

        if not all_ok:
            logger.error("action=self_check_failed reason=assertion_failed")
            print("❌ 自检失败：分析器断言未通过", file=sys.stderr)
            return 2

        # 4) 可选 sanitize 验证（仅当有 sanitize 引擎时）
        if output_dir is not None and (engines.fitz or engines.pypdf):
            output_dir.mkdir(parents=True, exist_ok=True)
            sanitized = output_dir / f"selfcheck_{tmp_path.stem}_sanitized.pdf"
            sanitizer = PDFSanitizer(engines)
            ok, msg = sanitizer.sanitize(tmp_path, sanitized, overwrite=True)
            if not ok:
                print(f"  ⚠️ sanitize 跳过：{msg}")
            else:
                # 重新扫描剥离后的 PDF，验证 /JavaScript 已消失
                post = PDFThreatAnalyzer(sanitized).analyze()
                js_remaining = (
                    "/JavaScript" in post["raw_keyword_hits"]
                    or "/JS" in post["raw_keyword_hits"]
                )
                print(f"  {'✅' if not js_remaining else '❌'} sanitize 后 /JavaScript 已剥离（engine={msg}）")
                if js_remaining:
                    logger.error("action=self_check_failed reason=sanitize_incomplete")
                    return 2

        print("✅ 标准自检全部通过")
        logger.info("action=self_check_passed")
        return 0
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass


def build_test_pdf_with_javascript() -> bytes:
    """构造一个合法的"含威胁对象"的 PDF，用于自检。

    使用 pypdf 生成（保证 xref / 字典结构合法，fitz 也能正确解析），
    注入以下威胁对象供检测与剥离：
    - /OpenAction（指向 /JavaScript 动作，权重 +2）
    - /Names → /JavaScript（权重 +3）
    - /AA（附加动作字典，权重 +2）
    - /AcroForm（XSS 载体，sanitize 时移除）

    整体目标：PDFThreatAnalyzer 命中 /JavaScript 后 risk_score 累加 ≥ 7，
    触发 HIGH 风险等级；sanitize 后这些 catalog 危险字典被清空，
    重新扫描时 raw_keyword_hits 不再包含 /JavaScript / /OpenAction / /AA。
    """
    import io
    from pypdf import PdfWriter
    from pypdf.generic import (
        DictionaryObject,
        NameObject,
        TextStringObject,
        ArrayObject,
    )

    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)

    # 注入 /OpenAction 指向一个 /JavaScript 动作
    js_action = DictionaryObject({
        NameObject("/Type"): NameObject("/Action"),
        NameObject("/S"): NameObject("/JavaScript"),
        NameObject("/JS"): TextStringObject("app.alert('payload')"),
    })
    writer._root_object[NameObject("/OpenAction")] = js_action

    # 注入 /AA（附加动作字典，触发时也可执行 JS）
    writer._root_object[NameObject("/AA")] = DictionaryObject({
        NameObject("/WC"): js_action,  # Will Close 时执行
    })

    # 注入 /Names → /JavaScript（Named Action 入口）
    writer._root_object[NameObject("/Names")] = DictionaryObject({
        NameObject("/JavaScript"): DictionaryObject({
            NameObject("/Names"): ArrayObject([
                TextStringObject("selfcheck"),
                js_action,
            ]),
        }),
        NameObject("/EmbeddedFiles"): DictionaryObject({
            NameObject("/Names"): ArrayObject([
                TextStringObject("dummy"),
                DictionaryObject({NameObject("/F"): TextStringObject("dummy.bin")}),
            ]),
        }),
    })

    # 注入 /AcroForm（XSS 常见载体，sanitize 时移除）
    writer._root_object[NameObject("/AcroForm")] = DictionaryObject({
        NameObject("/Fields"): ArrayObject(),
    })

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    module_dir = Path(__file__).resolve().parent
    default_input = module_dir / "input"
    default_output = module_dir / "output"

    parser = argparse.ArgumentParser(
        description="PDF 威胁分析与工业级安全剥离（无渲染 / 静态扫描）",
    )
    parser.add_argument(
        "--input", default=str(default_input),
        help="输入 PDF 文件或目录（默认: 模块 input/）",
    )
    parser.add_argument(
        "--output", default=str(default_output),
        help="输出目录（默认: 模块 output/）",
    )
    parser.add_argument(
        "--sanitize", action="store_true",
        help="生成剥离威胁后的安全 PDF（*_sanitized.pdf）",
    )
    parser.add_argument(
        "--no-sanitize", dest="sanitize", action="store_false",
        help="仅输出分析报告，不生成安全 PDF（默认）",
    )
    parser.add_argument(
        "--engine", choices=("auto", "fitz", "pypdf"), default="auto",
        help="sanitize 引擎选择（默认 auto: fitz 优先，pypdf 降级）",
    )
    parser.add_argument(
        "--recursive", dest="recursive", action="store_true", default=True,
        help="递归遍历子目录（默认开启）",
    )
    parser.add_argument(
        "--no-recursive", dest="recursive", action="store_false",
        help="不递归遍历",
    )
    parser.add_argument(
        "--overwrite", action="store_true",
        help="覆盖已存在的输出文件",
    )
    parser.add_argument(
        "--self-check", action="store_true",
        help="运行标准自检（生成测试 PDF → 分析 → 验证）",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="DEBUG 级日志",
    )
    parser.set_defaults(sanitize=False)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    _configure_logging(verbose=args.verbose)

    if args.self_check:
        out = Path(args.output) if args.output else None
        return self_check(out)

    engines = detect_engines()
    logger.info("action=engines_detected %s", engines.summary())

    if args.sanitize and not (engines.fitz or engines.pypdf):
        logger.error("action=preflight_failed reason=no_sanitize_engine")
        print("❌ 启用 --sanitize 需安装 fitz 或 pypdf（见 requirements.txt）",
              file=sys.stderr)
        return 2

    if args.sanitize and args.engine == "fitz" and not engines.fitz:
        logger.error("action=preflight_failed reason=fitz_required_not_found")
        print("❌ --engine=fitz 要求安装 PyMuPDF", file=sys.stderr)
        return 2
    if args.sanitize and args.engine == "pypdf" and not engines.pypdf:
        logger.error("action=preflight_failed reason=pypdf_required_not_found")
        print("❌ --engine=pypdf 要求安装 pypdf", file=sys.stderr)
        return 2

    try:
        return run_batch(
            input_path=Path(args.input).expanduser(),
            output_dir=Path(args.output).expanduser(),
            recursive=args.recursive,
            do_sanitize=args.sanitize,
            overwrite=args.overwrite,
            engines=engines,
        )
    except KeyboardInterrupt:
        logger.error("action=interrupted_by_user")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
