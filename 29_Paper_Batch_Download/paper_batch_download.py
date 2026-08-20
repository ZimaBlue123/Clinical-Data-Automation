"""
文献批量下载与重塑矩阵 (Protocol V3.0 - 终极融合版)
Vibe: Academic Cyberpunk

功能：
 1. 根据 DOI / PMID / 标题 / URL 批量突围下载文献 (OA + Sci-Hub 矩阵)。
 2. 下载落地后，自动调用 PyMuPDF 与 OCR 引擎扫描 PDF 内容。
 3. 提取真实标题、切除副标题、挂载出版年份，完成物理重命名与归档。
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import random
import re
import sys
import time
import shutil
import urllib3
from collections import defaultdict
from pathlib import Path
from collections.abc import Iterable
from urllib.parse import urljoin, urlparse

import requests
import fitz  # PyMuPDF
from PIL import Image
from tqdm import tqdm

# 屏蔽可能因为关闭 SSL 验证而产生的烦人警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logger = logging.getLogger(__name__)

# --- 视觉矩阵依赖 (容错加载) ---
try:
    import pytesseract
    # [!] Windows 用户如果未配置环境变量，请修改这里的路径
    # pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    logger.warning("action=ocr_backend_unavailable backend=pytesseract")

# 标题中默认保持小写的功能词（首词/尾词除外）
LOWERCASE_TITLE_WORDS = {
    "a", "an", "the", "and", "but", "or", "nor", "so", "yet",
    "as", "at", "by", "for", "in", "of", "on", "per", "to", "via", "vs", "v",
}

DANGLING_TOXINS = {
    "and", "or", "of", "for", "to", "in", "on", "with", "by", "from",
    "the", "a", "an", "at", "as", "what", "which", "that", "is", "are",
}


# ==========================================
# 模块一：本地文献重塑协议 (PDF Sanitizer)
# ==========================================
class PDFSanitizer:
    def __init__(self, temp_dir: Path, output_dir: Path):
        self.temp_dir = temp_dir
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _smart_title_case(text: str) -> str:
        """英文标题格式化：中间功能词保持小写，首尾词强制首字母大写。"""
        raw_words = text.split()
        if not raw_words:
            return ""

        formatted_words = []
        last_idx = len(raw_words) - 1
        for idx, word in enumerate(raw_words):
            lower_word = word.lower()
            if idx not in (0, last_idx) and lower_word in LOWERCASE_TITLE_WORDS:
                formatted_words.append(lower_word)
            else:
                formatted_words.append(lower_word.capitalize())
        return " ".join(formatted_words)

    @staticmethod
    def _simplify_filename(name: str, max_words: int = 15, max_chars: int = 40) -> str:
        """核心文本手术刀：副标题截断，双语自适应解析，全角标点粉碎，边缘修剪"""
        # 0. 副标题物理切除
        name = re.split(r"[:：]", name)[0]

        # 1. 阵营嗅探：检测是否包含中文字符
        is_chinese = bool(re.search(r"[\u4e00-\u9fa5]", name))

        # 2. 物理切除：切除各类括号及内部噪点
        cleaned = re.sub(r"[\[\(（【《].*?[\]\)）】》]", "", name)

        # 3. 抹除非法路径字符，替换为空格
        cleaned = re.sub(r"[^\w\s\-\u4e00-\u9fa5]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if is_chinese:
            # 中文高压压缩
            cleaned_cjk = cleaned.replace(" ", "")
            if not cleaned_cjk:
                return "未命名文献_Untitled"
            return cleaned_cjk[:max_chars]
        # 纯英文边缘修剪
        cleaned_en = PDFSanitizer._smart_title_case(cleaned)
        words = cleaned_en.split()[:max_words]
        # 切除坏死边缘
        while words and words[-1].lower() in DANGLING_TOXINS:
            words.pop()
        while words and words[0].lower() in DANGLING_TOXINS:
            words.pop(0)

        if not words:
            fallback = cleaned_en[:50].strip().replace(" ", "_")
            return fallback if fallback else "Untitled_Document"

        return "_".join(words)

    def _dedupe_filename(self, base_name: str) -> str:
        """量子态文件覆盖防御"""
        candidate = base_name
        counter = 1
        while (self.output_dir / f"{candidate}.pdf").exists():
            candidate = f"{base_name}_{counter}"
            counter += 1
        return candidate

    @staticmethod
    def _optical_scan(doc: fitz.Document) -> str:
        """视觉皮层：抽取首页渲染为图像，交由 Tesseract 识别"""
        if not OCR_AVAILABLE or len(doc) == 0:
            return ""
        try:
            pix = doc[0].get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(img, lang="chi_sim+eng")
            return text.strip()
        except Exception:
            return ""

    @staticmethod
    def _extract_title_by_visual_hierarchy(page: fitz.Page) -> str:
        """视觉层级引擎：通过物理字号锁定真实标题"""
        try:
            blocks = page.get_text("dict").get("blocks", [])
            text_sizes = []

            for b in blocks:
                if "lines" in b:
                    for line in b["lines"]:
                        for span in line["spans"]:
                            text = span.get("text", "").strip()
                            size = span.get("size", 0)
                            if text:
                                text_sizes.append((size, text))

            if not text_sizes:
                return ""

            size_map = {}
            for size, text in text_sizes:
                s = round(size, 1)
                if s not in size_map:
                    size_map[s] = []
                size_map[s].append(text)

            sorted_sizes = sorted(size_map.keys(), reverse=True)

            # 黑名单：免疫期刊页眉噪点
            blacklist = {"majorarticle", "researcharticle", "reviewarticle", "clinicalinfectiousdiseases"}

            for s in sorted_sizes:
                candidate = " ".join(size_map[s]).strip()
                compressed_cand = candidate.lower().replace(" ", "")
                is_toxic = any(noise in compressed_cand for noise in blacklist)

                # 若无毒且长度合理，直接将其判定为标题
                if len(candidate) >= 8 and not is_toxic:
                    return candidate
        except Exception:
            pass
        return ""

    def _scan_payload(self, pdf_path: Path) -> tuple[str, str]:
        """神经元探针：融合元数据、拓扑探测与 OCR 后备能源"""
        title = ""
        year = "XXXX"
        text_payload = ""
        hierarchy_title = ""

        try:
            with fitz.open(pdf_path) as doc:
                meta_title = doc.metadata.get("title", "").strip()

                if len(doc) > 0:
                    text_payload = doc[0].get_text("text").strip()
                    hierarchy_title = self._extract_title_by_visual_hierarchy(doc[0])

                if len(text_payload) < 20:
                    text_payload = self._optical_scan(doc)

                # --- 标题窃取 ---
                if hierarchy_title:
                    title = hierarchy_title
                elif meta_title and len(meta_title) > 2 and "microsoft word" not in meta_title.lower():
                    title = meta_title
                else:
                    lines = [line.strip() for line in text_payload.splitlines() if line.strip()]
                    for line in lines[:15]:
                        if len(line) >= 5:
                            title = line
                            break

                # --- 时间线锚定 ---
                text_head = text_payload[:2000]
                pattern = r"(?:©|copyright|published|vol\.|date|年|出版|收稿).*?\b(19[5-9]\d|20[0-2]\d)\b"
                explicit_year = re.search(pattern, text_head, re.IGNORECASE)

                if explicit_year:
                    year = explicit_year.group(1)
                else:
                    creation_date = doc.metadata.get("creationDate", "")
                    meta_year_match = re.search(r"D:(\d{4})", creation_date)
                    if meta_year_match:
                        year = meta_year_match.group(1)
                    else:
                        fallback_year = re.search(r"\b(19[5-9]\d|20[0-2]\d)\b", text_head)
                        if fallback_year:
                            year = fallback_year.group(1)

        except Exception:
            pass # 如果文件损坏，静默失败，使用原文件名兜底

        return title or pdf_path.stem, year

    def execute(self) -> None:
        """执行重命名装配流水线"""
        pdf_targets = list(self.temp_dir.glob("*.pdf"))
        if not pdf_targets:
            return

        print(f"\n[+] 物理重塑协议激活 | 扫描目标: {len(pdf_targets)}")
        print(f"[+] 视觉引擎(OCR): {'在线' if OCR_AVAILABLE else '离线'}\n")

        for pdf_path in tqdm(pdf_targets, desc="Sanitizing & Moving", unit="file", ascii=" ▖▘▝▗▚▞█"):
            # 扫描与解析
            raw_title, year = self._scan_payload(pdf_path)
            simplified_name = self._simplify_filename(raw_title)

            # 挂载年份
            chronological_name = f"{simplified_name}-{year}"

            # 去重并移动
            final_safe_name = self._dedupe_filename(chronological_name)
            target_path = self.output_dir / f"{final_safe_name}.pdf"

            try:
                shutil.move(str(pdf_path), str(target_path))
            except Exception as e:
                tqdm.write(f"[!] {pdf_path.name} 移动失败: {e}")


# ==========================================
# 模块二：网络突围系统 (Paper Downloader)
# ==========================================
HEADERS_REQ = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

DOI_PATTERN = re.compile(r"^10\.\d{4,9}/[-._;()/:A-Z0-9]+$", re.I)
PMID_PATTERN = re.compile(r"^\d+$")

class PaperDownloader:
    def __init__(
        self,
        output_dir: Path,
        mailto: str | None = None,
        sleep_seconds: float = 1.2,
        proxy: str | None = None,
        safe_mode: bool = True,
        min_interval: float = 2.0,
        max_retries: int = 4,
        backoff_base: float = 1.0,
        mirror_cooldown_seconds: float = 120.0,
    ):
        self.output_dir = output_dir
        # 下载阶段先将文件放入专属缓存区
        self.temp_dir = output_dir / ".temp_downloads"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

        self.mailto = mailto
        self.sleep_seconds = max(sleep_seconds, 0.0)
        self.safe_mode = safe_mode
        self.min_interval = max(min_interval, 0.0)
        self.max_retries = max(max_retries, 0)
        self.backoff_base = max(backoff_base, 0.2)
        self.mirror_cooldown_seconds = max(mirror_cooldown_seconds, 10.0)
        self.session = requests.Session()
        self.session.headers.update(HEADERS_REQ)
        self.last_request_ts: dict[str, float] = defaultdict(float)
        self.host_blocked_until: dict[str, float] = defaultdict(float)
        self._cache_dirty = False
        self.cache_path = self.output_dir / ".request_cache.json"
        self.cache = self._load_cache()

        if proxy:
            self.session.proxies = {"http": proxy, "https": proxy}

        self.scihub_mirrors = [
            "https://sci-hub.se", "https://sci-hub.ru", "https://sci-hub.st",
            "https://sci-hub.red", "https://sci-hub.box", "https://sci-hub.ee",
            "https://sci-hub.mk", "https://sci-hub.al"
        ]

    def _load_cache(self) -> dict[str, dict[str, str]]:
        default_cache = {"title_to_doi": {}, "pmid_to_doi": {}, "doi_to_pdf": {}}
        if not self.cache_path.exists():
            return default_cache
        try:
            payload = json.loads(self.cache_path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                return default_cache
            for key in default_cache:
                if not isinstance(payload.get(key), dict):
                    payload[key] = {}
            return payload
        except Exception:
            return default_cache

    def _save_cache(self) -> None:
        if not self._cache_dirty:
            return
        try:
            self.cache_path.write_text(
                json.dumps(self.cache, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _wait_for_host_slot(self, host: str) -> None:
        if not self.safe_mode:
            return
        now = time.time()
        blocked_until = self.host_blocked_until.get(host, 0.0)
        if blocked_until > now:
            time.sleep(blocked_until - now)
            now = time.time()
        gap = now - self.last_request_ts.get(host, 0.0)
        if gap < self.min_interval:
            time.sleep(self.min_interval - gap)

    def _compute_retry_delay(self, attempt: int, retry_after: str | None) -> float:
        if retry_after:
            try:
                return max(float(retry_after), self.backoff_base)
            except ValueError:
                pass
        exp = self.backoff_base * (2 ** max(attempt - 1, 0))
        jitter = random.uniform(0.1, 0.8)
        return exp + jitter

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        host = urlparse(url).netloc.lower()
        retries = self.max_retries if self.safe_mode else 0
        timeout = kwargs.pop("timeout", 15)
        kwargs["timeout"] = timeout
        kwargs.setdefault("verify", False)

        last_exc: Exception | None = None
        for attempt in range(1, retries + 2):
            self._wait_for_host_slot(host)
            try:
                print(f"[DEBUG] Requesting {method} {url} (Attempt {attempt})", flush=True)
                response = self.session.request(method, url, **kwargs)
                print(f"[DEBUG] Response {response.status_code} for {url}", flush=True)
                self.last_request_ts[host] = time.time()

                if response.status_code in (429, 403) or response.status_code >= 500:
                    if attempt <= retries:
                        delay = self._compute_retry_delay(attempt, response.headers.get("Retry-After"))
                        response.close()
                        time.sleep(delay)
                        continue
                return response
            except requests.RequestException as exc:
                self.last_request_ts[host] = time.time()
                last_exc = exc
                if attempt <= retries:
                    time.sleep(self._compute_retry_delay(attempt, None))
                    continue
                raise
        if last_exc:
            raise last_exc
        raise RuntimeError(f"Request failed unexpectedly: {url}")

    @staticmethod
    def _classify_query(query: str) -> str:
        if DOI_PATTERN.match(query):
            return "doi"
        if PMID_PATTERN.match(query):
            return "pmid"
        if query.lower().startswith("http"):
            return "url"
        return "title"

    @staticmethod
    def _normalize_query(query: str) -> tuple[str, str | None]:
        cleaned = query.strip()
        
        # 提取 URL 中的 DOI
        doi_url_match = re.search(r"(?:https?://(?:dx\.)?doi\.org/)(10\.\d{4,9}/[-._;()/:A-Z0-9]+)", cleaned, re.I)
        if doi_url_match:
            return doi_url_match.group(1), "已从 URL 中提取 DOI"
            
        # 提取 URL 中的 PMID
        pubmed_url_match = re.search(r"(?:https?://pubmed\.ncbi\.nlm\.nih\.gov/)(\d+)", cleaned, re.I)
        if pubmed_url_match:
            return pubmed_url_match.group(1), "已从 URL 中提取 PMID"

        if cleaned.startswith("10") and not cleaned.startswith("10."):
            cleaned = f"10.{cleaned[2:]}"
        plos_match = re.match(r"10\.1371/(?P<suffix>.+)$", cleaned, re.I)
        if plos_match:
            suffix = plos_match.group("suffix")
            digits_match = re.search(r"(\d{6,})$", suffix)
            if digits_match and ("pone" in suffix.lower() or "one" in suffix.lower()):
                digits = digits_match.group(1)
                return f"10.1371/journal.pone.{digits}", "已修正常见 PLOS DOI 误写格式"
        if cleaned != query:
            return cleaned, "已修正 DOI 前缀缺失"
        return cleaned, None

    def _title_to_doi(self, title: str) -> str | None:
        cached = self.cache["title_to_doi"].get(title)
        if cached:
            return cached
        try:
            res = self._request("GET", "https://api.crossref.org/works", params={"query.title": title, "rows": 1}, timeout=12)
            res.raise_for_status()
            items = res.json().get("message", {}).get("items", [])
            if items:
                doi = items[0].get("DOI")
                if doi:
                    self.cache["title_to_doi"][title] = doi
                    self._cache_dirty = True
                return doi
        except Exception:
            pass
        return None

    def _pmid_to_doi(self, pmid: str) -> str | None:
        cached = self.cache["pmid_to_doi"].get(pmid)
        if cached:
            return cached
        try:
            url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
            res = self._request("GET", url, params={"db": "pubmed", "id": pmid, "retmode": "json"}, timeout=12)
            res.raise_for_status()
            doc = res.json().get("result", {}).get(pmid, {})
            for item in doc.get("articleids", []):
                if item.get("idtype") == "doi":
                    doi = item.get("value")
                    if doi:
                        self.cache["pmid_to_doi"][pmid] = doi
                        self._cache_dirty = True
                    return doi
        except Exception:
            pass
        return None

    def _unpaywall_pdf(self, doi: str) -> str | None:
        if not self.mailto:
            return None
        cached = self.cache["doi_to_pdf"].get(doi)
        if cached:
            return cached
        try:
            res = self._request("GET", f"https://api.unpaywall.org/v2/{doi}", params={"email": self.mailto}, timeout=12)
            res.raise_for_status()
            best = res.json().get("best_oa_location") or {}
            resolved = best.get("url_for_pdf") or best.get("url")
            if resolved:
                self.cache["doi_to_pdf"][doi] = resolved
                self._cache_dirty = True
            return resolved
        except Exception:
            return None

    @staticmethod
    def _plos_pdf_from_doi(doi: str) -> str | None:
        match = re.search(r"10\.1371/journal\.([a-z]+)\.", doi, re.I)
        if not match:
            return None
        journal_map = {"pone": "plosone", "pbio": "plosbiology", "pmed": "plosmedicine", "pcbi": "ploscompbiol", "pgen": "plosgenetics", "pntd": "plosntds", "ppat": "plospathogens"}
        journal = journal_map.get(match.group(1).lower())
        if not journal:
            return None
        return f"https://journals.plos.org/{journal}/article/file?id={doi}&type=printable"

    def _find_pdf_in_html(self, html: str, base_url: str) -> str | None:
        candidates = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', html, flags=re.I)
        if not candidates:
            candidates = re.findall(r'src=["\']([^"\']+\.pdf[^"\']*)["\']', html, flags=re.I)
        if candidates:
            return urljoin(base_url, candidates[0])
        return None

    def _doi_landing_pdf(self, doi: str) -> str | None:
        try:
            res = self._request("GET", f"https://doi.org/{doi}", timeout=15, allow_redirects=True)
            return self._find_pdf_in_html(res.text, res.url)
        except Exception:
            return None

    def _scihub_pdf(self, doi: str) -> str | None:
        for mirror in self.scihub_mirrors:
            try:
                host = urlparse(mirror).netloc.lower()
                if self.safe_mode and self.host_blocked_until.get(host, 0.0) > time.time():
                    continue
                res = self._request("GET", f"{mirror}/{doi}", timeout=12)
                res.raise_for_status()
                match = re.search(r'(?:iframe|embed|object)[^>]+(?:id=["\']pdf["\']|type=["\']application/pdf["\'])[^>]*src=["\'](.*?)["\']', res.text, re.I)
                if not match:
                    match = re.search(r"location\.href=['\"](.*?)['\"]", res.text, re.I)
                if match:
                    pdf_url = match.group(1)
                    if pdf_url.startswith('//'):
                        pdf_url = 'https:' + pdf_url
                    elif pdf_url.startswith('/'):
                        pdf_url = f"{mirror}{pdf_url}"
                    elif not pdf_url.startswith('http'):
                        pdf_url = f"{mirror}/{pdf_url}"
                    self.scihub_mirrors.remove(mirror)
                    self.scihub_mirrors.insert(0, mirror)
                    return pdf_url
            except Exception:
                if self.safe_mode:
                    self.host_blocked_until[host] = time.time() + self.mirror_cooldown_seconds
                continue
        return None

    def _url_to_pdf(self, url: str) -> str | None:
        try:
            with self._request("GET", url, timeout=15, stream=True, allow_redirects=True) as res:
                content_type = res.headers.get('Content-Type', '').lower()
                if 'application/pdf' in content_type:
                    return res.url
                return self._find_pdf_in_html(res.text, res.url)
        except Exception:
            return None

    def _download_pdf(self, pdf_url: str, filename: str) -> None:
        # 下载到临时缓存区
        path = self.temp_dir / f"{filename}.pdf"
        with self._request("GET", pdf_url, stream=True, timeout=20) as res:
            res.raise_for_status()
            with path.open("wb") as f:
                for chunk in res.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

    def download(self, queries: Iterable[str]) -> dict[str, list[tuple[str, str]]]:
        results = {"success": [], "failed": []}
        query_list = [q.strip() for q in queries if q.strip()]

        proxy_status = self.session.proxies.get('http', '未使用 (直接连接)')
        logger.info("action=download_start targets=%s proxy=%s", len(query_list), proxy_status)

        for query in tqdm(query_list, desc="Downloading", unit="paper"):
            normalized, note = self._normalize_query(query)
            if note:
                tqdm.write(f"[i] 输入修正: {query} -> {normalized}")
            query = normalized
            q_type = self._classify_query(query)
            target = query

            if q_type == "title":
                target = self._title_to_doi(query) or ""
                if not target:
                    results["failed"].append((query, "Title to DOI mapping failed"))
                    time.sleep(self.sleep_seconds)
                    continue

            if q_type == "pmid":
                target = self._pmid_to_doi(query) or ""
                if not target:
                    results["failed"].append((query, "PMID to DOI mapping failed"))
                    time.sleep(self.sleep_seconds)
                    continue
                q_type = "doi"

            pdf_url = None
            if q_type == "url":
                pdf_url = self._url_to_pdf(target)
            elif q_type == "doi":
                pdf_url = (self._unpaywall_pdf(target) or self._plos_pdf_from_doi(target) or self._doi_landing_pdf(target) or self._scihub_pdf(target))
                if pdf_url:
                    self.cache["doi_to_pdf"][target] = pdf_url
                    self._cache_dirty = True

            if pdf_url:
                try:
                    # 临时文件名，使用 DOI 或原 URL 替换非法字符
                    temp_name = re.sub(r"[\\/*?:\"<>|]", "_", target)[:150]
                    self._download_pdf(pdf_url, temp_name)
                    results["success"].append((query, "Downloaded"))
                except Exception as e:
                    logger.warning("action=download_failed query=%s reason=%s", query, e)
                    results["failed"].append((query, f"Download failed: {e}"))
            else:
                results["failed"].append((query, "PDF not found in OA or Sci-Hub Matrix"))

            time.sleep(self.sleep_seconds)

        # ---------------------------------------------
        # 下载阶段结束，移交控制权给清洗模块 (Sanitizer)
        # ---------------------------------------------
        sanitizer = PDFSanitizer(temp_dir=self.temp_dir, output_dir=self.output_dir)
        sanitizer.execute()

        # 清理无用的临时缓存区
        try:
            shutil.rmtree(self.temp_dir)
        except Exception:
            logger.debug("action=temp_cleanup_failed dir=%s", self.temp_dir, exc_info=True)

        self._save_cache()
        self._print_report(results)
        return results

    @staticmethod
    def _print_report(results: dict[str, list[tuple[str, str]]]) -> None:
        logger.info(
            "action=download_summary success=%s failed=%s",
            len(results["success"]),
            len(results["failed"]),
        )
        if results["failed"]:
            for item, reason in results["failed"]:
                logger.warning("action=download_failure_detail query=%s reason=%s", item[:40], reason)


# ==========================================
# CLI 入口与多行输入协议
# ==========================================
def normalize_path(value: str) -> Path:
    return Path(value.strip().strip('"').strip("'"))

def load_queries_from_file(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"未找到文件: {path}")
    return [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip() and not line.strip().startswith("#")]

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="终极文献下载与重塑装配线")
    parser.add_argument("--queries", nargs="*", help="查询列表（DOI/PMID/Title/URL）")
    parser.add_argument("--file", default=None, help="包含查询列表的文本文件")
    parser.add_argument("--output", default=None, help="输出归档目录（默认 output/）")
    parser.add_argument("--mailto", default=None, help="Unpaywall 邮箱（用于提升 OA 命中率）")
    parser.add_argument("--sleep", type=float, default=1.2, help="请求间隔秒数")
    parser.add_argument("--proxy", default=None, help="网络代理地址，例如：http://127.0.0.1:15715")
    parser.add_argument(
        "--safe-mode",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="启用防限流安全模式（主机限速、重试退避、镜像冷却；默认开启）",
    )
    parser.add_argument("--min-interval", type=float, default=2.0, help="同一主机请求最小间隔秒数（safe-mode 下生效）")
    parser.add_argument("--max-retries", type=int, default=4, help="单请求最大重试次数（safe-mode 下生效）")
    parser.add_argument("--backoff-base", type=float, default=1.0, help="指数退避基础秒数（safe-mode 下生效）")
    parser.add_argument("--mirror-cooldown", type=float, default=120.0, help="Sci-Hub 镜像失败冷却秒数（safe-mode 下生效）")
    return parser.parse_args()

def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"
    args = parse_args()
    queries: list[str] = []

    if args.queries:
        queries.extend(args.queries)
    if args.file:
        file_path = normalize_path(args.file)
        if not file_path.is_absolute():
            file_path = Path.cwd() / file_path
        queries.extend(load_queries_from_file(file_path))

    if not queries:
        if not sys.stdin.isatty():
            logger.error("action=input_missing mode=non_interactive")
            sys.exit(1)

        logger.info("action=await_stdin_queries")
        while True:
            try:
                line = input().strip()
                if not line:
                    break
                queries.append(line)
            except EOFError:
                break

        if not queries:
            logger.error("action=input_missing mode=interactive")
            sys.exit(1)

    out_dir = Path(args.output) if args.output else output_dir
    if not out_dir.is_absolute():
        out_dir = output_dir / out_dir

    downloader = PaperDownloader(
        out_dir,
        mailto=args.mailto,
        sleep_seconds=args.sleep,
        proxy=args.proxy,
        safe_mode=args.safe_mode,
        min_interval=args.min_interval,
        max_retries=args.max_retries,
        backoff_base=args.backoff_base,
        mirror_cooldown_seconds=args.mirror_cooldown,
    )
    results = downloader.download(queries)
    raise SystemExit(0 if not results["failed"] else 1)

if __name__ == "__main__":
    main()
