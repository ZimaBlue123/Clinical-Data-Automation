# -*- coding: utf-8 -*-
"""
文献批量下载（全矩阵火力版 v2.2 - 协议穿透版）
Vibe: Academic Cyberpunk

功能：根据 DOI / PMID / 标题 / URL 批量下载文献。
逻辑：优先匹配 Open Access -> 兜底遍历 Sci-Hub 全节点矩阵。
特性：支持通过 CLI 动态挂载本地/远程代理隧道。
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import urllib3
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin

import requests
from tqdm import tqdm

# 屏蔽可能因为关闭 SSL 验证而产生的烦人警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

DOI_PATTERN = re.compile(r"^10\.\d{4,9}/[-._;()/:A-Z0-9]+$", re.I)
PMID_PATTERN = re.compile(r"^\d+$")


class PaperDownloader:
    def __init__(
        self, 
        output_dir: Path, 
        mailto: str | None = None, 
        sleep_seconds: float = 1.2,
        proxy: str | None = None
    ):
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.mailto = mailto
        self.sleep_seconds = max(sleep_seconds, 0.0)
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        
        # [!] 战术代理挂载 (Proxy Injection)
        if proxy:
            self.session.proxies = {
                "http": proxy,
                "https": proxy
            }
            # 如果你的代理软件拦截并重写了 SSL 证书导致报错，可以解除下方注释：
            # self.session.verify = False 

        # [!] Sci-Hub 战术节点矩阵 (已显式声明 https://)
        self.scihub_mirrors = [
            "https://sci-hub.se", 
            "https://sci-hub.ru", 
            "https://sci-hub.st",
            "https://sci-hub.red", 
            "https://sci-hub.box", 
            "https://sci-hub.ee",
            "https://sci-hub.mk", 
            "https://sci-hub.al"
        ]

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
        if cleaned.startswith("10") and not cleaned.startswith("10."):
            cleaned = f"10.{cleaned[2:]}"

        plos_match = re.match(r"10\.1371/(?P<suffix>.+)$", cleaned, re.I)
        if plos_match:
            suffix = plos_match.group("suffix")
            digits_match = re.search(r"(\d{6,})$", suffix)
            if digits_match and ("pone" in suffix.lower() or "one" in suffix.lower()):
                digits = digits_match.group(1)
                cleaned = f"10.1371/journal.pone.{digits}"
                return cleaned, "已修正常见 PLOS DOI 误写格式"

        if cleaned != query:
            return cleaned, "已修正 DOI 前缀缺失小数点"

        return cleaned, None

    def _title_to_doi(self, title: str) -> str | None:
        try:
            url = "https://api.crossref.org/works"
            res = self.session.get(url, params={"query.title": title, "rows": 1}, timeout=12)
            res.raise_for_status()
            items = res.json().get("message", {}).get("items", [])
            if items:
                return items[0].get("DOI")
        except Exception:
            return None
        return None

    def _pmid_to_doi(self, pmid: str) -> str | None:
        try:
            url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
            res = self.session.get(
                url,
                params={"db": "pubmed", "id": pmid, "retmode": "json"},
                timeout=12,
            )
            res.raise_for_status()
            data = res.json()
            doc = data.get("result", {}).get(pmid, {})
            for item in doc.get("articleids", []):
                if item.get("idtype") == "doi":
                    return item.get("value")
        except Exception:
            return None
        return None

    def _fetch_title_from_crossref(self, doi: str) -> str | None:
        try:
            url = f"https://api.crossref.org/works/{doi}"
            res = self.session.get(url, timeout=12)
            res.raise_for_status()
            title_list = res.json().get("message", {}).get("title", [])
            if title_list:
                return title_list[0]
        except Exception:
            return None
        return None

    @staticmethod
    def _simplify_title(title: str, max_words: int = 8) -> str:
        stopwords = {
            "the", "a", "an", "and", "or", "of", "for",
            "to", "in", "on", "with", "by", "from",
        }
        cleaned = re.sub(r"[\[\(].*?[\]\)]", "", title)
        cleaned = re.sub(r"[^\w\s-]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        words = [w for w in cleaned.split() if w.lower() not in stopwords]
        if not words:
            return cleaned
        return " ".join(words[:max_words])

    def _dedupe_filename(self, base_name: str) -> str:
        candidate = base_name
        counter = 1
        while (self.output_dir / f"{candidate}.pdf").exists():
            counter += 1
            candidate = f"{base_name}_{counter}"
        return candidate

    def _unpaywall_pdf(self, doi: str) -> str | None:
        if not self.mailto:
            return None
        try:
            url = f"https://api.unpaywall.org/v2/{doi}"
            res = self.session.get(url, params={"email": self.mailto}, timeout=12)
            res.raise_for_status()
            data = res.json()
            best = data.get("best_oa_location") or {}
            pdf_url = best.get("url_for_pdf") or best.get("url")
            if pdf_url:
                return pdf_url
        except Exception:
            return None
        return None

    @staticmethod
    def _plos_pdf_from_doi(doi: str) -> str | None:
        match = re.search(r"10\.1371/journal\.([a-z]+)\.", doi, re.I)
        if not match:
            return None
        journal_map = {
            "pone": "plosone", "pbio": "plosbiology",
            "pmed": "plosmedicine", "pcbi": "ploscompbiol",
            "pgen": "plosgenetics", "pntd": "plosntds",
            "ppat": "plospathogens",
        }
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
            res = self.session.get(f"https://doi.org/{doi}", timeout=15, allow_redirects=True)
            res.raise_for_status()
            return self._find_pdf_in_html(res.text, res.url)
        except Exception:
            return None

    def _scihub_pdf(self, doi: str) -> str | None:
        """[!] 暴力嗅探：遍历 Sci-Hub 矩阵寻找真实 PDF 地址"""
        for mirror in self.scihub_mirrors:
            try:
                # 矩阵头已自带 https://，直接拼接 DOI
                url = f"{mirror}/{doi}" 
                res = self.session.get(url, timeout=12)
                res.raise_for_status()
                
                # 正则剥离嵌套的 PDF 链接 (iframe / embed)
                match = re.search(
                    r'(?:iframe|embed|object)[^>]+(?:id=["\']pdf["\']|type=["\']application/pdf["\'])[^>]*src=["\'](.*?)["\']',
                    res.text, re.I
                )
                if not match:
                    # 尝试匹配备用按钮重定向
                    match = re.search(r"location\.href=['\"](.*?)['\"]", res.text, re.I)
                
                if match:
                    pdf_url = match.group(1)
                    # 处理 Sci-Hub 各种奇葩的相对路径
                    if pdf_url.startswith('//'):
                        pdf_url = 'https:' + pdf_url
                    elif pdf_url.startswith('/'):
                        pdf_url = f"{mirror}{pdf_url}"
                    elif not pdf_url.startswith('http'):
                        pdf_url = f"{mirror}/{pdf_url}"
                    
                    # 动态节点提权：将成功节点移至首位，榨干它的价值
                    self.scihub_mirrors.remove(mirror)
                    self.scihub_mirrors.insert(0, mirror)
                    
                    return pdf_url
            except Exception:
                continue # 当前节点阵亡，静默切下一个
        return None

    def _url_to_pdf(self, url: str) -> str | None:
        if url.lower().endswith(".pdf"):
            return url
        try:
            res = self.session.get(url, timeout=15)
            res.raise_for_status()
            return self._find_pdf_in_html(res.text, res.url)
        except Exception:
            return None

    @staticmethod
    def _safe_filename(text: str) -> str:
        cleaned = re.sub(r"[\\/*?:\"<>|]", "", text).strip()
        return cleaned[:160] if cleaned else "paper"

    def _download_pdf(self, pdf_url: str, filename: str) -> None:
        path = self.output_dir / f"{filename}.pdf"
        with self.session.get(pdf_url, stream=True, timeout=20) as res:
            res.raise_for_status()
            with path.open("wb") as f:
                for chunk in res.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

    def download(self, queries: Iterable[str]) -> dict[str, list[tuple[str, str]]]:
        results = {"success": [], "failed": []}
        query_list = [q.strip() for q in queries if q.strip()]

        proxy_status = self.session.proxies.get('http', '未使用 (直接连接)')
        print(f"\n[+] 任务开始 | 目标数量: {len(query_list)} | 输出: {self.output_dir}")
        print(f"[+] 隧道状态: {proxy_status}\n")

        for query in tqdm(query_list, desc="Downloading", unit="paper"):
            normalized, note = self._normalize_query(query)
            if note:
                tqdm.write(f"[i] 输入修正: {query} -> {normalized}（{note}）")
            query = normalized
            q_type = self._classify_query(query)
            target = query
            title_for_name = None

            if q_type == "title":
                title_for_name = query
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

            if q_type == "doi" and not title_for_name:
                title_for_name = self._fetch_title_from_crossref(target)

            pdf_url = None
            if q_type == "url":
                pdf_url = self._url_to_pdf(target)
            elif q_type == "doi":
                # [!] 核心路由池：OA API 优先 -> HTML解析 -> SciHub 全节点兜底破解
                pdf_url = (
                    self._unpaywall_pdf(target)
                    or self._plos_pdf_from_doi(target)
                    or self._doi_landing_pdf(target)
                    or self._scihub_pdf(target) 
                )

            if pdf_url:
                try:
                    if title_for_name:
                        simplified = self._simplify_title(title_for_name)
                        base = self._safe_filename(simplified)
                    else:
                        base = self._safe_filename(target)
                    name = self._dedupe_filename(base)
                    
                    self._download_pdf(pdf_url, name)
                    results["success"].append((query, f"ok -> {name}.pdf"))
                except Exception as e:
                    results["failed"].append((query, f"Download failed: {e}"))
            else:
                results["failed"].append((query, "PDF not found in OA or Sci-Hub Matrix"))

            time.sleep(self.sleep_seconds)

        self._print_report(results)
        return results

    @staticmethod
    def _print_report(results: dict[str, list[tuple[str, str]]]) -> None:
        print("\n" + "=" * 40)
        print(f"[*] 任务结束 | 成功: {len(results['success'])} | 失败: {len(results['failed'])}")
        if results["failed"]:
            print("\n[!] 失败明细:")
            for item, reason in results["failed"]:
                print(f" - {item[:40]}... : {reason}")
        print("=" * 40 + "\n")


def normalize_path(value: str) -> Path:
    cleaned = value.strip().strip('"').strip("'")
    return Path(cleaned)


def load_queries_from_file(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"未找到文件: {path}")
    lines = path.read_text(encoding="utf-8").splitlines()
    return [line.strip() for line in lines if line.strip() and not line.strip().startswith("#")]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量下载文献 (支持 OA, Sci-Hub 矩阵与代理隧道)")
    parser.add_argument("--queries", nargs="*", help="查询列表（DOI/PMID/Title/URL）")
    parser.add_argument("--file", default=None, help="包含查询列表的文本文件（每行一个）")
    parser.add_argument("--output", default=None, help="输出目录（默认 output/）")
    parser.add_argument("--mailto", default=None, help="Unpaywall 邮箱（建议填，用于提升 OA 命中率）")
    parser.add_argument("--sleep", type=float, default=1.2, help="请求间隔秒数（默认 1.2）")
    # [!] 新增代理参数，并在 Help 中以用户的截图端口作为示例
    parser.add_argument("--proxy", default=None, help="网络代理地址，例如：http://127.0.0.1:15715")
    return parser.parse_args()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

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
            print("未提供 --queries 或 --file，且当前为非交互环境。", file=sys.stderr)
            sys.exit(1)
        user_input = input("请输入查询（逗号分隔）：").strip()
        if not user_input:
            print("未输入查询。", file=sys.stderr)
            sys.exit(1)
        # 修改后 (同时兼容英文逗号和中文全角逗号)
        import re
        queries = [part.strip() for part in re.split(r'[,，]', user_input) if part.strip()]

    out_dir = Path(args.output) if args.output else output_dir
    if not out_dir.is_absolute():
        out_dir = output_dir / out_dir

    # 实例化时传入代理参数
    downloader = PaperDownloader(
        out_dir, 
        mailto=args.mailto, 
        sleep_seconds=args.sleep,
        proxy=args.proxy
    )
    downloader.download(queries)


if __name__ == "__main__":
    main()
