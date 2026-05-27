# -*- coding: utf-8 -*-
"""
本地文献重塑协议 (PDF Sanitizer v6.9 - 题目前缀剥离与缩写保留)
Vibe: Academic Cyberpunk
Engine: PyMuPDF | Chrono-Tracker | Visual Hierarchy | Subtitle Severance | OCR
"""
import re
import shutil
from pathlib import Path
import io
import argparse
import logging

# --- 基础依赖 ---
import fitz  # pyright: ignore[reportMissingImports]
from tqdm import tqdm
from PIL import Image

# 标题中默认保持小写的功能词（首词/尾词除外）
LOWERCASE_TITLE_WORDS = {
    "a",
    "an",
    "the",
    "and",
    "but",
    "or",
    "nor",
    "so",
    "yet",
    "as",
    "at",
    "by",
    "for",
    "in",
    "of",
    "on",
    "per",
    "to",
    "via",
    "vs",
    "v",
}

DANGLING_TOXINS = {
    "and",
    "or",
    "of",
    "for",
    "to",
    "in",
    "on",
    "with",
    "by",
    "from",
    "the",
    "a",
    "an",
    "at",
    "as",
    "what",
    "which",
    "that",
    "is",
    "are",
}

# FDA / ICH 等封面常见泛化大标题：最大字号往往是这一行，需剥离或跳过以便落到具体题目
GUIDANCE_COVER_PREFIXES: tuple[str, ...] = (
    "guidance for industry",
    "guidance for clinical investigators",
    "guidance for clinical trial sponsors",
    "guidance for sponsors",
    "draft guidance for industry",
    "guidance for industry and clinical investigators",
)

# 期刊封面页眉：字号常最大但不是论文题目，视觉层级需跳过
_JOURNAL_MASTHEAD_COMPRESSED: frozenset[str] = frozenset(
    {
        "plosone",
        "plosone.",
        "plosbiology",
        "plosmedicine",
        "plospathogens",
        "plosgenetics",
        "ploscompbiol",
        "nature",
        "science",
        "cell",
        "lancet",
        "nejm",
        "bmj",
        "jama",
        "pnas",
        "elife",
        "sciadv",
        "naturecommunications",
        "naturemedicine",
        "vaccine",
        "immunity",
        "lancet",
        "bmjopen",
        "plosone",
        "elsevier",
        "sciencedirect",
    }
)

# 文章类型行（全大写短行），其后才是正文标题
_ARTICLE_TYPE_LINE = re.compile(
    r"^(research article|review article|systematic review|meta-analysis|"
    r"editorial|correction|methods|resources|brief report|case report|"
    r"original research|original article|open access|peer-reviewed research|"
    r"short communication|letter|commentary|perspective|"
    r"review|research)\s*$",
    re.IGNORECASE,
)

# Elsevier / Springer 等「校样 / 待刊」页眉横幅（字号常最大，但不是论文题目）
_PUBLISHER_STATUS_RE = re.compile(
    r"^(article\s+in\s+press|accepted\s+manuscript|uncorrected\s+proof|"
    r"author'?s?\s+(accepted\s+)?manuscript|e[\-\s]?proof|"
    r"preprint|in\s+press|draft\s+manuscript|"
    r"available\s+online|just\s+accepted|forthcoming|"
    r"manuscript\s+in\s+press)\s*$",
    re.IGNORECASE,
)

_PUBLISHER_STATUS_COMPRESSED: frozenset[str] = frozenset(
    {
        "articleinpress",
        "acceptedmanuscript",
        "uncorrectedproof",
        "authorsacceptedmanuscript",
        "authorsmanuscript",
        "eproof",
        "preprint",
        "inpress",
        "justaccepted",
        "availableonline",
        "forthcoming",
    }
)

# 期刊卷期页眉行，如 Vaccine xxx (2014) xxx–xxx
_JOURNAL_VOL_HEADER = re.compile(
    r"^[A-Za-z][\w\s&.\-]{1,50}\s+[\dxX]{1,6}\s*\((19|20)\d{2}\)\s+[\dxX–\-—]+",
    re.IGNORECASE,
)

# 出版商页脚/页眉套话（非标题）
_BOILERPLATE_LINE = re.compile(
    r"(?i)^(contents lists available|journal homepage|www\.elsevier\.com|"
    r"sciencedirect|g model|received in revised form|"
    r"available online|copyright\s*©|all rights reserved|"
    r"please\s+cite\s+this\s+article|cite\s+this\s+article\s+in\s+press|"
    r"crossmark\.crossref\.org)",
)

# Elsevier「Please cite this article in press as:」引用提示（整行或视觉拼接前缀）
_CITATION_INSTRUCTION_RE = re.compile(
    r"(?i)^please\s+cite\s+this\s+article",
)
_CITATION_INSTRUCTION_IN_TEXT = re.compile(
    r"(?i)please\s+cite\s+this\s+article|cite\s+this\s+article\s+in\s+press",
)

# 视觉/拼接标题中，正文题目常见起始词（用于剥离 cite 套话与作者前缀）
_TITLE_BODY_START = re.compile(
    r"(?is)\b(?:clinical\s+evaluation|clinical\s+trial|randomized\s+trial|"
    r"systematic\s+review|meta[\-\s]?analysis|original\s+article|"
    r"efficacy\s+and\s+safety|phase\s+[iIvV\d]+|open[\-\s]?label|"
    r"a\s+novel|the\s+role\s+of|immunogenicity\s+of|safety\s+and\s+immunogenicity)\b",
)

# 纯作者姓行（无逗号）：Scheiermann Klinman
_AUTHOR_SURNAMES_ONLY = re.compile(
    r"^[A-Z][a-z'\-]{2,24}(?:\s+[A-Z][a-z'\-]{2,24}){1,4}$",
)

logger = logging.getLogger(__name__)

# --- 视觉矩阵依赖 (容错加载) ---
try:
    import pytesseract
    # 如果你是 Windows 用户且没有配置环境变量，请取消下方注释并修改路径
    # pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    logger.warning("action=ocr_backend_unavailable backend=pytesseract")


class PDFSanitizer:
    def __init__(
        self,
        input_dir: str = "input",
        output_dir: str = "output",
        *,
        recursive: bool = True,
        keep_structure: bool = True,
        overwrite: bool = False,
    ):
        self.base_dir = Path(__file__).resolve().parent
        self.input_dir = self.base_dir / input_dir
        self.output_dir = self.base_dir / output_dir
        self.recursive = recursive
        self.keep_structure = keep_structure
        self.overwrite = overwrite

        # 基础设施初始化
        self.input_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _normalize_ws_lower(s: str) -> str:
        return re.sub(r"\s+", " ", s.lower()).strip()

    @staticmethod
    def _strip_guidance_cover_prefix(raw: str) -> str:
        """若整段以泛化封面标题开头，去掉前缀，保留具体题目（同一行或拼接后的后半段）。"""
        t = raw.strip()
        if not t:
            return t
        norm = PDFSanitizer._normalize_ws_lower(t)
        for prefix in GUIDANCE_COVER_PREFIXES:
            if not norm.startswith(prefix):
                continue
            m = re.match(r"(?is)^\s*" + re.escape(prefix).replace(r"\ ", r"\s+"), t)
            if not m:
                continue
            rest = t[m.end() :].strip()
            if len(rest) >= 12:
                return rest
            return ""
        return t

    @staticmethod
    def _fda_specific_title_from_plain_text(text: str) -> str:
        """从首屏文本解析：Guidance for Industry 之后到固定套话前的具体标题。"""
        if not text:
            return ""
        m = re.search(
            r"(?is)Guidance\s+for\s+Industry[^\n]*\n+\s*(.+?)(?:\n\s*\n|This\s+guidance|FDA\s+is\s+issuing|Docket\s+No)",
            text[:6000],
        )
        if not m:
            return ""
        line = re.sub(r"\s+", " ", m.group(1).strip())
        return line if len(line) >= 12 else ""

    @staticmethod
    def _year_from_fda_docket(text: str) -> str | None:
        m = re.search(r"(?i)\bFDA-(\d{4})-[A-Z]-\d+\b", text)
        return m.group(1) if m else None

    @staticmethod
    def _compress_for_masthead(s: str) -> str:
        return re.sub(r"[^\w]+", "", s.lower())

    @staticmethod
    def _is_journal_masthead_only(candidate: str) -> bool:
        """是否为仅期刊名/页眉（非正文标题）。"""
        t = candidate.strip()
        if not t or len(t) < 4:
            return True
        comp = PDFSanitizer._compress_for_masthead(t)
        if comp in _JOURNAL_MASTHEAD_COMPRESSED:
            return True
        # 常见「Journal Name」短行：≤4 词且总长较短
        words = t.split()
        if len(words) <= 4 and len(t) <= 36 and not re.search(r"[.?:;]", t):
            if comp.endswith("journal") or "journalof" in comp:
                return True
        return False

    @staticmethod
    def _is_article_type_line(line: str) -> bool:
        return bool(_ARTICLE_TYPE_LINE.match(line.strip()))

    @staticmethod
    def _is_publisher_status_line(line: str) -> bool:
        """Elsevier 等「ARTICLE IN PRESS」横幅，非正文标题。"""
        s = line.strip()
        if not s:
            return False
        if _PUBLISHER_STATUS_RE.match(s):
            return True
        comp = PDFSanitizer._compress_for_masthead(s)
        if comp in _PUBLISHER_STATUS_COMPRESSED:
            return True
        # 视觉层级有时把多词横幅拼成一行
        if "articleinpress" in comp and len(comp) <= 24:
            return True
        return False

    @staticmethod
    def _is_citation_instruction_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return False
        if _CITATION_INSTRUCTION_RE.match(s):
            return True
        if _CITATION_INSTRUCTION_IN_TEXT.search(s) and len(s) < 200:
            return True
        return False

    @staticmethod
    def _is_toxic_title_candidate(text: str) -> bool:
        """视觉层级或拼接串含出版商套话 / 引用提示，不可直接作文件名。"""
        s = (text or "").strip()
        if not s:
            return False
        norm = PDFSanitizer._normalize_ws_lower(s)
        if _CITATION_INSTRUCTION_IN_TEXT.search(norm):
            return True
        if PDFSanitizer._is_publisher_status_line(s):
            return True
        if PDFSanitizer._is_citation_instruction_line(s):
            return True
        comp = PDFSanitizer._compress_for_masthead(s)
        if "pleasecitethisarticle" in comp or "citethisarticleinpress" in comp:
            return True
        return False

    @staticmethod
    def _looks_like_author_surnames_only(line: str) -> bool:
        """仅作者姓（2–5 个首字母大写词），无 of/for 等题目结构。"""
        s = line.strip()
        if not s or len(s) > 80 or " of " in s.lower():
            return False
        if not _AUTHOR_SURNAMES_ONLY.match(s):
            return False
        titleish = {
            "clinical",
            "evaluation",
            "vaccine",
            "cancer",
            "study",
            "trial",
            "review",
            "analysis",
            "disease",
            "infectious",
        }
        return not any(w.lower() in titleish for w in s.split())

    @staticmethod
    def _looks_like_scientific_title(text: str) -> bool:
        """是否已是正文题目（含 of/for 等），不可再按作者列表剥离前缀。"""
        s = (text or "").strip()
        if len(s) < 24:
            return False
        if re.search(r"\s+of\s+", s, re.I) or re.search(r"\s+for\s+", s, re.I):
            return True
        if re.search(
            r"(?i)\b(?:clinical|evaluation|efficacy|randomized|vaccine|trial|study|"
            r"oligonucleotide|adjuvant|infectious|cancer|immunogenicity)\b",
            s,
        ):
            return True
        return len(s.split()) >= 6

    @staticmethod
    def _strip_noise_prefix_from_title(raw: str) -> str:
        """去掉 Please cite… 等拼在题目前的 Elsevier 噪声（不剥离正文题目词）。"""
        t = re.sub(r"\s+", " ", (raw or "").strip())
        if not t:
            return t
        m = _TITLE_BODY_START.search(t)
        if m and m.start() > 0:
            head = t[: m.start()]
            if _CITATION_INSTRUCTION_IN_TEXT.search(head) or re.search(
                r"(?i)\bpress\b", head
            ) or re.search(
                r"(?i)\b(?:scheiermann|klinman|[A-Z][a-z]+,\s*[A-Z]\.)\b", head
            ):
                t = t[m.start() :].strip()
        t = re.sub(
            r"(?is)^(?:please\s+cite\s+this\s+article[^a-z]{0,40})+",
            "",
            t,
        ).strip()
        # 仅剥 cite 后、正文题目前的短作者块（禁止对含 of/CpG 的整段题目动刀）
        if not PDFSanitizer._looks_like_scientific_title(t):
            t = re.sub(
                r"^(?:(?:[A-Z][a-z'\-]{2,20})(?:\s*,\s*[A-Z]\.?){0,3}\s+){1,4}"
                r"(?=[A-Z][a-z]{5,}\s)",
                "",
                t,
            ).strip()
        return t

    @staticmethod
    def _should_prefer_academic(hierarchy: str, academic: str, text_payload: str) -> bool:
        """Elsevier 待刊：学术行解析优先于（更长的）视觉拼接。"""
        if not academic or len(academic) < 20:
            return False
        h = (hierarchy or "").strip()
        if not h:
            return True
        if PDFSanitizer._is_toxic_title_candidate(h):
            return True
        head = text_payload[:4000]
        if _CITATION_INSTRUCTION_IN_TEXT.search(head) and re.search(
            r"(?is)\barticle\s+in\s+press\b", head
        ):
            return True
        if len(h) > len(academic) + 8 and _CITATION_INSTRUCTION_IN_TEXT.search(h):
            return True
        return False

    @staticmethod
    def _looks_like_author_line(line: str) -> bool:
        """作者行：含通讯作者 *，或「名 姓, 名」模式。"""
        s = line.strip()
        if not s or len(s) > 240:
            return False
        if re.search(r"\*\s*$", s):
            return True
        if s.count(",") >= 1 and re.match(
            r"^[A-Z][\w\-\.']+(?:\s*,\s*[A-Z][\w\-\.']*)+",
            s,
        ) and len(s) < 120:
            return True
        if PDFSanitizer._looks_like_author_surnames_only(s):
            return True
        return False

    @staticmethod
    def _is_boilerplate_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return True
        if _BOILERPLATE_LINE.search(s):
            return True
        if _JOURNAL_VOL_HEADER.match(s):
            return True
        if re.match(r"^G\s+Model\b", s, re.I):
            return True
        if re.match(r"^[A-Z]{2,6}\s+\d{4,6}\s+\d", s):
            return True
        return False

    @staticmethod
    def _year_from_journal_header(text: str) -> str | None:
        """Elsevier 页眉：Vaccine xxx (2014) xxx–xxx"""
        m = re.search(
            r"(?i)\b[A-Za-z][\w\s&.\-]{2,48}\s+[\dxX]{0,6}\s*\((19[5-9]\d|20[0-2]\d)\)",
            text[:2500],
        )
        return m.group(1) if m else None

    @staticmethod
    def _academic_title_from_plain_text(text: str) -> str:
        """
        学术期刊首屏：跳过期刊名与 RESEARCH ARTICLE 等类型行，合并后续多行为标题，
        遇作者行（多逗号短名）、空段、邮箱/URL、纯数字起头的机构行时停止。
        """
        if not text:
            return ""
        raw_lines = text[:8000].splitlines()
        i = 0
        n = len(raw_lines)

        while i < n:
            s = raw_lines[i].strip()
            if not s:
                i += 1
                continue
            if (
                PDFSanitizer._is_journal_masthead_only(s)
                or PDFSanitizer._is_article_type_line(s)
                or PDFSanitizer._is_publisher_status_line(s)
                or PDFSanitizer._is_citation_instruction_line(s)
                or PDFSanitizer._is_boilerplate_line(s)
                or PDFSanitizer._looks_like_author_line(s)
                or PDFSanitizer._looks_like_author_surnames_only(s)
            ):
                i += 1
                continue
            break

        title_parts: list[str] = []
        while i < n:
            raw = raw_lines[i]
            s = raw.strip()
            if not s:
                if title_parts:
                    break
                i += 1
                continue
            if (
                PDFSanitizer._is_journal_masthead_only(s)
                or PDFSanitizer._is_article_type_line(s)
                or PDFSanitizer._is_publisher_status_line(s)
                or PDFSanitizer._is_citation_instruction_line(s)
                or PDFSanitizer._is_boilerplate_line(s)
                or (not title_parts and PDFSanitizer._looks_like_author_surnames_only(s))
            ):
                i += 1
                continue
            if title_parts and PDFSanitizer._looks_like_author_line(s):
                break
            if s.count(",") >= 2 and len(s) < 220 and re.search(r",\s*[A-Z][a-z]", s):
                break
            if "@" in s or re.search(r"https?://", s, re.I):
                break
            if re.match(r"^\d+\s+", s) and title_parts:
                break
            if s.lower().startswith("doi:") or s.lower().startswith("doi "):
                break
            if re.match(r"^citation\s*:", s, re.I):
                break
            if re.match(r"^Q\d+\b", s, re.I):
                s = re.sub(r"^Q\d+\s*", "", s, flags=re.I).strip()
                if not s:
                    i += 1
                    continue
            title_parts.append(s)
            if len(" ".join(title_parts)) > 520:
                break
            i += 1

        out = re.sub(r"\s+", " ", " ".join(title_parts)).strip()
        return out if len(out) >= 20 else ""

    @staticmethod
    def _preserve_scientific_token(word: str) -> str:
        """保留 CpG、mRNA、IL-6 等科学缩写原有大小写。"""
        w = word.strip()
        if not w:
            return w
        if re.match(r"^[A-Z][a-z]*[A-Z][\w\-]*$", w):
            return w
        if re.search(r"\d", w) and re.search(r"[A-Za-z]", w):
            return w
        return ""

    @staticmethod
    def _smart_title_case(text: str) -> str:
        """英文标题格式化：中间功能词保持小写，首尾词强制首字母大写。"""
        raw_words = text.split()
        if not raw_words:
            return ""

        formatted_words = []
        last_idx = len(raw_words) - 1
        for idx, word in enumerate(raw_words):
            sci = PDFSanitizer._preserve_scientific_token(word)
            if sci:
                formatted_words.append(sci)
                continue
            lower_word = word.lower()
            if idx not in (0, last_idx) and lower_word in LOWERCASE_TITLE_WORDS:
                formatted_words.append(lower_word)
            else:
                formatted_words.append(lower_word.capitalize())
        return " ".join(formatted_words)

    @staticmethod
    def _simplify_filename(name: str, max_words: int = 22, max_chars: int = 40) -> str:
        """核心文本手术刀：副标题截断，双语自适应解析，全角标点粉碎，边缘修剪"""
        # 0. [核心优化] 副标题物理切除：在遇到中英文冒号时进行硬截断
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
        else:
            # 纯英文边缘修剪
            cleaned_en = PDFSanitizer._smart_title_case(cleaned)
            words = cleaned_en.split()[:max_words]
            # 仅修剪首尾悬挂介词；保留题目内部的 of/for（如 evaluation of CpG）
            while len(words) > 2 and words[-1].lower() in DANGLING_TOXINS:
                words.pop()
            while len(words) > 2 and words[0].lower() in DANGLING_TOXINS:
                words.pop(0)

            if not words:
                fallback = cleaned_en[:50].strip().replace(" ", "_")
                return fallback if fallback else "Untitled_Document"
                
            return "_".join(words)

    def _dedupe_filename(self, target_dir: Path, base_name: str) -> str:
        """量子态文件覆盖防御"""
        candidate = base_name
        counter = 1
        while (target_dir / f"{candidate}.pdf").exists():
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
        except Exception as e:
            logger.warning("action=ocr_extract_failed reason=%s", e)
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
            
            if not text_sizes: return ""
            
            size_map = {}
            for size, text in text_sizes:
                s = round(size, 1)
                if s not in size_map: size_map[s] = []
                size_map[s].append(text)
                
            sorted_sizes = sorted(size_map.keys(), reverse=True)
            
            # 黑名单：免疫期刊页眉噪点 + 出版商待刊横幅
            blacklist = {
                "majorarticle",
                "researcharticle",
                "reviewarticle",
                "clinicalinfectiousdiseases",
                "articleinpress",
                "acceptedmanuscript",
                "uncorrectedproof",
                "sciencedirect",
                "elsevier",
                "pleasecitethisarticle",
                "citethisarticleinpress",
            }
            
            for s in sorted_sizes:
                candidate = " ".join(size_map[s]).strip()
                compressed_cand = candidate.lower().replace(" ", "")
                is_toxic = any(noise in compressed_cand for noise in blacklist)

                if len(candidate) < 8 or is_toxic:
                    continue

                if PDFSanitizer._is_journal_masthead_only(candidate):
                    continue
                if PDFSanitizer._is_article_type_line(candidate):
                    continue
                if PDFSanitizer._is_publisher_status_line(candidate):
                    continue
                if PDFSanitizer._is_citation_instruction_line(candidate):
                    continue
                if PDFSanitizer._is_toxic_title_candidate(candidate):
                    continue
                if PDFSanitizer._is_boilerplate_line(candidate):
                    continue

                refined = PDFSanitizer._strip_guidance_cover_prefix(candidate)
                refined = PDFSanitizer._strip_noise_prefix_from_title(refined)
                if refined and len(refined) >= 8 and not PDFSanitizer._is_toxic_title_candidate(refined):
                    return refined

                norm_one = PDFSanitizer._normalize_ws_lower(candidate)
                if any(norm_one == p or norm_one.startswith(p + " ") for p in GUIDANCE_COVER_PREFIXES):
                    continue
                cleaned = PDFSanitizer._strip_noise_prefix_from_title(candidate)
                if cleaned and len(cleaned) >= 8 and not PDFSanitizer._is_toxic_title_candidate(cleaned):
                    return cleaned
        except Exception:
            logger.debug("视觉层级提取失败", exc_info=True)
        return ""

    def _scan_payload(self, pdf_path: Path) -> tuple[str, str]:
        """神经元探针 V6.9：融合元数据、拓扑探测与 OCR 后备能源"""
        title = ""
        year = "XXXX"
        text_payload = ""
        hierarchy_title = ""
        
        try:
            with fitz.open(pdf_path) as doc:
                meta_title = doc.metadata.get("title", "").strip()
                
                # 尝试获取原生文本与视觉层级
                if len(doc) > 0:
                    text_payload = doc[0].get_text("text").strip()
                    hierarchy_title = self._extract_title_by_visual_hierarchy(doc[0])
                
                # 如果原生文本极少，触发 OCR
                if len(text_payload) < 20:
                    text_payload = self._optical_scan(doc)

                # --- 阶段 1: 标题窃取 (多维降维打击) ---
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

                text_head = text_payload[:4000]
                # FDA 封面：泛化层级已命中时，用首屏结构抽取具体题目
                if title and re.search(r"(?i)guidance\s+for\s+industry", title):
                    from_plain = self._fda_specific_title_from_plain_text(text_payload)
                    if from_plain:
                        title = from_plain

                # 出版商待刊 / 引用提示误命中时，清空以便后续学术解析接管
                if title and (
                    self._is_publisher_status_line(title)
                    or self._is_toxic_title_candidate(title)
                ):
                    title = ""

                # 学术期刊（PLOS / Nature / Elsevier 等）：跳过期刊页眉，用首屏纯文本抽取正文标题
                academic = self._academic_title_from_plain_text(text_payload)
                if academic:
                    t_strip = (title or "").strip()
                    prefer_academic = self._should_prefer_academic(
                        t_strip, academic, text_payload
                    )
                    if prefer_academic or (
                        not t_strip
                        or self._is_journal_masthead_only(t_strip)
                        or self._is_article_type_line(t_strip)
                        or self._is_publisher_status_line(t_strip)
                        or self._is_boilerplate_line(t_strip)
                        or self._is_toxic_title_candidate(t_strip)
                        or (
                            re.search(
                                r"(?is)\b(research article|review article|systematic review|"
                                r"article in press|please cite|review)\b",
                                text_payload[:3500],
                            )
                            and len(academic) > len(t_strip) + 10
                        )
                        or len(academic) > len(t_strip) + 15
                    ):
                        title = academic

                if title:
                    title = self._strip_noise_prefix_from_title(title)
                    if self._is_toxic_title_candidate(title) and academic:
                        title = academic

                # --- 阶段 2: 时间线锚定 ---
                header_year = self._year_from_journal_header(text_head)
                pattern = r"(?:©|copyright|published|vol\.|date|年|出版|收稿).*?\b(19[5-9]\d|20[0-2]\d)\b"
                explicit_year = re.search(pattern, text_head, re.IGNORECASE)
                
                if header_year:
                    year = header_year
                elif explicit_year:
                    year = explicit_year.group(1)
                else:
                    docket_year = self._year_from_fda_docket(text_payload[:8000])
                    if docket_year:
                        year = docket_year
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
            logger.exception("标题探测失败: file=%s", pdf_path.name)
            
        return title or pdf_path.stem, year

    def execute(self) -> None:
        """主控循环"""
        pattern = "**/*.pdf" if self.recursive else "*.pdf"
        pdf_targets = [p for p in self.input_dir.glob(pattern) if p.is_file()]

        if not pdf_targets:
            logger.warning("action=input_not_found input=%s", self.input_dir)
            return

        logger.info("action=rename_start targets=%s ocr=%s", len(pdf_targets), OCR_AVAILABLE)
        
        success_count = 0
        skipped_count = 0
        base_input_dir = self.input_dir.resolve()

        for pdf_path in tqdm(pdf_targets, desc="Sanitizing", unit="file", ascii=" ▖▘▝▗▚▞█"):
            raw_title, year = self._scan_payload(pdf_path)
            simplified_name = self._simplify_filename(raw_title)
            
            chronological_name = f"{simplified_name}-{year}"
            
            if self.keep_structure:
                rel_parent = pdf_path.resolve().relative_to(base_input_dir).parent
                target_dir = (self.output_dir / rel_parent).resolve()
            else:
                target_dir = self.output_dir.resolve()
            target_dir.mkdir(parents=True, exist_ok=True)

            target_path = target_dir / f"{chronological_name}.pdf"
            if target_path.exists():
                if not self.overwrite:
                    skipped_count += 1
                    continue
                try:
                    target_path.unlink()
                except Exception:
                    logger.warning("action=overwrite_delete_failed target=%s", target_path, exc_info=True)

            final_safe_name = self._dedupe_filename(target_dir, chronological_name)
            target_path = target_dir / f"{final_safe_name}.pdf"

            shutil.move(str(pdf_path), str(target_path))
            success_count += 1

        if skipped_count:
            logger.warning("action=skip_existing count=%s hint=use_overwrite", skipped_count)
        logger.info("action=rename_complete success=%s output=%s", success_count, self.output_dir)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="PDF 标题驱动重命名（支持递归遍历子文件夹）")
    parser.add_argument("--input", "-i", default="input", help="输入目录（相对 16_PDF_Title_Renamer/）")
    parser.add_argument("--output", "-o", default="output", help="输出目录（相对 16_PDF_Title_Renamer/）")
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
        help="是否在输出目录中保留相对目录结构（默认开启）",
    )
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    args = parser.parse_args()

    sanitizer = PDFSanitizer(
        input_dir=args.input,
        output_dir=args.output,
        recursive=args.recursive,
        keep_structure=args.keep_structure,
        overwrite=args.overwrite,
    )
    sanitizer.execute()
