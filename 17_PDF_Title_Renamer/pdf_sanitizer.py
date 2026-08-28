"""
本地文献重塑协议 (PDF Sanitizer v7.2 - 多源置信度仲裁与高精度版式解析)
Vibe: Academic Cyberpunk
Engine: PyMuPDF | Chrono-Tracker | Visual Hierarchy | Block Layout | Subtitle Severance | OCR
"""

from __future__ import annotations

import argparse
import io
import logging
from pathlib import Path
import re
import shutil
from typing import Any

# --- 基础依赖 ---
try:
    import fitz  # pyright: ignore[reportMissingImports]
except ImportError:  # pragma: no cover - 依赖仅在运行时检查
    fitz = None  # type: ignore[assignment]

try:
    from tqdm import tqdm  # pyright: ignore[reportMissingImports]
except ImportError:  # pragma: no cover

    def tqdm(iterable: Any, **_kwargs: Any) -> Any:  # type: ignore[no-untyped-def]
        return iterable


try:
    from PIL import Image  # noqa: F401
except ImportError:  # pragma: no cover
    Image = None  # type: ignore[assignment]

# 标题中默认保持小写的功能词（首词/尾词除外）
LOWERCASE_TITLE_WORDS: frozenset[str] = frozenset(
    {
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
)

DANGLING_TOXINS: frozenset[str] = frozenset(
    {
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
)

# 常见科学与医学缩写规范化映射表
_KNOWN_SCIENTIFIC_TOKENS: dict[str, str] = {
    "cpg": "CpG",
    "mrna": "mRNA",
    "trna": "tRNA",
    "rrna": "rRNA",
    "sirna": "siRNA",
    "cdna": "cDNA",
    "dna": "DNA",
    "rna": "RNA",
    "crispr": "CRISPR",
    "cas9": "Cas9",
    "covid": "COVID",
    "covid-19": "COVID-19",
    "sars-cov-2": "SARS-CoV-2",
    "pimd": "pIMD",
    "pimds": "pIMDs",
    "aesi": "AESI",
    "aesis": "AESIs",
    "hiv": "HIV",
    "hcv": "HCV",
    "hbv": "HBV",
    "hpv": "HPV",
    "gsk": "GSK",
    "fda": "FDA",
    "ich": "ICH",
    "who": "WHO",
    "ema": "EMA",
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

# 期刊封面页眉 / Masthead：字号常最大但不是论文题目，视觉层级需跳过
_JOURNAL_MASTHEAD_COMPRESSED: frozenset[str] = frozenset(
    {
        # PLOS 系列
        "plosone",
        "plosone.",
        "plosbiology",
        "plosmedicine",
        "plospathogens",
        "plosgenetics",
        "ploscompbiol",
        "ploswater",
        "plosclimate",
        "plossustainability",
        "plosdigitalhealth",
        "plosglobalpublichealth",
        # 顶级综合刊
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
        "naturebiotechnology",
        "naturemethods",
        "natureimmunology",
        "natureneuroscience",
        "naturegenetics",
        # 免疫 / 疫苗 / 微生物 / 临床医学
        "vaccine",
        "vaccines",
        "viruses",
        "antibodies",
        "biomedicines",
        "toxins",
        "pathogens",
        "microorganisms",
        "life",
        "biomolecules",
        "cells",
        "molecules",
        "materials",
        "sensors",
        "pharmaceutics",
        "pharmaceuticals",
        "diagnostics",
        "nutrients",
        "cancers",
        "genes",
        "ijms",
        "applsci",
        "jcm",
        "energies",
        "atmosphere",
        "agronomy",
        "antibiotics",
        "animals",
        "forests",
        "water",
        "remotesensing",
        "immunity",
        "bmjopen",
        "elsevier",
        "sciencedirect",
        "sciverse",
        "sciversesciencedirect",
        # Taylor & Francis 系
        "emergingmicrobesinfections",
        "tandfonline",
        "taylorfrancis",
        "taylorandfrancis",
        "virulence",
        "mabs",
        "expertopinion",
        "expertopinionondrugdelivery",
        "expertopiniononbiologicaltherapy",
        "expertopiniononemergingdrugs",
        "expertreview",
        "expertreviewvaccines",
        "expertreviewantiinfectivetherapy",
        "humanvaccines",
        "humanvaccinesimmunotherapeutics",
        # Wiley 系
        "advancedscience",
        "angewandtechemie",
        "angewandtechemieinternationaledition",
        "chemicalcommunications",
        "chemistryaeuropean",
        "europeanjournaloforganicchemistry",
        "europeanjournalofinorganicchemistry",
        # Springer Nature 系
        "scientificreports",
        "naturebiomedicalengineering",
        "naturechemicalbiology",
        "translationalmolecularmedicine",
        "cellandmolecularimmunology",
        # 出版商通用
        "journalhomepage",
        "wileyonlinelibrary",
        "springernature",
        "academicoup",
        "oupcom",
        "academicjournals",
        # 出版商角色行 / 元数据短语
        "academiceditor",
        "sectioneditor",
        "authorcontributions",
        "reviewingeditor",
        "guesteditor",
    }
)

# 出版商角色 / 元数据短语（命中即过滤）
_PUBLISHER_METADATA_PHRASES: tuple[str, ...] = (
    "academic editor",
    "section editor",
    "reviewing editor",
    "guest editor",
    "author contributions",
    "conflict of interest",
    "data availability",
    "funding",
    "acknowledgments",
    "supplementary materials",
    "institutional review board",
    "informed consent",
    "publisher's note",
    "article info",
    "article history",
    "keywords:",
    "abbreviations:",
    "corresponding author",
)

# 单独出现的文章类型短行
_ARTICLE_TYPE_SHORT = re.compile(
    r"^(article|editorial|communication|letter|commentary|perspective|"
    r"correction|erratum|review|protocol|hypothesis|preprint)\s*$",
    re.IGNORECASE,
)

# 文章类型行（全大写或规范名称短行）
_ARTICLE_TYPE_LINE = re.compile(
    r"^(research article|review article|systematic review|meta-analysis|"
    r"editorial|correction|methods|resources|brief report|case report|"
    r"original research|original article|open access|peer-reviewed research|"
    r"short communication|letter|commentary|perspective|"
    r"review|research|a\s+r\s+t\s+i\s+c\s+l\s+e\s+i\s+n\s+f\s+o|"
    r"a\s+b\s+s\s+t\s+r\s+a\s+c\s+t)\s*$",
    re.IGNORECASE,
)

# Elsevier / Springer 等「校样 / 待刊」页眉横幅
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

# 期刊卷期页眉行与独立页码模式
_JOURNAL_VOL_HEADER = re.compile(
    r"^[A-Za-z][\w\s&.\-]{1,50}\s+[\dxX]{1,6}\s*\((19|20)\d{2}\)\s+[\dxX–\-—]+",
    re.IGNORECASE,
)

_STANDALONE_VOL_HEADER = re.compile(
    r"^\s*[\dxX]{1,6}\s*\((19|20)\d{2}\)\s*[\dxX–\-—\s]+\s*$",
    re.IGNORECASE,
)

# 出版商页脚/页眉/链接套话
_BOILERPLATE_LINE = re.compile(
    r"(?i)^(contents lists available|journal homepage|"
    r"www\.elsevier\.com|www\.tandfonline\.com|www\.wiley\.com|"
    r"www\.wileyonlinelibrary\.com|www\.springer\.com|www\.nature\.com|"
    r"link\.springer\.com|sciverse|sciencedirect|g model|received in revised form|"
    r"available online|copyright\s*©|all rights reserved|"
    r"please\s+cite\s+this\s+article|cite\s+this\s+article\s+in\s+press|"
    r"crossmark\.crossref\.org|"
    r"taylor\s*&\s*francis|"
    r"wiley\s*&\s*sons|"
    r"published by\s+(?:elsevier|springer|wiley|taylor|mdpi|frontiers)|"
    r"©\s*\d{4}|"
    r"issn\s*:?\s*\d{4}[\-\s]?\d{3}[\dX]|"
    r"https?://(?:dx\.)?doi\.org/|"
    r"\d{4}-\d{3}[\dX]/\$|"
    r"this\s+article\s+is\s+(?:distributed|made)|"
    r"under\s+the\s+terms\s+of\s+the\s+creative\s+commons)",
)

# Elsevier 引用提示
_CITATION_INSTRUCTION_RE = re.compile(
    r"(?i)^please\s+cite\s+this\s+article",
)
_CITATION_INSTRUCTION_IN_TEXT = re.compile(
    r"(?i)please\s+cite\s+this\s+article|cite\s+this\s+article\s+in\s+press",
)

# 正文题目常见起始词
_TITLE_BODY_START = re.compile(
    r"(?is)(?:clinical\s+evaluation|clinical\s+trial|randomized\s+trial|"
    r"systematic\s+review|meta[\-\s]?analysis|original\s+article|"
    r"optimal\s+approaches|efficacy\s+and\s+safety|phase\s+[iIvV\d]+|open[\-\s]?label|"
    r"a\s+novel|the\s+role\s+of|immunogenicity\s+of|safety\s+and\s+immunogenicity)",
)

# 纯作者姓行（无逗号）
_AUTHOR_SURNAMES_ONLY = re.compile(
    r"^[A-Z][a-z'\-]{2,24}(?:\s+[A-Z][a-z'\-]{2,24}){1,4}$",
)

logger = logging.getLogger(__name__)

# --- 视觉矩阵依赖 (容错加载) ---
try:
    import pytesseract

    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    logger.warning("action=ocr_backend_unavailable backend=pytesseract")


def _is_skip_line(s: str) -> bool:
    """判断学术首屏文本流中是否为应跳过的行。"""
    return (
        PDFSanitizer._is_journal_masthead_only(s)
        or PDFSanitizer._is_article_type_line(s)
        or PDFSanitizer._is_short_article_type_line(s)
        or PDFSanitizer._is_publisher_status_line(s)
        or PDFSanitizer._is_citation_instruction_line(s)
        or PDFSanitizer._is_publisher_metadata_only(s)
        or PDFSanitizer._is_boilerplate_line(s)
        or PDFSanitizer._is_volume_header_line(s)
    )


class PDFSanitizer:
    def __init__(
        self,
        input_dir: str = "input",
        output_dir: str = "output",
        *,
        recursive: bool = True,
        keep_structure: bool = True,
        overwrite: bool = False,
        max_words: int = 40,
        max_chars: int = 200,
    ) -> None:
        self.base_dir = Path(__file__).resolve().parent
        self.input_dir = self.base_dir / input_dir
        self.output_dir = self.base_dir / output_dir
        self.recursive = recursive
        self.keep_structure = keep_structure
        self.overwrite = overwrite
        self.max_words = max_words
        self.max_chars = max_chars

        self.input_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _normalize_ws_lower(s: str) -> str:
        return re.sub(r"\s+", " ", s.lower()).strip()

    @staticmethod
    def _compress_for_masthead(s: str) -> str:
        return re.sub(r"[^\w]+", "", s.lower())

    @staticmethod
    def _strip_guidance_cover_prefix(raw: str) -> str:
        """若整段以泛化封面标题开头，去掉前缀，保留具体题目。"""
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
    def _is_journal_masthead_only(candidate: str) -> bool:
        """是否为仅期刊名/页眉（非正文标题）。"""
        t = candidate.strip()
        if not t:
            return True
        comp = PDFSanitizer._compress_for_masthead(t)
        if comp in _JOURNAL_MASTHEAD_COMPRESSED:
            return True
        words = t.split()
        return (
            len(words) <= 5
            and len(t) <= 48
            and not re.search(r"[.?:;]$", t)
            and (comp.endswith("journal") or "journalof" in comp or "infections" in comp or "&" in t)
        )

    @staticmethod
    def _is_publisher_metadata_only(candidate: str) -> bool:
        """出版商角色 / 元数据短语（Academic Editor / Author Contributions 等）。"""
        comp = PDFSanitizer._compress_for_masthead(candidate)
        if not comp:
            return False
        return any(phrase.replace(" ", "") in comp for phrase in _PUBLISHER_METADATA_PHRASES)

    @staticmethod
    def _is_short_article_type_line(line: str) -> bool:
        return bool(_ARTICLE_TYPE_SHORT.match(line.strip()))

    @staticmethod
    def _is_article_type_line(line: str) -> bool:
        return bool(_ARTICLE_TYPE_LINE.match(line.strip()))

    @staticmethod
    def _is_publisher_status_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return False
        if _PUBLISHER_STATUS_RE.match(s):
            return True
        comp = PDFSanitizer._compress_for_masthead(s)
        if comp in _PUBLISHER_STATUS_COMPRESSED:
            return True
        return "articleinpress" in comp and len(comp) <= 24

    @staticmethod
    def _is_citation_instruction_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return False
        if _CITATION_INSTRUCTION_RE.match(s):
            return True
        return _CITATION_INSTRUCTION_IN_TEXT.search(s) is not None and len(s) < 200

    @staticmethod
    def _is_volume_header_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return False
        if _JOURNAL_VOL_HEADER.match(s) or _STANDALONE_VOL_HEADER.match(s):
            return True
        comp = PDFSanitizer._compress_for_masthead(s)
        if re.search(r"^\d{1,6}\(19\d{2}|20\d{2}\)\d{1,6}", comp):
            return True
        return bool(re.search(r"\b\d+\s*\((19|20)\d{2}\)\s*[\dxX–\-—]+", s))

    @staticmethod
    def _is_boilerplate_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return True
        if _BOILERPLATE_LINE.search(s):
            return True
        if PDFSanitizer._is_volume_header_line(s):
            return True
        if re.match(r"^G\s+Model\b", s, re.I):
            return True
        comp = PDFSanitizer._compress_for_masthead(s)
        if any(
            noise in comp
            for noise in (
                "journalhomepage",
                "contentslistsavailable",
                "sciencedirect",
                "sciverse",
                "articleinfo",
                "articlehistory",
                "abstract",
                "keywords",
                "abbreviations",
                "correspondingauthor",
            )
        ):
            return True
        return bool(re.match(r"^[A-Z]{2,6}\s+\d{4,6}\s+\d", s))

    @staticmethod
    def _looks_like_author_surnames_only(line: str) -> bool:
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
            "optimal",
        }
        return not any(w.lower() in titleish for w in s.split())

    @staticmethod
    def _looks_like_author_line(line: str) -> bool:
        s = line.strip()
        if not s or len(s) > 280:
            return False
        if re.search(r"[\*∗]\s*$", s) or "@" in s:
            return True
        if s.count(",") >= 1 and re.match(r"^[A-Z][\w\-\.']+(?:\s*,\s*[A-Z][\w\-\.']*)+", s) and len(s) < 120:
            return True
        if re.search(r"\b[A-Z][a-z]+\s+[A-Z][a-z]+\s+[a-f1-9](?:,|$)", s):
            return True
        return bool(PDFSanitizer._looks_like_author_surnames_only(s))

    @staticmethod
    def _looks_like_scientific_title(text: str) -> bool:
        s = (text or "").strip()
        if len(s) < 20:
            return False
        if re.search(r"\s+of\s+", s, re.I) or re.search(r"\s+for\s+", s, re.I):
            return True
        if re.search(
            r"(?i)\b(?:clinical|evaluation|efficacy|randomized|vaccine|trial|study|"
            r"optimal|approaches|oligonucleotide|adjuvant|infectious|cancer|immunogenicity)\b",
            s,
        ):
            return True
        return len(s.split()) >= 5

    @staticmethod
    def _is_toxic_title_candidate(text: str) -> bool:
        s = (text or "").strip()
        if not s:
            return True
        norm = PDFSanitizer._normalize_ws_lower(s)
        comp = PDFSanitizer._compress_for_masthead(s)
        if any(
            bot_noise in comp
            for bot_noise in (
                "javascriptisdisabled",
                "redirecting",
                "verifythatyourenotarobot",
                "weneedtoverify",
                "skiptomaincontent",
                "officialwebsiteoftheunitedstates",
            )
        ):
            return True

        return (
            bool(_CITATION_INSTRUCTION_IN_TEXT.search(norm))
            or PDFSanitizer._is_publisher_status_line(s)
            or PDFSanitizer._is_citation_instruction_line(s)
            or PDFSanitizer._is_volume_header_line(s)
            or "pleasecitethisarticle" in comp
            or "citethisarticleinpress" in comp
            or PDFSanitizer._is_publisher_metadata_only(s)
            or PDFSanitizer._is_journal_masthead_only(s)
            or PDFSanitizer._is_short_article_type_line(s)
            or PDFSanitizer._is_article_type_line(s)
            or any(
                noise in comp
                for noise in (
                    "journalhomepage",
                    "contentslistsavailable",
                    "sciencedirect",
                    "sciverse",
                )
            )
            or bool(PDFSanitizer._looks_like_author_line(s))
        )

    @staticmethod
    def _strip_noise_prefix_from_title(raw: str) -> str:
        t = re.sub(r"\s+", " ", (raw or "").strip())
        if not t:
            return t
        m = _TITLE_BODY_START.search(t)
        if m and m.start() > 0:
            head = t[: m.start()]
            if (
                _CITATION_INSTRUCTION_IN_TEXT.search(head)
                or re.search(r"(?i)\bpress\b", head)
                or re.search(r"(?i)\b(?:scheiermann|klinman|[A-Z][a-z]+,\s*[A-Z]\.)\b", head)
            ):
                t = t[m.start() :].strip()
        t = re.sub(
            r"(?is)^(?:please\s+cite\s+this\s+article[^a-z]{0,40})+",
            "",
            t,
        ).strip()
        if not PDFSanitizer._looks_like_scientific_title(t):
            t = re.sub(
                r"^(?:(?:[A-Z][a-z'\-]{2,20})(?:\s*,\s*[A-Z]\.?){0,3}\s+){1,4}"
                r"(?=[A-Z][a-z]{5,}\s)",
                "",
                t,
            ).strip()

        # Strip trailing noise like " - PMC", " | PubMed", etc.
        return re.sub(r"(?i)(?:\s*[-|]\s*(?:PMC|PubMed|Europe\s+PMC|NCBI|NIH|medRxiv|bioRxiv))+$", "", t).strip()

    @staticmethod
    def _format_roman_token(w: str) -> str:
        """格式化罗马数字 token（如 IIa -> IIa, I_IIa -> I_IIa）。"""
        if re.fullmatch(r"^(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)[a-z]?$", w, re.IGNORECASE):
            m = re.match(r"^([IVXLCDM]+)([a-z]?)$", w, re.IGNORECASE)
            return f"{m.group(1).upper()}{m.group(2).lower()}" if m else ""
        roman_pats = r"^(?:I|II|III|IV|V|VI|VII|VIII|IX|X)_(?:I|II|III|IV|V|VI|VII|VIII|IX|X)[a-z]?$"
        if re.fullmatch(roman_pats, w, re.IGNORECASE):
            parts = w.split("_", 1)
            p1 = re.sub(r"[^A-Za-z]", "", parts[0]).upper()
            m2 = re.match(r"^([IVXLCDM]+)([a-z]?)$", parts[1], re.IGNORECASE)
            return f"{p1}_{m2.group(1).upper()}{m2.group(2).lower()}" if m2 else ""
        return ""

    @staticmethod
    def _preserve_scientific_token(word: str) -> str:
        """保留 CpG、mRNA、IL-6 等科学缩写原有大小写（含罗马数字如 IIa/IVb）。"""
        w = word.strip()
        if not w or w.lower() in LOWERCASE_TITLE_WORDS or w.lower() in {"is", "it", "am", "me", "do", "if", "be"}:
            return ""
        low = w.lower()
        if low in _KNOWN_SCIENTIFIC_TOKENS:
            return _KNOWN_SCIENTIFIC_TOKENS[low]
        if re.match(r"^[a-z]+[A-Z][\w\-]*$", w) or re.match(r"^[A-Z][a-z]*[A-Z][\w\-]*$", w):
            return w
        if re.search(r"\d", w) and re.search(r"[A-Za-z]", w):
            return w
        return PDFSanitizer._format_roman_token(w)

    @staticmethod
    def _smart_title_case(text: str) -> str:
        raw_words = text.split()
        if not raw_words:
            return ""

        formatted_words: list[str] = []
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
    def _split_on_smart_colon(name: str) -> str:
        if "：" in name:
            return name.split("：", 1)[0]
        if ":" in name:
            head, _, tail = name.partition(":")
            tail_clean = tail.strip()
            if re.search(
                r"(?i)\b(phase|randomized|trial|study|analysis|review|"
                r"double[\-\s]?blind|open[\-\s]?label|single[\-\s]?blind|"
                r"multicenter|multicentre|case\s+report|brief\s+report|"
                r"short\s+report|letter\s+to|preliminary)\b",
                tail_clean,
            ):
                return name
            return head
        return name

    @staticmethod
    def _smart_bracket_removal(name: str) -> str:
        out = re.sub(
            r"[\[\(（【《][^\[\]\(\)\u4e00-\u9fa5\w]*[\]\)）】》]",
            "",
            name,
        )
        return re.sub(
            r"[\[\(（【《]([^\[\]\(\)\u4e00-\u9fa5]*?[A-Za-z0-9\u4e00-\u9fa5][^\[\]\(\)\u4e00-\u9fa5]*?)[\]\)）】》]",
            r" \1 ",
            out,
        )

    @staticmethod
    def _split_roman_numerals(words: list[str]) -> list[str]:
        roman_token = re.compile(r"^([IVXLCDM]+)/([IVXLCDM]+[a-z]?)$", re.IGNORECASE)
        result: list[str] = []
        for w in words:
            if "/" in w and "." not in w and roman_token.match(w):
                parts = w.split("/", 1)
                result.append(f"{parts[0]}_{parts[1]}")
            else:
                result.append(w)
        return result

    def _simplify_filename(self, name: str) -> str:
        name = PDFSanitizer._split_on_smart_colon(name)
        is_chinese = bool(re.search(r"[\u4e00-\u9fa5]", name))

        name = re.sub(r"\s*&\s*", " And ", name)
        name = PDFSanitizer._smart_bracket_removal(name)

        cleaned = re.sub(r"[^\w\s\-\u4e00-\u9fa5/]", " ", name)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if not is_chinese:
            roman_words: list[str] = []
            for w in cleaned.split():
                roman_words.extend(PDFSanitizer._split_roman_numerals([w]))
            cleaned = " ".join(roman_words)

        if is_chinese:
            cleaned_cjk = cleaned.replace(" ", "").replace("/", "")
            if not cleaned_cjk:
                return "未命名文献_Untitled"
            return cleaned_cjk[:self.max_chars]

        cleaned_en = PDFSanitizer._smart_title_case(cleaned)
        words = cleaned_en.split()[:self.max_words]
        words = [w.replace("/", "") for w in words]

        metadata_blocklist = {
            "academic",
            "editor",
            "section",
            "reviewing",
            "guest",
            "author",
            "contributions",
            "funding",
            "acknowledgments",
            "supplementary",
            "materials",
            "contribution",
        }
        words = [w for w in words if w.lower() not in metadata_blocklist]

        while len(words) > 2 and words[-1].lower() in DANGLING_TOXINS:
            words.pop()
        while len(words) > 2 and words[0].lower() in DANGLING_TOXINS:
            words.pop(0)

        if not words:
            fallback = cleaned_en[:self.max_chars].strip().replace(" ", "_")
            return fallback if fallback else "Untitled_Document"

        joined_preview = " ".join(words)
        if PDFSanitizer._is_toxic_title_candidate(joined_preview):
            return "Untitled_Document"

        res = "_".join(words)
        if len(res) > self.max_chars:
            res = res[:self.max_chars]
            if "_" in res:
                res = res.rsplit("_", 1)[0]
        return res

    def _dedupe_filename(self, target_dir: Path, base_name: str) -> str:
        candidate = base_name
        counter = 1
        while (target_dir / f"{candidate}.pdf").exists():
            candidate = f"{base_name}_{counter}"
            counter += 1
        return candidate

    @staticmethod
    def _optical_scan(doc: Any) -> str:
        if not OCR_AVAILABLE or Image is None or fitz is None or doc is None or len(doc) == 0:
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
    def _extract_title_by_visual_hierarchy(page: Any) -> str:
        """视觉层级引擎：按物理字号和版面几何聚合标题。"""
        if fitz is None or page is None:
            return ""
        try:
            blocks = page.get_text("dict").get("blocks", [])
            size_map: dict[float, list[str]] = {}
            for b in blocks:
                if "lines" not in b:
                    continue
                for line in b["lines"]:
                    spans = line.get("spans", [])
                    line_text = " ".join(s.get("text", "").strip() for s in spans if s.get("text", "").strip()).strip()
                    if not line_text:
                        continue
                    max_size = max((s.get("size", 0) for s in spans), default=0)
                    s_key = round(max_size, 1)
                    if s_key not in size_map:
                        size_map[s_key] = []
                    size_map[s_key].append(line_text)

            for s in sorted(size_map.keys(), reverse=True):
                candidate = re.sub(r"\s+", " ", " ".join(size_map[s])).strip()
                if len(candidate) < 8 or PDFSanitizer._is_toxic_title_candidate(candidate):
                    continue
                if PDFSanitizer._is_boilerplate_line(candidate):
                    continue
                refined = PDFSanitizer._strip_guidance_cover_prefix(candidate)
                refined = PDFSanitizer._strip_noise_prefix_from_title(refined)
                if refined and len(refined) >= 8 and not PDFSanitizer._is_toxic_title_candidate(refined):
                    return refined

        except Exception:
            logger.debug("视觉层级提取失败", exc_info=True)
        return ""

    @staticmethod
    def _extract_title_from_blocks(page: Any) -> str:
        """基于 PyMuPDF 自然段落 Block 的结构化提取。"""
        if fitz is None or page is None:
            return ""
        try:
            blocks = page.get_text("blocks")
            page_height = page.rect.height
            for b in blocks:
                if len(b) < 5 or b[6] != 0 or b[1] > page_height * 0.55:
                    continue
                text = re.sub(r"\s+", " ", b[4]).strip()
                if len(text) < 12 or len(text) > 400:
                    continue
                if PDFSanitizer._is_toxic_title_candidate(text) or PDFSanitizer._is_boilerplate_line(text):
                    continue
                refined = PDFSanitizer._strip_guidance_cover_prefix(text)
                refined = PDFSanitizer._strip_noise_prefix_from_title(refined)
                if refined and len(refined) >= 12 and not PDFSanitizer._is_toxic_title_candidate(refined):
                    return refined
        except Exception:
            logger.debug("Block 层级提取失败", exc_info=True)
        return ""

    @staticmethod
    def _academic_title_from_plain_text(text: str) -> str:
        """学术期刊首屏文本流解析。"""
        if not text:
            return ""
        raw_lines = text[:8000].splitlines()
        i = 0
        n = len(raw_lines)

        while i < n:
            line_str = raw_lines[i].strip()
            if not line_str or _is_skip_line(line_str) or PDFSanitizer._looks_like_author_line(line_str):
                i += 1
                continue
            break

        title_parts: list[str] = []
        while i < n:
            s = raw_lines[i].strip()
            if not s:
                if title_parts:
                    break
                i += 1
                continue
            if _is_skip_line(s):
                i += 1
                continue
            if (
                (title_parts and PDFSanitizer._looks_like_author_line(s))
                or (s.count(",") >= 2 and len(s) < 220 and re.search(r",\s*[A-Z][a-z]", s))
                or ("@" in s or re.search(r"https?://", s, re.I))
                or (re.match(r"^\d+\s+", s) and title_parts)
                or s.lower().startswith(("doi:", "doi ", "citation:"))
            ):
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
    def _score_title_candidate(cand: str, meta_title: str) -> float:
        """评估候选标题的置信度得分。"""
        if not cand or len(cand.strip()) < 8 or PDFSanitizer._is_toxic_title_candidate(cand):
            return -100.0
        s = cand.strip()
        score = 10.0
        if 25 <= len(s) <= 250:
            score += 25.0
        elif 15 <= len(s) < 25 or 250 < len(s) <= 350:
            score += 10.0
        else:
            score -= 15.0

        if PDFSanitizer._looks_like_scientific_title(s):
            score += 20.0

        if meta_title and len(meta_title) >= 15:
            norm_cand = PDFSanitizer._compress_for_masthead(s)
            norm_meta = PDFSanitizer._compress_for_masthead(meta_title)
            if norm_cand == norm_meta:
                score += 60.0
            elif norm_meta in norm_cand:
                extra_len = len(norm_cand) - len(norm_meta)
                score += max(10.0, 40.0 - extra_len * 0.5)
            elif norm_cand in norm_meta:
                score += 30.0
            else:
                words_cand = set(s.lower().split())
                words_meta = set(meta_title.lower().split())
                intersection = len(words_cand & words_meta)
                if intersection >= 3:
                    score += 15.0 + intersection * 2.0

        return score

    @staticmethod
    def _arbitrate_title(
        hierarchy_title: str,
        block_title: str,
        plain_title: str,
        meta_title: str,
    ) -> str:
        """多源标题仲裁器：选择置信度最高的高质量标题。"""
        clean_meta = meta_title.strip()
        if "microsoft word" in clean_meta.lower() or clean_meta.lower().startswith("untitled"):
            clean_meta = ""

        candidates = [
            ("hierarchy", hierarchy_title, 10.0),
            ("block", block_title, 8.0),
            ("meta", clean_meta, 5.0),
            ("plain", plain_title, 0.0),
        ]

        scored = [
            (source, text, PDFSanitizer._score_title_candidate(text, clean_meta) + bonus)
            for source, text, bonus in candidates
            if text and not PDFSanitizer._is_toxic_title_candidate(text)
        ]

        if not scored:
            return clean_meta

        scored.sort(key=lambda x: x[2], reverse=True)
        best_source, best_text, best_score = scored[0]

        if best_score > 0:
            return best_text
        return clean_meta

    @staticmethod
    def _extract_year_from_sources(doc: Any, text_head: str, text_payload: str) -> str:
        """多层降级年份抽取。"""
        patterns = [
            r"(?i)\b[A-Za-z][\w\s&.\-]{0,48}\s*[\dxX]{0,6}\s*\((19[5-9]\d|20[0-2]\d)\)",
            r"\b\d{1,6}\s*\((19[5-9]\d|20[0-2]\d)\)\s*[\dxX–\-—]+",
            r"(?:©|copyright|published|vol\.|date|年|出版|收稿).*?\b(19[5-9]\d|20[0-2]\d)\b",
        ]
        for pat in patterns:
            m = re.search(pat, text_head[:2500], re.IGNORECASE)
            if m:
                return m.group(1)

        docket_year = PDFSanitizer._year_from_fda_docket(text_payload[:8000])
        if docket_year:
            return docket_year

        if doc is not None:
            meta_year_match = re.search(r"D:(\d{4})", doc.metadata.get("creationDate", ""))
            if meta_year_match:
                return meta_year_match.group(1)

        fallback_year = re.search(r"\b(19[5-9]\d|20[0-2]\d)\b", text_head)
        return fallback_year.group(1) if fallback_year else "XXXX"

    def _scan_payload(self, pdf_path: Path) -> tuple[str, str]:
        """融合元数据、视觉层级、Block 版面与 OCR 的综合探针。"""
        if fitz is None:
            logger.warning("action=fitz_missing file=%s hint=pip install pymupdf", pdf_path.name)
            return pdf_path.stem, "XXXX"

        title = ""
        year = "XXXX"

        try:
            with fitz.open(pdf_path) as doc:
                meta_title = doc.metadata.get("title", "").strip()
                text_payload = ""
                hierarchy_title = ""
                block_title = ""

                if len(doc) > 0:
                    first_page = doc[0]
                    text_payload = first_page.get_text("text").strip()
                    hierarchy_title = self._extract_title_by_visual_hierarchy(first_page)
                    block_title = self._extract_title_from_blocks(first_page)

                if len(text_payload) < 20:
                    text_payload = self._optical_scan(doc)

                plain_title = self._academic_title_from_plain_text(text_payload)
                if re.search(r"(?is)guidance\s+for\s+industry", text_payload[:2000]):
                    fda_title = self._fda_specific_title_from_plain_text(text_payload)
                    if fda_title:
                        plain_title = fda_title

                title = self._arbitrate_title(
                    hierarchy_title=hierarchy_title,
                    block_title=block_title,
                    plain_title=plain_title,
                    meta_title=meta_title,
                )

                if title:
                    title = self._strip_noise_prefix_from_title(title)
                    title = self._strip_guidance_cover_prefix(title)

                text_head = text_payload[:4000]
                year = self._extract_year_from_sources(doc, text_head, text_payload)

        except Exception:
            logger.exception("标题探测失败: file=%s", pdf_path.name)

        final_title = (title or "").strip()
        if not final_title or PDFSanitizer._is_toxic_title_candidate(final_title):
            stem = pdf_path.stem
            # 去除原文件名中已经包含的 -XXXX 或 -YYYY 后缀，防止重复
            stem = re.sub(r"-(?:XXXX|19[5-9]\d|20[0-2]\d)(?:_\d+)?$", "", stem)
            final_title = stem

        return final_title, year

    def _process_single_pdf(self, pdf_path: Path, base_input_dir: Path) -> tuple[bool, bool]:
        """处理单个 PDF 文件的重命名与移动，返回 (moved, skipped)。"""
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
                return False, True
            try:
                target_path.unlink()
            except Exception:
                logger.warning("action=overwrite_delete_failed target=%s", target_path, exc_info=True)

        final_safe_name = self._dedupe_filename(target_dir, chronological_name)
        target_path = target_dir / f"{final_safe_name}.pdf"

        for attempt in (1, 2):
            try:
                shutil.copy2(str(pdf_path), str(target_path))
                try:
                    Path(pdf_path).unlink()
                except OSError as e:
                    logger.warning(
                        "action=copy_succeeded_delete_failed src=%s reason=%s hint=output_written_close_locked_input",
                        pdf_path,
                        e,
                    )
                return True, False
            except OSError as e:
                if attempt == 1:
                    logger.debug("action=move_retry src=%s reason=%s", pdf_path, e)
                    continue
                logger.warning(
                    "action=move_failed src=%s target=%s reason=%s hint=close_locked_input_then_retry",
                    pdf_path,
                    target_path,
                    e,
                )
        return False, False

    def execute(self) -> None:
        """主控重命名循环。"""
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
            moved, skipped = self._process_single_pdf(pdf_path, base_input_dir)
            if moved:
                success_count += 1
            if skipped:
                skipped_count += 1

        if skipped_count:
            logger.warning("action=skip_existing count=%s hint=use_overwrite", skipped_count)
        logger.info("action=rename_complete success=%s output=%s", success_count, self.output_dir)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="PDF 标题驱动重命名（支持递归遍历子文件夹）")
    parser.add_argument("--input", "-i", default="input", help="输入目录（相对 17_PDF_Title_Renamer/）")
    parser.add_argument(
        "--output",
        "-o",
        default="output",
        help="输出目录（相对 17_PDF_Title_Renamer/）",
    )
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
    parser.add_argument("--max-words", type=int, default=40, help="英文标题保留的最大单词数（默认 40）")
    parser.add_argument("--max-chars", type=int, default=200, help="总文件名长度的安全截断字符数（默认 200）")
    args = parser.parse_args()

    sanitizer = PDFSanitizer(
        input_dir=args.input,
        output_dir=args.output,
        recursive=args.recursive,
        keep_structure=args.keep_structure,
        overwrite=args.overwrite,
        max_words=args.max_words,
        max_chars=args.max_chars,
    )
    sanitizer.execute()
