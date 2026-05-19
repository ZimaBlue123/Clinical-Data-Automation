# -*- coding: utf-8 -*-
"""
本地文献重塑协议 (PDF Sanitizer v6.6 - 副标题截断版)
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
    }
)

# 文章类型行（全大写短行），其后才是正文标题
_ARTICLE_TYPE_LINE = re.compile(
    r"^(research article|review article|systematic review|meta-analysis|"
    r"editorial|correction|methods|resources|brief report|case report|"
    r"original research|open access|peer-reviewed research)\s*$",
    re.IGNORECASE,
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
            if PDFSanitizer._is_journal_masthead_only(s):
                i += 1
                continue
            if PDFSanitizer._is_article_type_line(s):
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
            if PDFSanitizer._is_journal_masthead_only(s) or PDFSanitizer._is_article_type_line(s):
                i += 1
                continue
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
            title_parts.append(s)
            if len(" ".join(title_parts)) > 520:
                break
            i += 1

        out = re.sub(r"\s+", " ", " ".join(title_parts)).strip()
        return out if len(out) >= 20 else ""

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
            # 切除坏死边缘
            while words and words[-1].lower() in DANGLING_TOXINS:
                words.pop()
            while words and words[0].lower() in DANGLING_TOXINS:
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
            
            # 黑名单：免疫期刊页眉噪点
            blacklist = {"majorarticle", "researcharticle", "reviewarticle", "clinicalinfectiousdiseases"}
            
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

                refined = PDFSanitizer._strip_guidance_cover_prefix(candidate)
                if refined and len(refined) >= 8:
                    return refined

                norm_one = PDFSanitizer._normalize_ws_lower(candidate)
                if any(norm_one == p or norm_one.startswith(p + " ") for p in GUIDANCE_COVER_PREFIXES):
                    continue
                return candidate
        except Exception:
            logger.debug("视觉层级提取失败", exc_info=True)
        return ""

    def _scan_payload(self, pdf_path: Path) -> tuple[str, str]:
        """神经元探针 V6.6：融合元数据、拓扑探测与 OCR 后备能源"""
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

                # 学术期刊（PLOS / Nature 等）：跳过期刊页眉，用首屏纯文本抽取正文标题
                academic = self._academic_title_from_plain_text(text_payload)
                if academic:
                    t_strip = (title or "").strip()
                    if (
                        not t_strip
                        or self._is_journal_masthead_only(t_strip)
                        or self._is_article_type_line(t_strip)
                        or (
                            re.search(r"(?is)\b(research article|review article|systematic review)\b", text_payload[:3500])
                            and len(academic) > len(t_strip) + 10
                        )
                    ):
                        title = academic

                # --- 阶段 2: 时间线锚定 ---
                pattern = r"(?:©|copyright|published|vol\.|date|年|出版|收稿).*?\b(19[5-9]\d|20[0-2]\d)\b"
                explicit_year = re.search(pattern, text_head, re.IGNORECASE)
                
                if explicit_year:
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
    parser.add_argument("--input", "-i", default="input", help="输入目录（相对 15_PDF_Title_Renamer/）")
    parser.add_argument("--output", "-o", default="output", help="输出目录（相对 15_PDF_Title_Renamer/）")
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
