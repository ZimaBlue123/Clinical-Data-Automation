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

# --- 基础依赖 ---
import fitz  # pyright: ignore[reportMissingImports]
from tqdm import tqdm
from PIL import Image

# --- 视觉矩阵依赖 (容错加载) ---
try:
    import pytesseract
    # 如果你是 Windows 用户且没有配置环境变量，请取消下方注释并修改路径
    # pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("[!] 警告: pytesseract 未安装，视觉矩阵(OCR)已离线。")


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
    def _simplify_filename(name: str, max_words: int = 15, max_chars: int = 40) -> str:
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
            cleaned_en = cleaned.title()
            words = cleaned_en.split()[:max_words]
            dangling_toxins = {
                "And", "Or", "Of", "For", "To", "In", "On", "With", "By", "From", 
                "The", "A", "An", "At", "As", "What", "Which", "That", "Is", "Are"
            }
            # 切除坏死边缘
            while words and words[-1] in dangling_toxins:
                words.pop()
            while words and words[0] in dangling_toxins:
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
            print(f"\n[!] 视觉皮层短路: {e}")
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
                
                # 若无毒且长度合理，直接将其判定为标题
                if len(candidate) >= 8 and not is_toxic:
                    return candidate
        except Exception:
            pass
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
                            
                # --- 阶段 2: 时间线锚定 ---
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
                                
        except Exception as exc:
            print(f"\n[!] 探针受损: {pdf_path.name} | 异常: {exc}")
            
        return title or pdf_path.stem, year

    def execute(self) -> None:
        """主控循环"""
        pattern = "**/*.pdf" if self.recursive else "*.pdf"
        pdf_targets = [p for p in self.input_dir.glob(pattern) if p.is_file()]

        if not pdf_targets:
            print(f"\n[!] 雷达静默。{self.input_dir.name}/ 区块未扫描到目标。请装填弹药。\n")
            return

        print(f"\n[+] Protocol V6.6 (Severance) 启动 | 锁定目标: {len(pdf_targets)}")
        print(f"[+] 视觉矩阵(OCR): {'在线' if OCR_AVAILABLE else '离线'}\n")
        
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
                    pass

            final_safe_name = self._dedupe_filename(target_dir, chronological_name)
            target_path = target_dir / f"{final_safe_name}.pdf"

            shutil.move(str(pdf_path), str(target_path))
            success_count += 1

        print("\n" + "=" * 55)
        if skipped_count:
            print(f"[*] 跳过已存在输出: {skipped_count} 个文件（可用 --overwrite 覆盖）")
        print(f"[*] 战场清理完毕 | 成功重塑并跃迁: {success_count} 个文件")
        print(f"[*] 终极归档坐标: {self.output_dir}/")
        print("=" * 55 + "\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF 标题驱动重命名（支持递归遍历子文件夹）")
    parser.add_argument("--input", default="input", help="输入目录（相对 15_PDF_Sanitizer/）")
    parser.add_argument("--output", default="output", help="输出目录（相对 15_PDF_Sanitizer/）")
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
