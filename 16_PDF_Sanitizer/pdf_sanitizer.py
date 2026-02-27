# -*- coding: utf-8 -*-
"""
本地文献重塑协议 (PDF Sanitizer v6.0 - 终极全视版)
Vibe: Academic Cyberpunk
Engine: PyMuPDF | Chrono-Tracker | Bilingual Edge-Pruning | OCR Vision Matrix
"""
import re
import shutil
from pathlib import Path

# --- 基础依赖 ---
import fitz  # pyright: ignore[reportMissingImports]
from tqdm import tqdm
from PIL import Image
import io

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
    def __init__(self, input_dir: str = "input", output_dir: str = "output"):
        self.base_dir = Path(__file__).resolve().parent
        self.input_dir = self.base_dir / input_dir
        self.output_dir = self.base_dir / output_dir

        # 基础设施初始化
        self.input_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _simplify_filename(name: str, max_words: int = 12, max_chars: int = 40) -> str:
        """核心文本手术刀 V6：双语自适应解析，全角标点粉碎，边缘修剪"""
        # 1. 阵营嗅探：检测是否包含中文字符
        is_chinese = bool(re.search(r"[\u4e00-\u9fa5]", name))

        # 2. 物理切除：切除各类括号及内部噪点
        cleaned = re.sub(r"[\[\(（【《].*?[\]\)）】》]", "", name)
        
        # 3. 抹除非法路径字符，替换为空格
        cleaned = re.sub(r"[^\w\s\-\u4e00-\u9fa5]", " ", cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if is_chinese:
            # --- 中文高压压缩逻辑 ---
            cleaned_cjk = cleaned.replace(" ", "")
            if not cleaned_cjk:
                return "未命名文献_Untitled"
            return cleaned_cjk[:max_chars] 
        else:
            # --- 纯英文边缘修剪逻辑 ---
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
            # 渲染首页为 200 DPI 的高精度图像
            pix = doc[0].get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            
            # 启动中英双语识别
            text = pytesseract.image_to_string(img, lang="chi_sim+eng")
            return text.strip()
        except Exception as e:
            print(f"\n[!] 视觉皮层短路: {e}")
            return ""

    def _scan_payload(self, pdf_path: Path) -> tuple[str, str]:
        """神经元探针 V6：融合元数据、原生文本与 OCR 后备能源"""
        title = ""
        year = "XXXX"
        text_payload = ""
        
        try:
            with fitz.open(pdf_path) as doc:
                meta_title = doc.metadata.get("title", "").strip()
                
                # 尝试获取原生文本
                if len(doc) > 0:
                    text_payload = doc[0].get_text("text").strip()
                
                # --- 核心判断：如果原生文本极少（比如扫描版），触发 OCR 视觉矩阵 ---
                if len(text_payload) < 20:
                    text_payload = self._optical_scan(doc)

                # --- 阶段 1: 标题窃取 ---
                if meta_title and len(meta_title) > 2:
                    title = meta_title
                else:
                    lines = [line.strip() for line in text_payload.splitlines() if line.strip()]
                    for line in lines[:15]: # 扩大搜索范围防噪
                        if len(line) >= 5: # 兼容极短中文标题
                            title = line
                            break
                            
                # --- 阶段 2: 时间线锚定 (双语 Chrono-Tracking) ---
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
        pdf_targets = list(self.input_dir.glob("*.pdf"))

        if not pdf_targets:
            print(f"\n[!] 雷达静默。{self.input_dir.name}/ 区块未扫描到目标。请装填弹药。\n")
            return

        print(f"\n[+] Omni-Sight Protocol V6 启动 | 锁定目标: {len(pdf_targets)}")
        print(f"[+] 视觉矩阵(OCR): {'在线' if OCR_AVAILABLE else '离线'}\n")
        
        success_count = 0

        for pdf_path in tqdm(pdf_targets, desc="Sanitizing", unit="file", ascii=" ▖▘▝▗▚▞█"):
            raw_title, year = self._scan_payload(pdf_path)
            simplified_name = self._simplify_filename(raw_title)
            
            chronological_name = f"{simplified_name}-{year}"
            
            final_safe_name = self._dedupe_filename(chronological_name)
            target_path = self.output_dir / f"{final_safe_name}.pdf"

            shutil.move(str(pdf_path), str(target_path))
            success_count += 1

        print("\n" + "=" * 55)
        print(f"[*] 战场清理完毕 | 成功重塑并跃迁: {success_count} 个文件")
        print(f"[*] 终极归档坐标: {self.output_dir}/")
        print("=" * 55 + "\n")


if __name__ == "__main__":
    sanitizer = PDFSanitizer()
    sanitizer.execute()
