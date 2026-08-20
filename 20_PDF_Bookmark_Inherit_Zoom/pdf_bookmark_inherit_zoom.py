"""
PDF 书签「承前缩放」强健版模块 - Vibe Coder Edition
优化项：
1) 引擎降维：弃用 PyPDF2，切换至工业级 PyMuPDF (fitz)
2) 规范对齐：通过重写 TOC (目录树) 注入 XYZ & zoom=0 实现真正的承前缩放
3) 结构增强：引入 ThreadPoolExecutor 处理高并发 I/O
4) 边界防御：自动跳过加密文件，使用 pathlib 替换老旧的 os.path
5) 体积优化：附带高级垃圾回收机制 (garbage=3) 与流压缩 (deflate)
"""

from __future__ import annotations

import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import fitz  # 核心依赖：PyMuPDF

BASE = Path(__file__).resolve().parent

# 极简且专业的日志风格
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S",
)

logger = logging.getLogger(__name__)


def fix_pdf_zoom(input_path: Path, output_path: Path) -> tuple[bool, str]:
    """核心处理：深度遍历 PDF 目录树并强制注入承前缩放参数"""
    doc = None
    try:
        doc = fitz.open(input_path)

        # 边界防御：文档加密锁
        if doc.is_encrypted:
            return False, "文档已加密，拒绝访问"

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 提取深层目录树 (simple=False 允许获取完整的目标跳转属性)
        toc = doc.get_toc(simple=False)
        if not toc:
            # 没有书签？顺手帮你做个无损压缩
            doc.save(output_path, garbage=3, deflate=True)
            return True, "无书签，已优化体积并转存"

        # 逻辑雕刻：修改跳转属性 (XYZ + zoom=0 = 承前缩放)
        new_toc = []
        modified = False
        for item in toc:
            lvl, title, page, dest = item
            if isinstance(dest, dict):
                dest["kind"] = fitz.LINK_XYZ
                dest["zoom"] = 0.0  # 核心：0.0 指示阅读器保持当前缩放比例
                modified = True
            new_toc.append([lvl, title, page, dest])

        if modified:
            doc.set_toc(new_toc)

        # 存档并执行 Level 3 垃圾回收，同时压缩内容流（达成你原代码的目的）
        doc.save(output_path, garbage=3, deflate=True)
        return True, "承前缩放已注入"

    except Exception as e:
        return False, f"崩溃异常: {str(e)}"
    finally:
        if doc:
            doc.close()


def batch_set_scaling(input_dir: str, output_dir: str, max_workers: int = 4) -> None:
    """并发调度层：优雅榨干 CPU 的 I/O 性能"""
    in_path, out_path = Path(input_dir), Path(output_dir)

    if not in_path.exists():
        logging.error(f"致命错误：输入目录不存在 -> {in_path}")
        return

    pdf_files = list(in_path.glob("*.pdf"))
    if not pdf_files:
        logging.warning("未扫描到任何 PDF 文件。")
        return

    logging.info(f"已锁定 {len(pdf_files)} 个 PDF，启动 {max_workers} 线程并发修正...")

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 构建 Future 映射
        future_to_pdf = {executor.submit(fix_pdf_zoom, pdf, out_path / pdf.name): pdf for pdf in pdf_files}

        success_cnt = 0
        for future in as_completed(future_to_pdf):
            pdf = future_to_pdf[future]
            ok, msg = future.result()
            if ok:
                success_cnt += 1
                logging.info(f"[  OK  ] {pdf.name} -> {msg}")
            else:
                logging.error(f"[ FAIL ] {pdf.name} -> {msg}")

    logging.info(f"批处理终了。成功率: {success_cnt}/{len(pdf_files)}")


if __name__ == "__main__":
    # 默认使用模块目录下 input / output；亦可改为 Raw String 绝对路径（Windows 注意转义）
    INPUT_DIRECTORY = str(BASE / "input")
    OUTPUT_DIRECTORY = str(BASE / "output")

    batch_set_scaling(INPUT_DIRECTORY, OUTPUT_DIRECTORY, max_workers=6)
