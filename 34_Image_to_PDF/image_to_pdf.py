"""
批量将图片转换为 PDF，支持单图转单 PDF 或多图合并为一个 PDF。

默认目录：
- 输入：34_Image_to_PDF/input/
- 输出：34_Image_to_PDF/output/

用法：
  python image_to_pdf.py               # 单图转单 PDF
  python image_to_pdf.py --merge       # 合并 input 目录下所有图片为一个 PDF
  python image_to_pdf.py --input "D:\\Images" --output "D:\\PDFs"
"""

from __future__ import annotations

import argparse
import datetime
import logging
import re
import shutil
from pathlib import Path

from PIL import Image

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tiff"}


def process_image_for_pdf(img_path: Path) -> Image.Image | None:
    """读取图片，并进行色彩空间兼容性处理，返回适用于保存为 PDF 的 Image 对象。"""
    try:
        img = Image.open(img_path)
        # 统一转为 RGB 以兼容 PDF。如果含有 alpha 通道（如 PNG），则用白底替换透明背景。
        if img.mode in ("RGBA", "LA", "P"):
            if img.mode == "P" and "transparency" in img.info:
                img = img.convert("RGBA")
            if img.mode in ("RGBA", "LA"):
                # 创建全白背景
                background = Image.new("RGB", img.size, (255, 255, 255))
                # 将原图带 alpha 通道的内容贴到全白背景上
                background.paste(img, mask=img.split()[-1])
                return background
            return img.convert("RGB")
        if img.mode != "RGB":
            return img.convert("RGB")
        return img
    except Exception as e:
        logger.warning("action=read_image_failed file=%s error=%s", img_path.name, str(e))
        return None


def natural_sort_key(path: Path):
    """用于自然排序的 key 函数：提取文件名中的数字转为整数对比。"""
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', path.name)]

def collect_images(input_path: Path) -> list[Path]:
    """收集支持的图片文件并按照自然数字顺序排序，确保页面顺序。"""
    if input_path.is_file() and input_path.suffix.lower() in SUPPORTED_EXTS:
        return [input_path]
    if input_path.is_dir():
        files = [f for f in input_path.iterdir() if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS]
        return sorted(files, key=natural_sort_key)
    return []


def convert_single(img_path: Path, output_dir: Path, overwrite: bool = False) -> bool:
    """将单张图片转换为单一 PDF。"""
    pdf_path = output_dir / f"{img_path.stem}.pdf"
    if pdf_path.exists() and not overwrite:
        logger.info("action=skip_exists file=%s", pdf_path.name)
        return False

    img = process_image_for_pdf(img_path)
    if not img:
        return False

    try:
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(pdf_path, "PDF", resolution=100.0, save_all=True)
        logger.info("action=convert_success src=%s dest=%s", img_path.name, pdf_path.name)
        return True
    except Exception as e:
        logger.warning("action=convert_failed file=%s error=%s", img_path.name, str(e))
        return False
    finally:
        if img:
            img.close()


def merge_images(
    img_paths: list[Path],
    output_dir: Path,
    output_filename: str = "merged_output.pdf",
    overwrite: bool = False
) -> bool:
    """将多张图片合并为一个 PDF。"""
    if not img_paths:
        return False

    pdf_path = output_dir / output_filename
    if pdf_path.exists() and not overwrite:
        logger.info("action=skip_exists file=%s", pdf_path.name)
        return False

    images = []
    first_img = None
    try:
        for p in img_paths:
            img = process_image_for_pdf(p)
            if img:
                if not first_img:
                    first_img = img
                else:
                    images.append(img)

        if not first_img:
            logger.warning("action=merge_failed reason=no_valid_images")
            return False

        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        first_img.save(
            pdf_path,
            "PDF",
            resolution=100.0,
            save_all=True,
            append_images=images
        )
        logger.info("action=merge_success dest=%s total_pages=%d", pdf_path.name, len(images) + 1)
        return True
    except Exception as e:
        logger.warning("action=merge_failed dest=%s error=%s", pdf_path.name, str(e))
        return False
    finally:
        if first_img:
            first_img.close()
        for i in images:
            i.close()


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(description="批量将图片转换为 PDF，支持单张转换或多张合并。")
    parser.add_argument(
        "--input",
        default=str(default_input),
        help="输入目录或单个图片文件",
    )
    parser.add_argument(
        "--output",
        default=str(default_output),
        help="输出目录",
    )
    parser.add_argument(
        "--merge",
        action="store_true",
        help="是否将输入目录下的所有图片合并为一个 PDF",
    )
    parser.add_argument(
        "--merge-name",
        default="merged_output.pdf",
        help="合并后的 PDF 文件名（仅在 --merge 下生效）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖已存在的输出文件",
    )
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    img_files = collect_images(input_path)
    if not img_files:
        logger.info("action=exit reason=no_images_found path=%s", input_path)
        return

    logger.info("action=start total_images=%d mode=%s", len(img_files), "merge" if args.merge else "single")

    if args.merge:
        success = merge_images(img_files, output_dir, output_filename=args.merge_name, overwrite=args.overwrite)
        is_success = success
    else:
        success_count = 0
        for p in img_files:
            if convert_single(p, output_dir, overwrite=args.overwrite):
                success_count += 1
        logger.info("action=finish success=%d total=%d", success_count, len(img_files))
        is_success = success_count > 0

    if is_success:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = output_dir / f"backup_{timestamp}"
        backup_dir.mkdir(parents=True, exist_ok=True)

        move_count = 0
        for p in img_files:
            try:
                shutil.move(str(p), str(backup_dir / p.name))
                move_count += 1
            except Exception as e:
                logger.warning("action=backup_failed file=%s error=%s", p.name, str(e))

        logger.info("action=backup_complete total_moved=%d backup_path=%s", move_count, backup_dir.name)


if __name__ == "__main__":
    main()
