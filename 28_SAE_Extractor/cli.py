import argparse
import logging
from pathlib import Path

from parsers import configure_ocr
from pipelines import PipelineIO, run_omni_batch, run_pdf_batch
from sae_extractor import ClinicalDataGuard
from settings import check_environment, ensure_dirs, load_settings, resolve_default_output


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def _build_engine():
    settings = load_settings()
    ensure_dirs(settings)
    configure_ocr(settings)

    env = check_environment(settings)
    logging.info(
        "action=env_check tesseract_ok=%s poppler_ok=%s token_set=%s",
        env["tesseract"]["ok"],
        env["poppler"]["ok"],
        env["api"]["token_set"],
    )

    if not settings.api_token:
        raise SystemExit("未检测到 SAE_API_TOKEN，请先配置环境变量后再运行。")

    engine = ClinicalDataGuard(
        base_url=settings.api_base_url,
        token=settings.api_token,
        model_id=settings.model_id,
    )
    return settings, engine


def main() -> int:
    parser = argparse.ArgumentParser(description="SAE 抽取：自检 / 多格式批处理 / PDF 批处理")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("self-check", help="只执行环境自检并退出")

    p_omni = sub.add_parser("batch", help="多格式批处理（TXT/PDF/DOCX/Excel）")
    p_omni.add_argument("--input-dir", default=None, help="输入目录（默认：SAE_INPUT_DIR 或本模块 input/）")
    p_omni.add_argument("--output", default=None, help="输出 Excel 路径（默认：output/SAE_Omni_Listing.xlsx）")
    p_omni.add_argument("--fast-docx", action="store_true", help="DOCX 极速模式（大文件含图时推荐）")

    p_pdf = sub.add_parser("pdf-batch", help="仅 PDF 批处理（文本层 + OCR 回退）")
    p_pdf.add_argument("--input-dir", default=None, help="输入目录（默认：SAE_INPUT_DIR 或本模块 input/）")
    p_pdf.add_argument("--output", default=None, help="输出 Excel 路径（默认：output/SAE_Structured_Listing_PDF.xlsx）")

    args = parser.parse_args()

    settings = load_settings()
    ensure_dirs(settings)

    if args.cmd == "self-check":
        logging.info("action=self_check result=%s", check_environment(settings))
        return 0

    settings, engine = _build_engine()

    if args.cmd == "batch":
        io = PipelineIO(
            input_dir=Path(args.input_dir) if args.input_dir else settings.input_dir,
            output_file=Path(args.output) if args.output else resolve_default_output(settings, "SAE_Omni_Listing.xlsx"),
        )
        run_omni_batch(engine=engine, io=io, settings=settings, fast_docx=bool(args.fast_docx))
        return 0

    if args.cmd == "pdf-batch":
        io = PipelineIO(
            input_dir=Path(args.input_dir) if args.input_dir else settings.input_dir,
            output_file=Path(args.output) if args.output else resolve_default_output(settings, "SAE_Structured_Listing_PDF.xlsx"),
        )
        run_pdf_batch(engine=engine, io=io, settings=settings)
        return 0

    raise SystemExit(f"未知命令: {args.cmd}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        logging.error("action=interrupted_by_user")
        raise SystemExit(130)
    except Exception:
        logging.exception("action=cli_failed")
        raise SystemExit(1)
