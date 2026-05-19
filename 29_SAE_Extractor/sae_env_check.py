import importlib

from settings import check_environment, load_settings


def main() -> int:
    settings = load_settings()
    env = check_environment(settings)

    print("=" * 60)
    print("SAE Extractor 环境自检")
    print("=" * 60)
    print(f"Project root : {env['paths']['project_root']}")
    print(f"Input dir    : {env['paths']['input_dir']}")
    print(f"Output dir   : {env['paths']['output_dir']}")
    print("-" * 60)
    print(f"Tesseract OK : {env['tesseract']['ok']} | source={env['tesseract']['source']} | {env['tesseract']['detail']}")
    print(f"Poppler  OK  : {env['poppler']['ok']} | source={env['poppler']['source']} | {env['poppler']['detail']}")
    print(f"API base_url : {env['api']['base_url']}")
    print(f"API token set: {env['api']['token_set']}")
    print(f"Model id     : {env['api']['model_id']}")
    print("-" * 60)

    required_modules = [
        "requests",
        "pandas",
        "openpyxl",
        "pdfplumber",
        "pytesseract",
        "pdf2image",
        "docx",
        "PIL",
    ]
    missing: list[str] = []
    for m in required_modules:
        try:
            importlib.import_module(m)
        except Exception:
            missing.append(m)

    if missing:
        print("[Python依赖] 缺失/不可导入：", ", ".join(missing))
        return 2

    print("[Python依赖] OK（关键依赖均可导入）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
