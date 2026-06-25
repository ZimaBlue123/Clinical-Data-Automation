# Clinical Data Automation Toolkit

Clinical data automation toolkit for PDF extraction/normalization, Excel chart generation, PPT integration, document translation, network diagnostics, and more.

> **English version**: `README_EN.md` (this file)  
> **Chinese version**: `README.md`  
> **Maintenance rule**: the two README files should be updated **in sync**.

## Requirements

- Python 3.8+
- Dependencies: `requirements.txt` (full); `requirements-ci.txt` + `pytest` for CI smoke tests
- Windows-only dependency (`pywin32`) is required only for Windows + Microsoft Office COM automation modules.

## Install

```bash
pip install -r requirements.txt
```

> Notes in `requirements.txt` explain optional installs:
> - You may comment Paddle-related lines if you do not need **16_PPTX_PDF_to_PPT**
> - You may remove/comment `pywin32` if you do not use Office automation modules
>
> **Optional pre-commit**: install hooks to run YAML checks and a local secret scanner (`scripts/check_secrets.py`, mainly for **33_SAE_Extractor**-style API tokens) before each commit:
>
> ```bash
> pip install pre-commit
> pre-commit install
> ```

## Architecture overview

This repository is organized as **numbered, mostly-independent tool modules** (`01_` … `33_`).  
Two rules define the structure:

1. **Clear layering**: root-level shared resources (`src/`, `config*.yaml`) are separated from module folders.
2. **Fixed module order**: module numbering is the single source of truth, always read/maintain in ascending order (`01_` → `33_`).

| Layer | Description |
|------|-------------|
| **Entry points** | `*.py` scripts under each `NN_*/` folder or the folder `README.md` instructions. |
| **Shared library** | `src/`: `pdf_reader`, `excel_writer`, `color_theme`, etc. |
| **Config** | `config.yaml` / `config.example.yaml` for **rule-driven extraction** via `13_PDF_to_Excel_Rule_Extract`. |
| **I/O convention** | Default `input/` → script → `output/`, with CLI overrides in some modules. |
| **Runtime** | Pure Python libraries (openpyxl/pandas/PyMuPDF…) + optional Windows Office COM automation (`pywin32`). |

### Module groups and order (strictly by number)

- **Excel (01-02)**: `01_Excel_Charts` → `02_Excel_Chart_Colors`
- **PowerPoint (03-05)**: `03_PPT_Merge` → `04_PPT_Watermark_Removal` → `05_PPT_to_PDF`
- **Word (06-11)**: `06_Word_to_PDF` → … → `10_Word_Style_Cleaner` → `11_Word_Text_Replace`
- **PDF (12-23)**: `12_PDF_Batch_to_Excel` → `13_PDF_to_Excel_Rule_Extract` → `14_PDF_to_PPT` → `15_PDF_XSS` → `16_PPTX_PDF_to_PPT` → `17_PDF_Title_Renamer` → `18_PDF_eCTD_Converter` → `19_PDF_Merge` → `20_PDF_Bookmark_Inherit_Zoom` → `21_PDF_Watermark_Removal` → `22_PDF_Duplicate_Analyzer` → `23_PDF_Threat_Analyzer`
- **Utilities (22-32)**: `22_PDF_Duplicate_Analyzer` → `24_File_Translator` → `25_Py_to_EXE` → `26_C_Drive_Cleanup` → `27_WiFi_Passwords` → `28_Folder_File_Count` → `29_Paper_Batch_Download` → `30_Proxy_Config_Export` → `31_DNS_Leak_Detector` → `32_Network_Speed_Test` → `33_SAE_Extractor`

## Project structure

```
Clinical Data Automation/
├── # —— Excel as main input ——
├── 01_Excel_Charts/          # Excel chart generation + clinical table filler (GMC/GMI/seroconversion)
├── 02_Excel_Chart_Colors/    # Excel chart recoloring (clinical palettes)
│
├── # —— PowerPoint as main input ——
├── 03_PPT_Merge/             # PPT merge & narrative restructuring
├── 04_PPT_Watermark_Removal/ # Remove corner logo/watermark in PPTX screenshots
├── 05_PPT_to_PDF/            # Batch PPT/PPTX → PDF (PowerPoint COM)
│
├── # —— Word as main input ——
├── 06_Word_to_PDF/           # Batch Word → PDF (Word COM)
├── 07_Word_to_Excel_to_Figure/     # Word → Excel (tables + chart data) replication
├── 08_Word_Tables_to_Excel/         # Export selected Word tables to Excel
├── 09_Word_All_Tables_to_Excel/     # Export all top-level Word tables to Excel; also merge to one list
│
├── 10_Word_Style_Cleaner/         # Word styles cleanup & normalization (custom-only, keep built-ins)
├── 11_Word_Text_Replace/          # Batch OOXML text replace in docx (dates, IDs, literals)
│
├── # —— PDF as main input (incl. PDF/PPTX → editable PPT) ——
├── 12_PDF_Batch_to_Excel/         # Specialized/batch PDF → Excel (serology, ADR, audits)
├── 13_PDF_to_Excel_Rule_Extract/  # Generic rule-driven PDF → Excel (config.yaml)
├── 14_PDF_to_PPT/                 # PDF → PPT
├── 15_PDF_XSS/                    # PDF XSS/script/link sanitization
├── 16_PPTX_PDF_to_PPT/            # PDF/PPTX → native editable PPTX (table reconstruction)
├── 17_PDF_Title_Renamer/          # PDF title-driven rename (Sanitizer v6.9, move not copy)
│   ├── pdf_sanitizer.py           # Visual hierarchy + academic first-page + optional OCR
│   └── README.md                  # Module docs, noise filters, CLI
├── 18_PDF_eCTD_Converter/         # eCTD sanitize: 6.26 fonts, bookmark repair, audit Excel (see module README)
├── 19_PDF_Merge/                  # Merge PDFs in natural sort order
├── 20_PDF_Bookmark_Inherit_Zoom/  # Bookmark inherit-zoom (XYZ zoom=0)
├── 21_PDF_Watermark_Removal/      # Detect watermark/interference zones + audit masks + clean text
│
├── # —— Multi-format / Utilities ——
├── 22_PDF_Duplicate_Analyzer/   # PDF duplicate scan across subfolders (no input/)
├── 23_PDF_Threat_Analyzer/   # PDF threat analysis + industrial-grade sanitization (PyMuPDF/pypdf fallback)
│   ├── input/                        # Input: PDFs to analyze
│   ├── output/                       # Output: threat_report_*.json + *_sanitized.pdf
│   ├── pdf_threat_analyzer.py        # Main: static scan + risk score + optional sanitize + standard self-check
│   └── README.md                     # Module docs, CLI, fallback chain, secondary-scan recommendations
│
├── 24_File_Translator/
├── 25_Py_to_EXE/
├── 26_C_Drive_Cleanup/
├── 27_WiFi_Passwords/
├── 28_Folder_File_Count/
├── 29_Paper_Batch_Download/
├── 30_Proxy_Config_Export/
├── 31_DNS_Leak_Detector/
├── 32_Network_Speed_Test/      # Speed test + LAN device survey (menu 4)
├── 33_SAE_Extractor/           # SAE extraction (LLM + multi-format → Excel)
│
├── scripts/                    # Repo-wide helpers
│   ├── check_secrets.py        # Secret-pattern scan (pre-commit)
│   ├── set_env.ps1             # PowerShell: sample env for 33_SAE_Extractor
│   └── start_tunnel.ps1        # PowerShell: SSH port forward for API gateway
├── .pre-commit-config.yaml
├── src/
├── config.example.yaml
├── requirements.txt
├── LICENSE.md
└── README.md
```

## Modules

### Top Navigation (group jump)

- [Excel (01-02)](#modules-excel)
- [PowerPoint (03-05)](#modules-ppt)
- [Word (06-11)](#modules-word)
- [PDF (12-23)](#modules-pdf)
- [Utilities (24-33)](#modules-others)

### 01. Excel chart generation (`01_Excel_Charts`)<span id="modules-excel"></span>
Purpose: generate ADR combo charts (bar + line), with optional clinical palettes.

```bash
cd 01_Excel_Charts
python build_charts_xlsxwriter.py
```

**Clinical table data filler (unified entry: GMC + GMI + seroconversion rate):**
```bash
# Process all supported sheets (auto-detect GMC/GMI/seroconversion)
python fill_clinical_table.py input/TVAX-006.xlsx

# Process only GMC
python fill_clinical_table.py input/TVAX-006.xlsx --type gmc

# Process only GMI
python fill_clinical_table.py input/TVAX-006.xlsx --type gmi

# Process only seroconversion rate
python fill_clinical_table.py input/TVAX-006.xlsx --type yangzhuai

# Specify sheets
python fill_clinical_table.py input/TVAX-006.xlsx --sheets "Total GMC,40-59 GMC"
```
Purpose: GMC/GMI parses data from detailed stats (LS GMC/GMI 95%CI format like `768.17(507.87, 1161.89)`); seroconversion auto-scans to locate rows; GMI auto-adapts to different row structures (e.g. "60+ GMI" uses 51 rows instead of 52).

Common flags: `--type`, `--sheets`, `--output-dir`, `--verbose`.
Output: `01_Excel_Charts/output/`.
See module `README.md` and `CHANGELOG.md` for details.

---

### 02. Excel chart recoloring (`02_Excel_Chart_Colors`)
Purpose: recolor existing Excel charts without modifying original TFL files.

```bash
cd 02_Excel_Chart_Colors
python apply_clinical_colors.py --palette Lancet --n-colors 3
```

Batch mode: `--batch --input "input" --output "output"`.

---

### 03. PPT merge (`03_PPT_Merge`)<span id="modules-ppt"></span>
Purpose: physical merge + narrative restructuring for CSR-style decks.

```bash
cd 03_PPT_Merge
python merge_ppt.py
python ppt_engine.py
```

---

### 04. PPT corner watermark removal (`04_PPT_Watermark_Removal`)
Purpose: remove repeated corner logos from screenshot-heavy PPTX slides.

```bash
cd 04_PPT_Watermark_Removal
python pptx_corner_logo_patch.py
```

---

### 05. PPT batch to PDF (`05_PPT_to_PDF`)
Purpose: batch export PPT/PPTX to PDF (Windows + PowerPoint).

```bash
cd 05_PPT_to_PDF
python ppt_to_pdf.py
```

---

### 06. Word batch to PDF (`06_Word_to_PDF`)<span id="modules-word"></span>
Purpose: batch export Word documents to PDF (Windows + Word + `pywin32`).

```bash
cd 06_Word_to_PDF
python word_to_pdf.py
```

---

### 07. Word to Excel (tables + charts) replication (`07_Word_to_Excel_to_Figure`)
Purpose: write Word/RTF table values back to chart-linked template ranges in Excel.

```bash
cd 07_Word_to_Excel_to_Figure
python word_to_excel_to_figure.py --input-dir "input" --plan-only
python word_to_excel_to_figure.py --input-dir "input" --table-map-json "output/table_mapping_plan_<template>.json"
```

---

### 08. Word selected tables to Excel (`08_Word_Tables_to_Excel`)
Purpose: export selected Word tables by title/index/header keywords.

```bash
cd 08_Word_Tables_to_Excel
python word_tables_to_excel.py --help
```

---

### 09. Word all tables batch export (`09_Word_All_Tables_to_Excel`)
Purpose: export all top-level tables from each Word file (one table per sheet).

```bash
cd 09_Word_All_Tables_to_Excel
python word_all_tables_to_excel.py
```

---

### 10. Word style cleanup & normalization (`10_Word_Style_Cleaner`)
Purpose: batch-remove **unused custom styles** (keep built-in styles unchanged) and normalize remaining style names (Chinese tags for headings/body/tables/captions). Preserves original `styles.xml` namespace declarations, expands style dependency closure (`basedOn`/`link`/`next`), and strips orphan style references to reduce Word repair prompts.

```bash
cd 10_Word_Style_Cleaner
python word_style_cleaner.py --input "input" --output "output" --overwrite
```

Output: `10_Word_Style_Cleaner/output/` (includes `*_styles_cleaned.docx` + report). See `10_Word_Style_Cleaner/README.md`.

---

### 11. Word batch text replace (`11_Word_Text_Replace`)
Purpose: batch-replace CSR/report placeholders (dates, study IDs, literals) in body, tables, and headers/footers (OOXML, cross-run safe).

```bash
cd 11_Word_Text_Replace
python replace_docx.py --yes
python util_check_docx.py --latest
```

Input/output: `input/` → `output/` (`*_updated.docx`). Script roles: [`docs/script_roles.md`](docs/script_roles.md). See `11_Word_Text_Replace/README.md`.

---

### 12. PDF batch to Excel (`12_PDF_Batch_to_Excel`)<span id="modules-pdf"></span>
Purpose: specialized extraction/audit workflows (ADR grades, serology reports).

```bash
cd 12_PDF_Batch_to_Excel
python fill_adr_from_pdf.py
python util_audit_and_fix_consistency.py
python serology_report_pdf_to_excel.py --input "input" --output "output/serology_report_merged.xlsx" --ocr --ocr-dpi 110
```

---

### 13. PDF to Excel rule extraction (`13_PDF_to_Excel_Rule_Extract`)
Purpose: generic rule-based PDF extraction driven by `config.yaml`.

```bash
cd 13_PDF_to_Excel_Rule_Extract
python main.py
```

Optional exclusion boxes from module 19:
```bash
python main.py --config config.yaml --exclusion-json "../20_PDF_Watermark_Removal/output/your_boxes.json"
```

---

### 14. PDF to PPT (`14_PDF_to_PPT`)
Purpose: convert each PDF page into one PPT slide.

```bash
cd 13_PDF_to_PPT
python pdf_to_ppt.py
```

---

### 15. PDF XSS cleaning (`15_PDF_XSS`)
Purpose: sanitize script/protocol/link risks in PDFs.

```bash
cd 14_PDF_XSS
python pdf_xss_clean.py
```

---

### 16. PPTX/PDF to native editable PPT (`16_PPTX_PDF_to_PPT`)
Purpose: reconstruct editable tables from PDF/image-based PPTX into native PPT.

```bash
cd 16_PPTX_PDF_to_PPT
python convert_to_native_ppt.py
```

---

### 17. PDF title-driven rename (`17_PDF_Title_Renamer`)
Purpose: extract the **article title** and **year** from the first page, write `Title-Words-YYYY.pdf`, and **move** (not copy) PDFs from `input/` to `output/` (engine v6.9).

**Pipeline**: font-size visual hierarchy → academic first-page line merge → metadata/first lines → optional OCR (`pytesseract` + Tesseract).

**Noise filters**: journal mastheads, article-type banners, Elsevier “Article in Press” lines, “Please cite this article…” prompts, volume/issue headers, author-only lines; FDA “Guidance for Industry” covers strip generic prefixes. Filenames keep scientific tokens (e.g. CpG, mRNA) with smart English title case.

```bash
cd 16_PDF_Title_Renamer
python pdf_sanitizer.py
```

Common flags: `--input`, `--output`, `--no-recursive`, `--no-keep-structure`, `--overwrite`.  
**Warning**: source PDFs are removed from `input/` after a successful run—back up first. See `16_PDF_Title_Renamer/README.md`.

---

### 18. PDF eCTD converter (`18_PDF_eCTD_Converter`)
Purpose: validate, sanitize, and rewrite PDFs for common eCTD Annex 6 checks; includes **6.26** font mapping/embedding, and **6.5/6.6/6.8** bookmark fixes (assign GoTo to inactive entries, flatten multi-level TOC to avoid collapse parents, inherit zoom). Exports an Excel audit report (including a structure-warnings sheet).

```bash
cd 18_PDF_eCTD_Converter
python pdf_ectd_converter.py --input "./input" --output "./output" --report "./ectd_report.xlsx" --overwrite
```

Common flags: `--validate-only`, `--overwrite`, `--add-auto-bookmarks` (default `outline`), `--no-add-auto-bookmarks`.  
Dependencies: `pymupdf`, `pandas`, `openpyxl`, `fonttools` (font embedding). See `18_PDF_eCTD_Converter/README.md`.

---

### 19. PDF merge (`19_PDF_Merge`)
Purpose: merge PDFs in natural-sort order.

```bash
cd 18_PDF_Merge
python merge_pdf.py
```

---

### 20. PDF bookmark inherit zoom (`20_PDF_Bookmark_Inherit_Zoom`)
Purpose: force bookmark jumps to keep current zoom (`XYZ + zoom=0`).

```bash
cd 19_PDF_Bookmark_Inherit_Zoom
python pdf_bookmark_inherit_zoom.py
```

---

### 21. PDF watermark/interference detection & audit (`21_PDF_Watermark_Removal`)
Purpose: detect interference zones and output exclusion JSON + audit overlays + cleaned text.

```bash
cd 20_PDF_Watermark_Removal
python main.py --input "input" --output "output"
```

---

### 22. PDF duplicate analysis (`22_PDF_Duplicate_Analyzer`)<span id="modules-others"></span>
Purpose: detect duplicate PDFs across subfolders under one root path (same filename or same first-page text). No `input/` — specify external paths via CLI or JSON config.

```bash
cd 22_PDF_Duplicate_Analyzer
python pdf_duplicate_analyzer.py --root "D:\References" --folders "FolderA,FolderB" --label "batch1"
python pdf_duplicate_analyzer.py --config jobs.json
```

Output: `22_PDF_Duplicate_Analyzer/output/duplicate_report_*.txt`.

---

### 24 Bidirectional file translation (`24_File_Translator`)
Purpose: translate Excel/CSV/Word/PDF with fallback providers.

```bash
cd 24_File_Translator
python file_translator.py --self-test
python file_translator.py
```

---

### 25 Python script to EXE (`25_Py_to_EXE`)
Purpose: package `.py` files into Windows executables via PyInstaller.

```bash
cd 25_Py_to_EXE
python py_to_exe.py
```

---

### 26 C drive cleanup (`26_C_Drive_Cleanup`)
Purpose: scan/clean common temporary and cache files on C drive.

```bash
cd 26_C_Drive_Cleanup
python c_drive_cleanup.py
python c_drive_cleanup.py --delete --days 7
```

---

### 27 WiFi passwords export (`27_WiFi_Passwords`)
Purpose: export saved Windows WiFi credentials to CSV.

```bash
cd 27_WiFi_Passwords
python wifi_passwords.py
```

---

### 28 Folder file count (`28_Folder_File_Count`)
Purpose: recursively count files and export TXT/Excel reports.

```bash
cd 28_Folder_File_Count
python folder_file_count.py --path "D:\data"
```

---

### 29 Paper batch download (`29_Paper_Batch_Download`)
Purpose: batch download OA papers by DOI/PMID/title/URL with default safe-mode throttling and retry backoff to reduce IP rate-limit risk.

```bash
cd 29_Paper_Batch_Download
python paper_batch_download.py --queries "10.1038/s41586-020-2649-2" "32788730"
```

File-input mode: `python paper_batch_download.py --file "D:\papers.txt" --mailto "your_email@example.com"`.

Safe-mode options (enabled by default): `--safe-mode`, `--min-interval`, `--max-retries`, `--backoff-base`, `--mirror-cooldown`.

---

### 30 Proxy config export (`30_Proxy_Config_Export`)
Purpose: export current Windows proxy settings (registry + env vars).

```bash
cd 30_Proxy_Config_Export
python proxy_config_export.py
```

---

### 31 DNS leak diagnostics (`31_DNS_Leak_Detector`)
Purpose: compare egress IP and DNS upstream location for leak/split-routing checks.

```bash
cd 31_DNS_Leak_Detector
python dns_leak_detector.py --mode tun
python dns_leak_detector.py --mode socks --socks-port 10808
python dns_leak_detector.py --save-json
```

---

### 32 Network speed test & LAN occupancy survey (`32_Network_Speed_Test`)
Purpose: domestic/international download probes and VPN comparison (direct route vs SOCKS). **Menu 4** scans online LAN devices (IP/MAC) and explains router QoS rate limits (no ARP attack tooling).

```bash
cd 32_Network_Speed_Test
python network_speed_test.py
```

Interactive: **4** = LAN device survey (recommended when the network feels slow). CLI flags: `--socks-port`, `--skip-vpn`, `--save-json`. Reports: `output/speed_test_*.json`, `output/lan_survey_*.json`. Requires `PySocks` for SOCKS tests.

---

### 33 SAE structured extraction (`33_SAE_Extractor`)
Purpose: extract serious adverse event (SAE) fields from clinical PDFs, text, DOCX, or Excel via an OpenAI-compatible Chat Completions API and export to Excel.

```bash
cd 33_SAE_Extractor
python cli.py self-check
python cli.py batch
python cli.py pdf-batch
```

Requires `SAE_API_TOKEN`; optional `SAE_API_BASE_URL`, `SAE_MODEL_ID`, `TESSERACT_CMD`, `POPPLER_PATH`. Default I/O: `input/` and `output/`. See `33_SAE_Extractor/README.md`.

Helper scripts at repo root (`scripts/`, Windows PowerShell):

- `set_env.ps1`: sets `SAE_API_BASE_URL` and `SAE_OUTPUT_DIR` (defaults to `33_SAE_Extractor/output`); **you must still set `SAE_API_TOKEN`**. Run from repo root: `.\scripts\set_env.ps1` (may require execution policy).
- `start_tunnel.ps1`: SSH local port forward. **Requires `SAE_TUNNEL_SSH_HOST`**; optional `SAE_TUNNEL_SSH_USER`, `SAE_TUNNEL_LOCAL_PORT`, `SAE_TUNNEL_REMOTE_BIND`. Run: `powershell -ExecutionPolicy Bypass -File .\scripts\start_tunnel.ps1`

Manual secret scan: `pre-commit run detect-sensitive-secrets --all-files`.

## Configuration

Copy `config.example.yaml` to `config.yaml` and edit as needed (paths are relative to project root):

```yaml
pdf_path: "13_PDF_to_Excel_Rule_Extract/input/your.pdf"
excel_path: "13_PDF_to_Excel_Rule_Extract/output/out.xlsx"
rules:
  - name: "Rule 1"
    search:
      keyword: "keyword"
      page: 1
    excel:
      sheet: "Sheet1"
      cell: "B3"
```

## Extension points

- PDF reader/exclusion/mapping audit: `src/pdf_reader.py`
- Watermark/interference pipeline: `20_PDF_Watermark_Removal/main.py` and `20_PDF_Watermark_Removal/steps/`
- Excel writer helpers: `src/excel_writer.py`
- Chart styles/palettes: `01_Excel_Charts/build_charts_xlsxwriter.py`, `src/color_theme.py`
- Word→Excel replication: `07_Word_to_Excel_to_Figure/`

## Notes

1. Put inputs under each module's `input/` and check results in `output/`
2. XlsxWriter is recommended for chart generation
3. Module dependencies are documented in `requirements.txt`

### Serology reconciliation (PDF vs Word)

1. Generate Word merged list: `09_Word_All_Tables_to_Excel/output/word_tables_merged.xlsx`
2. Generate PDF merged list (and optionally backfill missing markers using `--reference-excel`): `12_PDF_Batch_to_Excel/serology_report_pdf_to_excel.py`
3. Compare and export diffs:

```bash
python compare_serology_outputs.py --pdf-excel <PDF.xlsx> --word-excel <WORD.xlsx> --out-csv <diff.csv>
```

## License

MIT License. See `LICENSE.md` (includes attribution for **SAE-Extractor**-derived portions such as `33_SAE_Extractor/`).

## Maintenance Audit (2026-05)

- A repository-wide Python syntax sweep (`compileall`) was completed; no syntax-level failures were found.
- Most modules already follow a clean single-folder/single-entry pattern. The highest-value cross-cutting improvements are logging consistency, exception granularity, and CLI argument validation.
- Recommended priority:
  - P1: normalize CLI flags and exit behavior (`--input/--output/--overwrite`) for easier pipeline chaining.
  - P1: reduce broad `except Exception` handling and include actionable error context.
  - P2: add minimal smoke regression samples (1-2 per critical module).
  - P3: gradually move duplicated I/O/logging helpers into `src/`.

