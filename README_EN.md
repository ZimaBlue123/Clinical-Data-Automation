# Clinical Data Automation Toolkit

Clinical data automation toolkit for PDF extraction/normalization, Excel chart generation, PPT integration, document translation, network diagnostics, and more.

> **English version**: `README_EN.md` (this file)  
> **Chinese version**: `README.md`  
> **Maintenance rule**: the two README files should be updated **in sync**.

## Requirements

- Python 3.8+
- Dependencies: `requirements.txt`
- Windows-only dependency (`pywin32`) is required only for Windows + Microsoft Office COM automation modules.

## Install

```bash
pip install -r requirements.txt
```

> Notes in `requirements.txt` explain optional installs:
> - You may comment Paddle-related lines if you do not need **14_PPTX_PDF_to_PPT**
> - You may remove/comment `pywin32` if you do not use Office automation modules

## Architecture overview

This repository is organized as **numbered, mostly-independent tool modules** (`01_` … `27_`).  
Two rules define the structure:

1. **Clear layering**: root-level shared resources (`src/`, `config*.yaml`) are separated from module folders.
2. **Fixed module order**: module numbering is the single source of truth, always read/maintain in ascending order (`01_` → `27_`).

| Layer | Description |
|------|-------------|
| **Entry points** | `*.py` scripts under each `NN_*/` folder or the folder `README.md` instructions. |
| **Shared library** | `src/`: `pdf_reader`, `excel_writer`, `color_theme`, etc. |
| **Config** | `config.yaml` / `config.example.yaml` for **rule-driven extraction** via `11_PDF_to_Excel_Rule_Extract`. |
| **I/O convention** | Default `input/` → script → `output/`, with CLI overrides in some modules. |
| **Runtime** | Pure Python libraries (openpyxl/pandas/PyMuPDF…) + optional Windows Office COM automation (`pywin32`). |

### Module groups and order (strictly by number)

- **Excel (01-02)**: `01_Excel_Charts` → `02_Excel_Chart_Colors`
- **PowerPoint (03-05)**: `03_PPT_Merge` → `04_PPT_Watermark_Removal` → `05_PPT_to_PDF`
- **Word (06-09)**: `06_Word_to_PDF` → `07_Word_to_Excel_to_Figure` → `08_Word_Tables_to_Excel` → `09_Word_All_Tables_to_Excel`
- **PDF (10-19)**: `10_PDF_Batch_to_Excel` → `11_PDF_to_Excel_Rule_Extract` → `12_PDF_to_PPT` → `13_PDF_XSS` → `14_PPTX_PDF_to_PPT` → `15_PDF_Sanitizer` → `16_PDF_eCTD_Converter` → `17_PDF_Merge` → `18_PDF_Bookmark_Inherit_Zoom` → `19_PDF_Watermark_Removal`
- **Utilities (20-27)**: `20_File_Translator` → `21_Py_to_EXE` → `22_C_Drive_Cleanup` → `23_WiFi_Passwords` → `24_Folder_File_Count` → `25_Paper_Batch_Download` → `26_Proxy_Config_Export` → `27_DNS_Leak_Detector`

## Project structure

```
Clinical Data Automation/
├── # —— Excel as main input ——
├── 01_Excel_Charts/          # Excel chart generation
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
├── # —— PDF as main input (incl. PDF/PPTX → editable PPT) ——
├── 10_PDF_Batch_to_Excel/          # Specialized/batch PDF → Excel (serology, ADR, audits)
├── 11_PDF_to_Excel_Rule_Extract/   # Generic rule-driven PDF → Excel (config.yaml)
├── 12_PDF_to_PPT/                 # PDF → PPT
├── 13_PDF_XSS/                    # PDF XSS/script/link sanitization
├── 14_PPTX_PDF_to_PPT/            # PDF/PPTX → native editable PPTX (table reconstruction)
├── 15_PDF_Sanitizer/              # PDF filename/title-based sanitization & move
├── 16_PDF_eCTD_Converter/         # PDF eCTD compliance conversion
├── 17_PDF_Merge/                  # Merge PDFs in natural sort order
├── 18_PDF_Bookmark_Inherit_Zoom/  # Bookmark inherit-zoom (XYZ zoom=0)
├── 19_PDF_Watermark_Removal/      # Detect watermark/interference zones + audit masks + clean text
│
├── # —— Multi-format / Utilities ——
├── 20_File_Translator/
├── 21_Py_to_EXE/
├── 22_C_Drive_Cleanup/
├── 23_WiFi_Passwords/
├── 24_Folder_File_Count/
├── 25_Paper_Batch_Download/
├── 26_Proxy_Config_Export/
├── 27_DNS_Leak_Detector/
│
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
- [Word (06-09)](#modules-word)
- [PDF (10-19)](#modules-pdf)
- [Utilities (20-27)](#modules-others)

### 01. Excel chart generation (`01_Excel_Charts`)<span id="modules-excel"></span>
Purpose: generate ADR combo charts (bar + line), with optional clinical palettes.

```bash
cd 01_Excel_Charts
python build_charts_xlsxwriter.py
```

Common flags: `--input`, `--output`, `--clinical-colors`.  
Output: `01_Excel_Charts/output/`.

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

### 10. PDF batch to Excel (`10_PDF_Batch_to_Excel`)<span id="modules-pdf"></span>
Purpose: specialized extraction/audit workflows (ADR grades, serology reports).

```bash
cd 10_PDF_Batch_to_Excel
python fill_adr_from_pdf.py
python audit_and_fix_consistency.py
python serology_report_pdf_to_excel.py --input "input" --output "output/serology_report_merged.xlsx" --ocr --ocr-dpi 110
```

---

### 11. PDF to Excel rule extraction (`11_PDF_to_Excel_Rule_Extract`)
Purpose: generic rule-based PDF extraction driven by `config.yaml`.

```bash
cd 11_PDF_to_Excel_Rule_Extract
python main.py
```

Optional exclusion boxes from module 19:
```bash
python main.py --config config.yaml --exclusion-json "../19_PDF_Watermark_Removal/output/your_boxes.json"
```

---

### 12. PDF to PPT (`12_PDF_to_PPT`)
Purpose: convert each PDF page into one PPT slide.

```bash
cd 12_PDF_to_PPT
python pdf_to_ppt.py
```

---

### 13. PDF XSS cleaning (`13_PDF_XSS`)
Purpose: sanitize script/protocol/link risks in PDFs.

```bash
cd 13_PDF_XSS
python pdf_xss_clean.py
```

---

### 14. PPTX/PDF to native editable PPT (`14_PPTX_PDF_to_PPT`)
Purpose: reconstruct editable tables from PDF/image-based PPTX into native PPT.

```bash
cd 14_PPTX_PDF_to_PPT
python convert_to_native_ppt.py
```

---

### 15. PDF title-driven rename (`15_PDF_Sanitizer`)
Purpose: extract canonical titles and rename/move PDF files to `output/`.

```bash
cd 15_PDF_Sanitizer
python pdf_sanitizer.py
```

---

### 16. PDF eCTD converter (`16_PDF_eCTD_Converter`)
Purpose: convert PDFs to eCTD-friendly outputs with audit report.

```bash
cd 16_PDF_eCTD_Converter
python pdf_ectd_converter.py --report "output/ectd_report.xlsx"
```

---

### 17. PDF merge (`17_PDF_Merge`)
Purpose: merge PDFs in natural-sort order.

```bash
cd 17_PDF_Merge
python merge_pdf.py
```

---

### 18. PDF bookmark inherit zoom (`18_PDF_Bookmark_Inherit_Zoom`)
Purpose: force bookmark jumps to keep current zoom (`XYZ + zoom=0`).

```bash
cd 18_PDF_Bookmark_Inherit_Zoom
python pdf_bookmark_inherit_zoom.py
```

---

### 19. PDF watermark/interference detection & audit (`19_PDF_Watermark_Removal`)
Purpose: detect interference zones and output exclusion JSON + audit overlays + cleaned text.

```bash
cd 19_PDF_Watermark_Removal
python main.py --input "input" --output "output"
```

---

### 20. Bidirectional file translation (`20_File_Translator`)<span id="modules-others"></span>
Purpose: translate Excel/CSV/Word/PDF with fallback providers.

```bash
cd 20_File_Translator
python file_translator.py --self-test
python file_translator.py
```

---

### 21. Python script to EXE (`21_Py_to_EXE`)
Purpose: package `.py` files into Windows executables via PyInstaller.

```bash
cd 21_Py_to_EXE
python py_to_exe.py
```

---

### 22. C drive cleanup (`22_C_Drive_Cleanup`)
Purpose: scan/clean common temporary and cache files on C drive.

```bash
cd 22_C_Drive_Cleanup
python c_drive_cleanup.py
python c_drive_cleanup.py --delete --days 7
```

---

### 23. WiFi passwords export (`23_WiFi_Passwords`)
Purpose: export saved Windows WiFi credentials to CSV.

```bash
cd 23_WiFi_Passwords
python wifi_passwords.py
```

---

### 24. Folder file count (`24_Folder_File_Count`)
Purpose: recursively count files and export TXT/Excel reports.

```bash
cd 24_Folder_File_Count
python folder_file_count.py --path "D:\data"
```

---

### 25. Paper batch download (`25_Paper_Batch_Download`)
Purpose: batch download OA papers by DOI/PMID/title/URL.

```bash
cd 25_Paper_Batch_Download
python paper_batch_download.py --queries "10.1038/s41586-020-2649-2" "32788730"
```

---

### 26. Proxy config export (`26_Proxy_Config_Export`)
Purpose: export current Windows proxy settings (registry + env vars).

```bash
cd 26_Proxy_Config_Export
python proxy_config_export.py
```

---

### 27. DNS leak diagnostics (`27_DNS_Leak_Detector`)
Purpose: compare egress IP and DNS upstream location for leak/split-routing checks.

```bash
cd 27_DNS_Leak_Detector
python dns_leak_detector.py --mode tun
python dns_leak_detector.py --mode socks --socks-port 10808
python dns_leak_detector.py --save-json
```

## Configuration

Copy `config.example.yaml` to `config.yaml` and edit as needed (paths are relative to project root):

```yaml
pdf_path: "11_PDF_to_Excel_Rule_Extract/input/your.pdf"
excel_path: "11_PDF_to_Excel_Rule_Extract/output/out.xlsx"
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
- Watermark/interference pipeline: `19_PDF_Watermark_Removal/main.py` and `19_PDF_Watermark_Removal/steps/`
- Excel writer helpers: `src/excel_writer.py`
- Chart styles/palettes: `01_Excel_Charts/build_charts_xlsxwriter.py`, `src/color_theme.py`
- Word→Excel replication: `07_Word_to_Excel_to_Figure/`

## Notes

1. Put inputs under each module's `input/` and check results in `output/`
2. XlsxWriter is recommended for chart generation
3. Module dependencies are documented in `requirements.txt`

### Serology reconciliation (PDF vs Word)

1. Generate Word merged list: `09_Word_All_Tables_to_Excel/output/word_tables_merged.xlsx`
2. Generate PDF merged list (and optionally backfill missing markers using `--reference-excel`): `10_PDF_Batch_to_Excel/serology_report_pdf_to_excel.py`
3. Compare and export diffs:

```bash
python compare_serology_outputs.py --pdf-excel <PDF.xlsx> --word-excel <WORD.xlsx> --out-csv <diff.csv>
```

## License

MIT License. See `LICENSE.md`.

