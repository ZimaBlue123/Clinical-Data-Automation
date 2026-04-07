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

This repository is organized as **numbered, mostly-independent tool modules** (`01_` … `27_`), grouped by the main input file type:
**Excel → PowerPoint → Word → PDF → Others**.

| Layer | Description |
|------|-------------|
| **Entry points** | `*.py` scripts under each `NN_*/` folder or the folder `README.md` instructions. |
| **Shared library** | `src/`: `pdf_reader`, `excel_writer`, `color_theme`, etc. |
| **Config** | `config.yaml` / `config.example.yaml` for **rule-driven extraction** via `11_PDF_to_Excel_Rule_Extract`. |
| **I/O convention** | Default `input/` → script → `output/`, with CLI overrides in some modules. |
| **Runtime** | Pure Python libraries (openpyxl/pandas/PyMuPDF…) + optional Windows Office COM automation (`pywin32`). |

**Module 07_Word_to_Excel_to_Figure data flow (high level)**: Word/RTF in `input/` → `word_to_excel_to_figure.py` reads chart series ranges from `Template/*.xlsx` → matches Word tables and writes values back via Excel COM; mapping plan is generated/parsed by `table_mapping_logic.py`. If openpyxl saves break pivot/OLAP structures, use `repair_output_by_patch.py` to patch chart ranges back into a clean template copy.

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

### 1. Excel chart generation (`01_Excel_Charts`)

Generates publication/CSR-style ADR combo charts (bars + lines).

**Steps**

1. Put the source Excel into `01_Excel_Charts/input/`
2. Run:

```bash
cd 01_Excel_Charts
python build_charts_xlsxwriter.py
```

3. Output: `01_Excel_Charts/output/...xlsx`

### 1.1 Excel chart recoloring (`02_Excel_Chart_Colors`)

Apply clinical/journal palettes to existing Excel charts without changing the original file.

```bash
cd 02_Excel_Chart_Colors
python apply_clinical_colors.py
```

### 2. PPT merge (`03_PPT_Merge`)

Merge multiple PPTs and optionally restructure into a CSR narrative.

```bash
cd 03_PPT_Merge
python merge_ppt.py
python ppt_engine.py
```

### 3. PDF → Excel (rule-driven, generic) (`11_PDF_to_Excel_Rule_Extract`)

Extract text/tables from PDF into Excel according to `config.yaml`.

1. Copy config:

```bash
copy config.example.yaml config.yaml
```

2. Edit `config.yaml`:

```yaml
pdf_path: "11_PDF_to_Excel_Rule_Extract/input/your.pdf"
excel_path: "11_PDF_to_Excel_Rule_Extract/output/out.xlsx"
rules:
  - name: "Example rule"
    search:
      keyword: "Keyword"
      page: 1
    excel:
      sheet: "Sheet1"
      cell: "B3"
```

3. Run:

```bash
cd 11_PDF_to_Excel_Rule_Extract
python main.py
```

**Optional: watermark/interference exclusion + mapping audit**

Generate `*_boxes.json` via `19_PDF_Watermark_Removal`, then:

```bash
python main.py --config config.yaml --exclusion-json "../19_PDF_Watermark_Removal/output/your_boxes.json"
```

### 3.2 ADR grade extraction (specialized) (`10_PDF_Batch_to_Excel`)

```bash
cd 10_PDF_Batch_to_Excel
python fill_adr_from_pdf.py
```

### 3.3 Consistency audit/fix (`10_PDF_Batch_to_Excel`)

```bash
cd 10_PDF_Batch_to_Excel
python audit_and_fix_consistency.py
```

### 3.4 Serology report PDF → Excel (with OCR) (`10_PDF_Batch_to_Excel`)

```bash
cd 10_PDF_Batch_to_Excel
python serology_report_pdf_to_excel.py --input "input" --output "output/serology_report_merged.xlsx" --ocr --ocr-dpi 110
```

### 4. PDF → PPT (`12_PDF_to_PPT`)

```bash
cd 12_PDF_to_PPT
python pdf_to_ppt.py
```

### 6. PPT corner watermark removal (`04_PPT_Watermark_Removal`)

```bash
cd 04_PPT_Watermark_Removal
python pptx_corner_logo_patch.py
```

### 7. PDF XSS cleaning (`13_PDF_XSS`)

```bash
cd 13_PDF_XSS
python pdf_xss_clean.py
```

### 8. PPT batch to PDF (`05_PPT_to_PDF`)

```bash
cd 05_PPT_to_PDF
python ppt_to_pdf.py
```

### 9. Word batch to PDF (`06_Word_to_PDF`)

```bash
cd 06_Word_to_PDF
python word_to_pdf.py
```

### 24. Word selected tables → Excel (`08_Word_Tables_to_Excel`)

See `08_Word_Tables_to_Excel/README.md` for table-selection strategies and flags.

### 25. PDF watermark/interference detection & audit (`19_PDF_Watermark_Removal`)

Detect interference zones (watermark/header/footer) without physically removing content.
Outputs exclusion boxes JSON + audit-masked PDF + clean extracted text for downstream extraction.

```bash
cd 19_PDF_Watermark_Removal
python main.py --input "input" --output "output"
```

### 26. Word all tables batch export (`09_Word_All_Tables_to_Excel`)

```bash
cd 09_Word_All_Tables_to_Excel
python word_all_tables_to_excel.py
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

