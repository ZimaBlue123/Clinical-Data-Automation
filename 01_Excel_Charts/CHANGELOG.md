# 01_Excel_Charts Changelog

Records key changes and new features for this module.

## [1.0.0] - 2026-06-25

### Added
- **`fill_clinical_table.py`**: Unified entry script for clinical trial table data filling
  - Supports `GMC` (Geometric Mean Concentration), `GMI` (Geometric Mean Fold Increase), and `seroconversion rate` (阳转率) table types
  - Auto-detects sheet type via `detect_sheet_type()`, no manual specification needed
  - Dynamically scans source data row numbers, adapting to different sheet structures
  - Full type hints, logging, exception handling, and boundary checks
  - Adapts to "60+ GMI" 51-row structure (default 52 rows)

### Changed
- Consolidated `fill_gmc_table.py` and `fill_yangzhuai_table.py` into a single entry point
- Removed all `__pycache__` directories and `.pyc` cache files
- Cleaned up debug script residues

### Data Format

#### GMC (Geometric Mean Concentration)
- Table structure: rows 1-7 (group header, sub-header, pre-vaccination + 4 time points)
- Source data:
  - Row 12: `GMC (95%CI)` → for pre-vaccination
  - Rows 17/22/27/32: `LS GMC (95%CI)` → for 4 time points
- Source format: `"768.17(507.87, 1161.89)"` or `"644.46 (280.78, 1479.20)"`

#### GMI (Geometric Mean Fold Increase)
- Table structure: same as GMC
- Source data: script dynamically scans to locate `GMI (95%CI)` rows
  - "60+ GMI" uses 51-row structure (default 52 rows)
- Source format: same as GMC

#### Seroconversion Rate (阳转率)
- Table structure: rows 1-7 (group header, sub-header, pre-vaccination not filled + 4 time points)
- Source data: script auto-scans to locate:
  - Title row: column A contains "一免后" or "全免后"
  - "阳转例数（阳转率）" row: format `"24 (75.00)"`
  - "95%CI" row: format `"56.60, 88.54"`

#### Column Structure (5 groups × 3 columns)

| Column Range | Group |
|--------------|-------|
| B-D | Low-dose adjuvant group (mean, upper, lower) |
| E-G | High-dose adjuvant group |
| H-J | Low-dose test group |
| K-M | High-dose test group |
| N-P | Placebo group |

## Usage Examples

```bash
# Process all supported sheets
python fill_clinical_table.py input/TVAX-006.xlsx

# Process only GMI
python fill_clinical_table.py input/TVAX-006.xlsx --type gmi

# Specify output directory
python fill_clinical_table.py input/TVAX-006.xlsx -o ./output

# Show verbose logs
python fill_clinical_table.py input/TVAX-006.xlsx -v
```

## Dependencies

- `openpyxl>=3.1.0` (Excel I/O)