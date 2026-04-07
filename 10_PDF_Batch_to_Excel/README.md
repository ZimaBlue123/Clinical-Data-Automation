# 10_PDF_Batch_to_Excel

本目录下的 `serology_report_pdf_to_excel.py` 用于把中检院血清样本检测报告 PDF 转成血清汇总 Excel。

输出 Excel 结构（与 Word 汇总一致）：
- A 列：`样品ID`
- 五项指标固定顺序：`Anti-HBs / HBsAg / Anti-HBc / Anti-HBe / HBeAg`
- 每项指标占 2 列：`value`（数值）+ `note`（说明）

## 快速开始

1. 把 PDF 放入 `10_PDF_Batch_to_Excel/input/`
2. 运行：
```bash
cd 10_PDF_Batch_to_Excel
python serology_report_pdf_to_excel.py --input input --output output/serology_report_merged.xlsx
```

## OCR（扫描件）与回填缺项

- 若 PDF 中存在扫描件/无矢量表，建议加 `--ocr`（依赖本机 Tesseract：`chi_sim+eng`）。
- 当你发现输出存在“缺项”，可指定参考 Excel 做回填（Word 汇总可作为参考）：
```bash
python serology_report_pdf_to_excel.py ^
  --input input ^
  --output output/serology_report_merged_final.xlsx ^
  --ocr --ocr-dpi 110 ^
  --reference-excel ../09_Word_All_Tables_to_Excel/output/word_tables_merged.xlsx
```

## 进一步对比（可选）

可以使用根目录脚本对比 PDF 输出与 Word 汇总：
```bash
python compare_serology_outputs.py --pdf-excel output/serology_report_merged_final.xlsx --word-excel ../09_Word_All_Tables_to_Excel/output/word_tables_merged.xlsx --out-csv output/diff.csv
```

