# Clinical Data Automation Toolkit

临床数据自动化处理工具集，提供 PDF 数据提取与规范化、Excel 图表生成、PPT 整合、文档翻译、网络诊断等自动化分析处理功能。

> English version: [`README_EN.md`](README_EN.md)（中英文 README 需同步更新）

## 环境要求

- Python 3.8+
- 依赖见 `requirements.txt`
- Windows 专用依赖（`pywin32`）仅在 Windows + Office 自动化模块中需要

## 安装

```bash
pip install -r requirements.txt
```

> `requirements.txt` 文件头已注明：非 Windows 或仅使用非 Office 自动化模块时，可移除或注释 `pywin32` 行后再安装；不需要 **14_PPTX_PDF_to_PPT** 时也可注释 Paddle 相关行以减轻安装体积。
>
> **提交前检查（可选）**：安装预提交钩子可在 commit 前运行 YAML 检查与本仓库敏感信息扫描（`scripts/check_secrets.py`，主要为 **29_SAE_Extractor** 等场景的 API Token 防误提交）：
>
> ```bash
> pip install pre-commit
> pre-commit install
> ```

## 整体架构

本仓库为 **按编号目录划分的独立工具集**（`01_` … `29_`）。整体遵循两条规则：

1. **结构分层清晰**：根级公共资源（`src/`、`config*.yaml`）与各独立模块目录分离。
2. **模块顺序固定**：以目录编号作为唯一顺序基准，按 **`01_` → `29_`** 递增组织与阅读。

| 层次 | 说明 |
|------|------|
| **入口** | 各 `NN_*/` 目录内的 `*.py` 脚本或子目录 `README.md` 中的命令行说明。 |
| **共享库** | `src/`：`pdf_reader`、`excel_writer`、`color_theme` 等，供图表/PDF 等模块引用。 |
| **配置** | 根目录 `config.yaml` / `config.example.yaml`，主要服务 **11_PDF_to_Excel_Rule_Extract** 的规则驱动提取。 |
| **数据约定** | 默认 **`input/` → 脚本 → `output/`**；部分模块支持命令行覆盖路径。 |
| **运行时** | **纯 Python + 文件库**（openpyxl、pandas、PyMuPDF 等）与 **Windows + Microsoft Office COM**（`pywin32`，用于 Word/Excel/PowerPoint 自动化）两类；后者仅在使用对应模块时需要。 |

### 模块分组与顺序（严格按目录编号）

- **Excel 相关（01-02）**：`01_Excel_Charts` → `02_Excel_Chart_Colors`
- **PowerPoint 相关（03-05）**：`03_PPT_Merge` → `04_PPT_Watermark_Removal` → `05_PPT_to_PDF`
- **Word 相关（06-09）**：`06_Word_to_PDF` → `07_Word_to_Excel_to_Figure` → `08_Word_Tables_to_Excel` → `09_Word_All_Tables_to_Excel`
- **PDF 相关（10-19）**：`10_PDF_Batch_to_Excel` → `11_PDF_to_Excel_Rule_Extract` → `12_PDF_to_PPT` → `13_PDF_XSS` → `14_PPTX_PDF_to_PPT` → `15_PDF_Title_Renamer` → `16_PDF_eCTD_Converter` → `17_PDF_Merge` → `18_PDF_Bookmark_Inherit_Zoom` → `19_PDF_Watermark_Removal`
- **其他工具（20-29）**：`20_PDF_Duplicate_Analyzer` → `21_File_Translator` → `22_Py_to_EXE` → `23_C_Drive_Cleanup` → `24_WiFi_Passwords` → `25_Folder_File_Count` → `26_Paper_Batch_Download` → `27_Proxy_Config_Export` → `28_DNS_Leak_Detector` → `29_SAE_Extractor`

## 项目结构

```
Clinical Data Automation/
├── # —— 输入以 Excel 为主 ——
├── 01_Excel_Charts/          # Excel 图表生成模块
│   ├── input/                        # 输入：源 Excel 文件
│   ├── output/                       # 输出：生成的图表文件
│   ├── build_charts_xlsxwriter.py    # 主程序（推荐，支持持续时间+临床配色）
│   ├── build_charts_openpyxl.py      # 备用引擎
│   └── apply_template_charts.py      # 基于模板应用配色（依赖 src/color_theme）
│
├── 02_Excel_Chart_Colors/    # Excel 图表配色（画图时配色调整）
│   ├── input/                        # 输入：ADR TFL 源表
│   ├── output/                       # 输出：应用临床配色的 TFL
│   └── apply_clinical_colors.py      # 临床发表/CSR 配色重绘（多期刊预设）
│
├── # —— 输入以 PowerPoint 为主 ——
├── 03_PPT_Merge/             # PPT 整合模块
│   ├── input/                        # 输入：待合并的 PPT 文件
│   ├── output/                       # 输出：合并后的 PPT 和报告
│   ├── merge_ppt.py                  # 物理合并（TF-IDF 去重）
│   ├── ppt_engine.py                 # 叙事编辑（SLIDE_BLUEPRINT 严格筛选）
│   ├── csr_ppt_integrator.py         # CSR 规范整合（叙事+去重+视觉统一）
│   └── test_and_validate.py          # 测试验证工具
│
├── 04_PPT_Watermark_Removal/ # PPT 边角水印去除模块
│   ├── input/                        # 输入：带水印的 PPTX
│   ├── output/                       # 输出：处理后的 PPTX
│   └── pptx_corner_logo_patch.py     # 主程序：仅处理图片型页面的右下角 logo
│
├── 05_PPT_to_PDF/            # PPT 批量转 PDF 模块
│   ├── input/                        # 输入：PPT/PPTX 文件
│   ├── output/                       # 输出：PDF 文件
│   └── ppt_to_pdf.py                 # 主程序：PowerPoint 导出 PDF
│
├── # —— 输入以 Word 为主 ——
├── 06_Word_to_PDF/           # Word 批量转 PDF 模块
│   ├── input/                        # 输入：Word 文档
│   ├── output/                       # 输出：PDF 文件
│   └── word_to_pdf.py                # 主程序：Word 导出 PDF
│
├── 07_Word_to_Excel_to_Figure/     # Word→Excel(表+图) 自动复刻（需 Windows+Word+pywin32）
│   ├── input/                        # 输入：Word/RTF 原始数据
│   ├── Template/                     # 骨架 Excel（图表/透视等结构）
│   ├── output/                       # 输出：复刻 Excel、table_mapping_plan_*.json
│   ├── word_to_excel_to_figure.py    # 主程序：发现图表区间、抽取并 Excel COM 写回
│   ├── table_mapping_logic.py        # 子表↔Word 表映射与 plan 生成（唯一实现，主程序引用）
│   ├── repair_output_by_patch.py     # 将旧 output 的图表区间 patch 回模板副本（修结构/弹窗）
│   ├── fixed_word_to_excel.py        # [Legacy] 硬编码示例脚本
│   └── README.md                     # 模块说明与命令行
│
├── 08_Word_Tables_to_Excel/      # Word 指定表格抽取为 Excel（精准表头 + 兼容样式）
│   ├── input/                        # 输入：Word/RTF
│   ├── output/                       # 输出：xlsx
│   ├── word_tables_to_excel.py       # 主程序：按表题/序号/表头关键词定位并导出
│   └── README.md                     # 模块说明与参数
│
├── 09_Word_All_Tables_to_Excel/     # Word 全部表格批量转 Excel（每文档多 sheet）
│   ├── input/                        # 输入：Word/RTF
│   ├── output/                       # 输出：xlsx
│   ├── word_all_tables_to_excel.py  # 主程序：批量导出（全部顶层表格）
│   └── README.md                     # 模块说明与参数
│
├── # —— 输入以 PDF 为主（含 PDF/PPTX 转可编辑 PPT）——
├── 10_PDF_Batch_to_Excel/          # PDF 批量转 Excel（血清报告、ADR 等专用工具）
│   ├── input/                        # 输入：PDF 源文件（若有）
│   ├── output/                       # 输出：生成的 Excel
│   ├── serology_report_pdf_to_excel.py  # 中检院血清报告 PDF 批量汇总
│   ├── fill_adr_from_pdf.py          # ADR 分级表专用提取
│   └── audit_and_fix_consistency.py  # 一致性检查与自动修正
│
├── 11_PDF_to_Excel_Rule_Extract/    # 通用规则驱动 PDF→Excel 提取（config.yaml）
│   ├── input/                        # 输入：PDF（可选；也可在 config.yaml 中写绝对路径）
│   ├── output/                       # 输出：Excel（可选）
│   ├── main.py                       # 入口：按规则检索并写入 Excel
│   └── README.md                     # 用法与与去水印联动说明
│
├── 12_PDF_to_PPT/            # PDF 转 PPT 模块
│   ├── input/                        # 输入：PDF 源文件
│   ├── output/                       # 输出：转换后的 PPT 文件
│   └── pdf_to_ppt.py                 # PDF 转换主程序
│
├── 13_PDF_XSS/               # PDF XSS/脚本清理模块
│   ├── input/                        # 输入：待清理 PDF
│   ├── output/                       # 输出：清理后的 PDF
│   └── pdf_xss_clean.py              # 主程序：移除 JS/恶意链接与嵌入文件
│
├── 14_PPTX_PDF_to_PPT/       # PPTX/PDF 转原生 PPT 模块
│   ├── input/                        # 输入：PDF 或 PPTX
│   ├── output/                       # 输出：可编辑 PPTX
│   └── convert_to_native_ppt.py      # 主程序：表格识别重建
│
├── 15_PDF_Title_Renamer/     # PDF 文件名极简清洗（剪切模式）
│   ├── input/                        # 输入：待重命名 PDF
│   ├── output/                       # 输出：重命名后 PDF
│   └── pdf_sanitizer.py              # 主程序：文件名手术 + 剪切
│
├── 16_PDF_eCTD_Converter/    # PDF eCTD 合规装甲（校验+清理+重写+审计报告）
│   ├── input/                        # 输入：待转换 PDF
│   ├── output/                       # 输出：*_ectd.pdf
│   ├── pdf_ectd_converter.py         # 主程序：6.26 字体映射/嵌入、书签修复、XSS/链接清理
│   └── README.md                     # 模块说明与参数
│
├── 17_PDF_Merge/             # PDF 合并模块（自然排序）
│   ├── input/                        # 输入：待合并 PDF（支持子文件夹）
│   ├── output/                       # 输出：合并后的 PDF
│   └── merge_pdf.py                  # 主程序：按自然排序合并
│
├── 18_PDF_Bookmark_Inherit_Zoom/ # PDF 书签「承前缩放」批处理（PyMuPDF）
│   ├── input/                        # 输入：待处理 PDF
│   ├── output/                       # 输出：重写书签后的 PDF
│   ├── pdf_bookmark_inherit_zoom.py  # 主程序：TOC 注入 XYZ / zoom=0，并发批处理
│   └── README.md                     # 模块说明与快速开始
│
├── 19_PDF_Watermark_Removal/   # PDF 干扰区定位与审计（排除框 + 审计 PDF + 清洗文本）
│   ├── input/                        # 输入：待分析 PDF
│   ├── output/                       # 输出：boxes.json / audit_masked.pdf / clean_text / 报告
│   ├── steps/                        # 管线子步骤（triage、vector、ocr、merge、audit、extract）
│   ├── main.py                       # 主程序
│   └── README.md                     # 模块说明、v2 JSON、与 11_PDF_to_Excel_Rule_Extract 联动与 mapping_audit
│
├── # —— 多格式 / 其他 ——
├── 20_PDF_Duplicate_Analyzer/ # PDF 跨文件夹重复分析（文件名 / 首页文本）
│   ├── output/                       # 输出：duplicate_report_*.txt（无 input/）
│   ├── pdf_duplicate_analyzer.py     # 主程序：--root 指定外部 References 路径
│   ├── jobs.example.json             # 多批次配置示例
│   └── README.md                     # 模块说明与命令行用法
│
├── 21_File_Translator/        # 多格式文档翻译模块（双向，免费优先）
│   ├── input/                        # 输入：待翻译 Excel/CSV/Word/PDF
│   ├── output/                       # 输出：翻译副本（*_en2zh.* / *_zh2en.*）
│   ├── file_translator.py            # 主程序：Excel/CSV/Word/PDF 双向翻译
│   └── README.md                     # 模块说明与命令行用法
│
├── 22_Py_to_EXE/             # Python 脚本转 EXE 模块
│   ├── input/                        # 输入：.py 脚本
│   ├── output/                       # 输出：.exe 文件
│   └── py_to_exe.py                  # 主程序：PyInstaller 打包
│
├── 23_C_Drive_Cleanup/       # C 盘垃圾/空文件清理模块
│   ├── input/                        # 输入：targets.txt（可选）
│   ├── output/                       # 输出：清理报告
│   └── c_drive_cleanup.py            # 主程序：扫描/清理
│
├── 24_WiFi_Passwords/        # WiFi 密码查看模块
│   ├── output/                       # 输出：wifi_passwords.csv
│   └── wifi_passwords.py             # 主程序：读取 WiFi 配置
│
├── 25_Folder_File_Count/     # 目录文件数量统计模块
│   ├── output/                       # 输出：统计结果
│   └── folder_file_count.py          # 主程序：统计目录文件数
│
├── 26_Paper_Batch_Download/  # 文献批量下载模块（OA）
│   ├── output/                       # 输出：下载的 PDF
│   └── paper_batch_download.py       # 主程序：批量下载
│
├── 27_Proxy_Config_Export/   # 代理配置导出模块（注册表+环境变量）
│   ├── output/                       # 输出：代理配置文本
│   └── proxy_config_export.py        # 主程序：导出代理配置
│
├── 28_DNS_Leak_Detector/      # DNS 泄漏诊断模块（TUN/SOCKS）
│   ├── output/                       # 输出：诊断 JSON 报告（可选）
│   ├── dns_leak_detector.py          # 主程序：出口与上游 DNS 一致性检测
│   └── README.md                     # 模块说明与命令行用法
│
├── 29_SAE_Extractor/          # SAE 结构化抽取（LLM + 多格式解析 → Excel）
│   ├── input/                        # 输入：临床文本/PDF/DOCX/Excel
│   ├── output/                       # 输出：SAE 列表 xlsx
│   ├── cli.py                        # 主入口：self-check / batch / pdf-batch
│   ├── sae_extractor.py              # API 抽取引擎与单条试跑
│   └── README.md
│
├── src/                      # 核心库
│   ├── __init__.py
│   ├── pdf_reader.py                 # PDF 内容提取
│   ├── excel_writer.py               # Excel 写入
│   └── color_theme.py                # 颜色主题（get_series_color 等）
│
├── config.yaml               # 配置文件（可复制 config.example.yaml）
├── config.example.yaml       # 配置示例
├── .pre-commit-config.yaml   # 可选：pre-commit（含敏感信息本地扫描）
├── scripts/                  # 仓库级辅助脚本
│   ├── check_secrets.py      # pre-commit 调用的密钥模式扫描
│   ├── set_env.ps1           # PowerShell：示例环境变量（29_SAE_Extractor）
│   └── start_tunnel.ps1      # PowerShell：SSH 本地端口转发（API 网关）
├── requirements.txt
├── LICENSE.md
└── README.md
```

## 功能模块

### 顶部导航索引（分组跳转）

- [Excel（01-02）](#modules-excel)
- [PowerPoint（03-05）](#modules-ppt)
- [Word（06-09）](#modules-word)
- [PDF（10-19）](#modules-pdf)
- [其他工具（20-29）](#modules-others)

### 01. Excel 图表生成（`01_Excel_Charts`）<span id="modules-excel"></span>
用途：生成 ADR 组合图（柱 + 线），并支持临床配色。

```bash
cd 01_Excel_Charts
python build_charts_xlsxwriter.py
```

常用参数：`--input`、`--output`、`--clinical-colors`。  
输出目录：`01_Excel_Charts/output/`。

---

### 02. Excel 图表配色（`02_Excel_Chart_Colors`）
用途：不改原始 TFL 的前提下，对现有图表应用期刊/临床配色。

```bash
cd 02_Excel_Chart_Colors
python apply_clinical_colors.py --palette Lancet --n-colors 3
```

支持批量：`--batch --input "input" --output "output"`。  
输出目录：`02_Excel_Chart_Colors/output/`。

---

### 03. PPT 整合（`03_PPT_Merge`）<span id="modules-ppt"></span>
用途：多 PPT 去重合并与叙事重组（CSR 结构）。

```bash
cd 03_PPT_Merge
python merge_ppt.py
python ppt_engine.py
```

`merge_ppt.py` 负责物理去重，`ppt_engine.py` 负责叙事编排。  
输出目录：`03_PPT_Merge/output/`（含 `merge_report.xlsx`、`narrative_report.xlsx`）。

---

### 04. PPT 边角水印去除（`04_PPT_Watermark_Removal`）
用途：针对大图截图页清理右下角重复 logo，可编辑页面自动跳过。

```bash
cd 04_PPT_Watermark_Removal
python pptx_corner_logo_patch.py
```

可指定输入输出：`python pptx_corner_logo_patch.py "input/your.pptx" -o "output/your_clean.pptx"`。  
输出目录：`04_PPT_Watermark_Removal/output/`。

---

### 05. PPT 批量转 PDF（`05_PPT_to_PDF`）
用途：批量将 PPT/PPTX 导出为 PDF（Windows + PowerPoint）。

```bash
cd 05_PPT_to_PDF
python ppt_to_pdf.py
```

常用参数：`--input`、`--output`、`--overwrite`。  
输出目录：`05_PPT_to_PDF/output/`。

---

### 06. Word 批量转 PDF（`06_Word_to_PDF`）<span id="modules-word"></span>
用途：批量将 Word 转 PDF（Windows + Word + `pywin32`）。

```bash
cd 06_Word_to_PDF
python word_to_pdf.py
```

常用参数：`--input`、`--output`、`--recursive/--no-recursive`、`--keep-structure/--no-keep-structure`、`--overwrite`。  
输出目录：`06_Word_to_PDF/output/`。

---

### 07. Word → Excel（表+图）自动复刻（`07_Word_to_Excel_to_Figure`）
用途：把 Word/RTF 表格数据写回骨架 Excel 图表区间，生成高保真复刻结果。

推荐工作流：
```bash
cd 07_Word_to_Excel_to_Figure
python word_to_excel_to_figure.py --input-dir "input" --plan-only
python word_to_excel_to_figure.py --input-dir "input" --table-map-json "output/table_mapping_plan_<模板名>.json"
```

关键文件：`table_mapping_logic.py`（映射单一数据源）、`repair_output_by_patch.py`（图表区间补丁修复）。  
输出目录：`07_Word_to_Excel_to_Figure/output/`。

---

### 08. Word 指定表格转 Excel（`08_Word_Tables_to_Excel`）
用途：按表题/序号/表头关键词精准定位并导出目标表格。

```bash
cd 08_Word_Tables_to_Excel
python word_tables_to_excel.py --help
```

常用定位参数：`--table-title`、`--table-index/--table-indices`、`--header-keywords`、`--merge-tables-from/--merge-tables-to`。  
辅助参数：`--list-word-tables`、`--dry-run`。

---

### 09. Word 所有表格批量转 Excel（`09_Word_All_Tables_to_Excel`）
用途：批量导出每个 Word 的全部顶层表格（每表一个 sheet）。

```bash
cd 09_Word_All_Tables_to_Excel
python word_all_tables_to_excel.py
```

支持：`--header-rows`、`--dry-run`、`--skip-existing`。  
输出目录：`09_Word_All_Tables_to_Excel/output/`。

---

### 10. PDF 批量转 Excel（`10_PDF_Batch_to_Excel`）<span id="modules-pdf"></span>
用途：面向 ADR 分级与血清报告的专用提取、回填与一致性校验。

```bash
cd 10_PDF_Batch_to_Excel
python fill_adr_from_pdf.py
python audit_and_fix_consistency.py
python serology_report_pdf_to_excel.py --input "input" --output "output/serology_report_merged.xlsx" --ocr --ocr-dpi 110
```

血清报告场景建议开启 `--ocr`（需本机 Tesseract + `chi_sim+eng` 语言包）。

---

### 11. PDF → Excel 规则提取（`11_PDF_to_Excel_Rule_Extract`）
用途：按 `config.yaml` 规则从 PDF 检索内容并写入 Excel。

```bash
cd 11_PDF_to_Excel_Rule_Extract
python main.py
```

可与 `19_PDF_Watermark_Removal` 联动：
```bash
python main.py --config config.yaml --exclusion-json "../19_PDF_Watermark_Removal/output/你的文件_boxes.json"
```

输出目录：`11_PDF_to_Excel_Rule_Extract/output/`。

---

### 12. PDF 转 PPT（`12_PDF_to_PPT`）
用途：每页 PDF 转换为一张 PPT 幻灯片。

```bash
cd 12_PDF_to_PPT
python pdf_to_ppt.py
```

输出目录：`12_PDF_to_PPT/output/`。

---

### 13. PDF XSS 清理（`13_PDF_XSS`）
用途：清理 PDF 中脚本/恶意协议链接/注释/嵌入文件。

```bash
cd 13_PDF_XSS
python pdf_xss_clean.py
```

常用参数：`--input`、`--output`、`--overwrite`、`--no-recursive`、`--no-keep-structure`。

---

### 14. PPTX/PDF 转原生 PPT（`14_PPTX_PDF_to_PPT`）
用途：把 PDF 或图片型 PPTX 中的表格重建为可编辑原生 PPT。

```bash
cd 14_PPTX_PDF_to_PPT
python convert_to_native_ppt.py
```

常用参数：`--input`、`--output`、`--dpi`、`--lang`。  
依赖：`paddleocr`、`paddlepaddle`、`pymupdf`、`python-pptx`。

---

### 15. PDF 标题驱动重命名（`15_PDF_Title_Renamer`）
用途：多策略提取标题并规范命名，文件从 `input/` 剪切移动到 `output/`。

```bash
cd 15_PDF_Title_Renamer
python pdf_sanitizer.py
```

常用参数：`--input`、`--output`、`--no-recursive`、`--no-keep-structure`、`--overwrite`。  
注意：执行后 `input/` 中 PDF 会被迁移走。

---

### 16. PDF eCTD 转换（`16_PDF_eCTD_Converter`）
用途：按 eCTD 附件 6 常见条款批量校验、清洗并重写 PDF；含 **6.26 字体名映射与子集嵌入**、**书签修复**；输出 Excel 审计报告（含「结构警告」工作表）。

```bash
cd 16_PDF_eCTD_Converter
python pdf_ectd_converter.py --input "./input" --output "./output" --report "./ectd_report.xlsx" --overwrite
```

常用参数：`--validate-only`、`--overwrite`、`--keep-name`、`--no-recursive`、`--no-keep-structure`、`--add-auto-bookmarks`（默认 `outline`）、`--no-add-auto-bookmarks`。  
依赖：`pymupdf`、`pandas`、`openpyxl`、`fonttools`（字体嵌入）。详见 `16_PDF_eCTD_Converter/README.md`。

---

### 17. PDF 合并（`17_PDF_Merge`）
用途：按自然排序合并多个 PDF（支持子目录）。

```bash
cd 17_PDF_Merge
python merge_pdf.py
```

常用参数：`--input`、`--output`、`--output-name`、`--overwrite`。  
输出文件：`17_PDF_Merge/output/merged.pdf`。

---

### 18. PDF 书签承前缩放（`18_PDF_Bookmark_Inherit_Zoom`）
用途：重写书签目标为 `XYZ + zoom=0`，点击目录时保持当前缩放比例。

```bash
cd 18_PDF_Bookmark_Inherit_Zoom
python pdf_bookmark_inherit_zoom.py
```

依赖：`pymupdf`。  
输出目录：`18_PDF_Bookmark_Inherit_Zoom/output/`。

---

### 19. PDF 干扰区定位与审计（`19_PDF_Watermark_Removal`）
用途：定位页眉/水印等干扰区，输出排除框 JSON、审计叠加 PDF 与清洗文本。

```bash
cd 19_PDF_Watermark_Removal
python main.py --input "input" --output "output"
```

与 `11_PDF_to_Excel_Rule_Extract` 的 `--exclusion-json` 可直接联动。

---

### 20. PDF 跨文件夹重复分析（`20_PDF_Duplicate_Analyzer`）<span id="modules-others"></span>
用途：在同一根目录的多个子文件夹之间，按**文件名**或**首页文本**检测重复 PDF（无 `input/`，源文件在外部路径）。

```bash
cd 20_PDF_Duplicate_Analyzer
python pdf_duplicate_analyzer.py --root "D:\References" --folders "FolderA,FolderB" --label "批次1"
python pdf_duplicate_analyzer.py --config jobs.json
```

常用参数：`--root`、`--folders`（省略则扫描 root 下全部一级子目录）、`--label`、`--config`、`--output`。  
输出目录：`20_PDF_Duplicate_Analyzer/output/`（`duplicate_report_*.txt`）。

---

### 21. 文档双向翻译（`21_File_Translator`）
用途：翻译 Excel/CSV/Word/PDF，默认免费引擎优先并支持多级兜底。

```bash
cd 21_File_Translator
python file_translator.py --self-test
python file_translator.py
```

常用参数：`--direction en2zh|zh2en`、`--provider`、`--engine`、`--columns`、`--pdf-mode`、`--cache-file`、`--max-workers`。  
输出目录：`21_File_Translator/output/`。

---

### 22. Python 脚本转 EXE（`22_Py_to_EXE`）
用途：基于 PyInstaller 将 `.py` 打包为 Windows 可执行文件。

```bash
cd 22_Py_to_EXE
python py_to_exe.py
```

常用参数：`--input`、`--output`、`--name`、`--dir`、`--noconsole`、`--icon`、`--clean-artifacts`。

---

### 23. C 盘垃圾/空文件清理（`23_C_Drive_Cleanup`）
用途：清理临时/缓存垃圾文件；默认仅扫描，避免误删。

```bash
cd 23_C_Drive_Cleanup
python c_drive_cleanup.py
python c_drive_cleanup.py --delete --days 7
```

常用参数：`--targets`、`--remove-empty-dirs`。  
输出文件：`23_C_Drive_Cleanup/output/cleanup_report.csv`。

---

### 24. WiFi 密码查看（`24_WiFi_Passwords`）
用途：导出 Windows 本机已保存 WiFi 账号密码（需权限）。

```bash
cd 24_WiFi_Passwords
python wifi_passwords.py
```

常用参数：`--output`、`--encoding`、`--quiet`。  
输出文件：`24_WiFi_Passwords/output/wifi_passwords.csv`。

---

### 25. 目录文件数量统计（`25_Folder_File_Count`）
用途：递归统计目录文件数量并输出 TXT + Excel。

```bash
cd 25_Folder_File_Count
python folder_file_count.py --path "D:\data"
```

常用参数：`--output`。  
输出目录：`25_Folder_File_Count/output/`。

---

### 26. 文献批量下载（`26_Paper_Batch_Download`）
用途：按 DOI/PMID/标题/URL 批量下载 Open Access PDF，并在默认安全模式下进行限速与退避重试以降低 IP 限制风险。

```bash
cd 26_Paper_Batch_Download
python paper_batch_download.py --queries "10.1038/s41586-020-2649-2" "32788730"
```

文件输入模式：`python paper_batch_download.py --file "D:\papers.txt" --mailto "your_email@example.com"`。  
安全参数（默认开启）：`--safe-mode`、`--min-interval`、`--max-retries`、`--backoff-base`、`--mirror-cooldown`。  
如需临时关闭限速防护：`python paper_batch_download.py --queries "10.xxx/xxx" --no-safe-mode`。  
输出目录：`26_Paper_Batch_Download/output/`。

---

### 27. 代理配置导出（`27_Proxy_Config_Export`）
用途：导出当前系统代理配置（注册表 + 环境变量）到文本文件。

```bash
cd 27_Proxy_Config_Export
python proxy_config_export.py
```

输出文件：`27_Proxy_Config_Export/output/proxy_config_YYYYMMDD_HHMMSS.txt`。  
仅支持 Windows。

---

### 28. DNS 泄漏诊断（`28_DNS_Leak_Detector`）
用途：检测公网出口与上游 DNS 是否偏离，用于排查 DNS 泄漏与分流错误。

```bash
cd 28_DNS_Leak_Detector
python dns_leak_detector.py --mode tun
python dns_leak_detector.py --mode socks --socks-port 10808
python dns_leak_detector.py --save-json
```

报告文件（可选）：`28_DNS_Leak_Detector/output/dns_diagnostic_<mode>_<timestamp>.json`。

---

### 29. SAE 结构化抽取（`29_SAE_Extractor`）
用途：从 PDF/TXT/DOCX/Excel 临床材料中提取严重不良事件（SAE）字段，调用 OpenAI 兼容 Chat Completions 接口，汇总导出 Excel。

```bash
cd 29_SAE_Extractor
python cli.py self-check
python cli.py batch
python cli.py pdf-batch
```

需配置环境变量 `SAE_API_TOKEN`；可选 `SAE_API_BASE_URL`、`SAE_MODEL_ID`、`TESSERACT_CMD`、`POPPLER_PATH`。默认读写目录：`input/`、`output/`。详见模块内 `README.md`。

辅助脚本（仓库根目录 `scripts/`，Windows PowerShell）：

- `set_env.ps1`：在当前会话设置 `SAE_API_BASE_URL`、`SAE_OUTPUT_DIR`（指向 `29_SAE_Extractor/output`）等；**仍需自行设置 `SAE_API_TOKEN`**。用法：在仓库根执行 `.\scripts\set_env.ps1`（注意点开源执行策略）。
- `start_tunnel.ps1`：建立 SSH 本地端口转发；须设置 **`SAE_TUNNEL_SSH_HOST`**（及可选 `SAE_TUNNEL_SSH_USER`、`SAE_TUNNEL_LOCAL_PORT`、`SAE_TUNNEL_REMOTE_BIND`）。用法：`powershell -ExecutionPolicy Bypass -File .\scripts\start_tunnel.ps1`

手动全量扫描敏感信息：`pre-commit run detect-sensitive-secrets --all-files`（需已 `pre-commit install` 或直接安装 `pre-commit`）。

## 配置说明

复制 `config.example.yaml` 为 `config.yaml`，按需修改：

```yaml
pdf_path: "11_PDF_to_Excel_Rule_Extract/input/不同剂量组ADR分析 (TFL).pdf"
excel_path: "11_PDF_to_Excel_Rule_Extract/output/不同剂量组ADR分析 (TFL).xlsx"
rules:
  - name: "提取规则1"
    search:
      keyword: "关键词"
      page: 1
    excel:
      sheet: "Sheet1"
      cell: "B3"
```

## 扩展开发

- **PDF 解析扩展**：修改 [`src/pdf_reader.py`](src/pdf_reader.py)（含排除框过滤、旋转坐标映射、`mapping_audit`）
- **PDF 干扰区管线**：[`19_PDF_Watermark_Removal/main.py`](19_PDF_Watermark_Removal/main.py) 与 [`19_PDF_Watermark_Removal/steps/`](19_PDF_Watermark_Removal/steps/)
- **Excel 写入扩展**：修改 [`src/excel_writer.py`](src/excel_writer.py)
- **图表样式与临床配色**：修改 [`01_Excel_Charts/build_charts_xlsxwriter.py`](01_Excel_Charts/build_charts_xlsxwriter.py) 内 `COLOR_MAP` 或 [`src/color_theme.py`](src/color_theme.py)
- **独立配色模块**： [`02_Excel_Chart_Colors/apply_clinical_colors.py`](02_Excel_Chart_Colors/apply_clinical_colors.py) 支持多期刊预设（NPG、Lancet、NEJM 等）
- **Word→Excel 图表复刻**：主逻辑在 [`07_Word_to_Excel_to_Figure/word_to_excel_to_figure.py`](07_Word_to_Excel_to_Figure/word_to_excel_to_figure.py)；子表映射与 plan 在 [`07_Word_to_Excel_to_Figure/table_mapping_logic.py`](07_Word_to_Excel_to_Figure/table_mapping_logic.py)

## 注意事项

1. **文件组织**：所有输入文件放在各模块的 `input/` 文件夹，输出文件自动保存到 `output/` 文件夹
2. **自动备份**：Excel 图表生成会自动备份旧文件（.bak.xlsx）
3. **引擎选择**：Excel 图表生成推荐使用 XlsxWriter 引擎，避免 XML 结构问题
4. **路径配置**：PDF 提取依赖文件结构，需根据实际 PDF 调整配置
5. **模块独立**：各模块独立运行，互不干扰，便于维护和扩展
6. **PPT 合并**：推荐先运行 `merge_ppt.py` 做基础合并，再运行 `ppt_engine.py` 做叙事重组
7. **依赖安装**：见根目录 `requirements.txt` 分组注释；`scikit-learn`（03_PPT_Merge）、`pymupdf`（多 PDF 模块含 17/18 等）、`pytesseract`（14/18，需本机 Tesseract）、`requests`（24/26）、`pywin32`（仅 Windows，05/06/07/08/09 等 Office 自动化）、`paddleocr`（13，可选注释以减小体积）按所用模块生效

### 血清报告对账（PDF vs Word）
1. 生成 Word 汇总：`09_Word_All_Tables_to_Excel/output/word_tables_merged.xlsx`
2. 生成 PDF 汇总并回填缺项：在 `10_PDF_Batch_to_Excel/serology_report_pdf_to_excel.py` 使用 `--reference-excel`（可选但强烈建议）
3. 对比差异并导出明细：运行 `compare_serology_outputs.py --pdf-excel <PDF.xlsx> --word-excel <WORD.xlsx> --out-csv <diff.csv>`

## 项目优势

- ✅ **自动化**：一键完成数据处理、图表生成、文件合并
- ✅ **智能化**：TF-IDF 算法确保内容质量，自动去重和筛选
- ✅ **高保真**：完整保留原始视觉效果和格式
- ✅ **可追溯**：详细的处理记录便于审计和验证
- ✅ **专业级**：符合临床研究报告标准

## 维护说明

- 各模块独立运行，按需使用；输入放 `input/`，输出在 `output/`。
- 历史优化与代码规范已融入代码与本文档，重大变更见版本记录或 Git 历史。

### 代码巡检结论（2026-05）

- 全仓 Python 脚本已完成一次批量语法巡检（`compileall`），当前无语法级报错。
- 多数模块已保持单目录单入口模式，结构清晰；优先优化建议集中在“日志一致性、异常分级、参数校验”三个横切点。
- 建议优先级：
  - P1：统一 CLI 参数风格（`--input/--output/--overwrite`）与错误码返回，降低批处理串联成本。
  - P1：减少宽泛 `except Exception` 捕获，补充失败上下文，便于定位真实数据问题。
  - P2：为核心模块补充最小回归样例（每模块 1~2 个 smoke case），降低后续重构风险。
  - P3：逐步抽取公共 I/O 与日志助手到 `src/`，减少重复实现。

## 许可证

本项目采用 [MIT License](LICENSE.md) 开源协议。`29_SAE_Extractor` 模块自独立项目 **SAE-Extractor** 迁入，该部分版权信息见 `LICENSE.md` 中的附加声明。
