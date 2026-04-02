# Clinical Data Automation Toolkit

临床数据自动化处理工具集，提供 PDF 数据提取与规范化、Excel 图表生成、PPT 整合、文档翻译、网络诊断等自动化分析处理功能。

## 环境要求

- Python 3.8+
- 依赖见 `requirements.txt`
- Windows 专用依赖（`pywin32`）仅在 Windows + Office 自动化模块中需要

## 安装

```bash
pip install -r requirements.txt
```

> `requirements.txt` 文件头已注明：非 Windows 或仅使用非 Office 自动化模块时，可移除或注释 `pywin32` 行后再安装；不需要 **11** 时也可注释 Paddle 相关行以减轻安装体积。

## 整体架构

本仓库为 **按编号目录划分的独立工具集**（`01_` … `24_`），共享少量根级资源，各模块可单独使用、互不强制耦合。

| 层次 | 说明 |
|------|------|
| **入口** | 各 `NN_*/` 目录内的 `*.py` 脚本或子目录 `README.md` 中的命令行说明。 |
| **共享库** | `src/`：`pdf_reader`、`excel_writer`、`color_theme` 等，供图表/PDF 等模块引用。 |
| **配置** | 根目录 `config.yaml` / `config.example.yaml`，主要服务 **03_PDF_to_Excel** 等基于规则的提取。 |
| **数据约定** | 默认 **`input/` → 脚本 → `output/`**；部分模块支持命令行覆盖路径。 |
| **运行时** | **纯 Python + 文件库**（openpyxl、pandas、PyMuPDF 等）与 **Windows + Microsoft Office COM**（`pywin32`，用于 Word/Excel/PowerPoint 自动化）两类；后者仅在使用对应模块时需要。 |

**模块 20（Word→Excel 图表数据）数据流简述：** `input/` 中的 Word/RTF → `word_to_excel_to_figure.py` 读取骨架 `Template/*.xlsx` 中的图表 `series(cat/val)` 区间 → 匹配 Word 表格并写回数值；子表与 Word 表的对应关系由 **`table_mapping_logic.py`** 生成/解析 plan JSON（主程序 `import` 该文件，避免重复实现）。若曾用 openpyxl 直接保存破坏了透视/OLAP 结构，可用 **`repair_output_by_patch.py`** 将「数据已正确」的旧文件中的图表区间 patch 回模板副本（Excel COM 写入）。

## 项目结构

```
Clinical Data Automation/
├── 01_Excel_Charts/          # Excel 图表生成模块
│   ├── input/                        # 输入：源 Excel 文件
│   ├── output/                       # 输出：生成的图表文件
│   ├── build_charts_xlsxwriter.py    # 主程序（推荐，支持持续时间+临床配色）
│   ├── build_charts_openpyxl.py      # 备用引擎
│   └── apply_template_charts.py      # 基于模板应用配色（依赖 src/color_theme）
│
├── 02_PPT_Merge/             # PPT 整合模块
│   ├── input/                        # 输入：待合并的 PPT 文件
│   ├── output/                       # 输出：合并后的 PPT 和报告
│   ├── merge_ppt.py                  # 物理合并（TF-IDF 去重）
│   ├── ppt_engine.py                 # 叙事编辑（SLIDE_BLUEPRINT 严格筛选）
│   ├── csr_ppt_integrator.py         # CSR 规范整合（叙事+去重+视觉统一）
│   └── test_and_validate.py          # 测试验证工具
│
├── 03_PDF_to_Excel/          # PDF 数据提取模块
│   ├── input/                        # 输入：PDF 源文件
│   ├── output/                       # 输出：提取后的 Excel 文件
│   ├── main.py                       # 通用 PDF 提取
│   ├── fill_adr_from_pdf.py          # ADR 分级表专用提取
│   └── audit_and_fix_consistency.py  # 数据一致性检查
│
├── 04_PDF_to_PPT/            # PDF 转 PPT 模块
│   ├── input/                        # 输入：PDF 源文件
│   ├── output/                       # 输出：转换后的 PPT 文件
│   └── pdf_to_ppt.py                 # PDF 转换主程序
│
├── 05_Excel_Chart_Colors/    # Excel 图表配色（画图时配色调整）
│   ├── input/                        # 输入：ADR TFL 源表
│   ├── output/                       # 输出：应用临床配色的 TFL
│   └── apply_clinical_colors.py      # 临床发表/CSR 配色重绘（多期刊预设）
│
├── 06_PPT_Watermark_Removal/ # PPT 边角水印去除模块
│   ├── input/                        # 输入：带水印的 PPTX
│   ├── output/                       # 输出：处理后的 PPTX
│   └── pptx_corner_logo_patch.py     # 主程序：仅处理图片型页面的右下角 logo
│
├── 07_PDF_XSS/               # PDF XSS/脚本清理模块
│   ├── input/                        # 输入：待清理 PDF
│   ├── output/                       # 输出：清理后的 PDF
│   └── pdf_xss_clean.py              # 主程序：移除 JS/恶意链接与嵌入文件
│
├── 08_PPT_to_PDF/            # PPT 批量转 PDF 模块
│   ├── input/                        # 输入：PPT/PPTX 文件
│   ├── output/                       # 输出：PDF 文件
│   └── ppt_to_pdf.py                 # 主程序：PowerPoint 导出 PDF
│
├── 09_Word_to_PDF/           # Word 批量转 PDF 模块
│   ├── input/                        # 输入：Word 文档
│   ├── output/                       # 输出：PDF 文件
│   └── word_to_pdf.py                # 主程序：Word 导出 PDF
│
├── 10_C_Drive_Cleanup/       # C 盘垃圾/空文件清理模块
│   ├── input/                        # 输入：targets.txt（可选）
│   ├── output/                       # 输出：清理报告
│   └── c_drive_cleanup.py            # 主程序：扫描/清理
│
├── 11_PPTX_PDF_to_PPT/       # PPTX/PDF 转原生 PPT 模块
│   ├── input/                        # 输入：PDF 或 PPTX
│   ├── output/                       # 输出：可编辑 PPTX
│   └── convert_to_native_ppt.py      # 主程序：表格识别重建
│
├── 12_Py_to_EXE/             # Python 脚本转 EXE 模块
│   ├── input/                        # 输入：.py 脚本
│   ├── output/                       # 输出：.exe 文件
│   └── py_to_exe.py                  # 主程序：PyInstaller 打包
│
├── 13_WiFi_Passwords/        # WiFi 密码查看模块
│   ├── output/                       # 输出：wifi_passwords.csv
│   └── wifi_passwords.py             # 主程序：读取 WiFi 配置
│
├── 14_Folder_File_Count/     # 目录文件数量统计模块
│   ├── output/                       # 输出：统计结果
│   └── folder_file_count.py          # 主程序：统计目录文件数
│
├── 15_Paper_Batch_Download/  # 文献批量下载模块（OA）
│   ├── output/                       # 输出：下载的 PDF
│   └── paper_batch_download.py       # 主程序：批量下载
│
├── 16_PDF_Sanitizer/         # PDF 文件名极简清洗（剪切模式）
│   ├── input/                        # 输入：待重命名 PDF
│   ├── output/                       # 输出：重命名后 PDF
│   └── pdf_sanitizer.py              # 主程序：文件名手术 + 剪切
│
├── 17_PDF_eCTD_Converter/    # PDF eCTD 转换模块（校验+清理+重写保存）
│   ├── input/                        # 输入：待转换 PDF
│   ├── output/                       # 输出：转换后的 eCTD PDF
│   └── pdf_ectd_converter.py         # 主程序：eCTD 校验与转换
│
├── 18_Proxy_Config_Export/   # 代理配置导出模块（注册表+环境变量）
│   ├── output/                       # 输出：代理配置文本
│   └── proxy_config_export.py        # 主程序：导出代理配置
│
├── 19_PDF_Merge/             # PDF 合并模块（自然排序）
│   ├── input/                        # 输入：待合并 PDF（支持子文件夹）
│   ├── output/                       # 输出：合并后的 PDF
│   └── merge_pdf.py                  # 主程序：按自然排序合并
│
├── 20_Word_to_Excel_to_Figure/     # Word→Excel(表+图) 自动复刻（需 Windows+Word+pywin32）
│   ├── input/                        # 输入：Word/RTF 原始数据
│   ├── Template/                     # 骨架 Excel（图表/透视等结构）
│   ├── output/                       # 输出：复刻 Excel、table_mapping_plan_*.json
│   ├── word_to_excel_to_figure.py    # 主程序：发现图表区间、抽取并 Excel COM 写回
│   ├── table_mapping_logic.py        # 子表↔Word 表映射与 plan 生成（唯一实现，主程序引用）
│   ├── repair_output_by_patch.py     # 将旧 output 的图表区间 patch 回模板副本（修结构/弹窗）
│   ├── fixed_word_to_excel.py        # [Legacy] 硬编码示例脚本
│   └── README.md                     # 模块说明与命令行
│
├── 21_File_Translator/        # 多格式文档翻译模块（双向，免费优先）
│   ├── input/                        # 输入：待翻译 Excel/CSV/Word/PDF
│   ├── output/                       # 输出：翻译副本（*_en2zh.* / *_zh2en.*）
│   ├── file_translator.py            # 主程序：Excel/CSV/Word/PDF 双向翻译
│   └── README.md                     # 模块说明与命令行用法
│
├── 22_DNS_Leak_Detector/      # DNS 泄漏诊断模块（TUN/SOCKS）
│   ├── output/                       # 输出：诊断 JSON 报告（可选）
│   ├── dns_leak_detector.py          # 主程序：出口与上游 DNS 一致性检测
│   └── README.md                     # 模块说明与命令行用法
│
├── 23_PDF_Bookmark_Inherit_Zoom/ # PDF 书签「承前缩放」批处理（PyMuPDF）
│   ├── input/                        # 输入：待处理 PDF
│   ├── output/                       # 输出：重写书签后的 PDF
│   ├── pdf_bookmark_inherit_zoom.py  # 主程序：TOC 注入 XYZ / zoom=0，并发批处理
│   └── README.md                     # 模块说明与快速开始
│
├── 24_Word_Tables_to_Excel/      # Word 指定表格抽取为 Excel（精准表头 + 兼容样式）
│   ├── input/                        # 输入：Word/RTF
│   ├── output/                       # 输出：xlsx
│   ├── word_tables_to_excel.py       # 主程序：按表题/序号/表头关键词定位并导出
│   └── README.md                     # 模块说明与参数
│
├── src/                      # 核心库
│   ├── __init__.py
│   ├── pdf_reader.py                 # PDF 内容提取
│   ├── excel_writer.py               # Excel 写入
│   └── color_theme.py                # 颜色主题（get_series_color 等）
│
├── config.yaml               # 配置文件（可复制 config.example.yaml）
├── config.example.yaml       # 配置示例
├── requirements.txt
└── README.md
```

## 功能模块

### 1. Excel 图表生成（01_Excel_Charts）

生成符合临床规范的 ADR 组合图（柱状图 + 折线图）。

**📁 使用步骤：**

1. 将源 Excel 文件放入 `01_Excel_Charts/input/` 文件夹
   - 示例：`不同剂量组ADR分析 (TFL).xlsx`

2. 运行脚本生成图表（标准配色）：
   ```bash
   cd 01_Excel_Charts
   python build_charts_xlsxwriter.py
   ```

3. 查看输出文件：`01_Excel_Charts/output/不同剂量组ADR分析 (TFL).xlsx`

**特性：**
- ✅ 无网格线纯白背景
- ✅ 百分比格式（0.00%）
- ✅ 黑色实线坐标轴（0.75pt）
- ✅ 通过微软 Excel 校验
- ✅ 自动备份旧文件
- ✅ 支持 ADR 发生情况 + ADR 持续时间 两种表结构

**高级用法：**

- **指定输入/输出路径：**
  ```bash
  cd 01_Excel_Charts
  python build_charts_xlsxwriter.py ^
    --input  "input/不同剂量组ADR分析 (TFL).xlsx" ^
    --output "output/不同剂量组ADR分析 (TFL).xlsx"
  ```

- **启用临床发表/CSR 规范配色：**

  - 规则：
    - 低剂量试验组：#5B9BD5（柔和蓝）
    - 高剂量试验组：#254061（深海军蓝）
    - 低剂量佐剂组：#ED7D31（橙色）
    - 高剂量佐剂组：#C00000（砖红）
    - 安慰剂组：#7F7F7F（中灰）
  - 同一组的 **柱（例数）** 和 **线（发生率）** 使用 **同一颜色**；线宽 1.5pt，带圆形标记。

  ```bash
  cd 01_Excel_Charts
  python build_charts_xlsxwriter.py ^
    --input  "input/不同剂量组ADR分析 (TFL).xlsx" ^
    --output "output/不同剂量组ADR分析_clinical_colors.xlsx" ^
    --clinical-colors
  ```

### 1.1 Excel 图表配色独立模块（05_Excel_Chart_Colors）

当只想在**不改动原 TFL 文件**的前提下，对 Excel 图表做画图时的配色调整、生成一份应用临床/期刊配色的副本时，可使用该模块：

**📁 使用步骤：**

1. 将 `不同剂量组ADR分析 (TFL).xlsx` 放入 `05_Excel_Chart_Colors/input/`
2. 运行脚本：
   ```bash
   cd 05_Excel_Chart_Colors
   python apply_clinical_colors.py
   ```
3. 查看输出：`05_Excel_Chart_Colors/output/不同剂量组ADR分析_clinical_colors_{NPG/Lancet/NEJM}.xlsx`

**可选参数：**

- **指定配色方案（非交互/批量推荐）：**
  ```bash
  python apply_clinical_colors.py --palette NPG
  ```

- **同一组柱/线同色，且仅使用 3 种颜色（默认行为，可调整）：**
  - 脚本会从每个图表系列标题中提取“组名”，让同一组的 **柱状图** 和 **折线图** 始终保持同色。
  - 默认每个图表最多使用 **3** 种颜色（超过 3 个组则循环复用颜色）。
  ```bash
  python apply_clinical_colors.py --palette Lancet --n-colors 3
  ```

- **批量处理输入目录：**
  ```bash
  python apply_clinical_colors.py --batch --input "input" --output "output" --palette Lancet --n-colors 3
  ```

---

### 2. PPT 整合（02_PPT_Merge）

智能合并多个 PPT 文件，支持物理合并与叙事重组两种模式。

**📁 使用步骤：**

1. 将待合并的 PPT 文件放入 `02_PPT_Merge/input/` 文件夹

2. **模式 A：基础合并**（推荐先运行）
   ```bash
   cd 02_PPT_Merge
   python merge_ppt.py
   ```
   - 输出：`output/merged_presentation.pptx`、`output/merge_report.xlsx`
   - 功能：TF-IDF 相似度去重，保留高复杂度版式

3. **模式 B：叙事编辑**（按 CSR 标准重组）
   ```bash
   python ppt_engine.py
   ```
   - 输出：`output/merged_presentation.pptx`（覆盖）、`output/narrative_report.xlsx`
   - 功能：按 SLIDE_BLUEPRINT 严格筛选，每 Slot 仅保留一张优胜页，自动插入章节过渡页

**功能特性：**
- ✅ **智能去重**：TF-IDF + 余弦相似度算法
- ✅ **高保真复制**：完整保留图表、图片、组合图形（含 part 与 rId 重映射）
- ✅ **叙事重组**：按 5 章 CSR 结构 + 10 个 Slot 严格筛选
- ✅ **Tie-Breaker**：优先含表格、N=64/N=32、shape_count 更大等
- ✅ **详细报告**：Excel 格式的处理记录与叙事结构报告

**技术细节：**

**去重策略：**
- 相似度 < 0.6：保留为独立内容
- 0.6 ≤ 相似度 < 0.9：标记部分重叠，全部保留
- 相似度 ≥ 0.9：只保留版式更复杂的幻灯片

**SLIDE_BLUEPRINT 结构：**
- Chapter 1: 研究背景 (Context) - Slot 1-2
- Chapter 2: 核心结论 (Top-line Safety) - Slot 3
- Chapter 3: 安全性特征详述 (Detailed Profile) - Slot 4-7
- Chapter 4: 竞品对标 (Benchmark vs Shingrix) - Slot 8-9
- Chapter 5: 总结 (Conclusion) - Slot 10

**输出文件说明：**
- `merged_presentation.pptx`：合并后的 PPT 文件，包含去重后的精选幻灯片和章节过渡页
- `merge_report.xlsx`：详细处理记录，包含每张幻灯片的来源文件、相似度分析、处理决策、形状数量统计
- `narrative_report.xlsx`：叙事结构报告（ppt_engine 生成），包含每张幻灯片的章节归属、剧本匹配结果、是否入选最终 PPT

**使用建议：**
1. 首次使用：先运行 `merge_ppt.py` 获得基础合并结果
2. 叙事重组：运行 `ppt_engine.py` 按 CSR 标准重组；或使用 `csr_ppt_integrator.py` 一键完成叙事+去重+视觉统一
3. 质量检查：查看 Excel 报告了解去重详情
4. 自定义配置：修改 `ppt_engine.py` 中的 `SLIDE_BLUEPRINT` 调整剧本

---

### 3. PDF 数据提取（03_PDF_to_Excel）

从 PDF 文件中提取数据并写入 Excel。

#### 3.1 通用提取（基于配置）

**📁 使用步骤：**

1. 将 PDF 文件放入 `03_PDF_to_Excel/input/` 文件夹

2. 在项目根目录配置 `config.yaml`：
   ```yaml
   pdf_path: "03_PDF_to_Excel/input/不同剂量组ADR分析 (TFL).pdf"
   excel_path: "03_PDF_to_Excel/output/不同剂量组ADR分析 (TFL).xlsx"
   rules:
     - name: "提取发热数据"
       search:
         keyword: "发热"
         page: 1
       excel:
         sheet: "Sheet1"
         cell: "B3"
   ```

3. 运行脚本：
   ```bash
   cd 03_PDF_to_Excel
   python main.py
   ```

4. 查看输出：`03_PDF_to_Excel/output/` 文件夹

#### 3.2 ADR 分级表专用提取

**📁 使用步骤：**

1. 将以下文件放入 `03_PDF_to_Excel/input/` 文件夹：
   - `不同剂量组ADR分析 (TFL).pdf`（总表）
   - `不同剂量组ADR分析-分级 (TFL).pdf`（分级表）
   - `不同剂量组ADR分析 (TFL).xlsx`（目标 Excel）

2. 运行脚本：
   ```bash
   cd 03_PDF_to_Excel
   python fill_adr_from_pdf.py
   ```

3. 查看输出：`03_PDF_to_Excel/output/` 文件夹

**功能：** 从分级 PDF 提取 1级/2级/3级/Total 数据并自动填充到 Excel

#### 3.3 数据一致性检查

```bash
cd 03_PDF_to_Excel
python audit_and_fix_consistency.py
```

**功能：** 检查并修复 Excel 中的数据一致性问题

---

### 4. PDF 转 PPT（04_PDF_to_PPT）

将 PDF 文件转换为 PPT 格式。

**📁 使用步骤：**

1. 将 PDF 文件放入 `04_PDF_to_PPT/input/` 文件夹

2. 运行脚本：
   ```bash
   cd 04_PDF_to_PPT
   python pdf_to_ppt.py
   ```

3. 查看输出：`04_PDF_to_PPT/output/` 文件夹

**功能：** 每页 PDF 转换为一张 PPT 幻灯片

---

### 6. PPT 边角水印去除（06_PPT_Watermark_Removal）

针对 **PPTX 内嵌的大图截图页** 清理右下角重复 logo，**可编辑页面自动跳过**，避免遮挡正文内容。

**📁 使用步骤：**

1. 将 PPTX 放入 `06_PPT_Watermark_Removal/input/` 文件夹

2. 运行脚本（默认处理 input 目录下首个 PPTX）：
   ```bash
   cd 06_PPT_Watermark_Removal
   python pptx_corner_logo_patch.py
   ```

3. 查看输出：`06_PPT_Watermark_Removal/output/原文件名_clean.pptx`

**可选参数：**
- **指定输入/输出：**
  ```bash
  python pptx_corner_logo_patch.py "input/your.pptx" -o "output/your_clean.pptx"
  ```

**说明：**
- 仅处理 **全页截图类图片**（占页面积 > 80%）
- **可编辑页面不处理**（包含文本框/形状/表格/图表等）
- 遮盖区域 **自动缩放**，仅覆盖右下角小 logo，尽量不影响正文

---

### 7. PDF XSS 清理（07_PDF_XSS）

清理 PDF 中潜在的脚本与恶意协议链接，删除注释与嵌入文件，保留正常外部 URL。

**📁 使用步骤：**

1. 将待清理 PDF 放入 `07_PDF_XSS/input/` 文件夹

2. 运行脚本（默认 input → output）：
   ```bash
   cd 07_PDF_XSS
   python pdf_xss_clean.py
   ```

3. 查看输出（默认会递归遍历子文件夹，并保留相对目录结构）：
   - 输入：`input/A/B/test.pdf`
   - 输出：`output/A/B/test_cleaned.pdf`

**可选参数：**

- **指定输入/输出目录或单个文件：**
  ```bash
  python pdf_xss_clean.py --input "D:\\PDF" --output "D:\\PDF_clean"
  ```

- **覆盖已有输出：**
  ```bash
  python pdf_xss_clean.py --overwrite
  ```

- **仅遍历当前目录（关闭递归）：**
  ```bash
  python pdf_xss_clean.py --no-recursive
  ```

- **不保留目录结构（全部输出到同一层目录）：**
  ```bash
  python pdf_xss_clean.py --no-keep-structure
  ```

---

### 8. PPT 批量转 PDF（08_PPT_to_PDF）

批量将 PPT/PPTX 转换为 PDF（需 Windows + PowerPoint）。

**📁 使用步骤：**

1. 将 PPT/PPTX 放入 `08_PPT_to_PDF/input/` 文件夹

2. 运行脚本（默认 input → output）：
   ```bash
   cd 08_PPT_to_PDF
   python ppt_to_pdf.py
   ```

3. 查看输出：`08_PPT_to_PDF/output/文件名.pdf`

**可选参数：**

- **指定输入/输出目录或单个文件：**
  ```bash
  python ppt_to_pdf.py --input "D:\\PPT" --output "D:\\PDF"
  ```

- **覆盖已有输出：**
  ```bash
  python ppt_to_pdf.py --overwrite
  ```

---

### 9. Word 批量转 PDF（09_Word_to_PDF）

批量将 Word 文档转换为 PDF（需 Windows + Microsoft Word + `pywin32`）。

**📁 使用步骤：**

1. 将 Word 文档放入 `09_Word_to_PDF/input/` 文件夹

2. 运行脚本（默认 input → output）：
   ```bash
   cd 09_Word_to_PDF
   python word_to_pdf.py
   ```

3. 查看输出（默认会保留输入目录的相对结构）：
   - 输入：`input/A/B/test.docx`
   - 输出：`output/A/B/test.pdf`

**可选参数：**

- **指定输入/输出目录或单个文件：**
  ```bash
  python word_to_pdf.py --input "D:\\DOCS" --output "D:\\PDF"
  ```

- **递归遍历子文件夹（默认开启）：**
  ```bash
  python word_to_pdf.py --recursive
  ```

- **仅遍历当前目录（关闭递归）：**
  ```bash
  python word_to_pdf.py --no-recursive
  ```

- **保留目录结构输出（默认开启，避免同名文件互相覆盖）：**
  ```bash
  python word_to_pdf.py --keep-structure
  ```

- **不保留目录结构（全部输出到同一层目录）：**
  ```bash
  python word_to_pdf.py --no-keep-structure
  ```

- **覆盖已有输出：**
  ```bash
  python word_to_pdf.py --overwrite
  ```

---

### 10. C 盘垃圾/空文件清理（10_C_Drive_Cleanup）

清理 C 盘常见临时/缓存位置的垃圾文件、空文件与无用文件，**默认仅扫描**，避免误删。

**📁 使用步骤：**

1. 运行脚本（默认仅扫描，生成报告）：
   ```bash
   cd 10_C_Drive_Cleanup
   python c_drive_cleanup.py
   ```

2. 执行删除（仅处理修改时间早于 N 天的文件，默认 7 天）：
   ```bash
   python c_drive_cleanup.py --delete --days 7
   ```

3. 查看输出：`10_C_Drive_Cleanup/output/cleanup_report.csv`

**可选参数：**

- **自定义扫描目录：**
  ```bash
  python c_drive_cleanup.py --targets "C:\\Windows\\Temp" "C:\\Users\\YourName\\AppData\\Local\\Temp"
  ```

- **删除清理后的空目录：**
  ```bash
  python c_drive_cleanup.py --delete --remove-empty-dirs
  ```

- **使用 input/targets.txt 覆盖默认目录：**
  - 每行一个目录，可用 `#` 注释

**默认扫描目录（仅 C 盘）：**
- `C:\\Windows\\Temp`
- `C:\\Users\\<User>\\AppData\\Local\\Temp`
- `C:\\Users\\<User>\\AppData\\Local\\Microsoft\\Windows\\INetCache`
- `C:\\Users\\<User>\\AppData\\Local\\CrashDumps`

---

### 11. PPTX/PDF 转原生 PPT（11_PPTX_PDF_to_PPT）

将 PDF 或图片型 PPTX 中的表格识别为可编辑表格，并重建为原生 PPT。

**📁 使用步骤：**

1. 将 PDF 或 PPTX 放入 `11_PPTX_PDF_to_PPT/input/` 文件夹

2. 运行脚本：
   ```bash
   cd 11_PPTX_PDF_to_PPT
   python convert_to_native_ppt.py
   ```

3. 查看输出：`11_PPTX_PDF_to_PPT/output/原文件名_editable.pptx`

**可选参数：**

- **指定输入/输出文件：**
  ```bash
  python convert_to_native_ppt.py --input "input/raw.pdf" --output "output/editable.pptx"
  ```

- **调整 PDF 渲染精度与 OCR 语言：**
  ```bash
  python convert_to_native_ppt.py --dpi 300 --lang ch
  ```

**依赖说明：**
- 需安装 `paddleocr` 及其底层 `paddlepaddle`
- PDF 渲染依赖 `pymupdf`，表格重建依赖 `python-pptx` 与 `pandas`

---

### 12. Python 脚本转 EXE（12_Py_to_EXE）

将 .py 脚本打包为 Windows 可执行文件（基于 PyInstaller）。

**📁 使用步骤：**

1. 将 .py 脚本放入 `12_Py_to_EXE/input/` 文件夹

2. 运行脚本：
   ```bash
   cd 12_Py_to_EXE
   python py_to_exe.py
   ```

3. 查看输出：`12_Py_to_EXE/output/`

**可选参数：**

- **指定输入/输出与名称：**
  ```bash
  python py_to_exe.py --input "input/demo.py" --output "output" --name "demo"
  ```

- **目录模式（关闭 onefile）/ 隐藏控制台：**
  ```bash
  python py_to_exe.py --dir --noconsole
  ```

- **指定图标：**
  ```bash
  python py_to_exe.py --icon "input/app.ico"
  ```

- **打包后清理 `build` 与 `.spec`：**
  ```bash
  python py_to_exe.py --clean-artifacts
  ```

**说明：**
- `build/`：PyInstaller 的临时构建目录
- `.spec`：打包配置文件（可用于二次定制与复现打包）

**依赖说明：**
- 需安装 `pyinstaller`

---

### 13. WiFi 密码查看（13_WiFi_Passwords）

读取 Windows 本机已保存的 WiFi 配置并导出密码（需具备相应权限）。

**📁 使用步骤：**

1. 运行脚本：
   ```bash
   cd 13_WiFi_Passwords
   python wifi_passwords.py
   ```

2. 查看输出：`13_WiFi_Passwords/output/wifi_passwords.csv`

**可选参数：**

- **指定输出路径：**
  ```bash
  python wifi_passwords.py --output "output/wifi_passwords.csv"
  ```

- **指定编码 / 静默模式：**
  ```bash
  python wifi_passwords.py --encoding gbk --quiet
  ```

**说明：**
- 仅支持 Windows（依赖 `netsh`）
- 默认仅导出 **有密码** 的 WiFi，并在控制台显示（如需静默请使用 `--quiet`）

---

### 14. 目录文件数量统计（14_Folder_File_Count）

递归统计指定目录下所有文件数量，并输出 TXT 与 Excel 结果。

**📁 使用步骤：**

1. 运行脚本：
   ```bash
   cd 14_Folder_File_Count
   python folder_file_count.py --path "D:\\data"
   ```

   - 不传 `--path` 时，脚本会提示输入目录路径（非交互环境需显式传入 `--path`）

2. 查看输出：`14_Folder_File_Count/output/folder_file_count.txt` 与 `folder_file_count.xlsx`（含树形汇总与文件名列表）

**可选参数：**

- **指定输出目录：**
  ```bash
  python folder_file_count.py --path "D:\\data" --output "output"
  ```

---

### 15. 文献批量下载（15_Paper_Batch_Download）

根据 DOI / PMID / 标题 / URL 批量下载可公开获取（Open Access）的 PDF。

**📁 使用步骤：**

1. 直接传入查询：
   ```bash
   cd 15_Paper_Batch_Download
   python paper_batch_download.py --queries "10.1038/s41586-020-2649-2" "32788730" "Attention Is All You Need"
   ```

2. 从文本文件读取（每行一个，支持 `#` 注释）：
   ```bash
   python paper_batch_download.py --file "D:\\papers.txt" --mailto "your_email@example.com"
   ```

3. 查看输出：`15_Paper_Batch_Download/output/`

**功能说明：**
- 仅下载 **公开可获取** 的 PDF（Open Access），不支持非公开渠道
- 支持自动修正常见 DOI 误写
- 自动用论文标题的“主干部分”重命名文件名（提炼前若干关键词）


---

### 16. PDF 标题驱动重命名（16_PDF_Sanitizer）

读取 `input/` 下的 PDF，**多策略提取标题并极简命名**，随后 **直接剪切移动** 到 `output/`。

**📁 使用步骤：**

1. 将待处理 PDF 放入 `16_PDF_Sanitizer/input/`

2. 运行脚本：
   ```bash
   cd 16_PDF_Sanitizer
   python pdf_sanitizer.py
   ```

3. 查看输出：`16_PDF_Sanitizer/output/`

**补充说明：**
- 默认会**递归遍历子文件夹**，并在 `output/` 内**保留相对目录结构**
- 脚本会对 PDF 做“剪切移动”，执行完毕后 `input/` 下的 PDF 会被迁移走（目录可能保留为空）

**功能说明：**
- **标题提取优先级**：视觉层级标题（最大字号）→ 元数据 `Title` → 正文前 15 行
- 当正文文本过少时自动启用 OCR 作为后备识别（需 `pytesseract` + 本地 Tesseract）
- 副标题截断：遇到中英文冒号直接切断，只保留主标题
- 括号/引号等噪点清理，全角标点与非法路径字符规整
- 中英文双策略精简：中文压缩到不超过 40 字；英文保留前若干关键词
- 自动追加年份后缀（如 `标题-2021.pdf`，年份来源于正文或元数据）
- 同名冲突自动追加序号（如 `_2`）
- 执行完毕后 `input/` 将被清空

**可选参数：**

- **指定输入/输出目录（相对 `16_PDF_Sanitizer/`）：**
  ```bash
  python pdf_sanitizer.py --input "input" --output "output"
  ```

- **仅遍历当前目录（关闭递归）：**
  ```bash
  python pdf_sanitizer.py --no-recursive
  ```

- **不保留目录结构（全部输出到同一层目录）：**
  ```bash
  python pdf_sanitizer.py --no-keep-structure
  ```

- **覆盖已存在输出：**
  ```bash
  python pdf_sanitizer.py --overwrite
  ```

---

### 17. PDF eCTD 转换（17_PDF_eCTD_Converter）

将 `input/` 下的 PDF 做 eCTD 常见约束的“可提交化”处理：**可读性校验、拒绝密码锁定、移除附件、移除非超链接注释、清理外部/非法链接、重写保存（含 Fast Web View）并导出 Excel 审计报告**。

**📁 使用步骤：**

1. 将待处理 PDF 放入 `17_PDF_eCTD_Converter/input/`

2. 运行脚本（默认 input → output）：
   ```bash
   cd 17_PDF_eCTD_Converter
   python pdf_ectd_converter.py --report "output/ectd_report.xlsx"
   ```

3. 查看输出（默认会递归遍历子文件夹，并保留相对目录结构）：
   - 输出 PDF：`17_PDF_eCTD_Converter/output/.../文件名_ectd.pdf`
   - 审计报告：`17_PDF_eCTD_Converter/output/ectd_report.xlsx`

**可选参数：**

- **仅校验（不输出）**：
  ```bash
  python pdf_ectd_converter.py --validate-only --report "output/ectd_report.xlsx"
  ```

- **覆盖已有输出**：
  ```bash
  python pdf_ectd_converter.py --overwrite
  ```

- **输出保持原文件名（默认追加 _ectd）**：
  ```bash
  python pdf_ectd_converter.py --keep-name
  ```

- **仅遍历当前目录（关闭递归）：**
  ```bash
  python pdf_ectd_converter.py --no-recursive
  ```

- **不保留目录结构（全部输出到同一层目录）：**
  ```bash
  python pdf_ectd_converter.py --no-keep-structure
  ```

### 18. 代理配置导出（18_Proxy_Config_Export）

导出当前系统代理配置（**Windows 注册表 + 环境变量**）到文本文件，便于复制到终端、构建机或远程环境。

**📁 使用步骤：**

1. 运行脚本：
   ```bash
   cd 18_Proxy_Config_Export
   python proxy_config_export.py
   ```

2. 查看输出：`18_Proxy_Config_Export/output/proxy_config_YYYYMMDD_HHMMSS.txt`

**功能说明：**
- 读取注册表 `HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings`
  - `ProxyEnable`（是否启用）
  - `ProxyServer`（代理地址）
  - `ProxyOverride`（免代理地址）
- 读取环境变量：`HTTP_PROXY` / `HTTPS_PROXY` / `NO_PROXY`（含大小写兼容）
- 当系统代理已启用时，自动输出可直接复用的：
  - `http://<proxy_server>`
  - `https://<proxy_server>`
  - `NO_PROXY`（将 `;` 自动转换为 `,`）
- 输出文件带时间戳，默认写入 `18_Proxy_Config_Export/output/`

**说明：**
- 仅支持 Windows（依赖 `reg query` 读取注册表）
- 仅使用 Python 标准库，无需额外安装第三方依赖

---

### 19. PDF 合并（19_PDF_Merge）

将 `input/` 下多个 PDF 按“命名从低到高”的自然排序合并为一个 PDF（支持遍历子文件夹，按相对路径排序）。

**📁 使用步骤：**

1. 将待合并 PDF 放入 `19_PDF_Merge/input/`（可放在子文件夹中）

2. 运行脚本（默认 input → output）：
   ```bash
   cd 19_PDF_Merge
   python merge_pdf.py
   ```

3. 查看输出：`19_PDF_Merge/output/merged.pdf`

**可选参数：**

- **指定输入/输出目录或单个文件：**
  ```bash
  python merge_pdf.py --input "D:\\PDFS" --output "D:\\OUT"
  ```

- **自定义输出文件名：**
  ```bash
  python merge_pdf.py --output-name all.pdf
  ```

- **覆盖已有输出：**
  ```bash
  python merge_pdf.py --overwrite
  ```

---

### 20. Word -> Excel(表+图) 自动复刻（20_Word_to_Excel_to_Figure）

把 `20_Word_to_Excel_to_Figure/input/` 下的 Word（`.doc/.docx/.rtf`）表格数据抽取出来，填入 `20_Word_to_Excel_to_Figure/Template/` 的骨架 Excel 中指定的图表数据区间（`chart.series.cat/val` 引用区间），生成高保真复刻 Excel（尽量避免 Office/WPS 修复弹窗）。

**推荐工作流（需要你确认映射后再生成 output）：**

1. 先生成候选映射计划（plan only）：
   ```bash
   cd 20_Word_to_Excel_to_Figure
   python word_to_excel_to_figure.py --input-dir "input" --plan-only
   ```
   输出到：`output/table_mapping_plan_<骨架xlsx名>.json`

2. 打开 `table_mapping_plan_<骨架xlsx名>.json`，为你确认的候选项添加 `selected: true`

3. 用确认后的映射生成 output：
   ```bash
   python word_to_excel_to_figure.py --input-dir "input" --table-map-json "output/table_mapping_plan_<骨架xlsx名>.json"
   ```

**关键说明：**
- 骨架 Excel 放在 `Template/`（多骨架时可用 `--template-xlsx` 指定）。
- 排布变化时先用 `--plan-only` 生成 JSON，在候选项上标记 `selected: true` 后再跑正式生成（详见模块内 `README.md`）。
- **`table_mapping_logic.py`**：映射 plan 与 JSON 加载的单一数据源，勿维护两套逻辑。
- **`repair_output_by_patch.py`**：在已有「数据正确」的 xlsx 上，把图表 cat/val 区间拷回模板副本并自检，用于修复 openpyxl 直接保存导致的 Office 结构问题。

---

### 21. 文档双向翻译（21_File_Translator）

对 `input/` 下的 Excel/CSV/Word/PDF 做翻译处理，输出到 `output/`。默认策略为 **translators 免费通道优先**（如 bing/google），并支持 `DeepL` 与 `LibreTranslate` 兜底。

**📁 使用步骤：**

1. 将待翻译文件放入 `21_File_Translator/input/`（支持 `.xlsx/.csv/.docx/.doc/.pdf`）
2. 运行脚本：
   ```bash
   cd 21_File_Translator
   python file_translator.py --self-test
   python file_translator.py
   ```
3. 查看输出：`21_File_Translator/output/*_en2zh.*` 或 `*_zh2en.*`

**默认行为：**
- 默认翻译列（Excel/CSV）：`Term`、`SOC`、`Comment`、`PT Name`、`SOC Name`
- 默认翻译引擎：`auto`（translators -> DeepL -> LibreTranslate -> 原文）
- 支持方向：`--direction en2zh|zh2en`
- 支持术语词典优先替换（`--glossary`）
- 支持 JSON 持久化缓存（`--cache-file`）与文件级并发（`--max-workers`）
- 对 `.xlsx`：尽量保留原工作簿样式与结构，仅新增双语列
- 对 `.csv`：输出为 `.xlsx` 双语文件（IME 列表自动兼容前 11 行 metadata）
- 对 `.docx/.doc`：生成翻译副本，覆盖正文/表格/页眉页脚，并可扩展文本框/脚注
- 对 `.pdf`：支持覆盖重绘（`overlay`）与仅导出双语文本层（`bilingual-text-layer`）
- Windows 下默认优先 `COM` 写回引擎，最大化保留图像/对象/排版（避免 `openpyxl` 丢失 WMF）
- 默认读取 `21_File_Translator/.env`（支持脱敏日志）

**可选参数：**

- **API 自检：**
  ```bash
  python file_translator.py --self-test
  ```

- **指定翻译列（逗号分隔）：**
  ```bash
  python file_translator.py --columns "Term,SOC,Comment"
  ```

- **指定写回引擎：**
  ```bash
  python file_translator.py --engine com
  python file_translator.py --engine openpyxl
  ```

- **指定翻译引擎：**
  ```bash
  python file_translator.py --provider tsfree --ts-engine bing
  python file_translator.py --provider deepl
  python file_translator.py --provider libre
  ```

- **中译英：**
  ```bash
  python file_translator.py --direction zh2en
  ```

- **PDF 仅导出双语文本层：**
  ```bash
  python file_translator.py --pdf-mode bilingual-text-layer
  ```
  - 该模式会保留原 PDF，并额外输出同名 `*.bilingual.txt` 双语文本层文件。

- **并发与缓存：**
  ```bash
  python file_translator.py --max-workers 3 --cache-file "output/translation_cache.json"
  python file_translator.py --no-cache
  ```
  - 并发模式下若检测到 `com` 引擎会自动降级 `openpyxl`，避免 COM 线程安全问题。
  - 缓存采用“临时文件 + 原子替换”写回，降低意外中断导致缓存损坏的风险。

---

### 22. DNS 泄漏诊断（22_DNS_Leak_Detector）

用于检测“公网出口”与“上游 DNS 解析节点”是否异常偏离，辅助排查代理规则错误、DNS 泄漏与分流配置问题。

**📁 使用步骤：**

1. 运行脚本（默认 TUN 模式）：
   ```bash
   cd 22_DNS_Leak_Detector
   python dns_leak_detector.py --mode tun
   ```

2. 可选：切换 SOCKS 模式（本地端口）
   ```bash
   python dns_leak_detector.py --mode socks --socks-port 10808
   ```

3. 可选：保存诊断 JSON 报告
   ```bash
   python dns_leak_detector.py --save-json
   ```

**输出说明：**
- 终端实时日志：链路探活、出口 IP、DNS 区域、风险结论
- 报告文件（可选）：`22_DNS_Leak_Detector/output/dns_diagnostic_<mode>_<timestamp>.json`

---

### 23. PDF 书签承前缩放（23_PDF_Bookmark_Inherit_Zoom）

批量处理 PDF：通过 **PyMuPDF** 重写目录（TOC），为书签目标注入 **XYZ + zoom=0**，使阅读器在点击书签时 **保持当前缩放比例**（承前缩放）；无书签时仍会执行垃圾回收与流压缩以优化体积。加密 PDF 会跳过并记录失败。

**📁 使用步骤：**

1. 将待处理 PDF 放入 `23_PDF_Bookmark_Inherit_Zoom/input/`

2. 运行脚本（默认读取模块目录下 `input` → `output`，线程数 6）：
   ```bash
   cd 23_PDF_Bookmark_Inherit_Zoom
   python pdf_bookmark_inherit_zoom.py
   ```

3. 查看输出：`23_PDF_Bookmark_Inherit_Zoom/output/`（与源文件同名）

**可选调整：**

- 修改 `pdf_bookmark_inherit_zoom.py` 末尾的 `INPUT_DIRECTORY`、`OUTPUT_DIRECTORY` 为任意绝对路径（Windows 建议使用原始字符串 `r'C:\...'`）。
- 调整 `batch_set_scaling(..., max_workers=6)` 中的并发线程数。

**依赖说明：**

- 需安装 `pymupdf`（`fitz`），见根目录 `requirements.txt`。

### 24. Word 指定表格转 Excel（24_Word_Tables_to_Excel）

将 Word（`.doc/.docx/.rtf`）中指定表格导出为 Excel。定位方式与优先级（详见该目录 `README.md`）：

- **`--merge-tables-from` / `--merge-tables-to`**：多段**顶层** `Document.Tables` 纵向合并（与表题/序号互斥）
- **`--table-title`**：按表题文本定位
- **`--table-indices` / `--table-index`**：按 Word 顶层表序号（1-based）
- **`--header-keywords`**：全量读表后按表头关键字筛选

辅助与自检：`--list-word-tables`（COM，与序号一致）、`--list-docx-tables`（仅 XML，段数常多于顶层表）、`--dry-run`（只统计不写 xlsx）。大表抽取优先按行 `Range.Text`，减轻整表截断与逐格 COM 过慢问题。

导出时支持表头样式化（深色表头、冻结窗格、自动筛选、边框、列宽自适应）；默认 `--header-rows 1`，可按文档结构调整。

## 配置说明

复制 `config.example.yaml` 为 `config.yaml`，按需修改：

```yaml
pdf_path: "03_PDF_to_Excel/input/不同剂量组ADR分析 (TFL).pdf"
excel_path: "03_PDF_to_Excel/output/不同剂量组ADR分析 (TFL).xlsx"
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

- **PDF 解析扩展**：修改 [`src/pdf_reader.py`](src/pdf_reader.py)
- **Excel 写入扩展**：修改 [`src/excel_writer.py`](src/excel_writer.py)
- **图表样式与临床配色**：修改 [`01_Excel_Charts/build_charts_xlsxwriter.py`](01_Excel_Charts/build_charts_xlsxwriter.py) 内 `COLOR_MAP` 或 [`src/color_theme.py`](src/color_theme.py)
- **独立配色模块**： [`05_Excel_Chart_Colors/apply_clinical_colors.py`](05_Excel_Chart_Colors/apply_clinical_colors.py) 支持多期刊预设（NPG、Lancet、NEJM 等）
- **Word→Excel 图表复刻**：主逻辑在 [`20_Word_to_Excel_to_Figure/word_to_excel_to_figure.py`](20_Word_to_Excel_to_Figure/word_to_excel_to_figure.py)；子表映射与 plan 在 [`20_Word_to_Excel_to_Figure/table_mapping_logic.py`](20_Word_to_Excel_to_Figure/table_mapping_logic.py)

## 注意事项

1. **文件组织**：所有输入文件放在各模块的 `input/` 文件夹，输出文件自动保存到 `output/` 文件夹
2. **自动备份**：Excel 图表生成会自动备份旧文件（.bak.xlsx）
3. **引擎选择**：Excel 图表生成推荐使用 XlsxWriter 引擎，避免 XML 结构问题
4. **路径配置**：PDF 提取依赖文件结构，需根据实际 PDF 调整配置
5. **模块独立**：各模块独立运行，互不干扰，便于维护和扩展
6. **PPT 合并**：推荐先运行 `merge_ppt.py` 做基础合并，再运行 `ppt_engine.py` 做叙事重组
7. **依赖安装**：见根目录 `requirements.txt` 分组注释；`scikit-learn`（02）、`pymupdf`（多 PDF 模块含 23）、`requests`（15/22）、`pywin32`（仅 Windows，09/20/08 等 Office 自动化）、`paddleocr`（11，可选注释以减小体积）按所用模块生效

## 项目优势

- ✅ **自动化**：一键完成数据处理、图表生成、文件合并
- ✅ **智能化**：TF-IDF 算法确保内容质量，自动去重和筛选
- ✅ **高保真**：完整保留原始视觉效果和格式
- ✅ **可追溯**：详细的处理记录便于审计和验证
- ✅ **专业级**：符合临床研究报告标准

## 维护说明

- 各模块独立运行，按需使用；输入放 `input/`，输出在 `output/`。
- 历史优化与代码规范已融入代码与本文档，重大变更见版本记录或 Git 历史。

## 许可证

本项目采用 [MIT License](LICENSE.md) 开源协议。
