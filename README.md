# Clinical Data Automation Toolkit

临床数据自动化处理工具集，提供 PDF 数据提取、Excel 图表生成、PPT 整合等自动化分析处理功能。

## 环境要求

- Python 3.8+
- 依赖见 `requirements.txt`
- Windows 专用依赖（`pywin32`）仅在 Windows + Office 自动化模块中需要

## 安装

```bash
pip install -r requirements.txt
```

> 非 Windows 环境若只使用 PDF/Excel 相关模块，可移除或注释 `pywin32` 后再安装。

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

- **批量处理输入目录：**
  ```bash
  python apply_clinical_colors.py --batch --input "input" --output "output" --palette Lancet
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

## 注意事项

1. **文件组织**：所有输入文件放在各模块的 `input/` 文件夹，输出文件自动保存到 `output/` 文件夹
2. **自动备份**：Excel 图表生成会自动备份旧文件（.bak.xlsx）
3. **引擎选择**：Excel 图表生成推荐使用 XlsxWriter 引擎，避免 XML 结构问题
4. **路径配置**：PDF 提取依赖文件结构，需根据实际 PDF 调整配置
5. **模块独立**：各模块独立运行，互不干扰，便于维护和扩展
6. **PPT 合并**：推荐先运行 `merge_ppt.py` 做基础合并，再运行 `ppt_engine.py` 做叙事重组
7. **依赖安装**：确保安装所有依赖，特别是 `scikit-learn`（PPT 合并需要）

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
