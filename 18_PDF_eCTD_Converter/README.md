# 18_PDF_eCTD_Converter

eCTD 合规装甲与 PDF 深度清理：按《eCTD 验证标准 V1.1》附件 6 常见条款（可实现部分）批量校验、清洗并重写 PDF，导出 Excel 审计报告。

## 功能概览

| 规则 | 行为 |
|------|------|
| 6.1 | 可打开且页数 > 0 |
| 6.17 | 移除嵌入附件 |
| 6.18 | 移除除超文本链接外的注释 |
| 6.3 / 6.10 / 6.11 | 拦截外部/恶意链接 |
| 6.19 / 6.21 | 需密码打开的 PDF 拒收；输出未加密 |
| 6.20 | 初始视图 UseOutlines + OneColumn |
| 6.22 | 尝试线性化（不支持时自动降级） |
| 6.23 | 超过 5 页须有书签（默认自动补全） |
| 6.5 | 为无动作/仅容器（collapse）的书签补全 GoTo 跳转 |
| 6.6 | 修正越界或无法解析的书签目标 |
| 6.8 | 全部书签跳转使用承前缩放（zoom=0） |
| 6.25 | 无可搜索文本时报告预警 |
| 6.26 | 将 Times-Roman / Helvetica 等映射为认可字体名；`subset_fonts` 嵌入所用字体 |

## 目录约定（源码 / EXE 通用）

程序始终按**可执行文件所在目录**组织输入输出：

```text
18_PDF_eCTD_Converter/
├─ pdf_ectd_converter.exe  (或 pdf_ectd_converter.py)
├─ input/                  # 放待处理 PDF
├─ output/                 # 处理后文件归档目录
└─ ectd_report.xlsx        # 审计报告（默认）
```

- 运行时若不存在 `input/` 或 `output/`，会自动创建。  
- 输入文件与输出文件天然分离归档（`input/` 只放原文件，`output/` 放新文件）。

## 快速开始（Python 方式）

```bash
cd 18_PDF_eCTD_Converter
# 将 PDF 放入 input/ 后执行（Windows 建议用已安装 pymupdf 的 python 全路径）
python pdf_ectd_converter.py --input "./input" --output "./output" --report "./ectd_report.xlsx" --overwrite
```

输出默认命名为 `原文件名_ectd.pdf`，写入 `output/`。

## 常用参数

| 参数 | 说明 |
|------|------|
| `--input` / `-i` | 输入目录或单个 PDF（默认 `input/`） |
| `--output` / `-o` | 输出目录（默认 `output/`） |
| `--report` | Excel 审计报告路径（默认 `ectd_report.xlsx`） |
| `--overwrite` | 覆盖已存在输出 |
| `--validate-only` | 仅校验，不写出 PDF |
| `--add-auto-bookmarks [outline\|pages\|minimal]` | 超 5 页无书签时自动补全（默认 `outline`） |
| `--no-add-auto-bookmarks` | 禁用自动书签，严格按 6.23 拒收 |
| `--no-recursive` | 不递归子目录 |
| `--keep-name` | 输出文件名与源文件一致（不加 `_ectd`） |
| `--subset-fonts / --no-subset-fonts` | 是否尝试嵌入字体子集（可能增大体积；默认关闭） |
| `--linearize / --no-linearize` | 是否尝试线性化（Fast Web View；默认开启） |

## 审计报告

- 工作表 **全部**：逐文件处理结果  
- 工作表 **结构警告**：源 PDF 存在 MuPDF xref/syntax 等告警、建议人工复核的文件  

## 依赖

根目录 `requirements.txt`：`pymupdf`、`pandas`、`openpyxl`；字体子集嵌入建议安装 `fonttools`（`pip install fonttools`）。

## 封装为 EXE（PyInstaller）

在 `18_PDF_eCTD_Converter` 目录执行：

```powershell
python -m pip install --upgrade pyinstaller pymupdf pandas openpyxl
pyinstaller --noconfirm --clean --onefile --name pdf_ectd_converter pdf_ectd_converter.py
```

生成文件：`dist/pdf_ectd_converter.exe`

建议最终交付结构：

```text
任意目录/
├─ pdf_ectd_converter.exe
├─ input/
└─ output/
```

> 即使只复制 `pdf_ectd_converter.exe`，首次运行也会自动创建 `input/` 与 `output/`。

## EXE 使用方式（跨路径 / 跨电脑）

1. 把 `pdf_ectd_converter.exe` 复制到目标路径（或其他电脑）。  
2. 双击运行一次（会自动创建 `input/`、`output/`）。  
3. 把待处理 PDF 放入 `input/`。  
4. 重新运行 `exe`，处理结果输出到 `output/`，审计报告默认在同级 `ectd_report.xlsx`。  

如需命令行指定路径，也可使用与 Python 脚本相同参数：

```powershell
.\pdf_ectd_converter.exe --input ".\input" --output ".\output" --report ".\ectd_report.xlsx" --overwrite
```

## 说明

- 保存时使用 `encryption=fitz.PDF_ENCRYPT_NONE`，一般可去除与加密绑定的打印/编辑限制；**批注**按 6.18 **删除**（超链接除外），非「解除批注限制」。  
- 更完整的仓库说明见根目录 `README.md` 中「16. PDF eCTD 转换」一节。
