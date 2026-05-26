# 15_PDF_Title_Renamer

PDF 文献按**正文标题 + 年份**自动重命名，并从 `input/` **剪切**到 `output/`（非复制）。引擎 **PDF Sanitizer v6.9**，面向 Elsevier / PLOS / Nature 系、FDA Guidance 封面等常见版式。

## 快速开始

```bash
cd 15_PDF_Title_Renamer
# 将待处理 PDF 放入 input/（可含子文件夹）
python pdf_sanitizer.py
```

默认输出：`output/<相对路径>/标题词_用下划线连接-年份.pdf`  
示例：`Clinical_Evaluation_Oligonucleotides_Adjuvants_Vaccines_Targeting_Infectious_Diseases_Cancer-2014.pdf`

## 标题提取策略（概要）

| 优先级 | 手段 | 说明 |
|--------|------|------|
| 1 | 视觉层级 | 首页按字号取最大文本块，过滤期刊页眉、文章类型行、待刊横幅 |
| 2 | 学术首屏解析 | 跳过 `RESEARCH ARTICLE`、`ARTICLE IN PRESS`、`Please cite this article…` 等，合并多行正文标题 |
| 3 | 元数据 / 首行 | PDF `title` 字段或首屏前几行文本 |
| 4 | OCR 后备 | 原生文本过少时用 Tesseract（`chi_sim+eng`，可选） |

**噪声过滤（v6.9 增强）**：期刊 masthead、Elsevier 待刊状态行、引用提示行、卷期页眉、纯作者姓行；FDA `Guidance for Industry` 泛化封面会剥离前缀保留具体题目。

**文件名规则**：中英文冒号前截断副标题；英文 Smart Title Case（保留 CpG、mRNA 等缩写）；中文去空格压缩；非法路径字符替换；末尾悬挂介词修剪；重名自动 `_1`、`_2`…

**年份**：期刊页眉 `(2014)`、版权/出版关键词、FDA docket、元数据创建日期等，依次回退。

## 命令行参数

| 参数 | 默认 | 说明 |
|------|------|------|
| `--input` / `-i` | `input` | 输入目录（相对本模块） |
| `--output` / `-o` | `output` | 输出目录 |
| `--recursive` / `--no-recursive` | 开启 | 是否递归子文件夹 |
| `--keep-structure` / `--no-keep-structure` | 开启 | 输出是否保留相对目录结构 |
| `--overwrite` | 关闭 | 覆盖已存在的目标 PDF |

示例：

```bash
python pdf_sanitizer.py -i "D:\Papers\inbox" -o "output" --no-keep-structure --overwrite
```

## 依赖

- **必需**：`pymupdf`（fitz）、`Pillow`、`tqdm`（见根目录 `requirements.txt`）
- **可选 OCR**：`pytesseract` + 本机 [Tesseract-OCR](https://github.com/tesseract-ocr/tesseract)（`chi_sim+eng`）；未安装时仅禁用 OCR 后备，其余流程正常

Windows 下若未配置 PATH，可在 `pdf_sanitizer.py` 中取消注释并设置：

```python
# pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
```

## 注意事项

1. **剪切模式**：成功后 `input/` 内对应 PDF 会被移走，请先备份原件。
2. **扫描版 PDF**：无文本层时依赖 OCR，识别质量取决于 DPI 与版式。
3. **误识别**：极端版式（双栏、超长作者列表）仍可能需人工核对；可用 `--overwrite` 重跑前先放回 `input/`。

更完整的模块索引见仓库根目录 [`README.md`](../README.md) 与 [`README_EN.md`](../README_EN.md)。
