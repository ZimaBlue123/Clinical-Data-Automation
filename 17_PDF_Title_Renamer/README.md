# 16_PDF_Title_Renamer

PDF 文献按**正文标题 + 年份**自动重命名，并从 `input/` **剪切**到 `output/`（非复制）。引擎 **PDF Sanitizer v7.1**，面向 Elsevier / PLOS / Nature 系、MDPI 复数期刊、FDA Guidance 封面等常见版式。

## 快速开始

```bash
cd 16_PDF_Title_Renamer
# 将待处理 PDF 放入 input/（可含子文件夹）
python pdf_sanitizer.py
```

默认输出：`output/<相对路径>/标题词_用下划线连接-年份.pdf`  
示例：`Clinical_Evaluation_Oligonucleotides_Adjuvants_Vaccines_Targeting_Infectious_Diseases_Cancer-2014.pdf`

## 标题提取策略（概要）

| 优先级 | 手段 | 说明 |
|--------|------|------|
| 1 | 视觉层级 | 首页按字号取最大文本块，过滤期刊页眉、文章类型行、待刊横幅、出版商角色行 |
| 2 | 学术首屏解析 | 跳过 `RESEARCH ARTICLE`、`ARTICLE IN PRESS`、`Please cite this article…` 等，合并多行正文标题 |
| 3 | 元数据 / 首行 | PDF `title` 字段或首屏前几行文本 |
| 4 | OCR 后备 | 原生文本过少时用 Tesseract（`chi_sim+eng`，可选） |

**噪声过滤（v7.1 增强）**：
- 期刊 masthead：含 PLOS / Nature 系 / EMBO / Wiley / T&F / MDPI（含 `vaccines` / `Viruses` / `Antibodies` / `Biomedicines` / `Toxins` / `Pathogens` / `Microorganisms` / `Life` / `Biomolecules` / `Cells` 等 16 种 MDPI 复数刊名）。
- Elsevier 待刊状态、引用提示行、卷期页眉、纯作者姓行。
- 出版商角色 / 元数据短语：Academic Editor / Section Editor / Author Contributions / Funding / Acknowledgments / Supplementary Materials 等。
- MDPI 启发式：1–2 个标题化名词（Vaccines / Antibodies）、压缩形式命中黑名单 → 视为刊名而非正文。
- FDA `Guidance for Industry` 泛化封面会剥离前缀保留具体题目。

**文件名规则**：中英文冒号前截断副标题；英文 Smart Title Case（保留 CpG、mRNA 等缩写）；中文去空格压缩；非法路径字符替换；末尾悬挂介词修剪；元数据单词剔除；重名自动 `_1`、`_2`…

**Windows 文件锁防护**：被 IDE / PDF 阅读器 / 资源管理器预览持有时，先 `copy2 → unlink → 一次重试`，输出不可丢失；最终失败给出 `hint=close_locked_input_then_retry` 提示。

**依赖容错加载**：`pymupdf / Pillow / tqdm` 缺失时会优雅降级，缺哪个就跳过 OCR 或视觉层级环节；不再因部分环境直接崩溃。

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

- **必需**：`pymupdf`（fitz）、`Pillow`、`tqdm`（见根目录 `requirements.txt`，pymupdf≥1.27.0）
- **可选 OCR**：`pytesseract` + 本机 [Tesseract-OCR](https://github.com/tesseract-ocr/tesseract)（`chi_sim+eng`）；未安装时仅禁用 OCR 后备，其余流程正常

Windows 下若未配置 PATH，可在 `pdf_sanitizer.py` 中取消注释并设置：

```python
# pytesseract.pytesseract.tesseract_cmd = r'D:\Tesseract-OCR\tesseract.exe'
```

## 注意事项

1. **剪切模式**：成功后 `input/` 内对应 PDF 会被移走，请先备份原件。若重跑时 `output/` 已有同名目标，需加 `--overwrite`。
2. **Windows 文件锁**：被其他进程占用时，不要直接在 IDE / PDF 阅读器 / 文件管理器预览中运行该脚本。先关闭持锁程序。
3. **扫描版 PDF**：无文本层时依赖 OCR，识别质量取决于 DPI 与版式。
4. **误识别**：极端版式（双栏、超长作者列表）仍可能需人工核对；可用 `--overwrite` 重跑前先放回 `input/`。

更完整的模块索引见仓库根目录 [`README.md`](../README.md) 与 [`README_EN.md`](../README_EN.md)。
