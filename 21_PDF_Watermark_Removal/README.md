# 21_PDF_Watermark_Removal（PDF 干扰区定位与审计）

本模块用于在 **不物理抹除 PDF 内容** 的前提下，定位疑似水印/页眉页脚/重复干扰区域，输出 **排除框坐标**、**审计叠加 PDF** 与 **清洗后逐页文本**，供下游抽取（如 `12_PDF_to_Excel_Rule_Extract`）跳过干扰区，并保留可追溯审计材料。

## 能力边界

| 会做 | 不会做 |
|------|--------|
| 矢量关键词命中或 OCR 辅助定位干扰区 | 输出「无水印」成品 PDF（物理去水印） |
| 生成半透明红框审计版 `*_audit_masked.pdf` | 修改原始 PDF 字节流以删除水印对象 |
| 输出 v2 `*_boxes.json`（含 rotation / mediabox / cropbox） | 保证 100% 不误标正文（需结合 `mapping_audit` 与人工 spot-check） |

## 依赖与环境

- **Python**：与仓库一致（3.8+）。
- **pip**：`pymupdf`（fitz）、`pdfplumber`、`Pillow`、`pytesseract`（见根目录 `requirements.txt`）。
- **Tesseract-OCR**：本机可执行文件，非 pip 包。
  - **环境变量**（推荐）：`TESSERACT_CMD_PATH` 指向 `tesseract.exe` 绝对路径。
  - **Windows 兜底**：若未设置环境变量，代码会尝试 `D:\Tesseract-OCR\tesseract.exe`（见 `steps/utils.py` 中 `configure_tesseract`）。

## 目录约定

```
21_PDF_Watermark_Removal/
├── input/           # 默认输入 PDF
├── output/          # 默认输出目录
├── steps/           # 管线子步骤（triage / vector / ocr / merge / audit / extract）
├── main.py          # 入口
└── README.md
```

## 处理管线（概要）

1. **Triage**：前几页用 PyMuPDF `search_for` 试探矢量关键词；满足阈值则走 **vector**，否则 **ocr**。
2. **Vector**：全页 `search_for` 关键词，合并相邻框。
3. **OCR**：按 DPI 栅格化页面，`pytesseract.image_to_data` 取词级框；可叠加「跨页重复词」启发式（关键词未命中时）。
4. **Merge**：同页重叠框合并并钳制在页面范围内。
5. **Audit PDF**：在 **同一 `fitz` 文档** 上叠加矩形注释，另存为 `*_audit_masked.pdf`（底层内容仍在）。
6. **Safe text**：通过 `src.pdf_reader.extract_text_from_pdf` + 排除框，写出 `*_clean_text_by_page.json`。

## 输出文件（每个输入 PDF 一份）

| 文件 | 说明 |
|------|------|
| `{stem}_boxes.json` | **v2**：每页键为 **从 1 起的页码字符串** `"1"`,`"2"`,…，值为 `rotation`、`mediabox`、`cropbox`、`page_width`、`page_height`、`boxes` |
| `{stem}_audit_masked.pdf` | 审计用叠加框 PDF |
| `{stem}_clean_text_by_page.json` | 按排除区过滤后的逐页文本（字符串键同上） |
| `{stem}_watermark_report.json` | 摘要：模式、关键词、各页框数量等 |

### `*_boxes.json`（v2）结构示例

```json
{
  "1": {
    "rotation": 0,
    "mediabox": [0.0, 0.0, 595.28, 841.89],
    "cropbox": [0.0, 0.0, 595.28, 841.89],
    "page_width": 595.28,
    "page_height": 841.89,
    "boxes": [[10.0, 20.0, 100.0, 40.0]]
  }
}
```

> 若你手边仍有旧版 v1（值为 `[[x0,top,x1,bottom], ...]` 的列表），`12_PDF_to_Excel_Rule_Extract` 仍兼容；**建议统一用本模块最新输出 v2**，以便旋转页与 `mapping_audit` 一致。

## 命令行参数

在项目根目录或本模块目录执行均可；相对路径默认相对 **本模块目录**。

```bash
cd 21_PDF_Watermark_Removal
python main.py --input "input" --output "output"
```

| 参数 | 默认 | 说明 |
|------|------|------|
| `--input` | `input` | PDF 文件或目录 |
| `--output` | `output` | 输出根目录 |
| `--recursive` / `--no-recursive` | 递归开 | 是否递归子文件夹 |
| `--keep-structure` / `--no-keep-structure` | 保留结构开 | 输出是否镜像输入相对路径 |
| `--overwrite` | 关 | 覆盖已存在输出 |
| `--keywords` | 内置中英文保密类词 | 逗号分隔，矢量与 OCR 共用 |
| `--vector-min-hit-pages` | `1` | 矢量路径所需最少命中页数 |
| `--ocr-dpi` | `200` | OCR 渲染分辨率 |
| `--ocr-conf-thresh` | `50` | OCR 词置信度下限 |
| `--ocr-lang` | `chi_sim+eng` | Tesseract 语言包 |
| `--ocr-repeated-heuristic` / `--no-ocr-repeated-heuristic` | 开 | 关键词未命中时是否启用重复词启发式 |
| `--ocr-repeated-min-pages` | `3` | 重复词最少跨页数 |

**跳过逻辑**：在未加 `--overwrite` 时，若同目录下已存在 `*_boxes.json`、`*_audit_masked.pdf`、`*_clean_text_by_page.json` 三件套，则跳过该 PDF。

## 与 `12_PDF_to_Excel_Rule_Extract` 联动

1. 先用本模块生成 `{stem}_boxes.json`（建议 v2）。
2. 在 `12_PDF_to_Excel_Rule_Extract` 中：

```bash
cd ../12_PDF_to_Excel_Rule_Extract
python main.py --config "config.yaml" --exclusion-json "../21_PDF_Watermark_Removal/output/你的文件_boxes.json"
```

`src/pdf_reader.py` 在排除时 **仅过滤带 `text` 的对象**，避免破坏 `pdfplumber` 表格网格线；并在 v2 元数据存在时对 **旋转页** 做坐标映射与钳制。

### 坐标映射审计（`mapping_audit`）

使用 `--exclusion-json` 时，`03` 默认调用 `build_mapping_audit_for_pdf` 并：

- 若同目录存在 `{pdf_stem}_watermark_report.json`（`pdf_stem` 为 `*_boxes.json` 文件名去掉 `_boxes`），则 **合并** `mapping_audit` 字段；
- 否则写入 `{pdf_stem}_mapping_audit.json`。

| `03` 参数 | 说明 |
|-----------|------|
| `--no-mapping-audit` | 不生成映射审计 |
| `--mapping-audit-output path.json` | 指定审计 JSON 路径 |

审计内容包括：`area_retention_ratio`、钳制次数与 delta、丢弃框统计、`severe_area_distortion_pages` 等，详见 `src/pdf_reader.py` 中 `prepare_exclusion_boxes_with_audit` / `build_mapping_audit_for_pdf`。

## 故障排查简表

| 现象 | 建议 |
|------|------|
| OCR 不工作 | 检查 `TESSERACT_CMD_PATH` 或 `D:\Tesseract-OCR\tesseract.exe`；确认已安装对应 `--ocr-lang` 语言数据 |
| 横向页排除错位 | 确认使用 v2 `*_boxes.json`；查看 `03` 合并后的 `mapping_audit` |
| 表格抽取变差 | 检查是否误标过大排除框；适当收紧 `--keywords` 或关闭 `--ocr-repeated-heuristic` 做对比 |
