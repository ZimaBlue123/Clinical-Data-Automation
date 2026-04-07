# 11_PDF_to_Excel_Rule_Extract

通用的「规则驱动」PDF → Excel 提取模块。

## 作用

- 读取根目录的 `config.yaml`（或 `--config` 指定的 YAML）
- 按规则在 PDF 中检索关键词/表格/页面文本
- 将提取结果写入指定 Excel 的 sheet/cell

## 快速开始

1. 在项目根目录复制配置：

```bash
copy config.example.yaml config.yaml
```

2. 修改 `config.yaml` 中的 `pdf_path` / `excel_path` 与规则。

3. 运行：

```bash
cd 11_PDF_to_Excel_Rule_Extract
python main.py --config ../config.yaml
```

## 与 PDF 去水印/干扰区联动（可选）

若 PDF 存在页眉页脚/水印等干扰区，建议先运行 `19_PDF_Watermark_Removal` 生成 `*_boxes.json`，
再在本模块中传入 `--exclusion-json`，并自动生成 `mapping_audit` 审计信息：

```bash
python main.py --config ../config.yaml --exclusion-json ../19_PDF_Watermark_Removal/output/你的文件_boxes.json
```

