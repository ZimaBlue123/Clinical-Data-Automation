# 20_PDF_Duplicate_Analyzer

在同一**根目录**下的多个子文件夹之间，检测 PDF 是否重复。

## 重复判定

满足以下任一条件即视为重复（仅在**同一 `--root` 下的一级子文件夹**之间比对，不跨 root）：

1. **文件名相同**（不区分大小写）
2. **首页文本相同**（PyMuPDF 提取第一页文本，空白归一化后比对）

## 目录约定

| 目录 | 说明 |
|------|------|
| （无 `input/`） | 源 PDF 在外部路径，由命令行或配置文件指定 |
| `output/` | 重复分析报告（`.txt`） |

## 依赖

- Python 3.8+
- `pymupdf`（`import fitz`），见仓库根目录 `requirements.txt`

## 用法

### 单次扫描

```bash
cd 20_PDF_Duplicate_Analyzer
python pdf_duplicate_analyzer.py --root "D:\References" --folders "2_5_1,CDP,IB,Protocol,RMP" --label "IND-References"
```

省略 `--folders` 时，自动扫描 `root` 下全部一级子目录。

### 多批次（配置文件）

复制 `jobs.example.json` 为 `jobs.json`（`jobs.json` 已 gitignore，勿提交含真实路径的文件），然后：

```bash
python pdf_duplicate_analyzer.py --config jobs.json
```

### 自定义输出目录

```bash
python pdf_duplicate_analyzer.py --root "D:\References" --output "output"
```

## 输出格式

`output/duplicate_report_<标签>.txt` 中，按子文件夹列出每个涉及重复的 PDF，并注明与哪个文件夹中的哪个文件重复，例如：

```
## 文件夹: IB-20260518

  - 文件: example.pdf
    重复类型: 首页内容重复
    → 与 [Protocol-20260518] 中的「example2.pdf」重复
```
