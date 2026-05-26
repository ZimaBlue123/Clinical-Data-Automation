# 10. Word 样式清理工具（`10_Word_Style_Cleaner`）

用途：

- **清理未使用的自定义样式**（保留 Word 内置样式不动）
- 同时对**保留下来的自定义样式做命名规范化**：将常见样式按中文标签统一（如“标题1/标题2/正文/表格”等），并把同类型样式的阿拉伯数字编号理顺为连续序列

## 快速开始

将待处理的 `.docx` 放入 `input/`，运行：

```bash
cd 10_Word_Style_Cleaner
python word_style_cleaner.py --input "./input" --output "./output"
```

## 参数

- `--input`：输入文件夹（默认 `./input`）
- `--output`：输出文件夹（默认 `./output`）
- `--recursive/--no-recursive`：是否递归处理子目录（默认递归）
- `--overwrite`：输出已存在时覆盖（默认不覆盖）
- `--suffix`：输出文件名后缀（默认 `_styles_cleaned`）
- `--report-dir`：报告输出目录（默认与输出文件同目录）
- `--no-rename-styles`：仅清理未用自定义样式，不做样式命名规范化

## 输出

- 清理后的 docx：`<原名>_styles_cleaned.docx`
- 清理报告：`<原名>_styles_cleaned_report.txt`（包含“删除了哪些自定义样式”“哪些自定义样式被重命名”）

