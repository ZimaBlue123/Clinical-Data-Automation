# 10. Word 样式清理工具（`10_Word_Style_Cleaner`）

用途：

- **清理未使用的自定义样式**（保留 Word 内置样式不动）
- 对**保留下来的自定义样式做命名规范化**：将常见样式按中文标签统一（如“标题1/标题2/正文/表格/图标题”等），并把同类型样式的阿拉伯数字编号理顺为连续序列

## 快速开始

将待处理的 `.docx` 放入 `input/`，运行：

```bash
cd 10_Word_Style_Cleaner
python word_style_cleaner.py --input "./input" --output "./output" --overwrite
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
- 清理报告：`<原名>_styles_cleaned_report.txt`（删除/重命名明细、孤儿引用修复统计）

## 健壮性说明

脚本在写入 `styles.xml` 时会**保留原始根标签上的全部 `xmlns` 声明与 `mc:Ignorable`**，避免 Word 打开输出文件时出现“不可读取内容 / 需修复后另存为”的弹窗。

其他保护措施：

- **样式依赖闭包**：删除前会沿 `basedOn` / `link` / `next` 保留父级样式，避免继承链断裂
- **孤儿引用清理**：若文档 XML 仍引用已删 `styleId`，会自动移除对应 `pStyle` / `rStyle` 等节点
- **原子写入**：先写临时文件再替换，降低输出 docx 写一半损坏的风险
- **批处理汇总**：逐文件输出 `[处理]` / `[跳过]` 状态；有失败时进程退出码为 `1`

## 注意事项

- 仅处理 `.docx`（不支持旧版 `.doc`）
- 不会修改 `styleId`，只更新样式显示名称（`w:name/@w:val`）并删除未用自定义样式定义
- 输出文件默认不覆盖；批处理请加 `--overwrite`
