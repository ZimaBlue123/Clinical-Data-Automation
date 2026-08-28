# 34_Image_to_PDF

批量将图片转换为 PDF 文件。

## 功能特性

- **单图转单 PDF**：将目录中的每张图片分别转换为单独的 PDF 文件。
- **多图合并为单 PDF**：将目录中的所有图片按文件名排序后，合并生成一个多页 PDF 文件。
- **高兼容性**：自动处理透明背景（RGBA/PNG）和索引颜色模式图片，将其安全转为 RGB 格式，防止生成 PDF 时报错。
- **格式支持**：支持 `.jpg`, `.jpeg`, `.png`, `.bmp`, `.webp`, `.tiff` 等常见图片格式。

## 依赖说明

依赖已在根目录 `requirements.txt` 中声明。本模块使用 `Pillow`：
```bash
pip install Pillow
```

## 用法说明

进入当前模块目录：
```bash
cd 34_Image_to_PDF
```

### 1. 默认模式（每张图单独生成一个 PDF）
将需要转换的图片放入 `input/` 目录，然后执行：
```bash
python image_to_pdf.py
```
转换结果将保存在 `output/` 目录中。

### 2. 合并模式（所有图片合并为一个 PDF）
```bash
python image_to_pdf.py --merge
```
生成的默认文件名为 `output/merged_output.pdf`。

### 3. 指定输入/输出与覆盖现有文件
```bash
# 指定输入输出目录并覆盖同名文件，同时自定义合并后的文件名
python image_to_pdf.py --input "D:\MyImages" --output "D:\MyPDFs" --merge --merge-name "clinical_report.pdf" --overwrite
```
