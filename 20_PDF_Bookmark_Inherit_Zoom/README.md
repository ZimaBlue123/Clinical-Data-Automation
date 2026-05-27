# 19_PDF_Bookmark_Inherit_Zoom

PDF 书签「承前缩放」批处理：使用 PyMuPDF 重写目录（TOC），为书签注入 XYZ 目标且 `zoom=0`，使阅读器在跳转时保持当前缩放比例；无书签时仍做垃圾回收与流压缩。加密 PDF 会跳过。

## 快速开始

```bash
cd 19_PDF_Bookmark_Inherit_Zoom
# 将 PDF 放入 input/ 后执行
python pdf_bookmark_inherit_zoom.py
```

输出写入 `output/`，与源文件同名。

## 依赖

根目录 `requirements.txt` 中的 `pymupdf`（`import fitz`）。

## 说明

详细参数与可选路径见仓库根目录 `README.md` 中「23. PDF 书签承前缩放」一节。
