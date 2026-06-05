# Python 版本与环境选择

本仓库已在 **Python 3.10.11** 下完整测试通过。多数 Windows 环境会同时存在多个 Python 解释器，本文档说明如何在不同环境下选对版本。

## 1. 现状（2026-06）

本仓库最近一次完整自检（`23_PDF_Threat_Analyzer --self-check`）通过的环境：

| 项 | 值 |
|----|---|
| Python | **3.10.11**（`C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe`） |
| 操作系统 | Windows 10/11（PowerShell 5.1） |
| 关键包 | `pypdf` 6.10.2 · `pymupdf` 1.27.2.3 · `fonttools` 4.62.1 · `PySocks` 1.7.1 · `pyinstaller` 6.20.0 · `mmap`（标准库，零拷贝扫描） |

## 2. 常见的多 Python 共存场景

```
C:\msys64\ucrt64\bin\python.exe               ← 3.14.4（MSYS2 / UCRT64，PATH 默认）
C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe   ← 3.10.11（有依赖，✅）
C:\Windows\py.exe                              ← py launcher（推荐用 py -3.10）
```

`pip` 默认指向 3.10 的 site-packages，但 PowerShell 里 `python` 默认指向 3.14 → 装得对，跑不了。

## 3. 验证当前用哪个 Python

```powershell
# 1) python 实际指向
python -c "import sys; print(sys.executable)"
# C:\msys64\ucrt64\bin\python.exe   ← 如果是这条，说明 PATH 被 3.14 占了

# 2) pip 装在哪个 site-packages
pip --version
# pip 26.1.1 from ...\Python310\lib\site-packages\pip (python 3.10)

# 3) 3.10 路径下的核心包
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" -c "import pypdf, fitz; print(pypdf.__version__, fitz.__version__)"
# 6.10.2 1.27.2.3
```

## 4. 选对 Python 的 3 种方式

### 方式 A：用 `py` 启动器（推荐）

`py.exe` 是 Windows Python Launcher，按版本号选解释器：

```powershell
py -3.10 23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
py -3.10 23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --input "D:\inbox" --sanitize
```

列出本机所有 Python：
```powershell
py -0
# -V:3.12 *        C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe
# -V:3.10 *        C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe
# -V:3.14          C:\msys64\ucrt64\bin\python.exe
```

### 方式 B：完整路径

```powershell
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" `
    23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
```

### 方式 C：临时改 PATH（会话级，不改注册表）

```powershell
$env:Path = "C:\Users\Administrator\AppData\Local\Programs\Python\Python310;" + $env:Path
python 23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
# 注：仅当前 PowerShell 窗口有效，关闭即失效
```

## 5. 安装依赖

```powershell
# 装完整版（包含 paddleocr / paddlepaddle，1.5GB+ 较大）
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" -m pip install -r requirements.txt

# 装精简版（跳过 paddleocr / paddlepaddle，节省 1.5GB+）
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" -m pip install -r requirements.lite.txt
```

精简版来自根目录 `requirements.lite.txt`，与 `requirements.txt` 的区别仅在于去掉了：

```text
paddleocr>=2.6.0
paddlepaddle>=2.5.0
```

这两个包是 **16_PPTX_PDF_to_PPT** 模块做表格识别的可选依赖，不用 OCR 表格识别可省。

## 6. 故障排查

### 症状 A：`ModuleNotFoundError: No module named 'pypdf'`

**根因**：用错 Python（PATH 默认指向 MSYS2 3.14 等无依赖的版本）。

**修复**：
```powershell
# 显式用 3.10
py -3.10 23_PDF_Threat_Analyzer\pdf_threat_analyzer.py --self-check
```

### 症状 B：`pip install` 报 `ERROR: Could not open requirements file`

**根因**：在错的目录下跑（典型是 `C:\Users\Administrator>` 而不是项目根）。

**修复**：
```powershell
cd "E:\Cursor Project\1-Clinical Data Automation"
& "C:\Users\...\Python310\python.exe" -m pip install -r requirements.lite.txt
```

### 症状 C：自检跑到 `sanitize` 步骤失败

**原因**：测试 PDF 是手工伪 PDF（旧版代码）→ 已修复（2026-06 commit），请 `git pull` 拉到最新版。

### 症状 D：`pymupdf` 装不上 / `fitz` 导入失败

Windows 上 `pymupdf` wheel 自带预编译二进制，正常 `pip install pymupdf` 即可。如果失败：

```powershell
# 1) 升级 pip
& "C:\...\Python310\python.exe" -m pip install --upgrade pip

# 2) 单独装 pymupdf（看完整错误）
& "C:\...\Python310\python.exe" -m pip install pymupdf
```

## 7. 可选系统级依赖（不是 pip 包）

| 工具 | 用途 | 安装 | 缺失影响 |
|------|------|------|---------|
| **qpdf** | PDF 线性化校验（23 模块 sanitize 后置） | `choco install qpdf` 或 [官网下载](https://qpdf.sourceforge.io/) | graceful skip，sanitize 仍可正常输出 |
| **Tesseract** | OCR 引擎（12/16/17/18/19/21/32 等模块） | `choco install tesseract` | 这些模块的 OCR 路径不可用，但非 OCR 流程不受影响 |
| **Poppler (pdftoppm)** | PDF 转图片（16/17/32 等模块） | `choco install poppler` | pdf2image 不可用 |
| **mutool** | mupdf CLI（少数场景） | mupdf 工具包自带 | 同 qpdf，graceful skip |

`choco` 是 Windows 包管理器（Chocolatey），未装可用 `winget` 或手动下载。

## 8. 版本兼容矩阵

| 包 | 最低版本 | 仓库实际使用 | 备注 |
|----|----------|--------------|------|
| Python | 3.8+ | 3.10.11 | 类型注解从 3.8 兼容 |
| pypdf | 4.0+ | 6.10.2 | 6.x 改进了加密 PDF 处理 |
| pymupdf | 1.23+ | 1.27.2.3 | 1.24+ 改进了 `page.delete_link` 行为 |
| mmap | 标准库 | 内置 | 23_PDF_Threat_Analyzer 零拷贝扫描用，无需 pip |
| openpyxl | 3.1+ | 3.1.5 | 仅 1.x 模块需要 |
| pdfplumber | 0.10+ | 0.11.9 | 13_PDF_to_Excel_Rule_Extract |
| requests | 2.31+ | 2.33.1 | 24 / 32 / 29 等 |
| translators | 5.9+ | 6.0.4 | 24_File_Translator |

## 9. 升级 Python 时的注意

- 升级前先 `pip freeze > requirements_freeze.txt` 备份
- 升级后 `pip install -r requirements_freeze.txt` 恢复
- `pymupdf` 新版本偶尔会改 API，看 [CHANGELOG](https://pymupdf.readthedocs.io/en/latest/changes.html)
- 升级 Python 主版本（如 3.10 → 3.12）可能触发 `typing` 行为变化，跑 `python -m compileall .` 全仓编译一遍验证

## 10. 反馈

如果发现新的环境兼容问题，请更新本文件或提交 PR。
