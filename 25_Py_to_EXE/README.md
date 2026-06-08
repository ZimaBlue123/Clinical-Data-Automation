# 25_Py_to_EXE

Python 脚本转 Windows EXE（**单脚本模式 + 批量模式**）。基于 PyInstaller。

继承自 18_PDF_eCTD_Converter 单独打包验证过的做法（`sys.executable -m PyInstaller` + `collect_submodules` 解决动态子模块漏打包），并扩展为**整仓批量**能力。

## 双模式

### 1) 单脚本模式（旧用法，100% 向后兼容）

把要打包的 `.py` 放进 `input/` 目录，然后双击或命令行运行：

```bash
cd 25_Py_to_EXE
python py_to_exe.py
# 或显式指定
python py_to_exe.py --input "input/foo.py" --name foo --onefile
python py_to_exe.py --icon "input/app.ico"
```

旧 API 完全保留：`--input` `--output` `--name` `--onefile/--dir` `--noconsole` `--icon` `--clean-artifacts`/`--no-clean-artifacts`。

### 2) 批量模式（新核心能力）

**前提**：复制 `manifest.example.yaml` → `manifest.yaml`，按需增删模块。

```bash
# 走 manifest 清单（推荐）
python py_to_exe.py --batch --manifest ./manifest.yaml

# 自动扫描仓库各 NN_*/ 模块的主程序（启发式排除 util_/test_/lib_）
python py_to_exe.py --batch --auto-discover

# 只打指定模块（与上面两个任意一个配合，覆盖为子集）
python py_to_exe.py --batch --auto-discover --modules 18_PDF_eCTD_Converter,19_PDF_Merge

# 遇错即停（默认失败继续 + 末尾汇总）
python py_to_exe.py --batch --manifest ./manifest.yaml --fail-fast
```

## CLI 全表

| 参数 | 默认 | 说明 |
|------|------|------|
| **单脚本模式** | | |
| `--input` | `input/*.py`（首个） | 输入 .py 路径 |
| `--output` | `output/` | 输出目录 |
| `--name` | 文件 stem | EXE 文件名（不含扩展名） |
| `--onefile` / `--dir` | `--onefile` | 单文件 / 目录模式 |
| `--noconsole` | 关 | 隐藏控制台黑窗口 |
| `--icon` | 无 | .ico 图标路径 |
| `--clean-artifacts` / `--no-clean-artifacts` | `--clean-artifacts` | 清理 build/spec 中间产物 |
| **批量模式（加 `--batch` 启用）** | | |
| `--manifest <path>` | `./manifest.yaml` | manifest YAML 路径 |
| `--auto-discover` | 关 | 自动扫描仓库模块 |
| `--modules <a,b,c>` | 无 | 逗号分隔模块名列表（覆盖 manifest） |
| `--fail-fast` | 关 | 遇错即停（默认失败继续） |
| `--workers <N>` | `1` | 并发数（PyInstaller CPU 密集，建议 ≤ CPU 核数） |

## 输出组织

### 单脚本模式
```
25_PY_to_EXE/
├─ input/                    # 你的 .py
└─ output/
   └─ foo.exe                # 打包结果（默认在 output/ 根）
```

### 批量模式
```
25_PY_to_EXE/
└─ output/
   ├─ 18_PDF_eCTD_Converter/
   │  └─ pdf_ectd_converter.exe
   ├─ 19_PDF_Merge/
   │  └─ merge_pdf.exe
   ├─ 22_PDF_Duplicate_Analyzer/
   │  └─ pdf_duplicate_analyzer.exe
   ├─ ...（按模块分子目录）
   └─ _batch_report/
      ├─ batch_report_<时间戳>.json   # 完整结构化数据
      ├─ batch_report_<时间戳>.md     # 人可读汇总
      └─ batch_report_<时间戳>.xlsx   # Excel（4 sheet：汇总/全部/含警告/失败）
```

> 批量模式下，`build/` 和 `.spec` 中间产物默认自动清理；用 `--no-clean-artifacts`（单脚本）或 manifest `clean_artifacts: false` 关闭。

## 自动发现规则（`--auto-discover`）

```
NN_xxx/                # 必须以两位数字 + 下划线开头
├─ main.py             # 优先作为入口（兼容 13/21 等历史命名）
├─ foo.py              # 根目录顶层 .py 自动识别
├─ bar.py              # 根目录顶层 .py 自动识别
├─ util_xxx.py         # ✗ 排除（辅助工具）
├─ test_xxx.py         # ✗ 排除
├─ _internal.py        # ✗ 排除
└─ lib/                # ✗ 整目录排除
    └─ _core.py
```

`docs/script_roles.md` 定义的"主程序/核心库/辅助工具"规则一致。

## Manifest 编写（manifest.yaml）

```yaml
batch:
  default_pyinstaller:
    onefile: true
    console: true
    clean_artifacts: true
    collect_submodules: [pymupdf, openpyxl, pandas]  # 全局默认
    # excludes: [sklearn, scipy, matplotlib]          # 瘦身用
    icon: null
    datas: []
    hiddenimports: []

  modules:
    - module: 18_PDF_eCTD_Converter
      script: pdf_ectd_converter.py
      pyinstaller:
        # 这里可覆盖 default
        collect_submodules: [pymupdf, openpyxl, pandas]

    - module: 05_PPT_to_PDF
      script: ppt_to_pdf.py
      pyinstaller:
        hiddenimports: [win32com, win32com.client, pythoncom]
        collect_submodules: [win32com]
      warning: "Windows + PowerPoint 依赖；目标机需装 Office"
```

字段说明：
- `module`：模块目录名（必须在仓库根下）
- `script`：入口 .py 文件名（相对模块根）
- `output_name`：EXE 名（默认 = script stem）
- `pyinstaller`：覆盖 default；合并而非替换
- `warning`：不阻断打包，只在报告 + 日志里提示

## 报告字段

每条 BatchResult 包含：

| 字段 | 含义 |
|------|------|
| `module` | 18_PDF_eCTD_Converter |
| `script` | pdf_ectd_converter.py |
| `status` | `success` / `failed` / `skipped` |
| `output_exe` | 产物绝对路径 |
| `size_mb` | EXE 大小（MB） |
| `duration_sec` | 打包耗时（秒） |
| `exit_code` | PyInstaller 退出码 |
| `warning` | 兼容性/平台依赖提示 |
| `error` | 失败原因（截前 2000 字符） |
| `timestamp` | ISO 时间戳 |

Excel 报告 4 sheet：汇总 / 全部 / 含警告 / 失败。

## 失败重试

批量结束后，失败模块会打印类似：

```
重试命令：python py_to_exe.py --batch --auto-discover --modules 13_PDF_to_Excel_Rule_Extract,21_PDF_Watermark_Removal
```

直接复制运行即可只重打失败的。

## 跨路径使用（与 18_PDF_eCTD_Converter 一致）

`py_to_exe.py` 自身支持被 PyInstaller frozen 化（路径以 `sys.executable` 为基准），但本工具是开发者侧的元工具，**通常不打包成 EXE**——在开发环境直接用 Python 跑即可。

被它打出来的 EXE（位于 `25_PY_to_EXE/output/<模块名>/<exe名>.exe`）则与 18 模块一样，可以**单独复制**到任何路径/电脑运行（前提是目标平台兼容，见 manifest 的 `warning`）。

## 依赖

- `pyinstaller>=5.13.0`（见根 `requirements.txt`）
- 批量模式额外：`pyyaml`、`pandas`、`openpyxl`（仓库根 requirements 都有）

```bash
python -m pip install -r ../../requirements.txt
```

## 不做的事（明确边界）

- **不**给生成的 EXE 加版本号 / 数字签名 / 自定义图标（除非 manifest 显式指定）
- **不**做 GUI 界面（双击 EXE 走命令行风格）
- **不**强制把 25 自己打成 EXE（dogfooding 是可选的，不在本工具范围）
- **不**并发执行（PyInstaller CPU 密集，`--workers` 留作未来扩展，当前默认 1）
