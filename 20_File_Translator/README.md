# 20_File_Translator

多格式双向翻译模块（translators 首选，DeepL + LibreTranslate 兜底）。

## 功能

- 批量处理 `input/` 下的 `.xlsx` / `.csv` / `.docx` / `.doc` / `.pdf`
- Excel/CSV 仅新增双语列；Word/PDF 生成同格式翻译副本
- 默认翻译列：`Term`、`SOC`、`Comment`、`PT Name`、`SOC Name`
- 支持方向：`en2zh` 与 `zh2en`
- 默认翻译引擎：`auto`（translators -> DeepL -> LibreTranslate -> 原文）
- 支持术语词典优先替换（clinical 术语一致性）
- 支持 JSON 持久化缓存 + 文件级并发
- PDF 支持 `overlay` 与 `bilingual-text-layer` 两种模式
- 并发下自动规避 COM 线程风险（自动降级 `openpyxl`）
- 支持 API 连通自检

> 模块目录已统一为 `20_File_Translator`，主脚本为 `file_translator.py`。

## 快速开始

```bash
cd 20_File_Translator
python file_translator.py --self-test
python file_translator.py
```

输出到：`output/*_en2zh.*` 或 `output/*_zh2en.*`

## 常用参数

```bash
# 批量目录（默认 input -> output）
python file_translator.py

# 指定列（逗号分隔）
python file_translator.py --columns "Term,SOC,Comment"

# API 自检
python file_translator.py --self-test

# 强制使用 translators / DeepL / LibreTranslate
python file_translator.py --provider tsfree --ts-engine bing
python file_translator.py --provider deepl
python file_translator.py --provider libre

# 中译英
python file_translator.py --direction zh2en

# PDF 仅导出双语文本层（避免覆盖重绘带来的字体偏差）
python file_translator.py --pdf-mode bilingual-text-layer

# 并发 + 持久化缓存
python file_translator.py --max-workers 3 --cache-file "output/translation_cache.json"

# 禁用持久化缓存
python file_translator.py --no-cache

# 术语词典优先替换
python file_translator.py --glossary "input/glossary.json"

# 指定写回引擎（Windows 推荐 com，跨平台可用 openpyxl）
python file_translator.py --engine com
python file_translator.py --engine openpyxl

# Word 扩展对象翻译（需 COM 可用）
python file_translator.py --word-include-textboxes --word-include-footnotes

# 跳过预检强制执行（不建议，调试时可用）
python file_translator.py --skip-preflight
```

> 并发说明：`--max-workers > 1` 时若使用到 `com` 引擎，会自动切换到 `openpyxl`，避免 COM 线程安全问题。
>
> PDF 文本层说明：`--pdf-mode bilingual-text-layer` 会保留原 PDF，并额外输出同名 `*.bilingual.txt` 双语对照文本层。
>
> 缓存说明：翻译缓存采用“临时文件 + 原子替换”写回，避免进程中断导致 JSON 缓存损坏。

## 环境变量（可选）

- `TS_TRANSLATOR_ENGINE`（`bing`/`google` 等，默认 `bing`）
- `TS_SLEEP_SECONDS`（默认 `0.5`，防风控节流）
- `TRANSLATION_DIRECTION`（`en2zh`/`zh2en`，默认 `en2zh`）
- `PDF_TRANSLATE_MODE`（`overlay`/`bilingual-text-layer`）
- `MAX_WORKERS`（并发文件数）
- `TRANSLATION_CACHE_FILE`（持久化缓存路径）
- `NO_CACHE`（可通过命令行 `--no-cache` 临时关闭缓存）
- `GLOSSARY_FILE`（术语词典路径，json/csv/tsv）
- `DEEPL_API_KEY`（推荐，DeepL Free 每月 50 万字符）
- `DEEPL_API_BASE`（默认 `https://api-free.deepl.com/v2`）
- `LIBRETRANSLATE_API_BASE`（默认 `https://libretranslate.com`）
- `LIBRETRANSLATE_API_KEY`（若实例要求）
- `TRANSLATOR_PROVIDER`（`auto`/`tsfree`/`deepl`/`libre`，默认 `auto`）

建议先复制 `.env.example` 为 `.env` 并填值，然后直接运行脚本。  
脚本启动会打印脱敏后的配置摘要（不会泄露完整密钥）。

## 术语词典示例

`input/glossary.json`:

```json
{
  "adverse event": "不良事件",
  "serious adverse event": "严重不良事件",
  "MedDRA": "MedDRA"
}
```

