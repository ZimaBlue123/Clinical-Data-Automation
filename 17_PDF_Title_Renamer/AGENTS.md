# 17_PDF_Title_Renamer 模块长期记忆规则

## 1. 长标题安全截断原则
在修改或重构 `pdf_sanitizer.py` 时，**切勿恢复按单词数一刀切**的暴力截断逻辑。
必须保持 `max_words` 放宽（例如默认 40 词）并严格使用 `max_chars`（默认 200 字符左右）来做系统安全截断，防止 Windows 文件路径过长引发错误。

## 2. 智能重命名工作流 (AI Agent Workflow)
当用户通过对话界面下达类似于“帮我精炼并重命名长标题”或“帮我处理一下文献名”的命令时，AI Agent 必须严格按照以下三步执行协同逻辑：

### 第一步：生成草案
执行以下命令以提取完整文本并生成草案（不执行物理重命名）：
```bash
python pdf_sanitizer.py --export-plan rename_plan.json
```

### 第二步：AI 思考与改写
1. 读取生成的 `rename_plan.json`。
2. 针对每个条目的 `raw_title`（原始长标题），利用 AI 语义理解能力，提取出核心的医学/学术信息。
3. 将其浓缩为**简短、专业、一目了然**的英文短标题。
4. **⚠️ 年份强制拼接**：从 JSON 计划中读取 `year` 字段，并在精炼完成的标题末尾**固定拼接连字符和年份**（格式为 `-{year}`）。如果 `year` 为 `XXXX`，则拼接 `-XXXX`；如果原文件原本就带有正确的年份，必须保留。
5. 将改写后包含年份的最终名称，覆写回 JSON 的 `proposed_name` 字段，并保存。

### 第三步：执行物理重命名
当草案更新完毕，且用户无异议（或默认已确认）时，执行：
```bash
python pdf_sanitizer.py --apply-plan rename_plan.json
```
执行完毕后，原 PDF 文件将被重命名并附带正确的年份后缀。
