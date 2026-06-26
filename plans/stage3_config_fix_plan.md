# 阶段三：配置修复计划（P0+P1+P2 × 14 项）

> 用户确认：C — 修全部 P0+P1+P2（14 项，含 README 文档补充）

---

## P0 Critical（6 项）

### P0-1: `requirements.lite.txt` — 去除 BOM
- **问题**：第 1 行含 `﻿`（UTF-8 BOM），导致部分工具解析异常
- **修复**：重写文件内容（以纯 UTF-8 无 BOM 保存），逐行复制原内容
- **文件**：`requirements.lite.txt`

### P0-2: `requirements.txt` — 锁定 paddleocr / paddlepaddle 版本上限
- **问题**：`paddleocr>=2.6.0` + `paddlepaddle>=2.5.0` 无上限，Paddle 3.x API 不兼容
- **修复**：`paddleocr>=2.6.0,<3.0` + `paddlepaddle>=2.5.0,<3.0`
- **文件**：`requirements.txt`（第 51–52 行）

### P0-3: `requirements.txt` — 添加 pytest / ruff
- **问题**：`requirements.txt` 缺少 pytest（AGENTS.md 要求运行 pytest）和 ruff（AGENTS.md 4.1 节要求）
- **修复**：在 "开发/提交钩子" 区块添加 `pytest>=7.0.0` + `ruff>=0.15.0`
- **文件**：`requirements.txt`

### P0-4: `requirements-ci.txt` — 同步 pymupdf 版本
- **问题**：`pymupdf>=1.23.0`（ci）< `pymupdf>=1.27.0`（main），行为不一致
- **修复**：`pymupdf>=1.27.0`
- **文件**：`requirements-ci.txt`（第 11 行）

### P0-5: `requirements-ci.txt` — 添加 ruff
- **问题**：AGENTS.md 4.1 节要求 ruff，但 requirements-ci.txt 缺失
- **修复**：添加 `ruff>=0.15.0`
- **文件**：`requirements-ci.txt`

### P0-6: `requirements.lite.txt` — 同步 pypdf 版本
- **问题**：`pypdf>=4.0.0`（lite）< `pypdf>=6.0`（main）
- **修复**：`pypdf>=6.0`
- **文件**：`requirements.lite.txt`（第 22 行）

---

## P1 Important（4 项）

### P1-1: `requirements.txt` — 锁定 translators 版本上限
- **问题**：`translators>=5.9.2` 无上限，6.x 有 breaking changes
- **修复**：`translators>=5.9.2,<6.0`
- **文件**：`requirements.txt`（第 66 行）

### P1-2: `requirements.txt` — 为 paddlepaddle 添加平台限制
- **问题**：`paddlepaddle` 仅支持 Windows/Linux，macOS 会 pip install 失败
- **修复**：`paddlepaddle>=2.5.0,<3.0; platform_system != "Darwin"`
- **文件**：`requirements.txt`（第 52 行）

### P1-3: `.gitignore` — 添加 worktree 例外规则
- **问题**：`.worktrees/` 整目录屏蔽，导致 `.worktrees/wt-*/plans/*.md` 等合法文件无法跟踪
- **修复**：在 `.worktrees/` 行后添加：
  ```gitignore
  # 例外：plans/ 和文档可入库
  !.worktrees/*/plans/
  !.worktrees/*/*.md
  !.worktrees/*/CHANGELOG.md
  ```
- **文件**：`.gitignore`（第 163 行附近）

### P1-4: `.gitignore` — 收紧 `*verify_*.py` 模式
- **问题**：`*.md` 中出现 `verify_*.py` 时会误匹配
- **修复**：限定为根目录级别并排除 plans/：
  ```gitignore
  /*verify_*.py
  !plans/*verify_*.py
  ```
- **文件**：`.gitignore`（第 173 行）

---

## P2 Enhancement（4 项）

### P2-1: `README.md` — 补充 Python 版本要求
- **问题**：`Python 3.8+` 与 AGENTS.md 的 `Python 3.10+` 类型提示规范不一致
- **修复**：在 "环境要求" 区块补充：
  ```markdown
  - Python 3.10+（3.8 可运行但类型提示规范按 3.10+；部分模块依赖 3.10+ 语法糖）
  ```
- **文件**：`README.md`（第 9 行）

### P2-2: `README.md` — 补充 pymupdf/pypdf 兼容性说明
- **问题**：23_PDF_Threat_Analyzer 等模块对 pymupdf/pypdf 版本敏感
- **修复**：在 "注意事项" 或 "依赖安装" 段落补充：
  ```markdown
  **pymupdf / pypdf 版本注意**：≥1.27.0 / ≥6.0 为推荐版本；低于此版本可能导致 18/23 模块行为异常。
  ```
- **文件**：`README.md`（第 795 行附近）

### P2-3: `AGENTS.md` — 补充 plans/ 目录引用
- **问题**：`plans/` 目录已存在但 AGENTS.md 未提及
- **修复**：在 "2. 必读上下文" 表格添加一行：
  ```markdown
  | `plans/` | 阶段性设计文档（跨 worktree 重用） |
  ```
- **文件**：`AGENTS.md`（第 14–21 行表格区域）

### P2-4: `.pre-commit-config.yaml` — 同步 ruff 版本
- **问题**：`rev: v0.7.2` 与 `ruff>=0.15.0`（实际运行 v0.15.16）不一致
- **修复**：`rev: v0.15.0`（匹配实际安装版本）
- **文件**：`.pre-commit-config.yaml`（第 12 行）

---

## 执行顺序

1. `requirements.lite.txt` — BOM 去除 + pypdf 版本同步（P0-1 + P0-6）
2. `requirements.txt` — paddleocr/paddlepaddle 版本上限 + translators 上限 + pytest/ruff（P0-2 + P0-3 + P1-1 + P1-2）
3. `requirements-ci.txt` — pymupdf 同步 + ruff 添加（P0-4 + P0-5）
4. `.gitignore` — worktree 例外 + verify_*.py 收紧（P1-3 + P1-4）
5. `README.md` — Python 版本 + pymupdf/pypdf 兼容性说明（P2-1 + P2-2）
6. `AGENTS.md` — plans/ 目录补充（P2-3）
7. `.pre-commit-config.yaml` — ruff 版本同步（P2-4）

---

## 不修改项（留待后续迭代）

- `README.md` 中 "Python 3.8+" 不删除（保持向下兼容声明），仅补充 3.10+ 建议
- ruff format 检查发现的 59 个文件待 Stage 4 合并后统一处理
- P0×3 代码质量修复另开 worktree 执行

---

## 预期 Commit 信息

```
chore(config): sync dependency versions across requirements*.txt

- P0: fix BOM (requirements.lite.txt), lock paddleocr/paddlepaddle
  <3.0, add pytest/ruff to requirements.txt, sync pymupdf>=1.27.0
  and ruff across ci/lite txts, align pypdf>=6.0
- P1: lock translators<6.0, add paddlepaddle platform restriction,
  relax .gitignore for worktree plans/docs, tighten verify_*.py pattern
- P2: clarify Python 3.10+ in README.md, add pymupdf/pypdf compatibility
  note, reference plans/ in AGENTS.md, sync ruff pre-commit rev

Refs: stage3-config-audit
```
