# AGENTS.md — AI Agent 协作指引

> 本文件供 AI 编程助手（如 Mavis / Claude Code / Cursor / GitHub Copilot Workspace 等）
> 在本仓库工作时遵循的工程规范与上下文速查。

## 1. 项目性质

- **Monorepo** 临床数据自动化工具集：34 个编号子项目（`01_Excel_Charts/` ~ `33_SAE_Extractor/`）+ 共享 `src/` + `tests/`
  - ⚠️ 2026-06-26：原 08_Word_Tables_to_Excel 被 `08_Word_Tables_to_Graphpad` 替换；原 09_Word_All_Tables_to_Excel 改名为 `09_Word_Tables_to_Excel`
- 单一 `requirements.txt` / `requirements.lite.txt` / `requirements-ci.txt` 覆盖全部模块
- **不要** 跨子项目做隐式耦合（每个子项目应可独立运行）

## 2. 必读上下文

| 文件 | 必读理由 |
|---|---|
| `README.md` | 总览 |
| `SECURITY.md` | 安全策略（涉密数据防泄露） |
| `.gitignore` | 哪些路径**不能**动（input/output/Template 都不入库） |
| `requirements*.txt` | 真实依赖；不要凭空 `pip install` 未经声明的包 |
| `CHANGELOG.md` | 看历史变更风格，保持一致 |

## 3. 子项目结构约定

```
XX_模块名/
├── <main>.py           # CLI 入口（通常有 if __name__ == "__main__"）
├── <util_*.py>        # 工具函数
├── lib/                # 共享库（含 __init__.py）
├── input/              # 用户上传的原始数据（不提交）
├── output/             # 生成结果（不提交）
├── tests/              # 单元测试（子项目内或仓库根 tests/）
└── README.md           # 子项目使用说明（如果有）
```

## 4. 核心工程规范

### 4.1 静态检查

- **必跑**：`ruff check .`（项目根 `requirements-ci.txt` 已声明 ruff）
- 已用规则集：`E, F, W, B, UP, S, N, SIM, RET, ARG, PTH, ERA` + `PLR`
- 复杂函数（PLR0912/0913/0915）允许 `# noqa` + TODO，下个迭代重构
- Magic value 允许在 `_JOURNAL_MASTHEAD_COMPRESSED` 等黑名单内出现（文件级 `ruff: noqa: PLR2004`）

### 4.2 异常处理

- **禁止** `try: ... except: pass`（静默吞异常）
- 必须 `except Exception as e: logger.warning(...)` 或 `raise X from e` 保留异常链
- 探测库可用性（`import fitz` 等）允许 `except: pass` —— 用文件级 `ruff: noqa: S110, S112, SIM105`

### 4.3 类型提示

- Python 3.10+：用 `X | Y`（PEP 604）+ `list[X]`（PEP 585）
- 文件首行加 `from __future__ import annotations` 兼容低版本
- 公共 API 必须标注参数和返回类型

### 4.4 日志

- 模式：`logger.info("action=xxx key=%s", value)`（结构化键值对）
- 配置：模块级 `logger = logging.getLogger(__name__)`
- 不要 print() 调试残留到 commit

### 4.5 Git 提交

- 严格 **Conventional Commits**：
  - `feat:` 新功能
  - `fix:` Bug 修复
  - `refactor:` 重构
  - `perf:` 性能
  - `docs:` 文档
  - `test:` 测试
  - `chore:` 构建/工具/依赖
  - `style:` 格式
  - `ci:` CI
- 影响范围写进 body：`Refs: 17_PDF_Title_Renamer`
- 严禁在 main 分支直接 commit；用 worktree 流程

## 5. Worktree 流程（强制）

```bash
# 创建 worktree
git worktree add -b <type>/<scope> .worktrees/wt-<id> main

# 在 worktree 内工作
cd .worktrees/wt-<id>
# 改代码...

# 完成后
git add -A
git commit -m "feat: ..."
git push origin <branch>
# 在 GitHub 上开 PR

# 清理
git worktree remove .worktrees/wt-<id>
```

**禁止**直接在主 checkout 改文件（worktree toggle 必须保持 ON）。

## 6. 域知识速查（临床研究）

- **IB / CSR / SAP**：研究性新药临床试验的核心文档
- **AESI** (Adverse Event of Special Interest)：特别关注不良事件
- **剂量 / 免疫程序**：医学合理性审查关注点
- **eCTD** (Electronic Common Technical Document)：药品注册申报标准格式
- **Suvoda / Medidata / Oracle**: 临床试验常用 EDC 系统

## 7. 禁止行为

- ❌ 提交 `input/`、`output/`、`config.yaml`、`.env`、`*.pdf`、患者数据
- ❌ 在 monorepo 根做破坏性 `rm -rf` 操作（用 `mavis-trash` 走回收站）
- ❌ 修改 git 历史（force push、rebase 已推分支）
- ❌ 改 `requirements*.txt` 不在 commit body 说明业务理由
- ❌ 跨子项目耦合（一个子项目 import 另一个子项目）

## 8. 沟通风格

- 用户角色：医学背景 / 临床研究方向
- 偏好：中文沟通 + 严格审计式审查 + 风险信号优先
- 技术细节可以中文，但代码 / 命令保持英文
