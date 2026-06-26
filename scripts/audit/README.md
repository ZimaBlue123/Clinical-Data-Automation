# scripts/audit/ — 仓库级审计工具

本目录提供**可复用**的静态分析脚本，用于在重构/PR 之前对代码与依赖做系统性扫描。

## 工具列表

| 脚本 | 用途 | 调用方式 |
|------|------|---------|
| [`audit_py.py`](audit_py.py) | 鲁棒性深扫：标记裸 except、宽 except、未关闭文件、print() 残留、TODO/FIXME 等 | `python -m scripts.audit.audit_py` |
| [`audit_deps.py`](audit_deps.py) | 依赖一致性审计：解析 `requirements*.txt` 与代码 `import` 语句，识别漏声明/多声明 | `python -m scripts.audit.audit_deps` |

## 设计原则

- **不修改文件**：所有审计脚本只读，产生报告
- **零硬编码路径**：使用 `pathlib.Path(".")` 作为 ROOT，可任意目录运行
- **SKIP `.worktrees/`**：避免误扫 worktree 中的副本
- **输出可重定向**：报告默认 print，可重定向到 `*.txt` 留档

## 适用范围

- 大型重构前风险点定位
- 依赖声明审计（防止漏掉隐式 import）
- CI 失败后的快速根因分类

## 限制

- 静态分析无法替代运行时测试
- 误报需人工二次确认（每个风险规则都有豁免清单）
