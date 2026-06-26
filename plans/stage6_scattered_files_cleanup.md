# 阶段六：散落文件审计与整合

> 用户请求：遍历分析散落文件，归整到 scripts/ 或合并，删除无意义文件

---

## 扫描结果（根目录 20 个散落文件）

| 文件 | 大小 | 状态 | 处置 |
|------|------|------|------|
| `config.yaml` | 573B | 已忽略（.gitignore 52） | 保留（用户配置） |
| `list_py.py` | ? | VSCode 缓存 | 磁盘不存在 |
| `audit_py.py` | 3425B | 通用审计工具 | **移到 `scripts/audit/`** |
| `_audit_deps.py` | 4593B | 通用依赖审计 | **移到 `scripts/audit/`** |
| `_audit_cfg.py` | ? | 阶段四已清理 | 磁盘不存在 |
| `_scan_junk.py` | ? | 阶段四已清理 | 磁盘不存在 |
| `_delete_caches.py` | ? | 阶段四已清理 | 磁盘不存在 |
| `_patch_changelog.py` | ? | 阶段四已清理 | 磁盘不存在 |
| `_patch_gitignore.py` | ? | 阶段四已清理 | 磁盘不存在 |
| `_verify.py` | 570B | 一次性清点脚本 | **删除** |
| `_fix_excel.py` | 1933B | 一次性类型注解补丁 | **删除**（已应用） |
| `_fix_sae.py` | 1260B | 一次性 except 替换 | **删除**（已应用） |
| `_fix_lite_bom.py` | 1121B | 一次性 BOM 修复 | **删除**（已应用） |
| `_stage4_check.py` | 1396B | 阶段四辅助 | **删除** |
| `_stage4_cleanup.py` | 681B | 阶段四辅助 | **删除** |
| `_stage6_scan.py` | 1409B | 阶段六辅助 | **删除** |
| `_commit_msg.txt` | 915B | commit 消息文件 | **删除**（已使用） |
| `_stage6_scan_out.txt` | 1855B | 阶段六输出 | **删除** |
| `_audit_cfg_out.txt` | 387B | 过期审计输出 | **删除** |
| `delete_log.txt` | 198B | 阶段二删除日志 | **删除** |
| `deps_audit.txt` | 2783B | 过期依赖审计 | **删除** |
| `ruff_baseline.txt` | 2151B | 过期 ruff 基线 | **删除** |
| `scan_junk.txt` | 2835B | 过期 junk 扫描 | **删除** |
| `_conflict_chore.md` | 11308B | 合并冲突副本（GBK 乱码） | **删除** |
| `_conflict_v7.md` | 7400B | 合并冲突副本（GBK 乱码） | **删除** |

---

## 整合策略

### A. 有用脚本 → `scripts/audit/`

```
scripts/audit/
├── __init__.py
├── README.md         # 解释 audit/ 子目录用途
├── audit_py.py        # 鲁棒性深扫（从根 audit_py.py 移动）
└── audit_deps.py     # 依赖审计（从根 _audit_deps.py 移动，去掉下划线前缀）
```

- 去掉 `_` 前缀（"private/internal" 标记）因为这些工具是仓库级可复用的
- 创建 `scripts/audit/__init__.py` 标识这是一个独立子包
- 编写 `scripts/audit/README.md` 说明每个脚本的用途和调用方式

### B. 一次性脚本 → 删除

阶段一/三/四/六的 13 个临时脚本：
- `_fix_*.py` (3 个) — 已应用，价值为 0
- `_stage*_*.py` / `_stage*_out.txt` (5 个) — 阶段辅助
- `_verify.py` (1 个) — 一次性清点
- `_commit_msg.txt` (1 个) — 已使用
- 过期审计快照 (5 个) — 已过期
- 冲突副本 (2 个) — GBK 乱码

### C. md/txt 文件 → 评估

- `plans/pdf_sanitizer_v7_optimization.md` — 已入库（commit 62a56fd）
- `plans/stage3_config_fix_plan.md` — 已入库（commit 62a56fd）
- `_conflict_*.md` — GBK 乱码，删除

---

## 预期 commit 信息

```
chore(scripts): 整合审计工具到 scripts/audit/，删除 13 个临时脚本

- 移动 audit_py.py → scripts/audit/audit_py.py
- 移动 _audit_deps.py → scripts/audit/audit_deps.py
- 创建 scripts/audit/__init__.py + README.md
- 删除 13 个一次性临时脚本（_fix_*, _stage*_*, _verify.py 等）
- 删除过期审计快照（*.txt, _conflict_*.md）

Refs: stage6-cleanup
```
