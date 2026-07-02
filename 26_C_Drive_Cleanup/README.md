# 26_C_Drive_Cleanup

C 盘垃圾文件 / 空文件 / 无用目录清理器。

## 功能矩阵

| 维度 | 等级 | 说明 |
| --- | --- | --- |
| 文件级（空文件、无用扩展名） | `safe` | 默认 `--delete` 即可清理 |
| 文件级（7+ 天未动的临时文件） | `safe` | 默认 `--delete` 即可清理 |
| 空目录 | `safe` | `--remove-empty-dirs` 时清理 |
| 孤儿 Python 安装 (`C:\Python314\`、`C:\Anaconda3\`…) | `review` | 需要 `--include-review`，默认移动到隔离目录 |
| 孤儿 Node.js 安装 (`C:\node-v18.17.0-win-x64\`) | `review` | 同上 |
| 用户级大体积缓存 (`pip-cache` / `npm-cache` / `Yarn\Cache` / `nuget`) | `review` | 同上 |
| IDE 残留 (`Code - Insiders` / `Cursor` / `JetBrains` 体系) | `review` | 同上 |
| Windows 更新缓存 (`SoftwareDistribution\Download`、`$Windows.~BT`) | `dangerous` | **永不自动删除**，仅出现在报告里 |

## 使用方法

### 仅扫描（默认安全模式）

```bash
python c_drive_cleanup.py
```

会在 `output/` 下生成：

- `cleanup_report.csv`  — 文件级候选
- `directory_report.csv` — 目录级候选（含可回收大小、风险等级）

### 只清理 safe 级（空文件、无用扩展、空目录）

```bash
python c_drive_cleanup.py --delete --remove-empty-dirs
```

### 处理 review 级（孤儿 Python、IDE 缓存）

```bash
python c_drive_cleanup.py --delete --include-review
```

默认会把 review 候选**移动**到 `output/_quarantine_<时间戳>/` 目录而不是直接删除，方便核对。
加 `--no-quarantine` 则直接 `rmtree`，**不可恢复**。

### 只看 Python 孤儿

```bash
python c_drive_cleanup.py --python-only
```

典型输出：

```
Top 15 目录候选（按大小降序）:
  [review  ]    2.34 GB  C:\Python314
  [review  ]    1.10 GB  C:\Python312-32
  [review  ]  812.55 MB  C:\Anaconda3
  [dangerous]    5.40 GB  C:\Windows\SoftwareDistribution\Download
```

### 自定义根目录

```text
# input/targets.txt，每行一个
C:\Temp
C:\Users\me\AppData\Local\Temp
D:\LegacyAppCache
```

或者命令行 `--targets "C:\\Temp" "C:\\Python314"`。

## 关键参数

| 参数 | 默认值 | 含义 |
| --- | --- | --- |
| `--days` | 7 | 文件级仅处理 mtime 早于 N 天的 |
| `--min-dir-age-days` | 30 | 目录级最近 N 天未改动的才纳入 review |
| `--top` | 15 | 控制台展示 Top N 候选 |
| `--delete` | False | 是否执行删除 |
| `--include-review` | False | 是否处理 review 级 |
| `--include-dangerous` | False | 是否处理 dangerous 级（强烈不建议） |
| `--no-quarantine` | False | review 级是否跳过隔离，直接 rmtree |
| `--python-only` | False | 目录级仅扫描 Python 孤儿 |

## 安全保证

- 默认 `safe` 模式只读不改。
- `review` 级默认**移动而非删除**（可手动恢复）。
- `dangerous` 级永不自动处理，需要人工进控制面板 / `DISM` 操作。
- 不需要管理员权限就能扫描；`--delete` 时建议管理员运行，否则部分系统目录会跳过。
- 不会触碰 `C:\Windows\System32\`、`C:\Program Files\` 下的非明确白名单目录。

## 与旧版差异

| 旧版 | 新版 |
| --- | --- |
| 只能清单文件（按扩展名/空文件） | 增加**目录级**候选，可识别 `C:\Python314` 等孤儿安装 |
| 没有大小估算 | 输出 `format_size` 汇总 + Top N |
| `--delete` 直接删 | `--delete` 默认只动 safe；review 走隔离目录 |
| 无风险分层 | 三级风险：safe / review / dangerous |
| 不读注册表 | 通过 Windows 注册表获取已注册 Python 列表，避免误删在用版本 |