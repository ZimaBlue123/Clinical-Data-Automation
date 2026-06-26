"""pzfx 文件改写（基于纯字符串定位，不破坏 XML 结构）。

关键设计
--------
- 不使用 ET 序列化整文件（避免命名空间 / 属性顺序被破坏）
- Subcolumn 排列规则（Prism 约定）：
    每个 YColumn 内 3 个 Subcolumn，按 (mid 列, up 列, lo 列) 顺序
    每个 Subcolumn 内 3 个 <d>，对应 3 个时间点
        (免前, 一免后2个月, 全免后1个月)
  YFormat="upper-lower-limits" 暗示此约定
- 改写流程：
    1) 用正则找全部 <Subcolumn>...</Subcolumn> 区间（按文本出现顺序）
    2) 按 (table_id, yi, si) 决定改写目标
    3) 从后往前替换（避免位置偏移）
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

# (免前, 一免后2个月, 全免后1个月) —— pzfx 亚列内 <d> 顺序
TIME_POINTS: tuple[str, str, str] = ("免前", "一免后2个月", "全免后1个月")
# 亚列类型顺序: si=0 -> mid, si=1 -> up, si=2 -> lo
SUBCOL_TYPE: tuple[str, str, str] = ("mid", "up", "lo")


@dataclass(frozen=True)
class PzfxValue:
    """一张亚列上的数值。按 3 个时间点各提供一条 (lo, mid, up)。

    Attributes:
        by_time: list[Triple]，长度=3，依次对应 (免前, 一免后2个月, 全免后1个月)
    """

    by_time: tuple  # type: ignore[type-arg]

    def __init__(self, *triples) -> None:
        # 允许灵活构造: PzfxValue(t1, t2, t3) 或 PzfxValue([t1, t2, t3])
        if len(triples) == 1 and isinstance(triples[0], (list, tuple)):
            arr = list(triples[0])
        else:
            arr = list(triples)
        if len(arr) != 3:
            raise ValueError(f"PzfxValue 需要 3 条 Triple，收到 {len(arr)}")
        # frozen + __init__ 套路：使用 object.__setattr__
        object.__setattr__(self, "by_time", tuple(arr))

    def by_type_at(self, ti: int, col_type: str) -> float:
        t = self.by_time[ti]
        if col_type == "mid":
            return t.mid
        if col_type == "up":
            return t.up
        if col_type == "lo":
            return t.lo
        raise ValueError(f"未知 col_type: {col_type}")


def fmt_num(v: float) -> str:
    """Prism 内部 d 节点的数字格式化：整数不带 .0，浮点保留 4 位有效数字。"""
    if v == int(v):
        return str(int(v))
    return f"{v:.4f}".rstrip("0").rstrip(".")


def rewrite_pzfx_data(
    src: Path | str,
    dst: Path | str,
    *,
    age_band: str,
    metric: str,
    table_id: str,
    new_values: dict[tuple[int, int], PzfxValue | None],
    logger=None,
) -> int:
    """改写一张 pzfx 表的 6 个 Subcolumn（2 YColumn × 3 Subcolumn）。

    Args:
        src:       源 pzfx
        dst:       输出 pzfx
        age_band:  期望的年龄段标题前缀（用于校验）
        metric:    期望的指标标题后缀（GMC / GMI / 阳转率）
        table_id:  目标 <Table ID="...">
        new_values: 键 = (ycol_idx, sub_idx)；值 = PzfxValue（None 表示保持原值）

    Returns:
        实际改写 Subcolumn 数量。
    """
    src_p = Path(src)
    dst_p = Path(dst)
    text = src_p.read_bytes().decode("utf-8")

    # 1) 找目标 Table 块
    m_table = re.search(
        rf'<Table\s+ID="{re.escape(table_id)}"[^>]*>(.*?)</Table>',
        text,
        re.DOTALL,
    )
    if not m_table:
        raise LookupError(f"{src_p.name} 中找不到 Table ID={table_id!r}")
    t_inner = m_table.group(1)
    t_abs_start = m_table.start(1)  # 内层起点
    m_table.end(1)

    # 2) 校验 Title
    title_m = re.search(r"<Title>([^<]*)</Title>", t_inner)
    title = title_m.group(1) if title_m else ""
    if not title.startswith(age_band) or not title.endswith(metric):
        raise ValueError(
            f"{src_p.name} Table {table_id} Title={title!r} 不匹配 {age_band}{metric}"
        )

    # 3) 定位 2 个 YColumn × 3 Subcolumn（按文本出现顺序）
    yc_iter = list(re.finditer(r"<YColumn\s[^>]*>.*?</YColumn>", t_inner, re.DOTALL))
    if len(yc_iter) < 2:
        raise ValueError(
            f"{src_p.name} Table {table_id} 仅有 {len(yc_iter)} 个 YColumn"
        )

    # 收集待替换区间（绝对位置）
    replacements: list[
        tuple[int, int, str]
    ] = []  # (abs_start, abs_end, new_subcolumn_xml)
    audit: list[str] = []

    for yi, yc_m in enumerate(yc_iter[:2]):
        yc_inner = yc_m.group(0)
        yc_abs = t_abs_start + yc_m.start()
        yc_title = re.search(r"<Title>([^<]*)</Title>", yc_inner).group(1)

        sub_iter = list(re.finditer(r"<Subcolumn>.*?</Subcolumn>", yc_inner, re.DOTALL))
        if len(sub_iter) != 3:
            raise ValueError(
                f"{src_p.name} Table {table_id}/{yc_title} Subcolumn={len(sub_iter)} 应=3"
            )

        for si, sub_m in enumerate(sub_iter):
            old_sub = sub_m.group(0)
            abs_start = yc_abs + sub_m.start()
            abs_end = yc_abs + sub_m.end()
            old_ds = re.findall(r"<d>([^<]*)</d>", old_sub)
            col_type = SUBCOL_TYPE[si]

            new_val = new_values.get((yi, si))
            if new_val is None:
                # None 表示不替换（保留原值）
                audit.append(
                    f"{table_id} | {yc_title} | yi={yi} si={si}({col_type}) | 保持 {old_ds}"
                )
                continue

            # 构造新 d 序列（按 3 个时间点）
            new_ds: list[str] = []
            for ti, tp in enumerate(TIME_POINTS):
                if metric == "GMI" and tp == "免前":
                    new_ds.append("1")
                    continue
                if metric == "阳转率" and tp == "免前":
                    new_ds.append("0")
                    continue
                new_ds.append(fmt_num(new_val.by_type_at(ti, col_type)))

            # 重建 Subcolumn 文本，保留缩进风格
            inner = old_sub[len("<Subcolumn>") : -len("</Subcolumn>")]
            pre = re.match(r"(\s*)<d", inner)
            indent = pre.group(1) if pre else "\n"
            tail_nl = "\n" if inner.endswith("\n") else ""
            new_d_lines = "".join(f"{indent}<d>{v}</d>" for v in new_ds)
            new_sub_text = f"<Subcolumn>{new_d_lines}{tail_nl}</Subcolumn>"

            replacements.append((abs_start, abs_end, new_sub_text))
            audit.append(
                f"{table_id} | {yc_title} | yi={yi} si={si}({col_type}) | "
                f"old={old_ds} → new={new_ds}"
            )

    # 4) 从后往前替换
    replacements.sort(key=lambda x: x[0], reverse=True)
    new_text = text
    for s, e, sub in replacements:
        new_text = new_text[:s] + sub + new_text[e:]

    # 5) 写盘
    dst_p.parent.mkdir(parents=True, exist_ok=True)
    dst_p.write_bytes(new_text.encode("utf-8"))

    if logger is not None:
        for line in audit:
            logger.info("rewrite %s", line)
    return len(replacements)
