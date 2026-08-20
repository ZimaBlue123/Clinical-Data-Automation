"""抗体数据提取 + 跨抗体映射。

从 docx 临床小结正文中识别 gE / VZV 抗体在 4 个年龄段 × 3 个指标
（GMC / GMI / 阳转率）× 3 个时间点（免前 / 一免后2个月 / 全免后1个月）
× 2 个组别（试验组 / 阳性对照组）的数值（lo, mid, up）。

设计点
------
- 输入：docx 解析后的段落列表
- 输出：AntibodyDataset 数据类，包含每个 (age_band, metric, time_point, group) 的 (lo, mid, up)
- 数值提取：依赖 docx 正文中形如下面的"对照逻辑"句式：
    "试验组和阳性对照组1免前抗gE抗原特异性血清抗体GMC（95%CI）分别为
     1103.5（905.2，1345.2）mIU/mL、935.7（761.4，1150.0）mIU/mL"
  解析后：trial=(905.2, 1103.5, 1345.2)  control=(761.4, 935.7, 1150.0)
- 阳转率：括号内是百分数（如 "100.00%（94.73%，100.00%）"），统一按 0-100 数值处理
- GMI：免前 = 1（无变化），docx 不显式报告；如果脚本里出现 "免前 ... GMI" 视为常量
- 阳转率免前 = 0（基线无阳转）

本模块本身不依赖 docx 文件路径，便于被 `poc_replicate.py` 和单元测试复用。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# 年龄段 4 档（docx 正文用 ~ / 以上 两种分隔符，模块内部统一为 "40-49岁" 等）
AGE_BANDS: tuple[str, ...] = ("40-49岁", "50-59岁", "60岁以上", "50岁以上")
METRICS: tuple[str, ...] = ("GMC", "GMI", "阳转率")
TIME_POINTS: tuple[str, ...] = ("免前", "一免后2个月", "全免后1个月")
ANTIBODIES: tuple[str, ...] = ("gE", "VZV")

# docx 中的年龄段行标题（"40~49岁" / "50~59岁" / "60岁及以上" / "50岁及以上"）→ 统一
_AGE_NORM = {
    "40~49岁": "40-49岁",
    "50~59岁": "50-59岁",
    "60岁及以上": "60岁以上",
    "50岁及以上": "50岁以上",
    "40-49岁": "40-49岁",
    "50-59岁": "50-59岁",
    "60岁以上": "60岁以上",
    "50岁以上": "50岁以上",
}

# 数值 95%CI 模式：3 个数（lo, mid, up），lo ≤ mid ≤ up
_95CI_RE = re.compile(r"([0-9]+(?:\.[0-9]+)?)\s*[（(]\s*([0-9]+(?:\.[0-9]+)?)\s*[,，]\s*([0-9]+(?:\.[0-9]+)?)\s*[)）]")
# 阳转率 95%CI 模式：百分数
_95CI_PCT_RE = re.compile(
    r"([0-9]+(?:\.[0-9]+)?)%\s*[（(]\s*([0-9]+(?:\.[0-9]+)?)%\s*[,，]\s*([0-9]+(?:\.[0-9]+)?)%\s*[)）]"
)

# 句式：试验组和阳性对照组1/2 <抗体> 抗原血清抗体 <指标>（95%CI）分别为 X、Y
_SENT_RE = re.compile(
    r"试验组和阳性对照组([12]?)\s*"
    r"(?:免前|一免后\d+个月|全免后\d+个月)?\s*"
    r"抗(ge|vzv)抗原(?:特异性)?血清抗体\s*"
    r"(GMC|GMI|阳转率|LS\s*GMC|GMI|阳转率)"
    r"[（(]95%CI[)）]\s*分别为\s*"
    r"(.+?)。",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Triple:
    """一条 (lo, mid, up) 数值。"""

    lo: float
    mid: float
    up: float

    @classmethod
    def parse(cls, s: str) -> Triple:
        """解析 "lo（mid, up）mIU/mL" 或 "lo%（mid%, up%）"。

        docx 中通常写作 "1103.5（905.2，1345.2）mIU/mL"，其中
        - 第 1 个数（中位）= 1103.5
        - 括号内 2 个数 = (905.2, 1345.2) = (lo, up)
        """
        s = s.strip()
        # 先试百分数
        m_pct = _95CI_PCT_RE.search(s)
        if m_pct:
            mid_pct, lo_pct, up_pct = m_pct.groups()
            return cls(lo=float(lo_pct), mid=float(mid_pct), up=float(up_pct))
        # 普通数值
        m = _95CI_RE.search(s)
        if m:
            mid, lo, up = m.groups()
            return cls(lo=float(lo), mid=float(mid), up=float(up))
        # 仅有单值（阳转率 0% 等）
        m_single = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
        if m_single:
            v = float(m_single.group(1))
            return cls(lo=v, mid=v, up=v)
        raise ValueError(f"无法解析 Triple: {s!r}")

    def by_type(self, col_type: str) -> float:
        if col_type == "mid":
            return self.mid
        if col_type == "up":
            return self.up
        if col_type == "lo":
            return self.lo
        raise ValueError(f"未知 col_type: {col_type}")


@dataclass
class AntibodyDataPoint:
    """一组 (age_band, metric, time_point) 的 试验组 / 阳性对照组 数值。"""

    age_band: str
    metric: str
    time_point: str
    trial: Triple | None = None
    control: Triple | None = None
    control_label: str = "阳性对照组"  # 阳性对照组1 / 阳性对照组2
    docx_para_ref: str = ""  # 段落 P#，用于审计

    def is_complete(self) -> bool:
        return self.trial is not None and self.control is not None


@dataclass
class AntibodyDataset:
    """完整抗体数据集。"""

    antibody: str  # "gE" or "VZV"
    data: dict[tuple[str, str, str], AntibodyDataPoint] = field(default_factory=dict)

    def set(self, dp: AntibodyDataPoint) -> None:
        key = (dp.age_band, dp.metric, dp.time_point)
        self.data[key] = dp

    def get(self, age_band: str, metric: str, time_point: str) -> AntibodyDataPoint | None:
        return self.data.get((age_band, metric, time_point))

    def coverage(self) -> dict[str, int]:
        """返回 (age_band, metric) 下的时间点完整数。"""
        cov: dict[str, int] = {}
        for (ab, m, _tp), dp in self.data.items():
            if dp.is_complete():
                key = f"{ab}/{m}"
                cov[key] = cov.get(key, 0) + 1
        return cov


def _split_two_groups(s: str) -> tuple[str, str]:
    """把 'X、Y' 切成 ('X', 'Y')。"""
    # "、" 或 "和" 分割
    s = s.replace("和", "、")
    parts = [p.strip() for p in s.split("、") if p.strip()]
    if len(parts) < 2:
        # 可能用 "," 分隔
        parts = [p.strip() for p in re.split(r"[,，]", s) if p.strip()]
    if len(parts) < 2:
        raise ValueError(f"无法切分为两组: {s!r}")
    return parts[0], parts[1]


def extract_antibody_dataset(paragraphs: list[str], antibody: str) -> AntibodyDataset:  # noqa: PLR0915 - TODO: 下个迭代重构 # noqa: PLR0912 - TODO: 下个迭代重构
    """从 docx 段落列表中抽取指定抗体（gE / VZV）的完整数据集。

    句式识别条件：
      1) 段落里出现 "抗{antibody}抗原 ... 抗体 ... <metric>（95%CI）分别为 A、B。"
      2) 紧接的上一行（非空）是年龄段标题，如 "40~49岁" / "50~59岁" 等
    """
    if antibody not in ANTIBODIES:
        raise ValueError(f"antibody 必须是 {ANTIBODIES} 之一，收到 {antibody!r}")

    ds = AntibodyDataset(antibody=antibody)

    # 把段落逐条过：当前句为数值句 → 向上回溯找最近的年龄段行
    age_in_effect: str | None = None
    for i, p in enumerate(paragraphs):
        ps = p.strip()
        # 年龄段标题行
        if ps in _AGE_NORM:
            age_in_effect = _AGE_NORM[ps]
            continue
        # 数值句：兼容两种句式
        #   A) 试验组和阳性对照组1 免前 抗gE抗原特异性血清抗体GMC（95%CI）分别为...
        #   B) 试验组和阳性对照组1 gE抗原血清抗体SCR（95%CI）分别为...
        m = re.search(
            rf"试验组和阳性对照组([12]?)\s*"
            rf"(免前|一免后\d+个月|全免后\d+个月|免后\d+天|免后\d+个月)?\s*"
            rf"(?:抗)?{antibody}抗原(?:特异性)?血清抗体\s*"
            rf"(LS\s*GMC|GMC|GMI|阳转率|阳转|SCR)\s*[（(]95%CI[)）]\s*分别为"
            rf"\s*(.+?)。",
            ps,
        )
        if not m:
            continue
        if age_in_effect is None:
            # 找不到对应年龄段，跳过
            continue
        ctrl_label_num = m.group(1) or "1"
        tp_raw = m.group(2)
        # 没出现 "免前/一免后/全免后" 前缀时，默认使用上一行的上下文
        # 若上一行包含“第 2 剂接种后 30 天”则映射为全免后 1 个月
        if tp_raw is None:
            # 查上下五个段落的提示（最多 5 行回溯）
            for off in range(1, 6):
                if i - off < 0:
                    break
                prev = paragraphs[i - off]
                # “第 2 剂接种后 30 天” → 全免后 1 个月
                if "接种后" in prev and "30" in prev:
                    tp_raw = "全免后1个月"
                    break
                # “第 2 剂接种前” / “PPS-h1” → 一免后 2 个月
                if "接种前" in prev or "PPS-h1" in prev:
                    tp_raw = "一免后2个月"
                    break
                # “PPS-h2” / “第 2 剂接种后” → 全免后 1 个月
                if "PPS-h2" in prev or "接种后" in prev:
                    tp_raw = "全免后1个月"
                    break
                if "免前" in prev or "基线" in prev:
                    tp_raw = "免前"
                    break
        tp = tp_raw or "免前"
        metric_raw = m.group(3)
        # 归一化
        metric = metric_raw.replace("LS ", "")
        # 阳转/阳转率/SCR 统一为阳转率
        if metric in ("阳转", "SCR"):
            metric = "阳转率"
        if metric not in METRICS:
            continue
        # 归一化时间点
        if tp in ("免前", "一免后2个月", "全免后1个月"):
            pass  # 已经是归一化值
        elif tp.startswith("免前"):
            tp = "免前"
        elif re.match(r"一免后\d+个月", tp):
            tp = "一免后2个月"
        elif re.match(r"全免后\d+个月", tp):
            tp = "全免后1个月"
        elif re.match(r"免后\d+天", tp) or re.match(r"免后\d+个月", tp):
            # 免后30天 = 全免后1个月，免后2个月 = 一免后2个月
            day_match = re.search(r"免后(\d+)天", tp)
            mon_match = re.search(r"免后(\d+)个月", tp)
            if day_match and int(day_match.group(1)) >= 28:
                tp = "全免后1个月"
            elif mon_match and int(mon_match.group(1)) >= 1:
                tp = "一免后2个月" if int(mon_match.group(1)) <= 2 else "全免后1个月"
            else:
                continue
        else:
            continue

        # 拆两组
        try:
            trial_str, ctrl_str = _split_two_groups(m.group(4))
        except ValueError:
            continue
        try:
            trial_t = Triple.parse(trial_str)
            ctrl_t = Triple.parse(ctrl_str)
        except ValueError:
            continue

        ds.set(
            AntibodyDataPoint(
                age_band=age_in_effect,
                metric=metric,
                time_point=tp,
                trial=trial_t,
                control=ctrl_t,
                control_label=f"阳性对照组{ctrl_label_num}",
                docx_para_ref=f"P{i + 1}",  # 段落序号 1-based（与 [P#] 行号一致）
            )
        )

    return ds


# 抗体对照映射：把源抗体（gE）每个 (age, metric, tp) 映射到目标抗体（VZV）
# 仅在两抗体都具备完整数据时返回映射。
def build_cross_mapping(
    source: AntibodyDataset,
    target: AntibodyDataset,
) -> dict[tuple[str, str, str], dict]:
    """生成 (age, metric, tp) -> {source, target, control_label, docx_refs}。

    返回的每条 dict 形如：
      {
        "age_band": "40-49岁",
        "metric": "GMC",
        "time_point": "免前",
        "source_trial": Triple(lo, mid, up),   # gE
        "source_control": Triple(...),
        "target_trial": Triple(...),            # VZV
        "target_control": Triple(...),
        "control_label": "阳性对照组1",
        "docx_refs": ["P127", "P138"],          # 源 + 目标
      }
    """
    out: dict[tuple[str, str, str], dict] = {}
    for key, src_dp in source.data.items():
        if not src_dp.is_complete():
            continue
        tgt_dp = target.get(*key)
        if tgt_dp is None or not tgt_dp.is_complete():
            continue
        age, metric, tp = key
        out[key] = {
            "age_band": age,
            "metric": metric,
            "time_point": tp,
            "source_trial": src_dp.trial,
            "source_control": src_dp.control,
            "target_trial": tgt_dp.trial,
            "target_control": tgt_dp.control,
            "control_label": src_dp.control_label,
            "docx_refs": [src_dp.docx_para_ref, tgt_dp.docx_para_ref],
        }
    return out
