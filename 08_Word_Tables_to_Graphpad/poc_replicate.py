"""08_Word_Tables_to_Graphpad 主程序。

把 docx 中"源抗体"的免疫原性数据（GMC / GMI / 阳转率 × 4 年龄段 × 3 时间点 × 2 组别）
按 (年龄段 × 指标) 替换 pzfx 模板中"目标抗体"的数值，生成新的 pzfx。

数据流
------
1. 读 docx 段落，解析源抗体 (gE) 与目标抗体 (VZV) 的所有 Triple(lo, mid, up)
2. 读 pzfx 模板，校验其数据 == docx 中源抗体的数据
3. 用目标抗体的数值，按 (table_id, yi, si) 改写 pzfx 的 6 个 Subcolumn × 3 个 d
4. 写新 pzfx + 审计 + 校验报告

CLI 示例
--------
.. code-block:: bash

    cd 08_Word_Tables_to_Graphpad
    python poc_replicate.py \\
        --docx input/source.docx \\
        --pzfx input/template.pzfx \\
        --source-antibody gE \\
        --target-antibody VZV \\
        --out output/result.pzfx
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

MODULE_DIR = Path(__file__).resolve().parent
ROOT_DIR = MODULE_DIR.parent

if str(MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(MODULE_DIR))

from lib.docx_parser import parse_docx  # noqa: E402
from lib.pzfx_parser import PzfxFile, parse_pzfx  # noqa: E402
from lib.pzfx_writer import (  # noqa: E402
    PzfxValue,
    SUBCOL_TYPE,
    TIME_POINTS,
    rewrite_pzfx_data,
)
from lib.antibody_mapping import (  # noqa: E402
    AntibodyDataset,
    Triple,
    build_cross_mapping,
    extract_antibody_dataset,
)

logger = logging.getLogger("poc_replicate")

# (年龄段, 指标) → pzfx TableID 映射
# 默认对应"重组带状疱疹疫苗（CHO细胞）II期阶段性小结"模板的 12 张表
DEFAULT_AGE_METRIC_TO_TABLE: dict[tuple[str, str], str] = {
    ("40-49岁", "GMC"): "Table2",
    ("50-59岁", "GMC"): "Table0",
    ("60岁以上", "GMC"): "Table1",
    ("50岁以上", "GMC"): "Table9",
    ("40-49岁", "GMI"): "Table5",
    ("50-59岁", "GMI"): "Table3",
    ("60岁以上", "GMI"): "Table4",
    ("50岁以上", "GMI"): "Table11",
    ("40-49岁", "阳转率"): "Table6",
    ("50-59岁", "阳转率"): "Table7",
    ("60岁以上", "阳转率"): "Table8",
    ("50岁以上", "阳转率"): "Table10",
}


def _build_logger(verbose: bool) -> logging.Logger:
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    if not logger.handlers:
        h = logging.StreamHandler()
        h.setFormatter(
            logging.Formatter(
                fmt="%(asctime)s [%(levelname)s] %(message)s",
                datefmt="%Y-%m-%dT%H:%M:%S",
            )
        )
        logger.addHandler(h)
    return logger


def _verify_pzfx_against_source(
    pzfx: PzfxFile,
    source: AntibodyDataset,
    age_metric_to_table: dict[tuple[str, str], str],
    *,
    tolerance: float = 0.15,
) -> tuple[int, int]:
    """校验 pzfx 现有数值 == docx 中源抗体数据。"""
    ok = 0
    bad = 0
    for (age, metric), tid in age_metric_to_table.items():
        tbl = next((t for t in pzfx.tables if t.table_id == tid), None)
        if tbl is None:
            continue
        for yi in range(min(2, len(tbl.columns))):
            col_key = "trial" if yi == 0 else "control"
            for si in range(3):
                col_type = SUBCOL_TYPE[si]
                tp = TIME_POINTS[si]
                if metric == "GMI" and tp == "免前":
                    continue
                if metric == "阳转率" and tp == "免前":
                    continue
                src_dp = source.get(age, metric, tp)
                if src_dp is None or not src_dp.is_complete():
                    continue
                src_triple = getattr(src_dp, col_key)
                expected = src_triple.by_type(col_type)
                cell_values = tbl.get(yi, si)
                if len(cell_values) < 3:
                    continue
                tp_idx = TIME_POINTS.index(tp)
                try:
                    actual = float(cell_values[tp_idx])
                except (ValueError, IndexError):
                    bad += 1
                    continue
                if abs(actual - expected) > tolerance:
                    bad += 1
                    logger.warning(
                        "源值不一致 %s/%s yi=%d si=%d tp=%s: pzfx=%s docx=%s",
                        tid,
                        metric,
                        yi,
                        si,
                        tp,
                        actual,
                        expected,
                    )
                else:
                    ok += 1
    return ok, bad


def _build_pzfx_values_for_age_metric(
    target: AntibodyDataset,
    age: str,
    metric: str,
) -> dict[tuple[int, int], PzfxValue | None]:
    """构造 (yi, si) -> PzfxValue 的字典。

    每个 (yi, si) 对应一张亚列（mid/up/lo），该亚列内 3 个 d 依次对应 3 个时间点。
    PzfxValue 包装 3 个 Triple(免前, 一免后2个月, 全免后1个月) 各自的 (lo, mid, up)。
    """
    out: dict[tuple[int, int], PzfxValue | None] = {}
    for yi in range(2):
        col_key = "trial" if yi == 0 else "control"
        for si in range(3):
            tp = TIME_POINTS[si]
            if metric == "GMI" and tp == "免前":
                out[(yi, si)] = None
                continue
            if metric == "阳转率" and tp == "免前":
                out[(yi, si)] = None
                continue
            # 构造 3 条 Triple
            triples: list[Triple] = []
            ok = True
            for tpx in TIME_POINTS:
                dp = target.get(age, metric, tpx)
                if dp is None or not dp.is_complete():
                    ok = False
                    break
                triples.append(getattr(dp, col_key))
            if not ok:
                out[(yi, si)] = None
                continue
            out[(yi, si)] = PzfxValue(*triples)
    return out


def main(argv: list[str] | None = None) -> int:  # noqa: PLR0915 - TODO: 下个迭代重构
    parser = argparse.ArgumentParser(
        description="把 docx 中源抗体的免疫原性数据按 (年龄段 × 指标) 替换到 pzfx 模板，生成目标抗体的新 pzfx。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--docx", required=True, help="docx 路径（包含源抗体与目标抗体数据）")
    parser.add_argument("--pzfx", required=True, help="pzfx 模板路径")
    parser.add_argument("--out", required=True, help="输出 pzfx 路径")
    parser.add_argument(
        "--source-antibody",
        default="gE",
        choices=("gE", "VZV"),
        help="源抗体（默认 gE）",
    )
    parser.add_argument(
        "--target-antibody",
        default="VZV",
        choices=("gE", "VZV"),
        help="目标抗体（默认 VZV）",
    )
    parser.add_argument("--audit-log", help="审计日志输出路径（可选）")
    parser.add_argument("--verify-report", help="校验报告 JSON 输出路径（可选）")
    parser.add_argument(
        "--table-map",
        help="自定义 (年龄段, 指标) → TableID 的 JSON 文件（可选，覆盖默认映射）",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="详细输出")
    args = parser.parse_args(argv)

    _build_logger(args.verbose)
    logger.info(
        "start: docx=%s pzfx=%s source=%s target=%s",
        args.docx,
        args.pzfx,
        args.source_antibody,
        args.target_antibody,
    )

    age_metric_to_table = dict(DEFAULT_AGE_METRIC_TO_TABLE)
    if args.table_map:
        age_metric_to_table.update(json.loads(Path(args.table_map).read_text(encoding="utf-8")))
        logger.info("loaded custom table map from %s", args.table_map)

    # 1) 读 docx
    try:
        content = parse_docx(args.docx)
    except (FileNotFoundError, ValueError, KeyError) as e:
        logger.error("解析 docx 失败: %s", e)
        return 2
    logger.info("docx 段落=%d 表=%d", len(content.paragraphs), len(content.tables))

    # 2) 抽取抗体数据集
    source_ds = extract_antibody_dataset(content.paragraphs, args.source_antibody)
    target_ds = extract_antibody_dataset(content.paragraphs, args.target_antibody)
    src_cov = source_ds.coverage()
    tgt_cov = target_ds.coverage()
    logger.info("源抗体 %s 数据完整度: %s", args.source_antibody, src_cov)
    logger.info("目标抗体 %s 数据完整度: %s", args.target_antibody, tgt_cov)

    # 3) 交叉映射
    cross = build_cross_mapping(source_ds, target_ds)
    if not cross:
        logger.error("源与目标抗体无任何 (年龄段, 指标, 时间点) 的共同完整数据，终止")
        return 3
    logger.info("交叉映射数: %d", len(cross))

    # 4) 读 pzfx
    try:
        pzfx = parse_pzfx(args.pzfx)
    except (FileNotFoundError, ValueError) as e:
        logger.error("解析 pzfx 失败: %s", e)
        return 2
    logger.info("pzfx 表数=%d (zip=%s)", len(pzfx.tables), pzfx.is_zip)

    # 5) 校验 pzfx 现有数据 == docx 源抗体
    ok, bad = _verify_pzfx_against_source(pzfx, source_ds, age_metric_to_table)
    if bad > 0:
        logger.warning("源值校验: 通过=%d 不通过=%d（继续，但请人工检查）", ok, bad)
    else:
        logger.info("源值校验: 全部 %d 项通过", ok)

    # 6) 改写 pzfx
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    current = Path(args.pzfx)
    rewrite_log: list[dict] = []
    for (age, metric), tid in age_metric_to_table.items():
        key = (age, metric, "免前")
        if key not in cross:
            continue
        new_values = _build_pzfx_values_for_age_metric(target_ds, age, metric)
        n = rewrite_pzfx_data(
            src=current,
            dst=out_path,
            age_band=age,
            metric=metric,
            table_id=tid,
            new_values=new_values,
            logger=logger,
        )
        rewrite_log.append({"table_id": tid, "age": age, "metric": metric, "replaced_subcolumns": n})
        current = out_path

    # 7) 写审计
    if args.audit_log:
        Path(args.audit_log).write_text(
            json.dumps(
                {
                    "source": args.source_antibody,
                    "target": args.target_antibody,
                    "docx": args.docx,
                    "pzfx": args.pzfx,
                    "out": str(out_path),
                    "rewrite_log": rewrite_log,
                    "source_coverage": src_cov,
                    "target_coverage": tgt_cov,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        logger.info("审计: %s", args.audit_log)

    # 8) 写校验报告
    if args.verify_report:
        Path(args.verify_report).write_text(
            json.dumps(
                {
                    "source_consistency": {"ok": ok, "bad": bad},
                    "rewrite_log": rewrite_log,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        logger.info("校验: %s", args.verify_report)

    logger.info("done: out=%s size=%d", out_path, out_path.stat().st_size)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
