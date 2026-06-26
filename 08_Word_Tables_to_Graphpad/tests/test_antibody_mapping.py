"""08_Word_Tables_to_Graphpad 单元测试 — 抗体数据提取 + 跨抗体映射。

不依赖 pytest（用 unittest.TestCase 内置断言）。
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

MODULE_DIR = Path(__file__).resolve().parent.parent
if str(MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(MODULE_DIR))

from lib.antibody_mapping import (  # noqa: E402
    build_cross_mapping,
    extract_antibody_dataset,
    Triple,
)


class TestTriple(unittest.TestCase):
    def test_parse_value_with_ci(self):
        t = Triple.parse("1103.5（905.2，1345.2）mIU/mL")
        self.assertEqual(t.lo, 905.2)
        self.assertEqual(t.mid, 1103.5)
        self.assertEqual(t.up, 1345.2)

    def test_parse_percentage_with_ci(self):
        t = Triple.parse("100.00%（94.73%，100.00%）")
        self.assertEqual(t.lo, 94.73)
        self.assertEqual(t.mid, 100.00)
        self.assertEqual(t.up, 100.00)

    def test_parse_single_value(self):
        t = Triple.parse("0")
        self.assertEqual(t.lo, 0)
        self.assertEqual(t.mid, 0)
        self.assertEqual(t.up, 0)

    def test_by_type(self):
        t = Triple(lo=1.0, mid=2.0, up=3.0)
        self.assertEqual(t.by_type("mid"), 2.0)
        self.assertEqual(t.by_type("up"), 3.0)
        self.assertEqual(t.by_type("lo"), 1.0)

    def test_by_type_invalid(self):
        t = Triple(lo=1.0, mid=2.0, up=3.0)
        with self.assertRaises(ValueError):
            t.by_type("xxx")


class TestExtractDataset(unittest.TestCase):
    def test_extracts_ge_pair(self):
        paragraphs = [
            "gE抗原血清抗体水平",
            "40~49岁",
            "试验组和阳性对照组1免前抗gE抗原特异性血清抗体GMC（95%CI）分别为1103.5（905.2，1345.2）mIU/mL、935.7（761.4，1150.0）mIU/mL，组间差异无统计学意义。",
        ]
        ds = extract_antibody_dataset(paragraphs, "gE")
        self.assertIn(("40-49岁", "GMC", "免前"), ds.data)
        dp = ds.data[("40-49岁", "GMC", "免前")]
        self.assertEqual(dp.trial.mid, 1103.5)
        self.assertEqual(dp.control.mid, 935.7)
        self.assertEqual(dp.control_label, "阳性对照组1")

    def test_extracts_vzv_pair(self):
        paragraphs = [
            "VZV抗原血清抗体水平",
            "40~49岁",
            "试验组和阳性对照组1免前抗VZV抗原特异性血清抗体GMC（95%CI）分别为615.9（509.6，744.3）mIU/mL、595.4（483.5，733.3）mIU/mL，组间差异无统计学意义。",
        ]
        ds = extract_antibody_dataset(paragraphs, "VZV")
        dp = ds.get("40-49岁", "GMC", "免前")
        self.assertIsNotNone(dp)
        self.assertEqual(dp.trial.mid, 615.9)

    def test_extracts_one_month_after_full(self):
        paragraphs = [
            "第2剂接种后30天",
            "40-49岁",
            "试验组和阳性对照组1 gE抗原血清抗体SCR（95%CI）分别为100.00%（94.73%，100.00%）、20.00%（11.39%，31.27%），组间SCR率差。",
        ]
        ds = extract_antibody_dataset(paragraphs, "gE")
        dp = ds.get("40-49岁", "阳转率", "全免后1个月")
        self.assertIsNotNone(dp)
        self.assertEqual(dp.trial.mid, 100.0)
        self.assertEqual(dp.control.mid, 20.0)

    def test_normalizes_age_bands(self):
        paragraphs = [
            "50~59岁",
            "试验组和阳性对照组2免前抗gE抗原特异性血清抗体GMC（95%CI）分别为1145.9（937.1，1401.2）mIU/mL、1199.8（1011.8，1422.8）mIU/mL。",
            "60岁及以上",
            "试验组和阳性对照组2免前抗gE抗原特异性血清抗体GMC（95%CI）分别为1339.7（1103.4，1626.5）mIU/mL、1294.9（1074.5，1560.5）mIU/mL。",
        ]
        ds = extract_antibody_dataset(paragraphs, "gE")
        self.assertIsNotNone(ds.get("50-59岁", "GMC", "免前"))
        self.assertIsNotNone(ds.get("60岁以上", "GMC", "免前"))

    def test_skips_incomplete_sentences(self):
        paragraphs = [
            "40~49岁",
            "本段不含对照数据。",
        ]
        ds = extract_antibody_dataset(paragraphs, "gE")
        self.assertEqual(ds.data, {})

    def test_wrong_antibody_skipped(self):
        paragraphs = [
            "40~49岁",
            "试验组和阳性对照组1免前抗VZV抗原特异性血清抗体GMC（95%CI）分别为615.9（509.6，744.3）mIU/mL。",
        ]
        ds = extract_antibody_dataset(paragraphs, "gE")
        self.assertEqual(ds.data, {})


class TestCrossMapping(unittest.TestCase):
    def test_build_mapping(self):
        src = extract_antibody_dataset(
            [
                "40~49岁",
                "试验组和阳性对照组1免前抗gE抗原特异性血清抗体GMC（95%CI）分别为1103.5（905.2，1345.2）mIU/mL、935.7（761.4，1150.0）mIU/mL。",
            ],
            "gE",
        )
        tgt = extract_antibody_dataset(
            [
                "40~49岁",
                "试验组和阳性对照组1免前抗VZV抗原特异性血清抗体GMC（95%CI）分别为615.9（509.6，744.3）mIU/mL、595.4（483.5，733.3）mIU/mL。",
            ],
            "VZV",
        )
        cross = build_cross_mapping(src, tgt)
        self.assertIn(("40-49岁", "GMC", "免前"), cross)
        item = cross[("40-49岁", "GMC", "免前")]
        self.assertEqual(item["source_trial"].mid, 1103.5)
        self.assertEqual(item["target_trial"].mid, 615.9)
        self.assertEqual(item["control_label"], "阳性对照组1")

    def test_no_mapping_when_target_missing(self):
        src = extract_antibody_dataset(
            [
                "40~49岁",
                "试验组和阳性对照组1免前抗gE抗原特异性血清抗体GMC（95%CI）分别为1103.5（905.2，1345.2）mIU/mL、935.7（761.4，1150.0）mIU/mL。",
            ],
            "gE",
        )
        tgt = extract_antibody_dataset([], "VZV")
        cross = build_cross_mapping(src, tgt)
        self.assertEqual(cross, {})


if __name__ == "__main__":
    unittest.main()
