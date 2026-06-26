"""08_Word_Tables_to_Graphpad — pzfx 解析/写盘单元测试。"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

MODULE_DIR = Path(__file__).resolve().parent.parent
if str(MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(MODULE_DIR))

from lib.antibody_mapping import Triple  # noqa: E402
from lib.pzfx_parser import parse_pzfx  # noqa: E402
from lib.pzfx_writer import (  # noqa: E402
    PzfxValue,
    rewrite_pzfx_data,
)


SAMPLE_PZFX = """<?xml version="1.0" encoding="UTF-8"?>
<GraphPadPrismFile xmlns="http://graphpad.com/prism/Prism.htm" PrismXMLVersion="5.00">
<Created><OriginalVersion CreatedByProgram="GraphPad Prism" CreatedByVersion="10.1.2.324"/></Created>
<TableSequence><Ref ID="Table0"/></TableSequence>
<Table ID="Table0" XFormat="none" YFormat="upper-lower-limits" TableType="TwoWay" EVFormat="AsteriskAfterNumber">
<Title>40-49岁GMC</Title>
<RowTitlesColumn Width="164">
<Subcolumn>
<d>免前</d>
<d>一免后2个月</d>
<d>全免后1个月</d>
</Subcolumn>
</RowTitlesColumn>
<YColumn Width="210" Decimals="2" Subcolumns="3">
<Title>试验组</Title>
<Subcolumn>
<d>100</d>
<d>200</d>
<d>300</d>
</Subcolumn>
<Subcolumn>
<d>150</d>
<d>250</d>
<d>350</d>
</Subcolumn>
<Subcolumn>
<d>50</d>
<d>150</d>
<d>250</d>
</Subcolumn>
</YColumn>
<YColumn Width="210" Decimals="2" Subcolumns="3">
<Title>阳性对照组1</Title>
<Subcolumn>
<d>10</d>
<d>20</d>
<d>30</d>
</Subcolumn>
<Subcolumn>
<d>15</d>
<d>25</d>
<d>35</d>
</Subcolumn>
<Subcolumn>
<d>5</d>
<d>15</d>
<d>25</d>
</Subcolumn>
</YColumn>
</Table>
</GraphPadPrismFile>
"""


def _make_sample(tmp_path: Path) -> Path:
    p = tmp_path / "sample.pzfx"
    p.write_text(SAMPLE_PZFX, encoding="utf-8")
    return p


class TestParsePzfx(unittest.TestCase):
    def test_parse_extracts_table_and_columns(self):
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            p = _make_sample(Path(td))
            pzfx = parse_pzfx(p)
            self.assertFalse(pzfx.is_zip)
            self.assertEqual(len(pzfx.tables), 1)
            t = pzfx.tables[0]
            self.assertEqual(t.table_id, "Table0")
            self.assertEqual(t.title, "40-49岁GMC")
            self.assertEqual(t.age_band, "40-49岁")
            self.assertEqual(t.metric, "GMC")
            self.assertEqual(len(t.columns), 2)
            self.assertEqual(t.columns[0]["title"], "试验组")
            self.assertEqual(t.columns[1]["title"], "阳性对照组1")
            self.assertEqual(t.columns[0]["subcolumns"][0], ["100", "200", "300"])
            self.assertEqual(t.columns[0]["subcolumns"][1], ["150", "250", "350"])
            self.assertEqual(t.columns[0]["subcolumns"][2], ["50", "150", "250"])

    def test_get_table_helper(self):
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            p = _make_sample(Path(td))
            pzfx = parse_pzfx(p)
            t = pzfx.get_table("40-49岁", "GMC")
            self.assertIsNotNone(t)
            self.assertEqual(t.table_id, "Table0")
            self.assertIsNone(pzfx.get_table("60岁以上", "GMC"))


class TestRewritePzfx(unittest.TestCase):
    def test_rewrite_replaces_all_subcolumns(self):
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            src = _make_sample(tmp)
            out = tmp / "out.pzfx"
            # 每个 (yi, si) 对应一个亚列（mid/up/lo），亚列内 3 个 d 依次对应 3 个时间点。
            # PzfxValue(*triples) 包装 3 条 Triple(免前, 一免后2个月, 全免后1个月)。
            # 试验组免前: lo=10 mid=20 up=30; 一免后: 11/21/31; 全免后: 12/22/32
            t_trial = [
                Triple(lo=10.0, mid=20.0, up=30.0),
                Triple(lo=11.0, mid=21.0, up=31.0),
                Triple(lo=12.0, mid=22.0, up=32.0),
            ]
            t_ctrl = [
                Triple(lo=13.0, mid=23.0, up=33.0),
                Triple(lo=14.0, mid=24.0, up=34.0),
                Triple(lo=15.0, mid=25.0, up=35.0),
            ]
            new_values = {
                (0, 0): PzfxValue(*t_trial),  # 试验组 mid 列
                (0, 1): PzfxValue(*t_trial),  # 试验组 up 列
                (0, 2): PzfxValue(*t_trial),  # 试验组 lo 列
                (1, 0): PzfxValue(*t_ctrl),
                (1, 1): PzfxValue(*t_ctrl),
                (1, 2): PzfxValue(*t_ctrl),
            }
            n = rewrite_pzfx_data(
                src=src,
                dst=out,
                age_band="40-49岁",
                metric="GMC",
                table_id="Table0",
                new_values=new_values,
            )
            self.assertEqual(n, 6)
            new = parse_pzfx(out)
            t = new.tables[0]
            # yi=0 试验组 mid 列: 3 个时间点 mid 依次 20, 21, 22
            self.assertEqual(t.columns[0]["subcolumns"][0], ["20", "21", "22"])
            self.assertEqual(t.columns[0]["subcolumns"][1], ["30", "31", "32"])
            self.assertEqual(t.columns[0]["subcolumns"][2], ["10", "11", "12"])
            # yi=1 对照组
            self.assertEqual(t.columns[1]["subcolumns"][0], ["23", "24", "25"])
            self.assertEqual(t.columns[1]["subcolumns"][1], ["33", "34", "35"])
            self.assertEqual(t.columns[1]["subcolumns"][2], ["13", "14", "15"])

    def test_rewrite_none_keeps_original(self):
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            src = _make_sample(tmp)
            out = tmp / "out.pzfx"
            triples = [Triple(lo=10.0, mid=20.0, up=30.0)] * 3
            new_values = {
                (0, 0): PzfxValue(*triples),
            }
            rewrite_pzfx_data(
                src=src,
                dst=out,
                age_band="40-49岁",
                metric="GMC",
                table_id="Table0",
                new_values=new_values,
            )
            new = parse_pzfx(out)
            t = new.tables[0]
            self.assertEqual(t.columns[0]["subcolumns"][0], ["20", "20", "20"])
            self.assertEqual(t.columns[0]["subcolumns"][1], ["150", "250", "350"])
            self.assertEqual(t.columns[0]["subcolumns"][2], ["50", "150", "250"])

    def test_rewrite_preserves_xml_structure(self):
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            src = _make_sample(tmp)
            out = tmp / "out.pzfx"
            triples = [Triple(lo=1.0, mid=2.0, up=3.0)] * 3
            rewrite_pzfx_data(
                src=src,
                dst=out,
                age_band="40-49岁",
                metric="GMC",
                table_id="Table0",
                new_values={(0, 0): PzfxValue(*triples)},
            )
            text = out.read_text(encoding="utf-8")
            self.assertIn('xmlns="http://graphpad.com/prism/Prism.htm"', text)
            self.assertIn('CreatedByVersion="10.1.2.324"', text)
            self.assertIn("<Title>40-49岁GMC</Title>", text)
            self.assertIn("<d>免前</d>", text)
            self.assertIn('YFormat="upper-lower-limits"', text)
            self.assertIn('Decimals="2"', text)


if __name__ == "__main__":
    unittest.main()
