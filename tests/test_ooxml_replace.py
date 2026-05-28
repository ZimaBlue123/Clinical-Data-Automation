"""11_Word_Text_Replace/lib/ooxml_replace 单元测试。"""

from __future__ import annotations

import re

from ooxml_replace import ReplaceRuleSet, qn


def test_qn():
    assert qn("t").endswith("}t")


def test_replace_ruleset_literals():
    rs = ReplaceRuleSet(literals=[("foo", "bar")])
    assert rs.count_in_text("foo foo") == 2
    out, n = rs.apply("foo baz")
    assert out == "bar baz"
    assert n == 1


def test_replace_ruleset_regex():
    rx = re.compile(r"\d+")
    rs = ReplaceRuleSet(regexes=[(rx, "X")])
    out, n = rs.apply("a12b34")
    assert out == "aXbX"
    assert n == 2
