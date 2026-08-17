# -*- coding: utf-8 -*-
"""编码顺延规则测试：标准/子类 A~Z、产品 AA~ZZ。"""
import pytest


@pytest.mark.parametrize("codes,length,expected", [
    ([], 1, "A"),
    (["A"], 1, "B"),
    (["A", "B"], 1, "C"),
    (["A", "C"], 1, "D"),   # 顺延取最大值+1，不填补中间空缺
    (["Y"], 1, "Z"),
    (["Z"], 1, None),       # 单字母已满
    ([], 2, "AA"),
    (["AA"], 2, "AB"),
    (["AZ"], 2, "BA"),
    (["ZY"], 2, "ZZ"),
    (["ZZ"], 2, None),      # 双字母已满
])
def test_next_code(server, codes, length, expected):
    assert server.next_code(codes, length) == expected


def test_next_code_ignores_invalid(server):
    # 小写、数字、空串不计入已有编码
    assert server.next_code(["a", "1", ""], 1) == "A"
