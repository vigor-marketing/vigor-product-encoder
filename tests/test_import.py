# -*- coding: utf-8 -*-
"""批量导入 do_import：层级构建 / 名称匹配 / 参数组合并 / 错误处理。"""


def test_import_single_chain(server, data_file):
    rows = [{
        "api_cn": "套管和油管", "api_en": "API 5CT",
        "cat_cn": "套管", "cat_en": "Casing",
        "prod_cn": "套管", "prod_en": "Casing",
        "params": [],
    }]
    stats = server.do_import(rows)
    assert stats["errors"] == []
    assert (stats["api"], stats["cat"], stats["prod"]) == (1, 1, 1)

    a = server.load_store()["data"][0]
    assert a["code"] == "A" and a["name"] == "API 5CT" and a["cname"] == "套管和油管"
    c = a["categories"][0]
    assert c["code"] == "A" and c["name"] == "Casing" and c["cname"] == "套管"
    p = c["products"][0]
    assert p["code"] == "AA" and p["name"] == "Casing" and p["cname"] == "套管"


def test_import_reuses_by_name_and_continues_code(server, data_file):
    base = {"api_cn": "套管和油管", "api_en": "API 5CT",
            "cat_cn": "套管", "cat_en": "Casing", "params": []}
    server.do_import([dict(base, prod_cn="套管", prod_en="Casing")])
    stats = server.do_import([dict(base, prod_cn="油管", prod_en="Tubing")])

    # 标准、子类已存在，仅新增产品
    assert (stats["api"], stats["cat"], stats["prod"]) == (0, 0, 1)

    data = server.load_store()["data"]
    assert len(data) == 1
    prods = data[0]["categories"][0]["products"]
    assert [p["cname"] for p in prods] == ["套管", "油管"]
    assert [p["code"] for p in prods] == ["AA", "AB"]


def test_import_new_format_params_merge_options(server, data_file):
    rows = [{
        "api_cn": "套管和油管", "api_en": "API 5CT",
        "cat_cn": "套管", "cat_en": "Casing",
        "prod_cn": "套管", "prod_en": "Casing",
        "params": [
            {"label": "K1", "group_name_cn": "外径", "group_name_en": "OD",
             "code": "45", "desc": "4½in", "desc_en": "4½in"},
            {"label": "K1", "group_name_cn": "外径", "group_name_en": "OD",
             "code": "50", "desc": "5in", "desc_en": "5in"},
        ],
    }]
    stats = server.do_import(rows)
    assert stats["errors"] == []
    assert (stats["pg"], stats["opt"]) == (1, 2)

    pgs = server.load_store()["data"][0]["categories"][0]["products"][0]["paramGroups"]
    assert len(pgs) == 1
    assert pgs[0]["label"] == "K1"
    assert pgs[0]["name"] == "外径" and pgs[0]["name_en"] == "OD"
    assert [o["code"] for o in pgs[0]["options"]] == ["45", "50"]
    assert [o["desc_en"] for o in pgs[0]["options"]] == ["4½in", "5in"]


def test_import_legacy_format_params(server, data_file):
    rows = [{
        "api_cn": "套管和油管", "api_en": "API 5CT",
        "cat_cn": "套管", "cat_en": "Casing",
        "prod_cn": "套管", "prod_en": "Casing",
        "params": [{"label": "", "name_cn": "外径", "name_en": "OD",
                    "opts": [["45", "4½in"], ["50", "5in"]]}],
    }]
    stats = server.do_import(rows)
    assert stats["errors"] == []
    assert (stats["pg"], stats["opt"]) == (1, 2)

    pgs = server.load_store()["data"][0]["categories"][0]["products"][0]["paramGroups"]
    assert pgs[0]["name"] == "外径"
    assert [o["code"] for o in pgs[0]["options"]] == ["45", "50"]


def test_import_missing_api_reports_error(server, data_file):
    stats = server.do_import([{"api_cn": "", "api_en": "", "cat_cn": "",
                               "cat_en": "", "prod_cn": "", "prod_en": "",
                               "params": []}])
    assert len(stats["errors"]) == 1
    assert "缺少产品标准名称" in stats["errors"][0]
