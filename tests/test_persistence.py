# -*- coding: utf-8 -*-
"""持久化：data.json 为软链（CVM 部署形态）时，保存必须写真实文件且保留软链。"""
import pytest


def test_save_store_follows_symlink(server, tmp_path, monkeypatch):
    real_dir = tmp_path / "varlib"
    real_dir.mkdir()
    real_file = real_dir / "data.json"
    link = tmp_path / "data.json"
    try:
        link.symlink_to(real_file)
    except (OSError, NotImplementedError):
        pytest.skip("当前环境不支持软链")
    monkeypatch.setattr(server, "DATA_FILE", str(link))

    server.save_store({"data": [{"code": "A", "name": "x", "cname": "x", "categories": []}], "saved": []})

    assert link.is_symlink(), "软链应被保留，不能被替换成普通文件"
    assert real_file.exists(), "数据应写到软链指向的真实文件"
    store = server.load_store()
    assert store["data"][0]["code"] == "A"
