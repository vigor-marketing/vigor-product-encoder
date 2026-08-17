# -*- coding: utf-8 -*-
"""pytest 共享夹具：按路径加载 local/cloud 两个 server 模块，并隔离数据文件。"""
import importlib.util
import pathlib

import pytest

ROOT = pathlib.Path(__file__).resolve().parent.parent


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def local_server():
    return _load(ROOT / "local" / "server.py", "vigor_local_server")


@pytest.fixture(scope="session")
def cloud_server():
    return _load(ROOT / "cloud" / "server.py", "vigor_cloud_server")


@pytest.fixture(params=["local_server", "cloud_server"])
def server(request):
    """同一套业务逻辑在本地版与云端版上各跑一遍，防止两处代码漂移。"""
    return request.getfixturevalue(request.param)


@pytest.fixture
def data_file(server, tmp_path, monkeypatch):
    """把数据文件指向临时目录：测试互不影响，也不触碰真实 data.json。"""
    p = tmp_path / "data.json"
    monkeypatch.setattr(server, "DATA_FILE", str(p))
    return p
