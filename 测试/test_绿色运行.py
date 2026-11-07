# coding: utf-8
"""绿色版只改变运行位置，不改财务规则。"""
import sys
from pathlib import Path
import pytest

def test_冻结程序默认位置不跟随工作目录(tmp_path, monkeypatch):
    from 核心.设置 import 默认配置路径
    from 核心.共用记录簿 import 项目根
    根 = tmp_path / "中文 软件"
    根.mkdir()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(根 / "审计附注搬运.exe"))
    monkeypatch.chdir(tmp_path)
    assert 默认配置路径() == 根 / "模型配置.json"
    assert 项目根() == 根

def test_绿色版保存备份在软件内且原文不变(tmp_path, monkeypatch):
    from 服务.人工操作 import 备份已有
    根 = tmp_path / "绿色软件"
    根.mkdir()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(根 / "审计附注搬运.exe"))
    原 = 根 / "设置.json"
    原.write_bytes(b'{"old":1}')
    备份已有(原)
    副本 = list((根 / "备份").rglob("设置.json"))
    assert len(副本) == 1 and 副本[0].read_bytes() == 原.read_bytes() == b'{"old":1}'

def test_绿色版拒绝包外写入而源码不受限制(tmp_path, monkeypatch):
    from 运行位置 import 检查数据位置
    根 = tmp_path / "绿色软件"
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(根 / "审计附注搬运.exe"))
    assert 检查数据位置(根 / "运行" / "成果.xlsx") == 根 / "运行" / "成果.xlsx"
    with pytest.raises(ValueError, match="软件文件夹"):
        检查数据位置(tmp_path / "外部" / "成果.xlsx")
    monkeypatch.setattr(sys, "frozen", False)
    assert 检查数据位置(tmp_path / "外部" / "成果.xlsx") == tmp_path / "外部" / "成果.xlsx"
