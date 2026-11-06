# coding: utf-8
import socket
import uuid
from pathlib import Path
import httpx
import pytest


def test_本机服务就绪和停止释放端口(tmp_path):
    from 桌面窗口 import 本机服务
    页面 = tmp_path / "index.html"
    页面.write_text("<html><head></head><body>桌面验收</body></html>", encoding="utf-8")
    服务 = 本机服务(页面)
    try:
        服务.启动()
        assert httpx.get(服务.地址 + "/api/桌面/概况", trust_env=False).json()["服务就绪"]
        assert "桌面验收" in httpx.get(服务.地址 + "/桌面", trust_env=False).text
        assert httpx.post(服务.地址 + "/api/中断", trust_env=False).status_code == 403
        assert httpx.post(服务.地址 + "/api/中断", headers={"X-Audit-Token": 服务.令牌}, trust_env=False).status_code == 200
        端口 = 服务.端口
    finally:
        服务.停止()
    with socket.socket() as s:
        assert s.connect_ex(("127.0.0.1", 端口)) != 0


def test_缺少构建页面不启动(tmp_path):
    from 桌面窗口 import 本机服务
    with pytest.raises(FileNotFoundError, match="桌面页面缺失"):
        本机服务(tmp_path / "不存在.html").启动()


def test_两个入口共用实例锁():
    from 桌面窗口 import 实例锁
    名 = "GUI验收_" + uuid.uuid4().hex
    第一 = 实例锁(名)
    第二 = 实例锁(名)
    assert 第一.取得()
    try:
        assert not 第二.取得()
    finally:
        第一.关闭()
        第二.关闭()
    第三 = 实例锁(名)
    assert 第三.取得()
    第三.关闭()
