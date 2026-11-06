# coding: utf-8
"""GUI验收只替换模型答复，提取、Office写回及回读使用真实实现。"""
import time
from pathlib import Path

from fastapi.testclient import TestClient
from 服务 import web
from 测试.组合样本 import 造金标准, 读Excel格
from 测试.test_Excel公式 import _造A, _造B, _模型


def 准备(tmp_path, monkeypatch, 模型=True):
    monkeypatch.setattr(web, "运行根", tmp_path / "运行")
    web.应用.state.状态 = {"线程": None, "运行目录": None, "日志": []}
    web.应用.state.配置路径 = tmp_path / "配置.json"
    monkeypatch.setattr(web, "读取模型配置", lambda _: {"url": "固定答复", "model": "验收"} if 模型 else {})
    if 模型:
        monkeypatch.setattr("核心.识别.调用识别", _模型())
    A, B = _造A(tmp_path / "A.xlsx"), _造B(tmp_path / "B.xlsx")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"]},
            "金标准": str(造金标准(tmp_path / "金标准.jsonl")), "复用": False}
    return TestClient(web.应用), 输入


def 等待(c, 标识, 秒=90):
    截止 = time.monotonic() + 秒
    while time.monotonic() < 截止:
        响应 = c.get(f"/api/桌面/任务/{标识}/进度")
        assert 响应.status_code == 200, 响应.text
        状态 = 响应.json()
        if not 状态["活动"]:
            return 状态
        time.sleep(0.05)
    raise AssertionError("任务未在限时内退出")


def 开始(c, 输入, 请求标识="开始一"):
    r = c.post("/api/桌面/任务", json={"请求标识": 请求标识, "类型": "自动搬运", "原始输入": 输入})
    assert r.status_code == 200, r.text
    return r.json()


def 核验金额(目录):
    from 核心.接续 import 读索引
    索引 = 读索引(目录)
    输出 = Path(索引["产物"]["A端新副本"])
    assert 读Excel格(输出, "资金", "C2") == 901
    assert 读Excel格(输出, "资金", "D2") == 102
    assert 读Excel格(输出, "资金", "B2") == "=客户甲"
    return 输出
