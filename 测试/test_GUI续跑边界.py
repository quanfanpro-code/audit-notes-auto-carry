# coding: utf-8
from pathlib import Path
from 测试.GUI样本 import 准备, 开始, 等待
from 核心.接续 import 写索引, 读索引


def test_旧输入不全仍显示文件但不猜复用设置(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    from 服务 import web
    p = web.运行根 / "任务_旧"
    p.mkdir(parents=True)
    写索引(p, {"步骤": "用户中断", "输入指纹": {**输入, "金标准": {"路径": 输入["金标准"]}}})
    原 = (p / "接续索引.json").read_bytes()
    r = c.get(f"/api/桌面/任务/{p.name}").json()
    assert r["A"]["路径"] == 输入["A"]["路径"]
    assert not r["允许继续"] and "未完整保存" in r["不能继续原因"]
    assert c.post(f"/api/桌面/任务/{p.name}/继续", json={}).status_code == 409
    assert (p / "接续索引.json").read_bytes() == 原


def test_未知步骤只读且原件在文件清单中(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch, 模型=False)
    r = 开始(c, 输入)
    等待(c, r["task_id"])
    p = Path(r["运行目录"])
    索引 = 读索引(p)
    索引["步骤"] = "未来版本的新步骤"
    写索引(p, 索引)
    显示 = c.get(f'/api/桌面/任务/{r["task_id"]}').json()
    assert not 显示["允许继续"] and "未知" in 显示["显示状态"]
    assert c.post(f'/api/桌面/任务/{r["task_id"]}/继续', json={}).status_code == 409
    文件 = c.get(f'/api/桌面/任务/{r["task_id"]}/文件').json()["条目"]
    assert {x["路径"] for x in 文件 if x["名称"] in ("A输入原件", "B输入原件")} == {输入["A"]["路径"], 输入["B"]["路径"]}


def test_空闲关闭也先阻止新任务(tmp_path, monkeypatch):
    from 桌面窗口 import 窗口桥
    from 服务 import web
    准备(tmp_path, monkeypatch)
    桥 = 窗口桥()
    try:
        assert 桥.关闭事件() is True
        assert web.应用.state.退出中 is True
    finally:
        web.应用.state.退出中 = False
