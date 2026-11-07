# coding: utf-8
import json
import threading
from pathlib import Path

from 服务 import web
from 核心.接续 import 读索引, 写索引
from 测试.GUI样本 import 准备, 开始, 等待, 核验金额
from 测试.组合样本 import 文件哈希
from 测试.test_Excel公式 import _模型


def test_同请求不重复执行_真实副本与指定历史继续(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch, 模型=False)
    第一 = 开始(c, 输入)
    重发 = 开始(c, 输入)
    assert 第一["task_id"] == 重发["task_id"]
    assert 等待(c, 第一["task_id"])["等待"]
    assert c.post(f'/api/桌面/任务/{第一["task_id"]}/中断').json()["显示状态"] == "用户中断"
    monkeypatch.setattr(web, "读取模型配置", lambda _: {"url": "固定答复", "model": "验收"})
    monkeypatch.setattr("核心.识别.调用识别", _模型())
    第二 = 开始(c, 输入, "开始二")
    assert 等待(c, 第二["task_id"])["步骤"] == "done"
    列表 = c.get("/api/桌面/任务").json()
    assert 列表["总数"] == 2
    r = c.post(f'/api/桌面/任务/{第一["task_id"]}/继续', json={"请求标识": "继续第一"})
    assert r.status_code == 200, r.text
    assert r.json()["运行目录"] == 第一["运行目录"]
    assert r.json()["执行次数"] == 2
    assert 等待(c, 第一["task_id"])["步骤"] == "done"
    核验金额(第一["运行目录"])


def test_看旧任务不切换活动任务_线程未退不能新开(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    旧 = web.运行根 / "任务_旧任务"
    旧.mkdir(parents=True)
    写索引(旧, {"步骤": "done", "产物": {}})
    进入, 放行 = threading.Event(), threading.Event()
    def 慢任务(输入, 目录, 配置, 取消=None):
        写索引(目录, {"步骤": "done", "产物": {}})
        进入.set()
        放行.wait(15)
    monkeypatch.setattr(web, "运行任务", 慢任务)
    try:
        新 = 开始(c, 输入)
        assert 进入.wait(3)
        assert c.get(f"/api/桌面/任务/{旧.name}/进度").json()["步骤"] == "done"
        assert Path(web.应用.state.状态["运行目录"]) == Path(新["运行目录"])
        assert c.post(f"/api/桌面/任务/{旧.name}/中断", json={}).status_code == 409
        assert c.post("/api/桌面/任务", json={"请求标识": "另一份", "原始输入": 输入}).status_code == 409
        assert not web.应用.state.状态["取消"].is_set()
    finally:
        放行.set()
        if web.应用.state.状态.get("线程"):
            web.应用.state.状态["线程"].join(5)


def test_拒绝改变输入的请求和续跑_不改旧索引(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch, 模型=False)
    原件前 = 文件哈希(输入["A"]["路径"])
    初 = 开始(c, 输入)
    assert 等待(c, 初["task_id"])["等待"]
    说明 = json.loads((Path(初["运行目录"]) / "任务说明.json").read_text(encoding="utf-8"))
    assert "类型" not in 说明["原始输入"]["A"]
    assert "api_key" not in json.dumps(说明)
    重用 = c.post("/api/桌面/任务", json={"请求标识": "开始一", "原始输入": {**输入, "复用": True}})
    assert 重用.status_code == 409
    索引文件 = Path(初["运行目录"]) / "接续索引.json"
    前 = 索引文件.read_bytes()
    Path(输入["金标准"]).write_text("改过的依据", encoding="utf-8")
    r = c.post(f'/api/桌面/任务/{初["task_id"]}/继续', json={"请求标识": "继续"})
    assert r.status_code == 409 and "内容已变" in r.text
    assert 索引文件.read_bytes() == 前
    assert 文件哈希(输入["A"]["路径"]) == 原件前


def test_不把任意路径和验收目录当历史任务(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    验收 = web.运行根 / "验收" / "任务_内部"
    验收.mkdir(parents=True)
    写索引(验收, {"步骤": "done"})
    assert c.get("/api/桌面/任务").json()["总数"] == 0
    assert c.get("/api/桌面/任务/不存在/进度").status_code == 404
