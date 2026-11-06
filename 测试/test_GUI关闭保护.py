# coding: utf-8
"""窗口确认使用可控替身；任务识别和中断走真实入口，不冒充原生窗口验收。"""
import threading
from 桌面窗口 import 窗口桥
from 服务 import web
from 测试.GUI样本 import 准备, 开始, 等待
from 测试.组合样本 import 文件哈希


class 确认窗口:
    def __init__(self, 同意=True):
        self.同意 = 同意
        self.已关闭 = threading.Event()
        self.提示 = []

    def create_confirmation_dialog(self, 标题, 提示):
        self.提示.append(提示)
        return self.同意

    def set_title(self, 标题):
        self.提示.append(标题)

    def destroy(self):
        self.已关闭.set()


def test_有草稿拒绝退出后仍可操作(tmp_path, monkeypatch):
    准备(tmp_path, monkeypatch)
    桥 = 窗口桥()
    桥._窗口 = 确认窗口(False)
    桥.未保存 = True
    桥._确认退出()
    assert not 桥._窗口.已关闭.is_set()
    assert not web.应用.state.退出中
    assert "草稿" in 桥._窗口.提示[0]


def test_识别未退出时关窗必须等安全停止(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    import 核心.识别
    原模型 = 核心.识别.调用识别
    进入, 放行 = threading.Event(), threading.Event()
    def 等待模型(*参数):
        进入.set()
        assert 放行.wait(15)
        return 原模型(*参数)
    monkeypatch.setattr(核心.识别, "调用识别", 等待模型)
    前 = [文件哈希(输入[x]["路径"]) for x in ("A", "B")]
    桥 = 窗口桥()
    桥._窗口 = 确认窗口()
    try:
        r = 开始(c, 输入, "关闭中断")
        assert 进入.wait(10)
        assert 桥.关闭事件() is False
        assert web.应用.state.状态["取消"].wait(5)
        assert not 桥._窗口.已关闭.is_set()
        assert c.post("/api/桌面/任务", json={"请求标识": "退出时新任务", "原始输入": 输入}).status_code == 409
        放行.set()
        assert 等待(c, r["task_id"])["步骤"] == "用户中断"
        assert 桥._窗口.已关闭.wait(5)
        assert [文件哈希(输入[x]["路径"]) for x in ("A", "B")] == 前
    finally:
        放行.set()
        线程 = web.应用.state.状态.get("线程")
        if 线程:
            线程.join(15)
        web.应用.state.退出中 = False
