# coding: utf-8
import json
from pathlib import Path

from 核心.接续 import 读索引, 写索引
from 测试.GUI样本 import 准备, 开始, 等待, 核验金额
from 测试.组合样本 import 文件哈希


def test_已完成任务拒绝继续且已有真实成果不变(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    原件 = {侧: 文件哈希(输入[侧]["路径"]) for 侧 in ("A", "B")}
    初 = 开始(c, 输入)
    assert 等待(c, 初["task_id"])["步骤"] == "done"
    目录 = Path(初["运行目录"])
    输出 = 核验金额(目录)
    成品哈希 = 文件哈希(输出)
    说明 = (目录 / "任务说明.json").read_bytes()
    基址 = f'/api/桌面/任务/{初["task_id"]}'
    # 完成后的其他结论仅检查入口分流；实际金额由上面的真实Excel任务核验。
    for 步骤 in ("done", "部分保留", "公式待核", "回读不一致"):
        索引 = 读索引(目录)
        索引["步骤"] = 步骤
        写索引(目录, 索引)
        原索引 = (目录 / "接续索引.json").read_bytes()
        进度 = c.get(基址 + "/进度").json()
        响应 = c.post(基址 + "/继续", json={"请求标识": "误续跑" + 步骤})
        if 响应.status_code == 200:
            等待(c, 初["task_id"])
        assert 响应.status_code == 409, 响应.text
        assert not 进度["允许继续"] and "新建" in 进度["不能继续原因"]
        assert (目录 / "接续索引.json").read_bytes() == 原索引
        assert (目录 / "任务说明.json").read_bytes() == 说明
        assert 文件哈希(输出) == 成品哈希
    旧入口 = c.post("/api/任务", json={**输入, "续跑": True})
    assert 旧入口.status_code == 409
    assert 文件哈希(输出) == 成品哈希
    # 丢失响应后重发同一次开始请求仍返回原任务，不再执行。
    assert 开始(c, 输入) == 初
    assert {侧: 文件哈希(输入[侧]["路径"]) for 侧 in ("A", "B")} == 原件
    assert json.loads(说明)["执行次数"] == 1
    核验金额(目录)


def test_真实交付留证清空仍明确提示变化且成品不变(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    初 = 开始(c, 输入)
    assert 等待(c, 初["task_id"])["步骤"] == "done"
    目录 = Path(初["运行目录"])
    输出 = 核验金额(目录)
    原 = 文件哈希(输出)
    明细 = 目录 / "交付明细.jsonl"
    assert 明细.stat().st_size > 0
    明细.write_bytes(b"")
    数据 = c.get(f'/api/桌面/任务/{初["task_id"]}/结果', params={"view": "明细"}).json()
    assert "已变化" in 数据["明细完整性"]
    assert any("已变化" in x for x in 数据["提示"])
    assert 数据["条目"] and all(not x["A实际值"]["是否取得"] for x in 数据["条目"])
    assert 数据["计数"].get("已写入", 0) == 0
    assert 文件哈希(输出) == 原
    核验金额(目录)
