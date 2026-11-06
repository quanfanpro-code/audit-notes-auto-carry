# coding: utf-8
from pathlib import Path
import pytest
from 测试.GUI样本 import 准备, 开始, 等待, 核验金额
from 测试.test_Excel公式 import _模型


@pytest.mark.parametrize("中断侧", ["A", "B"])
def test_两端答复中断后使用本任务材料完成真实交付(tmp_path, monkeypatch, 中断侧):
    c, 输入 = 准备(tmp_path, monkeypatch, 模型=False)
    任务 = 开始(c, 输入)
    基址 = f'/api/桌面/任务/{任务["task_id"]}'
    已中断 = False
    答 = _模型()
    for 次 in range(12):
        末 = 等待(c, 任务["task_id"])
        if 末["步骤"] == "done":
            break
        assert 末["等待"], 末
        材料 = c.get(基址 + "/答复").json()["条目"][0]
        文 = c.get(基址 + "/答复/" + 材料["id"]).json()
        原 = (Path(文["路径"]) / "材料.txt").read_bytes()
        if not 已中断 and 末["步骤"] == "identify_" + 中断侧.lower():
            assert c.post(基址 + "/中断").json()["显示状态"] == "用户中断"
            assert (Path(文["路径"]) / "材料.txt").read_bytes() == 原
            assert c.post(基址 + "/继续", json={"请求标识": "中断后继续"}).status_code == 200
            末 = 等待(c, 任务["task_id"])
            assert 末["等待"]
            已中断 = True
        请求 = {"请求标识": f"答复{次}", "原答复哈希": 文["答复哈希"],
                "答复": 答({}, 文["指令"], 文["材料"], Path(文["路径"]))}
        assert c.post(基址 + "/答复/" + 材料["id"], json=请求).status_code == 200
        assert c.post(基址 + "/继续", json={"请求标识": f"继续{次}"}).status_code == 200
    assert 已中断 and 等待(c, 任务["task_id"])["步骤"] == "done"
    核验金额(任务["运行目录"])
