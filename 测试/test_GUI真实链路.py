# coding: utf-8
"""GUI共同入口真实交付；模型固定，文件写入、Excel重算和查询均用产品实现。"""
import json
from pathlib import Path
from 测试.GUI样本 import 准备, 开始, 等待
from 测试.组合样本 import 读Excel格, 文件哈希
from 测试.test_Excel公式 import _多表, _多表金, _多表模型, _多表输入, _改格
from 测试.test_复核五项修复 import _标题任务, 标题新文A
from 核心.接续 import 读索引


def 成品(c, r, 状态="done"):
    实得 = 等待(c, r["task_id"])
    assert 实得["步骤"] == 状态, 实得
    p = Path(r["运行目录"])
    return Path(读索引(p)["产物"]["A端新副本"])


def test_GUI多工作表实际输出与展示一致(tmp_path, monkeypatch):
    c, _ = 准备(tmp_path, monkeypatch)
    A, B = _多表(tmp_path / "多表A.xlsx"), _多表(tmp_path / "多表B.xlsx", 来源=True)
    输入 = _多表输入(A, B, _多表金(tmp_path / "多表金.jsonl"))
    monkeypatch.setattr("核心.识别.调用识别", _多表模型())
    前 = [文件哈希(A), 文件哈希(B)]
    r = 开始(c, 输入)
    p = 成品(c, r)
    assert [读Excel格(p, "资金", x) for x in ("B2", "C2", "D2")] == ["资金甲", 901, 102]
    assert [读Excel格(p, "往来", x) for x in ("B2", "C2", "D2")] == ["往来甲", 701, 710]
    assert 读Excel格(p, "未选", "A1") == "别动"
    assert 读Excel格(p, "未选", "B1", 公式=True) == "=资金!C2+1"
    行 = c.get(f'/api/桌面/任务/{r["task_id"]}/结果?view=明细').json()["条目"]
    assert {x["A位置"]["工作表"] for x in 行} == {"资金", "往来"}
    for x in 行:
        if x["A实际值"]["是否取得"]:
            v = 读Excel格(p, x["A位置"]["工作表"], x["A位置"]["格"])
            assert x["A实际值"]["原样文本"] == ("" if v is None else str(v))
    assert 前 == [文件哈希(A), 文件哈希(B)]


def test_GUI公式差额显示真实200对210(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    _改格(输入["A"]["路径"], {"C2": 10, "C3": 20})
    _改格(输入["B"]["路径"], {"C2": 120, "C3": 80, "C4": 210})
    前 = [文件哈希(输入[x]["路径"]) for x in ("A", "B")]
    r = 开始(c, 输入)
    p = 成品(c, r, "公式待核")
    assert 读Excel格(p, "资金", "C4") == 200
    assert 读Excel格(p, "资金", "C4", 公式=True) == "=SUM(C2:C3)"
    行 = c.get(f'/api/桌面/任务/{r["task_id"]}/结果?view=公式').json()["条目"]
    条 = next(x for x in 行 if x["a格"] == "资金!C4")
    assert [条[x]["原样文本"] for x in ("A计算值", "B核对值", "差额值")] == ["200", "210", "-10"]
    assert 条["结论"] == "金额不一致"
    assert 前 == [文件哈希(输入[x]["路径"]) for x in ("A", "B")]


def test_GUI连续三份复用调用数与标题金额一致(tmp_path, monkeypatch):
    c, _ = 准备(tmp_path, monkeypatch)
    输入, 模型 = _标题任务(tmp_path, "GUI复用三份")
    monkeypatch.setenv("审计附注_识别记录目录", 输入["识别记录目录"])
    monkeypatch.setenv("审计附注_复用", "1")
    次数 = []
    def 计次(配置, 指令, 材料, 目录):
        次数.append(str(目录))
        return 模型(配置, 指令, 材料, 目录)
    monkeypatch.setattr("核心.识别.调用识别", 计次)
    for n in range(3):
        前次 = len(次数)
        r = 开始(c, 输入, f"第{n}份")
        p = 成品(c, r)
        assert [读Excel格(p, "正常", x) for x in ("B1", "B3", "C3", "D3")] == [标题新文A, "本次甲", 901, 102]
        复用 = c.get(f'/api/桌面/任务/{r["task_id"]}/结果?view=复用').json()["条目"]
        assert sum(sum(x["调用"].values()) for x in 复用) == len(次数) - 前次
        if n == 2:
            assert len(次数) == 前次
            assert all(all(x["调用"][k] == 0 for k in ("第一遍", "第二遍", "第三遍", "重试")) for x in 复用)
    (tmp_path / "实际模型调用.json").write_text(json.dumps(次数, ensure_ascii=False, indent=2), encoding="utf-8")
