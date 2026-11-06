# coding: utf-8
import json
from pathlib import Path

import pytest

from 服务 import web
from 核心.接续 import 读索引
from 测试.GUI样本 import 准备, 开始, 等待, 核验金额
from 测试.组合样本 import 造Word, 造Excel, 造金标准, 读Word格, 文件哈希
from 测试.test_Excel公式 import _模型


def 事实(目录):
    路径 = Path(目录) / "交付明细.jsonl"
    assert 路径.is_file(), "缺少最终交付事实，不能拿配对清单当写入结果"
    return [json.loads(x) for x in 路径.read_text(encoding="utf-8").splitlines()]


def test_真实Excel名单及公式记录的是最终结果(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    初 = 开始(c, 输入)
    assert 等待(c, 初["task_id"])["步骤"] == "done"
    核验金额(初["运行目录"])
    明细 = {x["a格"]: x for x in 事实(初["运行目录"]) if x["类型"] == "格子"}
    assert 明细["资金!C2"]["实际值"] == 901
    assert 明细["资金!C2"]["最终动作"] == "已写入"
    assert 明细["资金!B2"]["b格"] == "余额!B2"
    assert 明细["资金!B2"]["实际值"] == "=客户甲"
    assert 明细["资金!C4"]["最终动作"] == "公式保留"
    assert 明细["资金!C4"]["实际值"] == 1902
    assert 明细["资金!C4"]["公式核对"]["结论"] == "本次金额一致"


@pytest.mark.parametrize("单位", ["元", "万元"])
def test_真实Word写入与整组保留事实(tmp_path, monkeypatch, 单位):
    c, 输入 = 准备(tmp_path, monkeypatch)
    A = 造Word(tmp_path / "A.docx", 100, 110, 表内单位=True, 表前单位=False, 单位=单位)
    输入["A"] = {"路径": str(A), "工作表": []}
    B = 造Excel(tmp_path / "Word来源.xlsx", {"余额": [
        ["项目（单位：人民币元）", "2025-01-01", "2025-12-31"],
        ["本次甲", 901, 102], ["本次乙", 1001, 202], ["合计", 1902, 304]]}, 起列=2)
    输入["B"]["路径"] = str(B)
    模型 = _模型()
    def 答复(配置, 指令, 材料, 目录):
        答 = 模型(配置, 指令, 材料, 目录)
        if Path(目录).name.startswith("A") and 单位 == "万元":
            答 = 答.replace('人民币元', '人民币万元').replace('"单位": "元"', '"单位": "万元"').replace('"片段": "元"', '"片段": "万元"')
        return 答
    monkeypatch.setattr("核心.识别.调用识别", 答复)
    前 = 文件哈希(A)
    初 = 开始(c, 输入)
    等待(c, 初["task_id"])
    索引 = 读索引(初["运行目录"])
    输出 = 索引["产物"]["Word新副本"]
    预期 = "901" if 单位 == "元" else "100"
    assert 读Word格(输出, 0, 1, 1).replace(",", "") == 预期
    行 = next(x for x in 事实(初["运行目录"]) if x.get("a格") == "sheet1!C3")
    assert str(行["实际值"]).replace(",", "") == 预期
    assert 行["实际值已取得"] is True
    assert 行["最终动作"] == ("已写入" if 单位 == "元" else "原值保留")
    assert 文件哈希(A) == 前


def test_展示记录保存失败不重跑也不丢成品(tmp_path, monkeypatch):
    c, 输入 = 准备(tmp_path, monkeypatch)
    原写 = Path.write_text
    def 写(路径, *args, **kwargs):
        if 路径.name.startswith("交付明细"):
            raise OSError("本例注入展示记录磁盘故障")
        return 原写(路径, *args, **kwargs)
    monkeypatch.setattr(Path, "write_text", 写)
    初 = 开始(c, 输入)
    assert 等待(c, 初["task_id"])["步骤"] == "done"
    核验金额(初["运行目录"])
    索引 = 读索引(初["运行目录"])
    assert "磁盘故障" in 索引.get("展示错误", "")
    assert "交付明细" not in 索引["产物"]
