# coding: utf-8
"""金标准三层索引测试（直接读真实主文件）。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from pathlib import Path
from 核心.金标准 import 加载金标准, 建索引, 科目清单, 表清单, 候选槽位, 编号存在

主文件 = Path(r"C:\Users\27651\Desktop\附注自动更新系统\金标准\精简新副本_20261101_010814\财务报表附注语义金标准_当前基准_以此为准.jsonl")


def test_真实主文件三层对数():
    金 = 加载金标准(主文件)
    assert len(金) == 12817
    索 = 建索引(金)
    assert len(科目清单(索)) == 103
    assert "应收账款" in 科目清单(索)
    表 = 表清单(索, "应收账款")
    assert len(表) == 18
    assert "按账龄披露" in 表


def test_候选槽位只带必要字段():
    索 = 建索引(加载金标准(主文件))
    候 = 候选槽位(索, "应收账款", "按账龄披露")
    assert len(候) == 57                       # 真实值
    for c in 候[:5]:
        assert set(c) == {"id", "槽位名", "行路径", "列路径", "别名"}
        assert c["id"].startswith("C-N054-T001")


def test_编号校验():
    索 = 建索引(加载金标准(主文件))
    assert 编号存在(索, "C-N054-T001-S001")
    assert not 编号存在(索, "C-N054-T001-S999")
    assert not 编号存在(索, "随便写的")


def test_不再有关键词挑片段():
    import 核心.金标准 as 模
    assert not hasattr(模, "相关片段"), "旧做法已废弃，不得回潮"


def test_顺序由金标准display_order决定():
    索 = 建索引(加载金标准(主文件))
    assert 科目清单(索)[0] == "货币资金"              # 全库 display_order 最小的科目
    assert 表清单(索, "应收账款")[0] == "按账龄披露"   # 该科目下 display_order 最小的表
