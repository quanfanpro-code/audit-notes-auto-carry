# coding: utf-8
"""金标准三层索引测试（直接读真实主文件）。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from pathlib import Path
import pytest
from 核心.金标准 import 加载金标准, 建索引, 科目清单, 表清单, 候选槽位, 编号存在, 期间必填

主文件 = Path(__file__).resolve().parent.parent / "金标准" / "精简新副本_20261101_010814" / "财务报表附注语义金标准_当前基准_以此为准.jsonl"


def test_真实主文件三层对数():
    金 = 加载金标准(主文件)
    assert len(金) == 12806
    索 = 建索引(金)
    assert len(索["编号"]) == 9196
    assert len(科目清单(索)) == 103
    assert "应收账款" in 科目清单(索)
    表 = 表清单(索, "应收账款")
    assert len(表) == 17
    assert "明细" not in 表                  # 该表仅剩停用条目，不进入运行目录
    assert "按账龄披露" in 表




def test_候选槽位只带必要字段():
    索 = 建索引(加载金标准(主文件))
    候 = 候选槽位(索, "应收账款", "按账龄披露")
    assert len(候) == 68                       # 当前有效条目，停用号不进入候选
    for c in 候[:5]:
        assert set(c) == {"id", "槽位名", "行路径", "列路径", "别名", "值类型", "维度"}
        assert c["id"].startswith("C-N054-T001")
        assert isinstance(c["维度"], dict)      # {维度名: 是否必填}




def test_编号校验():
    索 = 建索引(加载金标准(主文件))
    assert 编号存在(索, "C-N054-T001-S001")
    assert not 编号存在(索, "C-N054-T001-S999")
    assert not 编号存在(索, "随便写的")




def test_编号存在不限C开头():
    """有效P类同样可用；曾用于旧测试的停用P号必须拒绝。"""
    索 = 建索引(加载金标准(主文件))
    assert 编号存在(索, "P-N003-T002-S001")
    assert not 编号存在(索, "P-N003-T001-S001")
    assert not 编号存在(索, "C-N999-T999-S999")  # 虚构




def test_期间必填按编号查():
    """实测：11,395 条 period 必填；1,422 条没有 period 维度。"""
    索 = 建索引(加载金标准(主文件))
    assert 期间必填(索, "C-N087-T001-S001") is True    # 库存现金余额，来源里期初期末都有
    assert 期间必填(索, "C-N054-T003-S002") is False   # 名称类文字槽位
    assert 期间必填(索, "C-N019-T013-S006") is False   # 款项的性质，维度只有 counterparty
    with pytest.raises(ValueError, match="无效槽位编号"):
        期间必填(索, "C-N999-T999-S999")




def test_薪酬槽位素材完整():
    """槽位名没有尖括号，行路径却有：识别只能靠素材判断，不能靠名字。"""
    索 = 建索引(加载金标准(主文件))
    候 = 候选槽位(索, "应付职工薪酬", "应付职工薪酬列示")
    c = next(x for x in 候 if x["id"] == "C-N049-T001-S001")
    assert c["行路径"] == ["<薪酬类别或类别集合>"]
    assert c["值类型"] == "monetary"
    assert c["维度"].get("period") is True




def test_债权投资素材完整():
    """整表没有文字槽位，债券名只在行标上：素材里必须看得见这一点。"""
    索 = 建索引(加载金标准(主文件))
    候 = 候选槽位(索, "债权投资", "债权投资的情况")
    c = next(x for x in 候 if x["id"] == "C-N014-T001-S001")
    assert c["行路径"] == ["<披露项目>"]
    assert c["值类型"] == "monetary"
    文字类 = [x for x in 候 if x["值类型"] == "text"]
    assert 文字类 == []                        # 这张表确实一条文字槽位都没有




def test_不再有关键词挑片段():
    import 核心.金标准 as 模
    assert not hasattr(模, "相关片段"), "旧做法已废弃，不得回潮"




def test_顺序由金标准display_order决定():
    索 = 建索引(加载金标准(主文件))
    assert 科目清单(索)[0] == "货币资金"              # 全库 display_order 最小的科目
    assert 表清单(索, "应收账款")[0] == "按账龄披露"   # 该科目下 display_order 最小的表
