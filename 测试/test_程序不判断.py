# coding: utf-8
"""铁律：程序不得参与任何判断。

这里四条用例各盯一处此前由程序替模型做决定的地方：
1. 结构操作只能执行给定清单，程序不再自己比较两份文件推断插行、删行；
2. 标题更新只能执行给定结论，程序不再自己找段落、自己换年份；
3. 期间角色只认七种完整写法，程序不再把「本期」「期初」两半拼起来；
4. 列标题写在哪一行必须由识别给，程序不再拿唯一表头行顶上。
"""
import json

import pytest


def test_结构操作只执行给定清单(tmp_path):
    from 核心.结构维护 import 读结构操作

    p = tmp_path / "结构操作.json"
    p.write_text(json.dumps([
        {"range_name": "TA_001", "axis": "row", "action": "insert",
         "index": 3, "count": 1, "template_index": 2},
        {"range_name": "TA_002", "axis": "row", "action": "delete",
         "index": 1, "count": 2},
    ], ensure_ascii=False), encoding="utf-8")

    出 = 读结构操作(p)
    assert 出[0]["range_name"] == "TA_001" and 出[0]["action"] == "insert"
    assert 出[1]["action"] == "delete" and 出[1]["count"] == 2


@pytest.mark.parametrize("坏", [
    {"axis": "row", "action": "insert", "index": 1, "count": 1},         # 缺区域
    {"range_name": "TA_001", "axis": "row", "action": "插行", "index": 1, "count": 1},
    {"range_name": "TA_001", "axis": "row", "action": "insert", "index": -1, "count": 1},
    {"range_name": "TA_001", "axis": "row", "action": "insert", "index": 1, "count": 0},
])
def test_结构操作格式不对就报错(tmp_path, 坏):
    from 核心.结构维护 import 读结构操作

    p = tmp_path / "结构操作.json"
    p.write_text(json.dumps([坏], ensure_ascii=False), encoding="utf-8")
    with pytest.raises(RuntimeError):
        读结构操作(p)


def test_程序不再自己推断结构变化():
    """连同名函数都不许留：留一个入口，早晚有人再把它接回流程。"""
    import 核心.结构维护 as m

    assert not hasattr(m, "生成结构操作")


def test_标题更新只执行给定结论():
    from 核心.写回 import 读标题更新

    结构 = [
        {"区域": "TA_008", "科目": "货币资金", "表名": "余额构成",
         "标题": {"原文": "期末指2024年12月31日", "新文": "期末指2025年12月31日"}},
        {"区域": "TA_009", "科目": "应收账款", "表名": "按账龄披露"},      # 没给结论
        {"区域": "TA_010", "科目": "应收账款", "表名": "按坏账计提方法分类披露",
         "标题": {"原文": "同一句", "新文": "同一句"}},                    # 没变化
    ]
    出 = 读标题更新(结构)
    assert 出 == [{"range_name": "TA_008",
                 "old_text": "期末指2024年12月31日", "text": "期末指2025年12月31日"}]


def test_程序不再自己找段落换年份():
    import 核心.写回 as m

    assert not hasattr(m, "构造标题更新")


@pytest.mark.parametrize("半截", ["期末", "期初", "本期", "上期"])
def test_期间只认七种完整写法(半截):
    from 核心.识别 import _期间值

    with pytest.raises(RuntimeError):
        _期间值(半截, "列 C 的")
    assert _期间值("本期期末", "列 C 的") == "本期期末"


def test_没给标签行不再自动补():
    """只有一层表头也不许程序自己顶上去：那是识别该说的话。"""
    from 核心.识别 import 铺格子

    结构 = {
        "工作表": "sheet1", "科目": "固定资产", "表名": "固定资产情况",
        "列": [{"列": "B", "路径": "项目", "角色": "行名列", "期间角色": None, "序号": None},
              {"列": "C", "路径": "房屋", "角色": "数据列", "期间角色": "本期期末", "序号": 1}],
        "行": [{"行": 1, "名字": "项目", "角色": "表头"},
              {"行": 2, "名字": "金额", "角色": "数据", "序号": None,
               "期间角色": None, "槽位": {"C": "C-N038-T001-S001"}}],
        "作用范围": {}, "例外": [],
    }
    问题 = []
    铺格子(结构, 问题)
    assert 结构["项标签缺失"] is True
    assert any("标签行" in (x.get("原因") or "") for x in 问题)
