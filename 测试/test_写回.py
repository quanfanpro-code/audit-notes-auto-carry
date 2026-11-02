# coding: utf-8
"""写回测试：期间标题随表更新，old_text 为完整原标题、所属表 range_name 正确。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from 核心.写回 import 构造标题更新


def test_期间标题随表更新且旧文完整():
    上下文 = {"tables": [{"range_name": "TA_003",
               "chapter_context": "……资产负债表日 2024年12月31日 ……"}]}
    A记录 = {"格": "s1!A1", "槽位": "T", "槽位名": "表前日期标题",
             "维度值": {"期间角色": "期末", "期间实际": "2024-12-31"}}
    对应 = [{"a格": "s1!A1", "b格": "B1", "状态": "对应", "原因": None, "采用来源日期": "2025-12-31"}]
    ups, 失败 = 构造标题更新(对应, [A记录], 上下文)
    assert 失败 == []
    assert ups[0]["range_name"] == "TA_003"
    assert "2024年12月31日" in ups[0]["old_text"]            # 完整原标题在 old_text
    assert "2025年12月31日" in ups[0]["text"] and "2024" not in ups[0]["text"]


def test_标题找不到唯一原句就报失败():
    上下文 = {"tables": [{"range_name": "TA_003", "chapter_context": "本表没有任何日期"}]}
    A记录 = {"格": "s1!A1", "科目": "应收账款", "表名": "按账龄披露",
             "维度值": {"期间角色": "本期期末", "期间实际": "2024-12-31"}}
    对应 = [{"a格": "s1!A1", "b格": "B1", "状态": "对应", "采用来源日期": "2025-12-31"}]
    ups, 失败 = 构造标题更新(对应, [A记录], 上下文)
    assert ups == []                                  # 不许硬改
    assert 失败 and 失败[0]["原因"] == "原句找不到"     # 明确交代，好让调用方整表不换数
