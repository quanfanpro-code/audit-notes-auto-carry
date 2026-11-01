# coding: utf-8
"""接续索引测试：往返、原子写、待办步骤判断。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from 核心.接续 import 写索引, 读索引, 待办步骤


def test_索引往返与待办(tmp_path):
    写索引(tmp_path, {"步骤": "identify_a", "已完成范围": ["sheet1!A1:C3"],
                 "产物": {"表A": "表A.xlsx"}})
    idx = 读索引(tmp_path)
    assert idx["步骤"] == "identify_a"
    assert 待办步骤(idx) == "identify_a"      # 未完成步骤
    写索引(tmp_path, {**idx, "步骤": "done"})
    assert 待办步骤(读索引(tmp_path)) is None
