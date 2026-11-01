# coding: utf-8
"""提取编排测试：产物布局＋接续索引登记。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from pathlib import Path
from 核心 import 提取 as 提取模块
from 核心.接续 import 读索引


def test_产物布局与索引(tmp_path, monkeypatch):
    def 假桥接(word, excel, 带连接word, 上下文):
        Path(excel).write_text("x"); Path(带连接word).write_text("y"); Path(上下文).write_text("{}")
        return {"ok": True, "table_count": 1, "warnings": []}
    monkeypatch.setattr(提取模块, "调用桥接提取", 假桥接)
    out = 提取模块.提取任务("原.docx", tmp_path)
    assert Path(out["表A"]).is_file() and Path(out["带连接Word"]).is_file() and Path(out["上下文"]).is_file()
    idx = 读索引(tmp_path)
    assert idx["产物"]["表A"] == str(Path(out["表A"]))
    assert idx["步骤"] == "identify_a"          # 提取完成，下一步待办
