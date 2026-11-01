# coding: utf-8
"""Word 桥接封装测试：请求文件、响应解析、错误处理（桩程序）。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from 核心.桥接 import 提取


def test_请求文件与响应解析(tmp_path):
    # 假 exe：回显固定 JSON
    假exe = tmp_path / "假桥接.py"
    假exe.write_text("import sys,json;print(json.dumps({'ok':True,'action':'extract','table_count':2}))",
                     encoding="utf-8")
    req = {"word": "a.docx", "excel": "b.xlsx", "带连接word": "c.docx", "上下文": "d.json"}
    out = 提取(req["word"], req["excel"], req["带连接word"], req["上下文"],
              exe=假exe, 前缀=sys.executable, 工作目录=tmp_path)
    assert out["ok"] is True and out["table_count"] == 2
    # 真实冒烟（集成标记，见计划 Task 4 Step 4 用例）：
    # PYTHON -m pytest 测试/test_桥接.py -q -m 集成
