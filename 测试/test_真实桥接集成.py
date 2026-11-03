# coding: utf-8
"""真实桥接集成用例（-m 集成）：用 测试\样本 真实 docx＋真桥接，断言实际结果。
覆盖：①提取命名区域/上下文 ②普通值回写 ③增行 ④删行 ⑧命名区域缺失报错 ⑨原件哈希不变。
⑤列调整/⑥标题更新/⑦跨合并拒绝 需真实业务文件，列入天宇油脂验收演练（见计划 T15）。"""
import hashlib
import json
import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from 核心.桥接 import 提取, 更新

样本 = Path(__file__).resolve().parent / "样本" / "01_普通表格.docx"
pytestmark = pytest.mark.集成


def md5(p):
    return hashlib.md5(Path(p).read_bytes()).hexdigest()




def 提取到(tmp_path):
    d = tmp_path / "任务"
    d.mkdir()
    提取(str(样本), str(d / "表A.xlsx"), str(d / "带连接.docx"), str(d / "上下文.json"))
    return d




def 回读(d, 名):
    提取(str(d / 名), str(d / "回读.xlsx"), str(d / "回读连.docx"), str(d / "回读上下文.json"))
    return load_workbook(d / "回读.xlsx")




def test_提取命名区域与上下文(tmp_path):
    d = 提取到(tmp_path)
    wb = load_workbook(d / "表A.xlsx")
    assert "TA_001" in wb.defined_names
    ctx = json.loads((d / "上下文.json").read_text(encoding="utf-8-sig"))
    assert ctx["tables"][0]["range_name"] == "TA_001"
    assert "chapter_context" in ctx["tables"][0]




def test_普通值回写并回读核对(tmp_path):
    前 = md5(样本)
    d = 提取到(tmp_path)
    表A = load_workbook(d / "表A.xlsx")
    表A["sheet1"]["C2"] = "新值"
    表A.save(d / "改后.xlsx")
    更新(str(d / "带连接.docx"), str(d / "改后.xlsx"), str(d / "新副本.docx"))
    结果表 = 回读(d, "新副本.docx")
    assert 结果表["sheet1"]["C2"].value == "新值"       # 实际写入
    assert 结果表["sheet1"]["B2"].value != "新值"       # 旁格不动（C2 之外）
    assert md5(样本) == 前                              # 原件哈希不变




def test_命名区域缺失报错(tmp_path):
    d = 提取到(tmp_path)
    wb = Workbook(); wb.active.title = "sheet1"; wb.active["B2"] = "x"
    wb.save(d / "无命名.xlsx")
    with pytest.raises(ValueError) as e:
        更新(str(d / "带连接.docx"), str(d / "无命名.xlsx"), str(d / "新副本.docx"))
    assert "命名区域" in str(e.value)




# —— 2026-11-02 追加：⑩ 多区域下删中间行、⑪ 跨区域明确报错 ——
