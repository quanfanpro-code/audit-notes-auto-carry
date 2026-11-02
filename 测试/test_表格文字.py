# coding: utf-8
"""表格快照与保结构转文字测试。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import Workbook
from openpyxl.utils.cell import coordinate_to_tuple
from 核心.表格 import 读快照
from 核心.表格文字 import 表格转文字


def 表头链(快照, 表名, 格):
    """测试助手：按合并结构机械算该格上方表头链，不做语义判断。"""
    sheet = next(s for s in 快照["sheets"] if s["name"] == 表名)
    row, col = coordinate_to_tuple(格)
    链 = []
    for r in range(1, row):
        for addr, cell in sheet["cells"].items():
            if cell["row"] == r and cell["column"] == col:
                链.append(str(cell.get("value")))
                break
    return tuple(链)




def test_表头归属与值挂对(tmp_path):
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    ws["A1"] = "行\\列"; ws["B1"] = "账面余额"; ws["C1"] = "坏账准备"
    ws["B2"] = "金额"; ws["C2"] = "计提比例"; ws["A3"] = "按单项计提"
    ws["B3"] = 110800; ws["C3"] = 100.00
    p = tmp_path / "t.xlsx"; wb.save(p)
    快照 = 读快照(p)
    # 值必须挂对表头：110800 属于"账面余额→金额"，100.00 属于"坏账准备→计提比例"
    assert 表头链(快照, "s1", "B3") == ("账面余额", "金额")
    assert 表头链(快照, "s1", "C3") == ("坏账准备", "计提比例")
    text = 表格转文字(快照, {"sheet": "s1", "range": "A1:C3"})
    assert "110800" in text and "按单项计提" in text
    assert "并入" not in text or "合并" in text   # 呈现合并结构信息
