# coding: utf-8
"""搬运测试：0/真实空白/读取失败政策；基于提取A副本只改格，保格式与命名区域。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from 核心.搬运 import 生成更新表, 保存更新表


def test_搬运政策与格式保留(tmp_path):
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    ws["A1"] = 100; ws["A1"].number_format = "#,##0.00"
    ws["B1"] = 200; ws["C1"] = 300
    wb.defined_names.add(DefinedName("TA_001", attr_text="s1!$A$1:$C$1"))
    基 = tmp_path / "表A.xlsx"; wb.save(基)
    B记录 = {"格": "B1", "原值": 0, "原值状态": "值"}          # 数字 0，不是文字
    B空白 = {"格": "B2", "原值": None, "原值状态": "真实空白"}
    B失败 = {"格": "B3", "原值": None, "原值状态": "读取失败"}
    对应 = [{"a格": "s1!A1", "b格": "B1", "状态": "对应", "原因": None, "采用来源日期": "2025-12-31"},
            {"a格": "s1!B1", "b格": "B2", "状态": "对应", "原因": None, "采用来源日期": "2025-12-31"},
            {"a格": "s1!C1", "b格": "B3", "状态": "对应", "原因": None, "采用来源日期": "2025-12-31"}]
    变更, 问题 = 生成更新表(
        {"sheets": [{"name": "s1", "rows": [[100, 200, 300]], "merges": []}]},
        对应, [B记录, B空白, B失败])
    assert 变更 == {"s1!A1": 0, "s1!B1": ""}        # 0 照搬、真实空白清空
    assert 问题 and 问题[0]["原因"] == "读取失败"    # 失败格不进变更
    out = tmp_path / "更新后附注表格.xlsx"
    保存更新表(基, 变更, out)
    新 = load_workbook(out)
    v = 新["s1"]["A1"].value
    assert v == 0 and isinstance(v, (int, float))              # 实际写入数字类型
    assert 新["s1"]["C1"].value == 300                         # 未变格保留
    assert 新["s1"]["A1"].number_format == "#,##0.00"          # 格式保留
    assert "TA_001" in 新.defined_names                        # 命名区域保留（回写前提）


def test_百分比类型与显示(tmp_path):
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    ws["A1"] = 0.8621; ws["A1"].number_format = "0.00%"        # 底层 0.8621，显示 86.21%
    基 = tmp_path / "表A.xlsx"; wb.save(基)
    变更, _ = 生成更新表(
        {"sheets": [{"name": "s1", "rows": [[0.8621]], "merges": []}]},
        [{"a格": "s1!A1", "b格": "B1", "状态": "对应", "原因": None, "采用来源日期": "2025-12-31"}],
        [{"格": "B1", "原值": 0.9, "原值状态": "值"}])
    out = tmp_path / "更新后附注表格.xlsx"
    保存更新表(基, 变更, out)
    格 = load_workbook(out)["s1"]["A1"]
    assert 格.value == 0.9 and isinstance(格.value, float)      # 数字按数字写
    assert 格.number_format == "0.00%"                          # 显示仍按百分比（90.00%）
