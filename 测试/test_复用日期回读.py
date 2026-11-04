# coding: utf-8
"""复用解释里的期间日期：必须按标注位置从本次文件重读，绝不沿用上一份的值。

位置缺失就无法用本次文件重读；这时带着旧日期继续，会把上一年的日期写进标题。
"""
from openpyxl import Workbook

from 核心.表格 import 读快照
from 核心.复用接入 import 补本次日期
from 核心.复用证据 import 表视图


def _视图(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    for c, 文 in enumerate(("项目", "2026-01-01", "2026-12-31", "单位：元"), 1):
        ws.cell(1, c, 文)
    ws.cell(2, 1, "甲公司")
    ws.cell(2, 2, 1)
    ws.cell(2, 3, 2)
    p = tmp_path / "本次.xlsx"
    wb.save(p)
    wb.close()
    return 表视图(读快照(p), "sheet1", (1, 1, 4, 5))


def _解释(位置, 旧值=None):
    return {"列": [{"列": "C", "角色": "数据列", "期间角色": "本期期末",
                    "期间实际": 旧值, "期间实际位置": 位置}], "行": []}


def test_有位置时按本次文件重读(tmp_path):
    解释 = _解释("C1", 旧值="2024-12-31")
    assert 补本次日期(解释, _视图(tmp_path)) == ""
    assert 解释["列"][0]["期间实际"] == "2026-12-31"


def test_只有旧值没有位置必须拒绝(tmp_path):
    """位置缺失就无法重读；沿用旧值会把上一年写进标题。"""
    解释 = _解释(None, 旧值="2024-12-31")
    原因 = 补本次日期(解释, _视图(tmp_path))
    assert "位置" in 原因


def test_既无日期也无位置时放行(tmp_path):
    assert 补本次日期(_解释(None), _视图(tmp_path)) == ""


def test_位置本次为空必须拒绝(tmp_path):
    原因 = 补本次日期(_解释("C5"), _视图(tmp_path))
    assert "没有内容" in 原因


def test_位置本次不是日期必须拒绝(tmp_path):
    原因 = 补本次日期(_解释("A1"), _视图(tmp_path))
    assert "不是日期" in 原因
