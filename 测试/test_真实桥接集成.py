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
from 核心.结构维护 import 生成结构操作

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


def test_增行操作(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName
    d = 提取到(tmp_path)
    表A = load_workbook(d / "表A.xlsx")
    表A["sheet1"].insert_rows(3)
    表A["sheet1"]["B3"] = "新增"; 表A["sheet1"]["C3"] = "行值"
    # 模拟用户在 Excel 中插入行后命名区域随之扩展（openpyxl 不会自动扩）
    表A.defined_names["TA_001"] = DefinedName("TA_001", attr_text="sheet1!$B$2:$D$4")
    表A.save(d / "改后.xlsx")
    ops = 生成结构操作(d / "表A.xlsx", d / "改后.xlsx")
    assert ops == [{"range_name": "TA_001", "axis": "row", "action": "insert",
                    "index": 1, "count": 1, "template_index": 0}]
    更新(str(d / "带连接.docx"), str(d / "改后.xlsx"), str(d / "新副本.docx"), operations=ops)
    结果表 = 回读(d, "新副本.docx")
    assert 结果表["sheet1"]["B3"].value == "新增"       # 新行内容写入正确位置


def test_删行操作(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName
    d = 提取到(tmp_path)
    表A = load_workbook(d / "表A.xlsx")
    原行数 = 表A["sheet1"].max_row
    表A["sheet1"].delete_rows(3)
    # 模拟用户在 Excel 中删行后命名区域随之收缩
    表A.defined_names["TA_001"] = DefinedName("TA_001", attr_text="sheet1!$B$2:$D$2")
    表A.save(d / "改后.xlsx")
    ops = 生成结构操作(d / "表A.xlsx", d / "改后.xlsx")
    assert ops == [{"range_name": "TA_001", "axis": "row", "action": "delete",
                    "index": 1, "count": 1}]
    更新(str(d / "带连接.docx"), str(d / "改后.xlsx"), str(d / "新副本.docx"), operations=ops)
    结果表 = 回读(d, "新副本.docx")
    assert 结果表["sheet1"].max_row < 原行数          # 行确实少了

    改后 = load_workbook(d / "改后.xlsx")["sheet1"]
    期望 = {坐标: 改后[坐标].value for 坐标 in ("A2", "B2", "C2", "A3", "B3")}
    实际 = 结果表["sheet1"]
    for 坐标, 值 in 期望.items():
        assert 实际[坐标].value == 值, f"{坐标} 回读不一致：期望 {值!r}，实际 {实际[坐标].value!r}"

    改后 = load_workbook(d / "改后.xlsx")["sheet1"]
    期望 = {坐标: 改后[坐标].value for 坐标 in ("A2", "B2", "C2", "A3", "B3")}
    实际 = 结果表["sheet1"]
    for 坐标, 值 in 期望.items():
        assert 实际[坐标].value == 值, f"{坐标} 回读不一致：期望 {值!r}，实际 {实际[坐标].value!r}"


def test_命名区域缺失报错(tmp_path):
    d = 提取到(tmp_path)
    wb = Workbook(); wb.active.title = "sheet1"; wb.active["B2"] = "x"
    wb.save(d / "无命名.xlsx")
    with pytest.raises(ValueError) as e:
        更新(str(d / "带连接.docx"), str(d / "无命名.xlsx"), str(d / "新副本.docx"))
    assert "命名区域" in str(e.value)


# —— 2026-11-02 追加：⑩ 多区域下删中间行、⑪ 跨区域明确报错 ——


def test_六十六区域下删中间表一行_序号相对该表(tmp_path):
    """一张工作表装 66 个命名区域：删中间某张表的行，序号必须相对该表自己的区域。"""
    from openpyxl.workbook.defined_name import DefinedName
    wb = Workbook(); ws = wb.active; ws.title = "sheet1"
    顶们 = {}
    for i in range(1, 67):
        顶 = 2 + (i - 1) * 5
        顶们[f"TA_{i:03d}"] = 顶
        ws[f"B{顶}"] = "项目"; ws[f"C{顶}"] = "期末余额"; ws[f"D{顶}"] = "期初余额"
        ws[f"B{顶 + 1}"] = f"明细{i}"; ws[f"C{顶 + 1}"] = i; ws[f"D{顶 + 1}"] = i * 2
        ws[f"B{顶 + 2}"] = "合计"; ws[f"C{顶 + 2}"] = i; ws[f"D{顶 + 2}"] = i * 2
        wb.defined_names[f"TA_{i:03d}"] = DefinedName(
            f"TA_{i:03d}", attr_text=f"sheet1!$B${顶}:$D${顶 + 2}")
    基 = tmp_path / "基准.xlsx"
    改 = tmp_path / "改后.xlsx"
    wb.save(基)
    shutil.copy2(基, 改)
    wb2 = load_workbook(改)
    wb2["sheet1"].delete_rows(顶们["TA_030"] + 1)     # 删该表中间那行
    wb2.save(改)
    ops = 生成结构操作(基, 改)
    assert ops == [{"range_name": "TA_030", "axis": "row", "action": "delete",
                    "index": 1, "count": 1}]          # 相对区首行得 1，相对工作表会得 152


def test_跨两个命名区域的改动明确报错(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName
    def 建(路径):
        wb = Workbook(); ws = wb.active; ws.title = "s1"
        for 行, 文 in ((1, "甲"), (2, "乙"), (3, "丙"), (4, "丁")):
            ws[f"A{行}"] = 文
        wb.defined_names["TA_001"] = DefinedName("TA_001", attr_text="s1!$A$1:$A$2")
        wb.defined_names["TA_002"] = DefinedName("TA_002", attr_text="s1!$A$3:$A$4")
        wb.save(路径); return 路径

    基 = 建(tmp_path / "基准.xlsx")
    改 = 建(tmp_path / "改后.xlsx")
    wb = load_workbook(改)
    wb["s1"].delete_rows(2, 2)             # 连删第 2、3 行 → 横跨 TA_001 与 TA_002
    wb.save(改)
    with pytest.raises(RuntimeError) as e:
        生成结构操作(基, 改)
    assert "区域" in str(e.value)
