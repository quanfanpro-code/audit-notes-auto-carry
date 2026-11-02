# coding: utf-8
"""结构维护测试：先定位命名区域再算相对行号；区域外与跨区域明确失败。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from 核心.结构维护 import 生成结构操作, 区域归属


def 建表(路径, 行们, 区域=None):
    """行们：[(行号, A列文字, B列数字)]；区域：{名: 's1!$A$1:$B$3'}"""
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    for 行, 甲, 乙 in 行们:
        ws[f"A{行}"] = 甲; ws[f"B{行}"] = 乙
    for 名, 范围 in (区域 or {}).items():
        wb.defined_names[名] = DefinedName(名, attr_text=范围)
    wb.save(路径); return 路径




def test_只改名称直接放行(tmp_path):
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", [(1, "合计", 10), (2, "丙公司", 20), (3, "丁公司", 30)],
              {"TA_001": "s1!$A$1:$B$3"})
    assert 生成结构操作(基, 改) == []          # 无结构操作




def test_删行序号相对命名区域而非工作表(tmp_path):
    # TA_001 只覆盖第 2-4 行；删第 3 行 → index 应为 1（3−2），不是 2
    五行 = [(1, "抬头", 1), (2, "合计", 10), (3, "甲公司", 20), (4, "乙公司", 30), (5, "落款", 2)]
    基 = 建表(tmp_path / "基准.xlsx", 五行, {"TA_001": "s1!$A$2:$B$4"})
    改 = 建表(tmp_path / "改后.xlsx", 五行, {"TA_001": "s1!$A$2:$B$4"})
    wb2 = load_workbook(改); wb2["s1"].delete_rows(3); wb2.save(改)
    ops = 生成结构操作(基, 改)
    assert ops == [{"range_name": "TA_001", "axis": "row", "action": "delete",
                    "index": 1, "count": 1}]




def test_变更落在区域之外明确失败(tmp_path):
    五行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "游离行", 99), (4, "乙公司", 30), (5, "丙公司", 40)]
    区 = {"TA_001": "s1!$A$1:$B$2", "TA_002": "s1!$A$4:$B$5"}
    基 = 建表(tmp_path / "基准.xlsx", 五行, 区)
    改 = 建表(tmp_path / "改后.xlsx", 五行, 区)
    wb2 = load_workbook(改); wb2["s1"].delete_rows(3); wb2.save(改)   # 删的是区域外的游离行
    try:
        生成结构操作(基, 改)
        assert False, "应明确失败"
    except RuntimeError as e:
        assert "区域" in str(e)




def test_区域归属(tmp_path):
    p = 建表(tmp_path / "x.xlsx", [(1, "a", 1), (2, "b", 2), (3, "c", 3)],
             {"TA_001": "s1!$A$1:$B$2"})
    from 核心.表格 import 读快照
    快照 = 读快照(str(p))
    assert 区域归属(快照, "s1", 1, 2) == "TA_001"
    assert 区域归属(快照, "s1", 2, 1) == "TA_001"
    assert 区域归属(快照, "s1", 2, 2) is None      # 跨出区域




# —— R6：列变化明确拒绝、清空内容不算删行 ——



def test_加列明确报不支持(tmp_path):
    from openpyxl.styles import Font
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$C$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 行 in (1, 2, 3):
        ws[f"C{行}"] = 行 * 100                      # 真加了一列并填了值
        ws[f"C{行}"].font = Font(bold=True)
    wb2.save(改)
    try:
        生成结构操作(基, 改)
        assert False, "加列应明确报不支持，不能返回空操作"
    except RuntimeError as e:
        assert "列结构变化" in str(e)




def test_清空内容不算删行(tmp_path):
    from openpyxl.styles import Font
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 列 in ("A", "B"):
        ws[f"{列}3"] = None                          # 清空第三行内容，行本身还在
        ws[f"{列}3"].font = Font(bold=True)          # 样式仍在，行没有被删
    wb2.save(改)
    assert 生成结构操作(基, 改) == []                 # 没删行，就不产生删行操作




# —— R6：列变化明确拒绝、清空内容不算删行 ——


# —— 独立复核第三轮：空行插入、无样式清空、换列 ——

def test_中间插入空白行要认出增行(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName as DN
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    ws.insert_rows(2)                                  # 中间插一行空白，内容整体下移
    wb2.defined_names["TA_001"] = DN("TA_001", attr_text="s1!$A$1:$B$4")
    wb2.save(改)
    assert 生成结构操作(基, 改) == [{"range_name": "TA_001", "axis": "row", "action": "insert",
                                 "index": 1, "count": 1, "template_index": 0}]


def test_清空内容不带样式也不算删行(tmp_path):
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 列 in ("A", "B"):
        ws[f"{列}3"] = None                            # 只清空，不加任何样式
    wb2.save(改)
    assert 生成结构操作(基, 改) == []                   # 行还在，不许生成删行


def test_交换两列明确拒绝(tmp_path):
    三行 = [(1, "项目", "金额"), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 行 in (1, 2, 3):
        甲, 乙 = ws[f"A{行}"].value, ws[f"B{行}"].value
        ws[f"A{行}"], ws[f"B{行}"] = 乙, 甲            # 两列对调
    wb2.save(改)
    try:
        生成结构操作(基, 改)
        assert False, "换列应明确报不支持"
    except RuntimeError as e:
        assert "列结构变化" in str(e)


# —— 独立复核第三轮：空行插入、无样式清空、换列 ——

def test_中间插入空白行要认出增行(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName as DN
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    ws.insert_rows(2)                                  # 中间插一行空白，内容整体下移
    wb2.defined_names["TA_001"] = DN("TA_001", attr_text="s1!$A$1:$B$4")
    wb2.save(改)
    assert 生成结构操作(基, 改) == [{"range_name": "TA_001", "axis": "row", "action": "insert",
                                 "index": 1, "count": 1, "template_index": 0}]


def test_清空内容不带样式也不算删行(tmp_path):
    三行 = [(1, "合计", 10), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 列 in ("A", "B"):
        ws[f"{列}3"] = None                            # 只清空，不加任何样式
    wb2.save(改)
    assert 生成结构操作(基, 改) == []                   # 行还在，不许生成删行


def test_交换两列明确拒绝(tmp_path):
    三行 = [(1, "项目", "金额"), (2, "甲公司", 20), (3, "乙公司", 30)]
    基 = 建表(tmp_path / "基准.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    改 = 建表(tmp_path / "改后.xlsx", 三行, {"TA_001": "s1!$A$1:$B$3"})
    wb2 = load_workbook(改); ws = wb2["s1"]
    for 行 in (1, 2, 3):
        甲, 乙 = ws[f"A{行}"].value, ws[f"B{行}"].value
        ws[f"A{行}"], ws[f"B{行}"] = 乙, 甲            # 两列对调
    wb2.save(改)
    try:
        生成结构操作(基, 改)
        assert False, "换列应明确报不支持"
    except RuntimeError as e:
        assert "列结构变化" in str(e)
