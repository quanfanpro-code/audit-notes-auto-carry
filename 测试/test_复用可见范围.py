# coding: utf-8
"""B 端可见范围：打印区域并集、隐藏行列、跨边界合并，三遍材料与合集共用同一集合。"""
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName

from 核心.表格 import 读快照, 取可见集
from 核心.识别 import 全文文字, 铺格子, 补原值, _完整性


def _填(ws, 行们, 列们, 前缀=""):
    for r in 行们:
        for c in 列们:
            ws.cell(r, c, f"{前缀}{ws.cell(r, c).coordinate}")


def _空字符串公式(xlsx, 工作表, 坐标):
    """把公式格写成 t="str"、<v></v>：openpyxl 只按原包类型才认得出空字符串缓存。"""
    from xml.etree import ElementTree as ET
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with ZipFile(xlsx) as z:
        部件 = {n: z.read(n) for n in z.namelist()}
    wbxml = ET.fromstring(部件["xl/workbook.xml"])
    rels = ET.fromstring(部件["xl/_rels/workbook.xml.rels"])
    关系 = {r.get("Id"): r.get("Target") for r in rels}
    rid = None
    for s in wbxml.find(f"{{{ns}}}sheets"):
        if s.get("name") == 工作表:
            rid = s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    目标 = 关系[rid]
    路径 = 目标.lstrip("/") if 目标.startswith("/") else (
        目标 if 目标.startswith("xl/") else "xl/" + 目标)
    树 = ET.fromstring(部件[路径])
    for c in 树.iter(f"{{{ns}}}c"):
        if c.get("r") != 坐标:
            continue
        c.set("t", "str")
        v = c.find(f"{{{ns}}}v")
        if v is None:
            v = ET.SubElement(c, f"{{{ns}}}v")
        v.text = ""
    部件[路径] = ET.tostring(树, encoding="UTF-8", xml_declaration=True)
    with ZipFile(xlsx, "w", ZIP_DEFLATED) as z:
        for 名, 数据 in 部件.items():
            z.writestr(名, 数据)


def _带打印区域(tmp_path, 区域, 行们=range(1, 6), 列们=range(1, 10)):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    _填(ws, 行们, 列们)
    ws.print_area = 区域
    p = tmp_path / "B.xlsx"
    wb.save(p)
    wb.close()
    return p


def test_不相邻打印区域忽略中间与隐藏内容(tmp_path):
    p = _带打印区域(tmp_path, ["A1:C5", "G1:I5"])
    wb = load_workbook(p)
    ws = wb["word附注"]
    ws["E2"] = "标记-两块之间的内容"
    ws["G4"] = "标记-隐藏行内容"
    ws["H2"] = "标记-隐藏列内容"
    ws.row_dimensions[4].hidden = True
    ws.column_dimensions["H"].hidden = True
    wb.save(p)
    wb.close()

    快照 = 读快照(p)
    可见 = 取可见集(快照, "word附注")
    assert 可见.含(2, 5) is False          # E2 在两块打印区域之间
    assert 可见.含(4, 7) is False          # 第 4 行隐藏
    assert 可见.含(2, 8) is False          # H 列隐藏
    assert 可见.含(2, 3) is True and 可见.含(2, 7) is True

    文 = 全文文字(快照, 只跑表=["word附注"], 用可见范围=True)
    assert "标记-两块之间的内容" not in 文
    assert "标记-隐藏行内容" not in 文
    assert "标记-隐藏列内容" not in 文
    assert "A2" in 文 and "G2" in 文


def test_重叠打印区域同一坐标只出现一次(tmp_path):
    p = _带打印区域(tmp_path, ["A1:C5", "B3:D6"])
    wb = load_workbook(p)
    ws = wb["word附注"]
    ws["B3"] = "重叠格-B3"
    wb.save(p)
    wb.close()

    快照 = 读快照(p)
    文 = 全文文字(快照, 只跑表=["word附注"], 用可见范围=True)
    assert 文.count("重叠格-B3") == 1
    # 并集：A1:C5（15 格）＋ B3:D6（12 格）－ 重叠 3 行×2 列（6 格）＝ 21 格，重复坐标只算一次
    assert 取可见集(快照, "word附注").格数() == 21


def test_跨界合并锚点可见只记一次(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    _填(ws, range(1, 5), range(1, 5))
    ws["A1"] = "跨组合并标题"
    ws.merge_cells("A1:C1")
    ws.print_area = "A1:C4"
    p = tmp_path / "合并.xlsx"
    wb.save(p)
    wb.close()

    文 = 全文文字(读快照(p), 只跑表=["word附注"], 用可见范围=True)
    assert 文.count("跨组合并标题") == 1
    assert "并入A1" in 文


def test_跨边界合并锚点隐藏不取正文(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    _填(ws, range(1, 6), range(1, 5))
    ws["A1"] = "隐藏锚点正文"
    ws.merge_cells("A1:C3")
    ws.row_dimensions[1].hidden = True
    ws.print_area = "A1:C5"
    p = tmp_path / "隐藏锚点.xlsx"
    wb.save(p)
    wb.close()

    快照 = 读快照(p)
    文 = 全文文字(快照, 只跑表=["word附注"], 用可见范围=True)
    assert "隐藏锚点正文" not in 文
    assert "起点不可见" in 文
    assert "范围外合并起点原文" not in 文          # 旧渲染的补取路径不得用于 B


def test_无打印区域只取可见内容并保留公式位置(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    ws["A1"] = "项目"
    ws["B1"] = "期末余额"
    ws["A2"] = "库存现金"
    ws["B2"] = 100
    ws["A3"] = "银行存款"
    ws["B3"] = 0
    ws["C5"] = '=IF(1=1,"","")'
    ws["D6"] = "   "                      # 纯空格不扩大范围
    ws["Z30"].font = Font(name="宋体")     # 纯格式不扩大范围
    ws.merge_cells("E10:F11")             # 孤立空白合并
    ws.merge_cells("A2:A3")               # 表内合并属结构证据
    wb.defined_names.add(DefinedName("本表", attr_text="'word附注'!$A$1:$C$5"))
    p = tmp_path / "无打印区域.xlsx"
    wb.save(p)
    wb.close()
    _空字符串公式(p, "word附注", "C5")

    快照 = 读快照(p)
    可见 = 取可见集(快照, "word附注")
    assert 可见.含(5, 3) is True            # 公式返回空白仍占位置
    assert 可见.含(6, 4) is False            # 纯空格
    assert 可见.含(30, 26) is False          # 纯格式
    assert 可见.含(10, 5) is False           # 孤立空白合并
    assert 可见.含(2, 7) is False            # 不铺到内容列以外（Z 列）
    assert 可见.含(3, 3) is True             # 表内空白（内容行×内容列之间）仍保留

    文 = 全文文字(快照, 只跑表=["word附注"], 用可见范围=True)
    assert "公式无保存结果" not in 文
    assert "缓存为空字符串" in 文

    格 = 补原值([{"格": "word附注!C5"}, {"格": "word附注!B2"}], 快照)
    assert 格[0]["原值"] == "" and 格[0]["原值状态"] == "值"
    assert 格[1]["原值"] == 100


def test_缩小打印区域后不可见内容退出合集与完整性(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    _填(ws, range(1, 7), range(1, 5))
    ws.print_area = "A1:D6"
    p1 = tmp_path / "大区域.xlsx"
    wb.save(p1)
    ws.print_area = "A1:B3"
    ws.column_dimensions["C"].hidden = True
    p2 = tmp_path / "小区域.xlsx"
    wb.save(p2)
    wb.close()

    大 = 读快照(p1)
    小 = 读快照(p2)
    assert 取可见集(大, "word附注").含(6, 4) is True
    assert 取可见集(小, "word附注").含(6, 4) is False

    结构 = {"工作表": "word附注", "侧": "B", "行范围": [1, 6], "范围": "A1:D6",
            "科目": "货币资金", "表名": "余额构成",
            "列": [{"列": "A", "路径": "项目", "角色": "行名列"},
                  {"列": "B", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末"}],
            "行": [{"行": 1, "角色": "表头"}] + [
                {"行": r, "角色": "数据", "槽位": {"B": "C-N087-T001-S001"}} for r in range(2, 7)],
            "作用范围": {}, "例外": []}
    问题 = []
    格们 = 补原值(铺格子(结构, 问题, 可见集=取可见集(小, "word附注")), 小, 用可见范围=True)
    数据格 = [g["格"] for g in 格们 if g["角色"] == "数据格"]
    assert 数据格 == ["word附注!B2", "word附注!B3"]     # 区域外不铺
    assert 结构["越界解释"] is True
    assert any("可见范围" in (x.get("原因") or "") for x in 问题)

    完整性 = _完整性([{"行范围": [1, 6], "科目": "货币资金"}],
                   [{"工作表": "word附注", "范围": "A1:D6", "行范围": [1, 6], "科目": "货币资金",
                     "表名": "余额构成", "完成": True,
                     "列": [{"列": "A"}, {"列": "B"}],
                     "行": [{"行": r} for r in range(1, 7)]}], 小, 用可见范围=True)
    assert 完整性["第三遍逐行"] is True      # 区域外的行不再算漏行
    assert 完整性["漏行"] == {}


def test_打印区域没有可见内容给明确原因(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    _填(ws, range(1, 4), range(1, 3))
    ws.print_area = "A1:C3"
    for r in (1, 2, 3):
        ws.row_dimensions[r].hidden = True
    p = tmp_path / "全隐藏.xlsx"
    wb.save(p)
    wb.close()

    快照 = 读快照(p)
    记录 = 快照["sheets"][0]["可见范围"]
    assert 记录["原因"] and "可见内容" in 记录["原因"]
    assert 记录["行"] == {}
    try:
        全文文字(快照, 只跑表=["word附注"], 用可见范围=True)
        assert False, "打印区域没有可见内容，应明确报错"
    except RuntimeError as e:
        assert "可见内容" in str(e)


def test_无法解析的打印区域不回读整表(tmp_path):
    """openpyxl 遇到越界打印区域会丢掉定义再整张读；这里必须自己发现并报错。"""
    import re
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    ws["A1"] = "内容"
    ws.print_area = "A1:B2"
    原 = tmp_path / "原区域.xlsx"
    wb.save(原)
    wb.close()
    with ZipFile(原) as z:
        部件 = {n: z.read(n) for n in z.namelist()}
    文 = 部件["xl/workbook.xml"].decode("utf-8")
    assert "'word附注'!$A$1:$B$2" in 文
    文 = 文.replace("'word附注'!$A$1:$B$2", "'word附注'!$ZZZZ$1:$ZZZZ$2")
    部件["xl/workbook.xml"] = 文.encode("utf-8")
    p = tmp_path / "坏区域.xlsx"
    with ZipFile(p, "w", ZIP_DEFLATED) as z:
        for 名, 数据 in 部件.items():
            z.writestr(名, 数据)
    try:
        读快照(p)
        assert False, "打印区域越界应报错，不得回读整表"
    except ValueError as e:
        assert "打印区域" in str(e)
