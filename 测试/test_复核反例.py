# coding: utf-8
"""独立复核反例：期间角色、金额单位、列坐标、续表，先复现再修。

每条都直指复核报告的具体缺陷，不靠“任务完成”或“回读差异为零”判定。
"""
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from openpyxl import Workbook

from 核心.表格 import 读快照
from 核心.复用证据 import 表视图, 建证据_精读, 核对精读, _抽解释, 重定位
from 核心.流程 import 运行任务
from 测试.test_复用证据 import _表, _记录
from 测试.安全目录 import 安全记录目录

样本 = Path(__file__).resolve().parent / "样本" / "01_普通表格.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_列号 = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5}
_字母 = {v: k for k, v in _列号.items()}


def _单组Word(tmp_path, 名="单组.docx"):
    """一张表、一句日期标题：与真实桥接用例同一套造法。"""
    with ZipFile(样本) as z:
        树 = E.fromstring(z.read("word/document.xml"))
    body = 树.find(W + "body")
    for 子 in list(body):
        body.remove(子)

    p = E.Element(W + "p")
    r = E.SubElement(p, W + "r")
    E.SubElement(r, W + "t").text = "金额单位：人民币元"
    body.append(p)

    tbl = E.Element(W + "tbl")
    pr = E.SubElement(tbl, W + "tblPr")
    w = E.SubElement(pr, W + "tblW")
    w.set(W + "w", "0")
    w.set(W + "type", "auto")
    grid = E.SubElement(tbl, W + "tblGrid")
    for _ in range(3):
        E.SubElement(grid, W + "gridCol").set(W + "w", "2400")
    for 行 in (["项目", "2025-01-01", "2025-12-31"], ["甲公司", 100, 110], ["乙公司", 200, 210]):
        tr = E.SubElement(tbl, W + "tr")
        for 值 in 行:
            tc = E.SubElement(tr, W + "tc")
            tcPr = E.SubElement(tc, W + "tcPr")
            E.SubElement(tcPr, W + "tcW").set(W + "w", "2400")
            cell = E.SubElement(tc, W + "p")
            rr = E.SubElement(cell, W + "r")
            E.SubElement(rr, W + "t").text = str(值)
    body.append(tbl)

    word = tmp_path / 名
    with ZipFile(样本) as zin, ZipFile(word, "w") as zout:
        for part in zin.infolist():
            内容 = (E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
                  if part.filename == "word/document.xml" else zin.read(part.filename))
            zout.writestr(part, 内容)
    return word


def _从起列造B(路径, 起列="B", 期初=100, 期末=110):
    """B 端业务表从指定列开始：A 列留空，模拟“不是从 A 列起”。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    起 = _列号[起列]
    ws.cell(1, 起, "项目")
    ws.cell(1, 起 + 1, "2025-01-01")
    ws.cell(1, 起 + 2, "2025-12-31")
    ws.cell(1, 起 + 3, "单位：元")
    ws.cell(2, 起, "甲公司")
    ws.cell(2, 起 + 1, 期初)
    ws.cell(2, 起 + 2, 期末)
    ws.cell(3, 起, "乙公司")
    ws.cell(3, 起 + 1, 期初 + 1)
    ws.cell(3, 起 + 2, 期末 + 1)
    wb.save(路径)
    wb.close()
    return 路径


def _单组金标准(tmp_path):
    行 = {"id": "N001", "note": {"name": "货币资金"}, "table": {"name": "余额构成"},
         "slot": {"name": "余额"}, "row_path": ["甲公司"], "column_path": ["期末余额"],
         "dimensions": [{"name": "period", "required": True}]}
    p = tmp_path / "金标准单组.jsonl"
    p.write_text(json.dumps(行, ensure_ascii=False), encoding="utf-8")
    return p


def _单组模型(起列="B"):
    """固定答复；B 侧列按起列走，A 侧仍是桥接提取后的 B:D。"""
    起 = _列号[起列]

    def 调用(配置, 指令, 材料, 目录):
        侧 = Path(目录).name[:1]
        顶 = 2 if 侧 == "A" else 1          # A 端第 1 行是表前段落，表格从第 2 行起
        范围 = [顶, 顶 + 2]
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": 范围, "科目": "货币资金"}]},
                              ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表 = {"行范围": 范围, "科目": "货币资金", "表名": "余额构成",
                 "表组": "G1", "续表序": 1}
            return json.dumps({"表": [表]}, ensure_ascii=False)
        列们 = ([{"列": "B", "路径": "项目", "角色": "行名列"},
               {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
               {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
              if 侧 == "A" else
              [{"列": _字母[起], "路径": "项目", "角色": "行名列"},
               {"列": _字母[起 + 1], "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                "期间实际": {"值": "2025-01-01", "位置": f"{_字母[起 + 1]}{顶}"}},
               {"列": _字母[起 + 2], "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                "期间实际": {"值": "2025-12-31", "位置": f"{_字母[起 + 2]}{顶}"}},
               {"列": _字母[起 + 3], "路径": "单位说明", "角色": "说明列"}])
        槽 = {列们[1]["列"]: "N001", 列们[2]["列"]: "N001"}
        行们 = [{"行": 顶, "角色": "表头"},
               {"行": 顶 + 1, "角色": "数据", "序号": 1, "槽位": dict(槽)},
               {"行": 顶 + 2, "角色": "数据", "序号": 2, "槽位": dict(槽)}]
        return json.dumps({"列": 列们, "行": 行们,
                          "作用范围": {"币种": "人民币", "单位": "元", "单位依据": [
                              {"来源": "上下文" if 侧 == "A" else "表内",
                               "位置": ["unit_text_candidates", 0] if 侧 == "A" else f"{_字母[起+3]}1",
                               "原文": "金额单位：人民币元" if 侧 == "A" else "单位：元",
                               "片段": "元", "币种": "人民币", "单位": "元",
                               "列": [列们[1]["列"], 列们[2]["列"]]}]},
                          "空名单": [], "例外": []}, ensure_ascii=False)
    return 调用


class 读回Word:
    """重新打开输出的 Word：读段落与表格，断言最终金额用。"""

    def __init__(self, 路径):
        with ZipFile(路径) as z:
            body = E.fromstring(z.read("word/document.xml")).find(W + "body")
        self.段落 = ["".join(p.itertext()).strip() for p in body.findall(W + "p")]
        self._表 = [[["".join(tc.itertext()) for tc in tr.findall(W + "tc")]
                   for tr in tbl.findall(W + "tr")] for tbl in body.findall(W + "tbl")]

    def 表格(self, 序=0):
        return self._表[序]


def _跑一份(目录, word, b表, 金, 记录目录=None):
    输入 = {"word": str(word), "b表": str(b表), "金标准": str(金),
            "b工作表": "word附注"}
    if 记录目录 is not None:
        输入["识别记录目录"] = str(记录目录)
    状态 = 运行任务(输入, 目录, {"测试": True})
    return {"任务": Path(目录), "状态": 状态,
            "Word": Path(状态["产物"]["Word新副本"]) if 状态.get("产物") else None}


def test_复用后非A列起表_写回位置与金额都正确(tmp_path, monkeypatch):
    """第一份建立解释；第二份结构不变、B 端从 B 列起、金额不同。

    要求：B 端第三遍确实复用（没靠重认绕过错列），第二份 Word 里期初 901、期末 102，
    相邻格不被误写。
    """
    monkeypatch.setenv("审计附注_复用", "1")
    monkeypatch.setenv("审计附注_识别记录目录", str(安全记录目录()))
    monkeypatch.setattr("核心.识别.调用识别", _单组模型("B"))
    金 = _单组金标准(tmp_path)

    一 = _跑一份(tmp_path / "任务一", _单组Word(tmp_path, "第一份.docx"),
               _从起列造B(tmp_path / "第一份B.xlsx", "B", 100, 110), 金)
    assert 一["状态"]["步骤"] == "done", 一["状态"].get("问题")

    二 = _跑一份(tmp_path / "任务二", _单组Word(tmp_path, "第二份.docx"),
               _从起列造B(tmp_path / "第二份B.xlsx", "B", 901, 102), 金)
    assert 二["状态"]["步骤"] == "done", 二["状态"].get("问题")

    记录 = json.loads((二["任务"] / "复用记录_B.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第三遍"] == 0, 记录["回退"]
    assert 记录["复用"], "B 端没有命中复用，这轮没复现错列缺陷"

    Word = 读回Word(二["Word"])
    行们 = Word.表格(0)
    assert 行们[1][1] == "901", f"期初格实际是 {行们[1][1]}"
    assert 行们[1][2] == "102", f"期末格实际是 {行们[1][2]}"
    assert 行们[2][1] == "902" and 行们[2][2] == "103"


class _视图桩:
    """重定位只用到矩形左上角，用它验证坐标往返。"""
    顶 = 1
    左 = 3


def _换列表(tmp_path, 名, 首行, 行们):
    p = _表(tmp_path, 名, 行们, 首行=首行)
    return p


def test_期间列互换后必须回退(tmp_path):
    """B 把期初、期末两列的日期对调：角色已变，必须重认，不能按旧角色搬。"""
    上年 = _换列表(tmp_path, "上年", ("项目", "2025-01-01", "2025-12-31", "单位：元"),
                [("甲公司", 100, 110), ("乙公司", 200, 210), ("合计", 300, 320)])
    证据 = 建证据_精读("A", [_记录()], 读快照(上年))

    本年 = _换列表(tmp_path, "本年", ("项目", "2026-12-31", "2026-01-01", "单位：元"),
                [("甲公司", 901, 102), ("乙公司", 902, 103), ("合计", 1803, 205)])
    出 = 核对精读(证据, 表视图(读快照(本年), "sheet1", (1, 1, 4, 4)))
    assert 出["结果"] == "未命中"
    assert "期间" in 出["原因"] or "日期" in 出["原因"]


def test_单位变化后必须回退(tmp_path):
    """Word 表前说明从“元”改成“万元”：口径差一万倍，必须重认。"""
    表 = _换列表(tmp_path, "同表", ("项目", "2025-01-01", "2025-12-31", "单位：元"),
              [("甲公司", 100, 110), ("乙公司", 200, 210), ("合计", 300, 320)])
    证据 = 建证据_精读("A", [_记录()], 读快照(表))
    上下文 = {"tables": [{"range_name": "TA_001",
                       "unit_text_candidates": ["金额单位：人民币万元"]}]}
    出 = 核对精读(证据, 表视图(读快照(表), "sheet1", (1, 1, 4, 4)),
                上下文=上下文, 区域="TA_001")
    assert 出["结果"] == "未命中"
    assert "单位" in 出["原因"]


def test_单位能从上下文机械确认时仍可复用(tmp_path):
    """表前说明写的就是“元”：单位对得上，日期也对得上，应当命中。"""
    表 = _换列表(tmp_path, "同表", ("项目", "2025-01-01", "2025-12-31", "单位：元"),
              [("甲公司", 100, 110), ("乙公司", 200, 210), ("合计", 300, 320)])
    上下文 = {"tables": [{"range_name": "TA_001",
                       "unit_text_candidates": ["金额单位：人民币元"]}]}
    证据 = 建证据_精读("A", [_记录()], 读快照(表), 上下文=上下文)
    出 = 核对精读(证据, 表视图(读快照(表), "sheet1", (1, 1, 4, 4)),
                上下文=上下文, 区域="TA_001")
    assert 出["结果"] == "唯一"


def test_非A列起表坐标往返不变():
    """业务表从 C 列开始：抽相对位置再还原，列号必须原样，不能多加起始列。"""
    记录 = _记录()
    记录["列"][0].update({"列": "C"})
    记录["列"][1].update({"列": "D", "期间实际位置": "D1"})
    记录["列"][2].update({"列": "E", "期间实际位置": "E1"})
    记录["列"][3].update({"列": "F"})
    记录["行"] = [{"行": 1, "名字": "项目", "角色": "表头"}]
    解释 = _抽解释([记录], 顶=1, 左=3)
    还原 = 重定位(解释, _视图桩())
    assert [c["列"] for c in 还原["列"]] == ["C", "D", "E", "F"]
    assert 还原["列"][1]["期间实际位置"] == "D1"


def test_跨Z列坐标往返不变():
    """四列整体落在 AA:AD：往返必须逐列一致（四项对四项）。"""
    记录 = _记录()
    记录["列"][0].update({"列": "AA"})
    记录["列"][1].update({"列": "AB", "期间实际位置": "AB1"})
    记录["列"][2].update({"列": "AC", "期间实际位置": "AC1"})
    记录["列"][3].update({"列": "AD"})
    记录["行"] = [{"行": 1, "名字": "项目", "角色": "表头"}]
    解释 = _抽解释([记录], 顶=1, 左=27)          # AA 起
    还原 = 重定位(解释, type("视图", (), {"顶": 1, "左": 27})())
    assert [c["列"] for c in 还原["列"]] == ["AA", "AB", "AC", "AD"]


def test_续表的行不能丢():
    """主表＋续表共四行：抽出的解释必须四行都在，不能只留主表两行。"""
    主 = _记录(行数=2, 起=1, 组="G1", 续表序=1)
    续 = _记录(行数=2, 起=3, 组="G1", 续表序=2)
    解释 = _抽解释([主, 续], 顶=1, 左=1)
    assert len(解释["行"]) == 4
    还原 = 重定位(解释, type("视图", (), {"顶": 1, "左": 1})())
    assert [r["行"] for r in 还原["行"]] == [1, 2, 3, 4]


def test_续表列顺序不同时必须整组重认():
    """续表列顺序与主表不同：不能沿用主表列解释。"""
    主 = _记录(行数=2, 起=1, 组="G1", 续表序=1)
    续 = _记录(行数=2, 起=3, 组="G1", 续表序=2)
    续["列"] = [dict(c) for c in 续["列"]]
    续["列"][1], 续["列"][2] = 续["列"][2], 续["列"][1]
    assert _抽解释([主, 续], 顶=1, 左=1).get("组结构一致") is False


def test_两个12月31日正常换年可复用(tmp_path):
    """期初、期末都是 12-31：月日不变、年差都是 1，必须复用。"""
    上年 = _表(tmp_path, "上年", [("甲公司", 100, 110), ("乙公司", 200, 210)],
             首行=("项目", "2024-12-31", "2025-12-31", "单位：元"))
    记录 = _记录(行数=3)
    记录["列"][1].update({"期间实际": "2024-12-31", "期间实际位置": "B1"})
    记录["列"][2].update({"期间实际": "2025-12-31", "期间实际位置": "C1"})
    证据 = 建证据_精读("A", [记录], 读快照(上年))
    本年 = _表(tmp_path, "本年", [("甲公司", 1, 2), ("乙公司", 3, 4)],
             首行=("项目", "2025-12-31", "2026-12-31", "单位：元"))
    出 = 核对精读(证据, 表视图(读快照(本年), "sheet1", (1, 1, 4, 3)))
    assert 出["结果"] == "唯一", 出["原因"]


def test_年末变年中必须重认(tmp_path):
    """年差都是 1，但 12-31 变成 06-30：期间性质变了，必须重认。"""
    上年 = _表(tmp_path, "上年", [("甲公司", 100, 110), ("乙公司", 200, 210)],
             首行=("项目", "2024-12-31", "2025-12-31", "单位：元"))
    记录 = _记录(行数=3)
    记录["列"][1].update({"期间实际": "2024-12-31", "期间实际位置": "B1"})
    记录["列"][2].update({"期间实际": "2025-12-31", "期间实际位置": "C1"})
    证据 = 建证据_精读("A", [记录], 读快照(上年))
    本年 = _表(tmp_path, "本年", [("甲公司", 1, 2), ("乙公司", 3, 4)],
             首行=("项目", "2025-06-30", "2026-06-30", "单位：元"))
    出 = 核对精读(证据, 表视图(读快照(本年), "sheet1", (1, 1, 4, 3)))
    assert 出["结果"] == "未命中" and "月日" in 出["原因"]


def test_普通文字表头与合计变化仍可复用(tmp_path):
    """表头是“期初余额/期末余额”文字，合计金额变了：仍应命中，不按金额判身份。"""
    上年 = _表(tmp_path, "上年", [("甲公司", 100, 110), ("合计", 100, 110)],
             首行=("项目", "期初余额", "期末余额", "单位：元"))
    记录 = _记录(行数=3)
    记录["列"][1].update({"期间实际": "", "期间实际位置": None})
    记录["列"][2].update({"期间实际": "", "期间实际位置": None})
    记录["行"][1]["名字"] = "甲公司"
    记录["行"][2]["名字"] = "合计"
    记录["行"][2]["角色"] = "合计"
    证据 = 建证据_精读("A", [记录], 读快照(上年))
    本年 = _表(tmp_path, "本年", [("甲公司", 1, 2), ("合计", 999, 888)],
             首行=("项目", "期初余额", "期末余额", "单位：元"))
    出 = 核对精读(证据, 表视图(读快照(本年), "sheet1", (1, 1, 4, 3)))
    assert 出["结果"] == "唯一", 出["原因"]


def test_记录目录不安全时停用复用但完整识别照常(tmp_path, monkeypatch):
    """报告第 10 项：网络共享目录被拒，只停记录簿；识别、搬运、写回全部照常。"""
    monkeypatch.setenv("审计附注_复用", "1")
    monkeypatch.delenv("审计附注_识别记录目录", raising=False)
    monkeypatch.setattr("核心.识别.调用识别", _单组模型("B"))
    金 = _单组金标准(tmp_path)
    任务 = _跑一份(tmp_path / "任务", _单组Word(tmp_path, "不记录.docx"),
                 _从起列造B(tmp_path / "不记录B.xlsx", "B", 901, 102), 金,
                 记录目录=Path(r"\\localhost\C$\识别记录"))
    assert 任务["状态"]["步骤"] == "done", 任务["状态"].get("问题")
    记录 = json.loads((任务["任务"] / "复用记录_B.json").read_text(encoding="utf-8"))
    assert 记录["库"]["可用"] is False
    assert "允许的位置" in (记录["库"]["原因"] or "")
    assert all(x.get("步骤") for x in 记录["回退"])      # 三遍都按未命中处理
    Word = 读回Word(任务["Word"])
    assert Word.表格(0)[1][1] == "901"                  # 完整识别把金额照常搬了
