# coding: utf-8
import json
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from 测试.组合样本 import 造Word, 造金标准, 文件哈希, 读Excel格, 跑场景


def test_纯数字算式分类():
    from 核心.Excel交付 import 是纯数字算式

    assert 是纯数字算式("=100+200")
    assert 是纯数字算式("=-(100+20)/2")
    assert not 是纯数字算式("=C2+C3")
    assert not 是纯数字算式("=SUM(100,200)")
    assert not 是纯数字算式("=TODAY()")
    assert not 是纯数字算式('="100"+200')


def _造A(路径):
    wb = Workbook()
    ws = wb.active
    ws.title = "资金"
    ws["B1"] = "项目（单位：人民币元）"
    ws["C1"] = "2025-01-01"
    ws["D1"] = "2025-12-31"
    ws["B2"] = "甲公司"
    ws["C2"] = "=100+200"
    ws["D2"] = "=100+200"
    ws["C2"].number_format = "0.00"
    ws["D2"].number_format = "0.00"
    ws["B3"] = "乙公司"
    ws["C3"] = 1000
    ws["D3"] = 200
    ws["B4"] = "合计"
    ws["C4"] = "=SUM(C2:C3)"
    ws["D4"] = "=SUM(D2:D3)"
    wb.save(路径)
    wb.close()
    return 路径


def _造B(路径):
    wb = Workbook()
    ws = wb.active
    ws.title = "余额"
    for 行, 值 in enumerate([["项目（单位：人民币元）", "2025-01-01", "2025-12-31"],
                         ["=客户甲", 901, 102], ["=客户乙", 1001, 202], ["合计", 1902, 304]], 1):
        for 列, 值2 in enumerate(值, 2):
            c = ws.cell(行, 列)
            c.value = 值2
            if isinstance(值2, str) and 值2.startswith("=") and 行 in (2, 3):
                c.data_type = "s"
    wb.save(路径)
    wb.close()
    return 路径


def _模型():
    def 调用(配置, 指令, 材料, 目录):
        实际 = 材料.split("【二、", 1)[0]
        顶 = 2 if "以下为原文件二维网格" in 实际 else 1
        合计 = "合计" in 实际
        尾项 = "丙公司" in 实际
        底 = 顶 + (3 if 合计 or 尾项 else 2)
        划分 = []
        if 顶 == 2:
            划分.append({"行范围": [1, 1], "科目": "非表格"})
        划分.append({"行范围": [顶, 底], "科目": "货币资金"})
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": 划分}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [顶, 底], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}]
        for i in range(1, 4 if 尾项 else 3):
            行.append({"行": 顶 + i, "名字": "丙公司" if i == 3 else f"名单{i}", "角色": "数据", "序号": i,
                      "槽位": {"C": "N001", "D": "N001"}})
        if 合计:
            行.append({"行": 顶 + 3, "名字": "合计", "角色": "合计", "序号": None,
                      "槽位": {"C": "N001", "D": "N001"}})
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": f"B{顶}", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


@pytest.mark.parametrize("B格式", ["excel", "word"])
def test_Excel接收真实来源_数字公式和等号文字(tmp_path, B格式, monkeypatch):
    A = _造A(tmp_path / "A.xlsx")
    B = (_造B(tmp_path / "B.xlsx") if B格式 == "excel" else
         造Word(tmp_path / "B.docx", 901, 102, 名单=("=客户甲", "=客户乙"),
               表内单位=True, 表前单位=False))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"] if B格式 == "excel" else []},
            "金标准": str(金), "复用": False}
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景(输入, tmp_path / "任务", _模型(), monkeypatch)
    输出 = 结果["输出"]
    assert 输出 and 输出.is_file(), 结果["状态"]
    assert 读Excel格(输出, "资金", "C2") == 901
    assert 读Excel格(输出, "资金", "D2") == 102
    assert (读Excel格(输出, "资金", "C3"), 读Excel格(输出, "资金", "D3")) == (1001, 202)
    assert 读Excel格(输出, "资金", "C4", 公式=True) == "=SUM(C2:C3)"
    assert 读Excel格(输出, "资金", "C4") == 1902
    wb = load_workbook(输出, data_only=False)
    try:
        assert wb["资金"]["B2"].value == "=客户甲" and wb["资金"]["B2"].data_type == "s"
        assert wb["资金"]["B3"].value == "=客户乙" and wb["资金"]["B3"].data_type == "s"
        assert wb["资金"]["C2"].number_format == "0.00"
    finally:
        wb.close()
    assert 前 == (文件哈希(A), 文件哈希(B))


def _多表(路径, 来源=False, 保护=False):
    wb = Workbook()
    wb.remove(wb.active)
    表们 = ("来源一", "来源二") if 来源 else ("资金", "往来")
    数值 = ((901, 102), (701, 710)) if 来源 else ((100, 110), (200, 210))
    名称 = (("资金甲", "资金乙"), ("往来甲", "往来乙")) if 来源 else (("甲公司", "乙公司"), ("甲公司", "乙公司"))
    for i, 名 in enumerate(表们):
        ws = wb.create_sheet(名)
        ws["B1"] = "项目（单位：人民币元）"
        ws["C1"] = "2025-01-01"
        ws["D1"] = "2025-12-31"
        for j in (2, 3):
            ws.cell(j, 2, 名称[i][j - 2])
            ws.cell(j, 3, 数值[i][0] + (j - 2) * 100)
            ws.cell(j, 4, 数值[i][1] + (j - 2) * 100)
    if not 来源:
        if 保护:
            wb["资金"]["C2"] = "=100+200"
            wb["资金"].protection.sheet = True
        ws = wb.create_sheet("未选")
        ws["A1"] = "别动"
        ws["B1"] = "=资金!C2+1"
    wb.save(路径)
    wb.close()
    return 路径


def _多表金(路径):
    行们 = [{"id": "N001", "note": {"name": "货币资金"}, "table": {"name": "余额构成"},
            "slot": {"name": "余额"}, "dimensions": [{"name": "period", "required": True}]},
           {"id": "N003", "note": {"name": "存货"}, "table": {"name": "存货明细"},
            "slot": {"name": "余额"}, "dimensions": [{"name": "period", "required": True}]}]
    Path(路径).write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们), encoding="utf-8")
    return Path(路径)


def _多表模型():
    def 调用(配置, 指令, 材料, 目录):
        目录 = Path(目录)
        序 = next(int(x[1:]) - 1 for x in 目录.parts if x in ("P001", "P002"))
        科目, 表名, 编号 = (("货币资金", "余额构成", "N001"), ("存货", "存货明细", "N003"))[序]
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [1, 3], "科目": 科目}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [1, 3], "科目": 科目, "表名": 表名,
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C1"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
        行 = [{"行": 1, "名字": "项目", "角色": "表头"}] + [
            {"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
             "槽位": {"C": 编号, "D": 编号}} for i in (2, 3)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def _多表输入(A, B, 金):
    return {"A": {"路径": str(A), "工作表": ["资金", "往来"]},
            "B": {"路径": str(B), "工作表": ["来源一", "来源二"]},
            "金标准": str(金), "复用": False}


def test_多表Excel接收且不影响用户工作簿(tmp_path, monkeypatch):
    import pythoncom
    import win32com.client

    A = _多表(tmp_path / "A.xlsx")
    B = _多表(tmp_path / "B.xlsx", 来源=True)
    金 = _多表金(tmp_path / "金标准.jsonl")
    用户文件 = tmp_path / "用户.xlsx"
    wb0 = Workbook()
    wb0.active["A1"] = "用户原值"
    wb0.save(用户文件)
    wb0.close()
    pythoncom.CoInitialize()
    用户app = win32com.client.DispatchEx("Excel.Application")
    用户app.Visible = False
    用户wb = 用户app.Workbooks.Open(str(用户文件))
    try:
        结果 = 跑场景(_多表输入(A, B, 金), tmp_path / "任务", _多表模型(), monkeypatch)
        assert 结果["状态"]["步骤"] == "done", 结果["状态"]
        assert 读Excel格(结果["输出"], "资金", "C2") == 901
        assert 读Excel格(结果["输出"], "往来", "C2") == 701
        assert 读Excel格(结果["输出"], "资金", "B2") == "资金甲"
        assert 读Excel格(结果["输出"], "往来", "B2") == "往来甲"
        assert 读Excel格(结果["输出"], "未选", "A1") == "别动"
        assert 读Excel格(结果["输出"], "未选", "B1", 公式=True) == "=资金!C2+1"
        assert 用户wb.Name == "用户.xlsx" and 用户wb.Worksheets(1).Range("A1").Value2 == "用户原值"
        assert 用户app.Workbooks.Count == 1
    finally:
        用户wb.Close(False)
        用户app.Quit()
        pythoncom.CoUninitialize()


def test_受保护组部分保留且独立组写入(tmp_path, monkeypatch):
    A = _多表(tmp_path / "A.xlsx", 保护=True)
    B = _多表(tmp_path / "B.xlsx", 来源=True)
    金 = _多表金(tmp_path / "金标准.jsonl")
    结果 = 跑场景(_多表输入(A, B, 金), tmp_path / "任务", _多表模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "部分保留", 结果["状态"]
    assert 读Excel格(结果["输出"], "资金", "C2", 公式=True) == "=100+200"
    assert 读Excel格(结果["输出"], "资金", "B2") == "甲公司"
    assert 读Excel格(结果["输出"], "往来", "C2") == 701
    assert 读Excel格(结果["输出"], "往来", "B2") == "往来甲"
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "受保护" in 问题 and "资金!B2" in 问题 and "P001:G1" in 问题


def _改格(路径, 改动):
    wb = load_workbook(路径)
    try:
        ws = wb.active
        for 格, 值 in 改动.items():
            ws[格] = 值
        wb.save(路径)
    finally:
        wb.close()


def _重算保存(路径):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    wb = None
    try:
        wb = app.Workbooks.Open(str(路径), UpdateLinks=0)
        app.CalculateFullRebuild()
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(False)
        app.Quit()
        pythoncom.CoUninitialize()


def _跑公式(tmp_path, monkeypatch, A改=None, B改=None, B格式="excel", 未选=False, 重算B=False):
    tmp_path.mkdir(parents=True, exist_ok=True)
    A = _造A(tmp_path / "A.xlsx")
    if A改:
        _改格(A, A改)
    if 未选:
        wb = load_workbook(A)
        ws = wb.create_sheet("未选")
        ws["A1"] = "未选原值"
        ws["B1"] = "=资金!C2+1"
        wb.save(A)
        wb.close()
    B = (_造B(tmp_path / "B.xlsx") if B格式 == "excel" else
         造Word(tmp_path / "B.docx", 901, 102, 名单=("客户甲", "客户乙"),
               表内单位=True, 表前单位=False))
    if B改:
        _改格(B, B改)
    if 重算B:
        _重算保存(B)
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"] if B格式 == "excel" else []},
            "金标准": str(金), "复用": False}
    return 跑场景(输入, tmp_path / "任务", _模型(), monkeypatch)


def _记录(结果, 格):
    路径 = 结果["任务"] / "公式核对记录.jsonl"
    记录 = [json.loads(x) for x in 路径.read_text(encoding="utf-8").splitlines() if x.strip()]
    return next(x for x in 记录 if x["a格"] == 格)


def test_公式合计一致且未选表不纳入核对(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch, 未选=True)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    记录 = _记录(结果, "资金!C4")
    assert 记录["A结果"] == 1902 and 记录["B金额"] == 1902
    assert 记录["结论"] == "本次金额一致"
    assert 读Excel格(结果["输出"], "资金", "C4") == 1902
    assert 读Excel格(结果["输出"], "资金", "C4", 公式=True) == "=SUM(C2:C3)"
    assert 读Excel格(结果["输出"], "未选", "B1", 公式=True) == "=资金!C2+1"
    全部 = (结果["任务"] / "公式核对记录.jsonl").read_text(encoding="utf-8")
    assert "未选!" not in 全部


def test_公式合计差异保留待核副本(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch,
                  A改={"C2": 10, "C3": 20}, B改={"C2": 120, "C3": 80, "C4": 210})
    assert 结果["状态"]["步骤"] == "公式待核"
    assert 读Excel格(结果["输出"], "资金", "C2") == 120
    assert 读Excel格(结果["输出"], "资金", "C3") == 80
    assert 读Excel格(结果["输出"], "资金", "C4") == 200
    assert 读Excel格(结果["输出"], "资金", "C4", 公式=True) == "=SUM(C2:C3)"
    记录 = _记录(结果, "资金!C4")
    assert 记录["A结果"] == 200 and 记录["B金额"] == 210 and 记录["差额"] == -10
    assert 记录["结论"] == "金额不一致"


def test_公式引错范围与偶然一致分开证明(tmp_path, monkeypatch):
    差异 = _跑公式(tmp_path / "差异", monkeypatch,
                  A改={"C2": 10, "C3": 20, "C4": "=SUM(C2:C2)"},
                  B改={"C2": 120, "C3": 80, "C4": 200})
    assert _记录(差异, "资金!C4")["A结果"] == 120
    assert _记录(差异, "资金!C4")["差额"] == -80
    assert 读Excel格(差异["输出"], "资金", "C4", 公式=True) == "=SUM(C2:C2)"

    相同 = _跑公式(tmp_path / "相同", monkeypatch,
                  A改={"C2": 10, "C3": 20, "C4": "=SUM(C2:C2)"},
                  B改={"C2": 120, "C3": 0, "C4": 120})
    记录 = _记录(相同, "资金!C4")
    assert 记录["结论"] == "本次金额一致" and "逻辑正确" not in json.dumps(记录, ensure_ascii=False)


def test_公式没有对应金额不能通过(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch, B格式="word")
    assert 结果["状态"]["步骤"] == "公式待核"
    for 格 in ("资金!C4", "资金!D4"):
        记录 = _记录(结果, 格)
        assert 记录["结论"] == "未能核对" and "没有对应金额" in 记录["原因"]
        assert 读Excel格(结果["输出"], "资金", 格.split("!")[1], 公式=True).startswith("=SUM")


def test_公式错误与返回空字符串分别列明(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch,
                  A改={"C4": "=C2/0", "D4": '=IF(D2=102,"",D2)'})
    错误 = _记录(结果, "资金!C4")
    空白 = _记录(结果, "资金!D4")
    assert 错误["结论"] == "计算错误" and 错误["A结果"] == "#DIV/0!"
    assert 空白["结论"] == "未能核对" and 空白["A结果"] == "" and "空字符串" in 空白["原因"]
    assert 读Excel格(结果["输出"], "资金", "C4", 公式=True) == "=C2/0"
    assert 读Excel格(结果["输出"], "资金", "D4", 公式=True) == '=IF(D2=102,"",D2)'


@pytest.mark.parametrize("重算B", [False, True])
def test_B公式无保存值或错误均不可用(tmp_path, monkeypatch, 重算B):
    结果 = _跑公式(tmp_path, monkeypatch, B改={"C4": "=C2+C3" if not 重算B else "=1/0"},
                  重算B=重算B)
    记录 = _记录(结果, "资金!C4")
    assert 记录["结论"] == "未能核对" and "B金额不可用" in 记录["原因"]
    assert 结果["状态"]["步骤"] == "公式待核"


def test_清尾后保留引用公式并列无对象(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch,
                  A改={"B4": "丙公司", "C4": "=C2", "D4": "=D2"},
                  B改={"B4": None, "C4": None, "D4": None})
    assert 读Excel格(结果["输出"], "资金", "B4") is None
    assert 读Excel格(结果["输出"], "资金", "C4", 公式=True) == "=C2"
    记录 = _记录(结果, "资金!C4")
    assert 记录["结论"] == "未能核对" and "没有对应金额" in 记录["原因"]


def test_外部依赖不按旧缓存通过(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch, A改={"C4": "='[外部.xlsx]Sheet1'!C2"})
    记录 = _记录(结果, "资金!C4")
    assert 记录["结论"] == "未能核对" and "外部依赖" in 记录["原因"]
    assert 记录["原公式"] == "='[外部.xlsx]Sheet1'!C2"
    # Excel保存时把外部文件链接规范化成编号；公式仍是外部引用，不是B端金额。
    assert 读Excel格(结果["输出"], "资金", "C4", 公式=True) == "=[1]Sheet1!C2"
    assert 结果["状态"]["步骤"] in ("公式待核", "回读不一致")


@pytest.mark.parametrize("案例,B改,待核", [
    ("尾差", {"C2": 0.1, "C3": 0.2, "C4": 0.3}, False),
    ("大金额", {"C2": 1000000000, "C3": 0, "C4": 1000000000.01}, True),
    ("整数", {"D4": 200.4}, True),
])
def test_公式金额按实际数值和固定尾差比较(tmp_path, monkeypatch, 案例, B改, 待核):
    A改 = {"C2": 0.1, "C3": 0.2} if 案例 == "尾差" else None
    if 案例 == "大金额":
        A改 = {"C2": 1000000000, "C3": 0}
    结果 = _跑公式(tmp_path, monkeypatch, A改=A改, B改=B改)
    记录 = _记录(结果, "资金!D4" if 案例 == "整数" else "资金!C4")
    assert (记录["结论"] == "金额不一致") == 待核


def test_单位冲突先于公式和名单写入(tmp_path, monkeypatch):
    结果 = _跑公式(tmp_path, monkeypatch, A改={"B1": "项目（单位：人民币万元）"})
    assert 读Excel格(结果["输出"], "资金", "C2", 公式=True) == "=100+200"
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "单位" in 问题


def test_Excel不可用负向故障保留原件和候选(tmp_path, monkeypatch):
    """只注入启动失败验证错误路径，不代表真实Excel成功验收。"""
    from 核心.Excel交付 import _实际交付
    from 核心.表格 import 读快照
    import win32com.client

    A = _造A(tmp_path / "A.xlsx")
    前 = 文件哈希(A)
    输出 = tmp_path / "候选.xlsx"
    monkeypatch.setattr(win32com.client, "DispatchEx", lambda *a: (_ for _ in ()).throw(RuntimeError("Excel不可用")))
    请求 = {"原件": str(A), "输出": str(输出), "变更": {"资金!C2": 901},
            "A快照": 读快照(A), "A结构": [{"工作表": "资金", "范围": "B1:D4", "表组": "G1"}],
            "公式清单": [], "标题变更": {}, "状态": str(tmp_path / "实例.json")}
    with pytest.raises(RuntimeError, match="Excel不可用"):
        _实际交付(请求)
    assert 文件哈希(A) == 前 and 输出.is_file()


def test_取消控制只终止本次工作进程(tmp_path, monkeypatch):
    """控制层负向用例；不使用模拟文件证明Excel成功。"""
    import threading
    import 核心.Excel交付 as 模块

    A = _造A(tmp_path / "A.xlsx")
    前 = 文件哈希(A)
    收到 = []

    class 假进程:
        def poll(self):
            return None

        def terminate(self):
            收到.append("terminate")

        def wait(self, timeout=None):
            return 0

        def kill(self):
            收到.append("kill")

    monkeypatch.setattr(模块.subprocess, "Popen", lambda *a, **k: 假进程())
    取消 = threading.Event()
    取消.set()
    with pytest.raises(RuntimeError, match="用户中断"):
        模块.交付Excel(A, tmp_path / "不交付.xlsx", {}, {}, [], [], {}, 取消=取消)
    assert 收到 == ["terminate"] and 文件哈希(A) == 前


def test_写入异常撤回后公式仍是公式(tmp_path, monkeypatch):
    """仅注入一次写入失败，Excel打开、写入、回滚、重算和保存仍真实执行。"""
    import 核心.Excel交付 as 模块
    from 核心.表格 import 读快照

    A = _造A(tmp_path / "A.xlsx")
    输出 = tmp_path / "撤回.xlsx"
    原写入 = 模块._写入

    def 写入(cell, 值):
        if cell.Address == "$D$2" and 值 == 102:
            raise RuntimeError("注入一次写入失败")
        return 原写入(cell, 值)

    monkeypatch.setattr(模块, "_写入", 写入)
    请求 = {"原件": str(A), "输出": str(输出),
            "变更": {"资金!C2": 901, "资金!D2": 102},
            "A快照": 读快照(A), "A结构": [{"工作表": "资金", "范围": "B1:D4", "表组": "G1"}],
            "公式清单": [], "标题变更": {}, "状态": str(tmp_path / "实例.json")}
    结果 = 模块._实际交付(请求)
    assert 结果["撤回组"] == ["G1"]
    assert 读Excel格(输出, "资金", "C2", 公式=True) == "=100+200"
    assert 读Excel格(输出, "资金", "D2", 公式=True) == "=100+200"
    wb = load_workbook(输出, data_only=False)
    try:
        assert wb["资金"]["C2"].data_type == "f" and wb["资金"]["D2"].data_type == "f"
    finally:
        wb.close()
