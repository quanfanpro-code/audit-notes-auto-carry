# coding: utf-8
"""独立复核五项修复的业务用例：全部走正常任务入口，只固定模型答复，
断言读最终文件（金额、名单、标题、公式），不以状态或调用数代替。"""
import json
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from 测试.组合样本 import 造Word, 造金标准, 文件哈希, 读Excel格, 读Word格, 跑场景


def _关复用(monkeypatch, 记录=None):
    """复用开关同时设环境变量和任务入参，环境变量优先，两者保持一致。"""
    monkeypatch.setenv("审计附注_复用", "0")
    monkeypatch.setenv("审计附注_识别记录目录", str(记录 or 安全记录目录()))
    return {"复用": False}


def _开复用(monkeypatch, 记录=None):
    monkeypatch.setenv("审计附注_复用", "1")
    monkeypatch.setenv("审计附注_识别记录目录", str(记录 or 安全记录目录()))
    return {"复用": True, "识别记录目录": str(记录)}


def 安全记录目录():
    from 测试.安全目录 import 安全记录目录 as 建目录

    return 建目录("五项修复")


def _简单簿(路径, 表名, 行们, 隐藏行=()):
    """三行小表：表头＋两行名单，B1 表内单位；行们为（名称, 期初, 期末）。"""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(表名)
    ws["B1"] = "项目（单位：人民币元）"
    ws["C1"] = "2025-01-01"
    ws["D1"] = "2025-12-31"
    for 行号, (名, 期初, 期末) in enumerate(行们, 2):
        ws.cell(行号, 2, 名)
        ws.cell(行号, 3, 期初)
        ws.cell(行号, 4, 期末)
    for 行 in 隐藏行:
        ws.row_dimensions[行].hidden = True
    wb.save(路径)
    wb.close()
    return 路径


def _简单模型():
    """两侧同为三行小表的固定答复。"""
    def 调用(配置, 指令, 材料, 目录):
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [1, 3], "科目": "货币资金"}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [1, 3], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C1"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
        行 = [{"行": 1, "名字": "项目", "角色": "表头"}] + [
            {"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
             "槽位": {"C": "N001", "D": "N001"}} for i in (2, 3)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def _简单输入(tmp_path, A, B, 关):
    金 = 造金标准(tmp_path / "金标准.jsonl")
    return {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"]},
            "金标准": str(金), **关}


# —— 问题 1：合法工作表名含“!” ——

def _四行簿(路径, 表名, 行们, 合计):
    """表头＋两行名单＋普通数值合计（不是公式）；行们为（名称, 期初, 期末）。"""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(表名)
    ws["B1"] = "项目（单位：人民币元）"
    ws["C1"] = "2025-01-01"
    ws["D1"] = "2025-12-31"
    for 行号, (名, 期初, 期末) in enumerate(行们, 2):
        ws.cell(行号, 2, 名)
        ws.cell(行号, 3, 期初)
        ws.cell(行号, 4, 期末)
    ws.cell(4, 2, "合计")
    ws.cell(4, 3, 合计[0])
    ws.cell(4, 4, 合计[1])
    wb.save(路径)
    wb.close()
    return 路径


def _四行模型():
    def 调用(配置, 指令, 材料, 目录):
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [1, 4], "科目": "货币资金"}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [1, 4], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C1"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
        行 = [{"行": 1, "名字": "项目", "角色": "表头"}]
        for i in (2, 3):
            行.append({"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
                      "槽位": {"C": "N001", "D": "N001"}})
        行.append({"行": 4, "名字": "合计", "角色": "合计", "序号": None,
                  "槽位": {"C": "N001", "D": "N001"}})
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def test_叹号表名_Excel接收金额名单合计正确(tmp_path, monkeypatch):
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A = _四行簿(tmp_path / "A.xlsx", "资金!明细",
              [("甲公司", 100, 110), ("乙公司", 200, 210)], (300, 320))
    B = _四行簿(tmp_path / "B.xlsx", "余额!本期",
              [("本次甲", 901, 102), ("本次乙", 1001, 202)], (1902, 304))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    前 = (文件哈希(A), 文件哈希(B))
    输入 = {"A": {"路径": str(A), "工作表": ["资金!明细"]},
            "B": {"路径": str(B), "工作表": ["余额!本期"]},
            "金标准": str(金), **关}
    结果 = 跑场景(输入, tmp_path / "任务", _四行模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "资金!明细", "B2") == "本次甲"
    assert 读Excel格(输出, "资金!明细", "C2") == 901
    assert 读Excel格(输出, "资金!明细", "D2") == 102
    assert 读Excel格(输出, "资金!明细", "B3") == "本次乙"
    assert 读Excel格(输出, "资金!明细", "C3") == 1001
    assert 读Excel格(输出, "资金!明细", "D3") == 202
    assert 读Excel格(输出, "资金!明细", "C4") == 1902
    assert 读Excel格(输出, "资金!明细", "D4") == 304
    assert 前 == (文件哈希(A), 文件哈希(B))


def test_叹号表名_真实Word接收(tmp_path, monkeypatch):
    from 测试.组合样本 import 跨格式模型
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A = 造Word(tmp_path / "A.docx")
    B = _简单簿(tmp_path / "B.xlsx", "余额!本期",
              [("本次甲", 901, 102), ("本次乙", 1001, 202)])
    前 = (文件哈希(A), 文件哈希(B))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": []},
            "B": {"路径": str(B), "工作表": ["余额!本期"]},
            "金标准": str(金), **关}
    结果 = 跑场景(输入, tmp_path / "任务", 跨格式模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert (读Word格(结果["输出"], 0, 1, 0), 读Word格(结果["输出"], 0, 1, 1),
            读Word格(结果["输出"], 0, 1, 2)) == ("本次甲", "901", "102")
    assert (读Word格(结果["输出"], 0, 2, 0), 读Word格(结果["输出"], 0, 2, 1),
            读Word格(结果["输出"], 0, 2, 2)) == ("本次乙", "1001", "202")
    回读 = (结果["任务"] / "回读记录.jsonl")
    assert not 回读.exists() or not [x for x in 回读.read_text(encoding="utf-8").splitlines() if x.strip()]
    assert 前 == (文件哈希(A), 文件哈希(B))


def test_不可见状态注入整项保留并列原因(tmp_path, monkeypatch):
    """负向故障注入：真实读取后仅把指定B数据格改成“不可见”，其余不改，
    验证来源异常整项保留，不当真实空白清空A。"""
    from 核心 import 识别 as 识别模块
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A = _简单簿(tmp_path / "A.xlsx", "资金",
              [("甲公司", 100, 110), ("乙公司", 200, 210)])
    B = _简单簿(tmp_path / "B.xlsx", "余额",
              [("本次甲", 901, 102), ("本次乙", 1001, 202)])
    原 = 识别模块.补原值

    def 包装(记录们, 快照, 用可见范围=False):
        出 = 原(记录们, 快照, 用可见范围)
        for 记录 in 出:
            if str(记录.get("格")) in ("余额!C2", "余额!D2"):   # 注入边界：仅B这两个数据格
                记录["原值"] = None
                记录["原值状态"] = "不可见"
        return 出

    monkeypatch.setattr(识别模块, "补原值", 包装)
    结果 = 跑场景(_简单输入(tmp_path, A, B, 关), tmp_path / "任务", _简单模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "资金", "B2") == "甲公司"
    assert 读Excel格(输出, "资金", "C2") == 100
    assert 读Excel格(输出, "资金", "D2") == 110
    assert 读Excel格(输出, "资金", "B3") == "本次乙"
    assert 读Excel格(输出, "资金", "C3") == 1001
    assert 读Excel格(输出, "资金", "D3") == 202
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "不可见" in 问题
    问题们 = [json.loads(x) for x in (结果["任务"] / "问题清单.jsonl").read_text(
        encoding="utf-8").splitlines() if x.strip()]
    对应 = [json.loads(x) for x in (结果["任务"] / "对应清单.jsonl").read_text(
        encoding="utf-8").splitlines() if x.strip()]
    对应a格 = {x.get("a格"): x for x in 对应 if x.get("状态") == "对应"}
    条 = next(x for x in 问题们 if "不可见" in str(x.get("原因")) and x.get("a格") in 对应a格)
    条目 = 对应a格[条["a格"]]
    assert str(条目["b格"]).startswith("余额!")


def test_范围外来源A对应项保留独立项正常(tmp_path, monkeypatch):
    """真实范围限制：B 来源一第2行隐藏，固定答复仍带入该行，真实读取按可见集合拒绝。"""
    from 测试.test_Excel公式 import _多表, _多表金, _多表模型, _多表输入
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A = _多表(tmp_path / "A.xlsx")
    B = _多表(tmp_path / "B.xlsx", 来源=True)
    wb = load_workbook(B)
    wb["来源一"].row_dimensions[2].hidden = True
    wb.save(B)
    wb.close()
    金 = _多表金(tmp_path / "金标准.jsonl")
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景({**_多表输入(A, B, 金), **关}, tmp_path / "任务", _多表模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert 读Excel格(结果["输出"], "资金", "B2") == "甲公司"
    assert 读Excel格(结果["输出"], "资金", "C2") == 100
    assert 读Excel格(结果["输出"], "资金", "D2") == 110
    assert 读Excel格(结果["输出"], "往来", "B2") == "往来甲"
    assert 读Excel格(结果["输出"], "往来", "C2") == 701
    assert 读Excel格(结果["输出"], "往来", "D2") == 710
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "不在本次可见范围" in 问题
    assert 前 == (文件哈希(A), 文件哈希(B))


def test_来源零值空白和清尾仍按原规则(tmp_path, monkeypatch):
    """保护对照：0照搬、真实空白清空、合法清尾整项写空，三种不混淆。"""
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    wb = Workbook()
    wb.remove(wb.active)
    资金 = wb.create_sheet("资金")
    资金["B1"] = "项目（单位：人民币元）"
    资金["C1"] = "2025-01-01"
    资金["D1"] = "2025-12-31"
    for 行, (名, 期初, 期末) in enumerate(
            (("甲公司", 100, 110), ("乙公司", 200, 210), ("丙公司", 300, 310)), 2):
        资金.cell(行, 2, 名)
        资金.cell(行, 3, 期初)
        资金.cell(行, 4, 期末)
    A = tmp_path / "A.xlsx"
    wb.save(A)
    wb.close()
    wb = Workbook()
    wb.remove(wb.active)
    余额 = wb.create_sheet("余额")
    余额["B1"] = "项目（单位：人民币元）"
    余额["C1"] = "2025-01-01"
    余额["D1"] = "2025-12-31"
    余额["B2"], 余额["C2"], 余额["D2"] = "本次甲", 0, 102
    余额["B3"] = "本次乙"
    B = tmp_path / "B.xlsx"
    wb.save(B)
    wb.close()
    前 = (文件哈希(A), 文件哈希(B))

    def 模型():
        def 调用(配置, 指令, 材料, 目录):
            三行A = "丙公司" in 材料
            底 = 4 if 三行A else 3
            if 指令.startswith("把整份文件"):
                return json.dumps({"划分": [{"行范围": [1, 底], "科目": "货币资金"}]}, ensure_ascii=False)
            if 指令.startswith("在第一遍"):
                return json.dumps({"表": [{"行范围": [1, 底], "科目": "货币资金", "表名": "余额构成",
                    "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
            列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
                  {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                   "期间实际": {"值": "2025-01-01", "位置": "C1"}},
                  {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                   "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
            行 = [{"行": 1, "名字": "项目", "角色": "表头"}] + [
                {"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
                 "槽位": {"C": "N001", "D": "N001"}} for i in range(2, 底 + 1)]
            作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
                {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
                 "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
            return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                              ensure_ascii=False)
        return 调用

    结果 = 跑场景(_简单输入(tmp_path, A, B, 关), tmp_path / "任务", 模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "资金", "B2") == "本次甲"
    assert 读Excel格(输出, "资金", "C2") == 0
    assert 读Excel格(输出, "资金", "D2") == 102
    assert 读Excel格(输出, "资金", "B3") == "本次乙"
    assert 读Excel格(输出, "资金", "C3") is None
    assert 读Excel格(输出, "资金", "D3") is None
    assert 读Excel格(输出, "资金", "B4") is None
    assert 读Excel格(输出, "资金", "C4") is None
    assert 读Excel格(输出, "资金", "D4") is None
    assert 前 == (文件哈希(A), 文件哈希(B))


# —— 问题 2：第二遍复用后标题重建与整组保留状态 ——

标题原文A = "截至2024年12月31日／比较日2023年12月31日"
标题新文A = "截至2025年12月31日／比较日2024年12月31日"


def _标题任务(tmp_path, 名, B值=((901, 102), (1001, 202)), A标题=标题原文A, 记录=None):
    """A 带标题“截至2024…／比较日2023…”，B 本次 901/102；复用 test_组合搬运 的标题样本。"""
    from 测试.test_组合搬运 import _标题工作簿, _标题金, _标题模型

    记录 = 记录 or (tmp_path / "识别记录")
    A = _标题工作簿(tmp_path / f"A_{名}.xlsx")
    if A标题 != 标题原文A:
        wb = load_workbook(A)
        wb["正常"]["B1"] = A标题
        wb.save(A)
        wb.close()
    B源 = _标题工作簿(tmp_path / f"B源_{名}.xlsx", 来源=True)
    wb = load_workbook(B源)
    wb["来源正常"]["C2"], wb["来源正常"]["D2"] = B值[0]
    wb.save(B源)
    wb.close()
    B = tmp_path / f"B_{名}.xlsx"
    import shutil

    shutil.copy2(B源, B)
    金 = _标题金(tmp_path / f"金标准_{名}.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["正常"]},
            "B": {"路径": str(B), "工作表": ["来源正常"]},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    return 输入, _标题模型()


def test_复用第二三份仍搬本次金额名单标题(tmp_path, monkeypatch):
    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "1")
    输入, 模型 = _标题任务(tmp_path, "复用三份")
    结果们 = [跑场景(输入, tmp_path / f"任务{n}", 模型, monkeypatch) for n in (1, 2, 3)]
    for n, 结果 in enumerate(结果们, 1):
        assert 结果["状态"]["步骤"] == "done", (n, 结果["状态"])
        输出 = 结果["输出"]
        assert 读Excel格(输出, "正常", "B1") == 标题新文A, (n, "标题")
        assert 读Excel格(输出, "正常", "B3") == "本次甲", (n, "名单")
        assert 读Excel格(输出, "正常", "C3") == 901, (n, "期初")
        assert 读Excel格(输出, "正常", "D3") == 102, (n, "期末")
        assert 读Excel格(输出, "正常", "C4") == 200
        assert 读Excel格(输出, "正常", "D4") == 210
    份2调用 = {x["侧"]: (x["第一遍"], x["第二遍"], x["第三遍"]) for x in 结果们[1]["调用"]}
    份3调用 = {x["侧"]: (x["第一遍"], x["第二遍"], x["第三遍"]) for x in 结果们[2]["调用"]}
    for 侧 in ("A", "B"):
        assert 份2调用[侧][:2] == (0, 0), (侧, 份2调用)
        assert 份3调用[侧] == (0, 0, 0), (侧, 份3调用)
    if 份2调用["B"][2] == 1:      # 允许有真实原因的重认，不接受无理由回退
        记录B = json.loads((tmp_path / "任务2" / "复用记录_B.json").read_text(encoding="utf-8"))
        原因们 = [x.get("原因") for x in 记录B.get("回退", []) if x.get("步骤") == "第三遍"]
        assert any(原因们), 记录B


def test_金额相同仍换年_复用(tmp_path, monkeypatch):
    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "1")
    输入, 模型 = _标题任务(tmp_path, "同额换年", B值=((100, 110), (200, 210)))
    for n in (1, 2):
        结果 = 跑场景(输入, tmp_path / f"任务{n}", 模型, monkeypatch)
        assert 结果["状态"]["步骤"] == "done", 结果["状态"]
        输出 = 结果["输出"]
        assert 读Excel格(输出, "正常", "B1") == 标题新文A, (n, "标题未按本次日期更新")
        assert 读Excel格(输出, "正常", "C3") == 100
        assert 读Excel格(输出, "正常", "D3") == 110


def test_复用标题日期冲突保留组且状态部分保留(tmp_path, monkeypatch):
    """标题日期冲突整组保留（对照既有关闭复用场景），状态须如实为部分保留，
    保留组的引用公式不得宣称本次核对通过。复用路径本身由 A5/A16/A17/A19 覆盖。"""
    monkeypatch.setenv("审计附注_识别记录目录", str(tmp_path / "识别记录"))
    monkeypatch.setenv("审计附注_复用", "0")
    输入, 模型 = _标题任务(tmp_path, "冲突", B值=((901, 102), (1001, 202)))
    wb = load_workbook(输入["A"]["路径"])
    wb["冲突"]["C4"] = "=SUM(C3:C3)"          # 保留组内放一条引用公式
    wb.save(输入["A"]["路径"])
    wb.close()
    冲突输入 = {**输入, "复用": False,
              "A": {**输入["A"], "工作表": ["正常", "冲突"]},
              "B": {**输入["B"], "工作表": ["来源正常", "来源冲突"]}}
    结果 = 跑场景(冲突输入, tmp_path / "冲突任务", 模型, monkeypatch)
    assert 结果["状态"]["步骤"] == "部分保留", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "正常", "B1") == 标题新文A
    assert 读Excel格(输出, "正常", "C3") == 901
    assert 读Excel格(输出, "冲突", "B1") == "截至2024年12月31日"
    assert 读Excel格(输出, "冲突", "B3") == "冲突甲"
    assert 读Excel格(输出, "冲突", "C3", 公式=True) == "=100+200"
    assert 读Excel格(输出, "冲突", "C4", 公式=True) == "=SUM(C3:C3)"
    assert 读Excel格(输出, "冲突", "D3") == 310
    记录们 = [json.loads(x) for x in (tmp_path / "冲突任务" / "公式核对记录.jsonl")
            .read_text(encoding="utf-8").splitlines() if x.strip()]
    条 = next(x for x in 记录们 if x["a格"] == "冲突!C4")
    assert 条["结论"] == "未能核对" and "标题未能确认" in 条["原因"], 条
    问题 = (tmp_path / "冲突任务" / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "日期" in 问题 and "冲突" in 问题


def test_复用标题骨架对不上保留组且布局仍复用(tmp_path, monkeypatch):
    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "1")
    输入, 模型 = _标题任务(tmp_path, "骨架", B值=((901, 102), (1001, 202)))
    第一 = 跑场景(输入, tmp_path / "任务1", 模型, monkeypatch)
    assert 第一["状态"]["步骤"] == "done", 第一["状态"]
    改题 = {**输入, "A": {**输入["A"]}}
    import shutil

    A2 = tmp_path / "A_改题.xlsx"
    shutil.copy2(输入["A"]["路径"], A2)
    wb = load_workbook(A2)
    wb["正常"]["B1"] = "注：截至2024年12月31日"
    wb.save(A2)
    wb.close()
    改题["A"]["路径"] = str(A2)
    第二 = 跑场景(改题, tmp_path / "任务2", 模型, monkeypatch)
    assert 第二["状态"]["步骤"] == "部分保留", 第二["状态"]
    输出 = 第二["输出"]
    assert 读Excel格(输出, "正常", "B1") == "注：截至2024年12月31日"
    assert 读Excel格(输出, "正常", "C3") == 100
    assert 读Excel格(输出, "正常", "D3") == 110
    问题 = (tmp_path / "任务2" / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "无法按原骨架重建" in 问题
    # 标题行文字变化会令整侧第一遍重认；第二遍按表段身份（不含日期列）仍应复用。
    调用A = next(x for x in 第二["调用"] if x["侧"] == "A")
    assert 调用A["第二遍"] == 0, 调用A


def test_二遍复用三遍重认仍保留本次标题(tmp_path, monkeypatch):
    from 核心.复用管理 import 停用
    from 核心.共用记录簿 import 记录簿
    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "1")
    输入, 模型 = _标题任务(tmp_path, "停用三遍")
    第一 = 跑场景(输入, tmp_path / "任务1", 模型, monkeypatch)
    assert 第一["状态"]["步骤"] == "done", 第一["状态"]
    簿 = 记录簿(记录)
    快照 = 簿.快照()
    # 第三遍候选按结构摘要跨端共享：要构造“第三遍不可复用而前两遍仍有效”，
    # 必须停用两侧全部可匹配的第三遍候选；第一、二遍版本不动。
    三遍 = [版本 for 版本, 项 in (快照.get("解释") or {}).items()
          if 项.get("步骤") == "第三遍" and 项.get("状态") == "有效"]
    assert 三遍, 快照.get("解释")
    停用(簿, "测试停用第三遍", 三遍, "本轮用例", "验证第三遍重认路径")
    第二 = 跑场景(输入, tmp_path / "任务2", 模型, monkeypatch)
    assert 第二["状态"]["步骤"] == "done", 第二["状态"]
    输出 = 第二["输出"]
    assert 读Excel格(输出, "正常", "B1") == 标题新文A
    assert 读Excel格(输出, "正常", "B3") == "本次甲"
    assert 读Excel格(输出, "正常", "C3") == 901
    assert 读Excel格(输出, "正常", "D3") == 102
    调用 = {x["侧"]: (x["第一遍"], x["第二遍"], x["第三遍"]) for x in 第二["调用"]}
    assert 调用["A"] == (0, 0, 1), 调用
    assert 调用["B"] == (0, 0, 1), 调用


def _跨格式标题模型():
    """任务1：A1原生Excel（表内单位，1-3行）；B1 Word（表前单位段落占第1行，表2-4行）。
    任务2：A2 Word（标题段落占第1行，表2-4行，本次标题，表内单位）；B2原生Excel（1-3行）。"""
    def 调用(配置, 指令, 材料, 目录):
        目录 = Path(目录)
        侧 = 目录.name[0]
        任务二 = "任务2" in 目录.parts
        word = (侧 == "B") if not 任务二 else (侧 == "A")
        顶 = 2 if word else 1
        表内 = not (not 任务二 and 侧 == "B")          # 仅任务1的B1用上下文单位
        底 = 顶 + 2
        划分 = ([{"行范围": [1, 1], "科目": "非表格"}] if word else [])
        划分.append({"行范围": [顶, 底], "科目": "货币资金"})
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": 划分}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表 = {"行范围": [顶, 底], "科目": "货币资金", "表名": "余额构成",
                  "表组": "G1", "续表序": 1}
            if 任务二 and 侧 == "A":
                表["标题"] = {"原文": "截至2024年12月31日",
                              "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"}]}
            return json.dumps({"表": [表]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
            {"行": 顶 + i, "名字": f"名单{i}", "角色": "数据", "序号": i,
             "槽位": {"C": "N001", "D": "N001"}} for i in (1, 2)]
        if 表内:
            依据 = [{"来源": "表内", "位置": f"B{顶}", "原文": "项目（单位：人民币元）",
                    "片段": "元", "币种": "人民币", "单位": "元", "列": ["C", "D"]}]
        else:
            依据 = [{"来源": "上下文", "位置": ["unit_text_candidates", 0],
                    "原文": "金额单位：人民币元", "片段": "元", "币种": "人民币",
                    "单位": "元", "列": ["C", "D"]}]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": 依据}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def test_本次标题不被跨格式精读覆盖(tmp_path, monkeypatch):
    """A17：任务1形成Excel侧精读记录；任务2为带标题的Word，第二遍首认给出本次标题，
    第三遍跨格式复用旧解释，本次标题与位置不得被覆盖，旧格位置不得混入Word指令。"""
    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "1")
    金 = 造金标准(tmp_path / "金标准.jsonl")
    A1 = _简单簿(tmp_path / "A1.xlsx", "资金",
               [("甲公司", 100, 110), ("乙公司", 200, 210)])
    B1 = 造Word(tmp_path / "B1.docx", 901, 102, 名单=("客户甲", "客户乙"))
    输入1 = {"A": {"路径": str(A1), "工作表": ["资金"]}, "B": {"路径": str(B1), "工作表": []},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    结果1 = 跑场景(输入1, tmp_path / "任务1", _跨格式标题模型(), monkeypatch)
    assert 结果1["状态"]["步骤"] == "done", 结果1["状态"]

    A2 = 造Word(tmp_path / "A2.docx", 100, 110, 标题="截至2024年12月31日",
              表内单位=True, 表前单位=False)
    B2 = _简单簿(tmp_path / "B2.xlsx", "余额",
               [("本次甲", 901, 102), ("本次乙", 1001, 202)])
    输入2 = {"A": {"路径": str(A2), "工作表": []}, "B": {"路径": str(B2), "工作表": ["余额"]},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    结果2 = 跑场景(输入2, tmp_path / "任务2", _跨格式标题模型(), monkeypatch)
    assert 结果2["状态"]["步骤"] == "done", 结果2["状态"]
    assert 读Word格(结果2["输出"], 0, 1, 0) == "本次甲"
    assert 读Word格(结果2["输出"], 0, 1, 1) == "901"
    assert 读Word格(结果2["输出"], 0, 1, 2) == "102"
    from 测试.组合样本 import 读Word段落

    assert "截至2025年12月31日" in 读Word段落(结果2["输出"])
    证据 = json.loads((结果2["任务"] / "标题证据.json").read_text(encoding="utf-8"))
    assert 证据 and all("格" not in x for x in 证据), 证据
    调用A = next(x for x in 结果2["调用"] if x["侧"] == "A")
    assert 调用A["第三遍"] == 0, 调用A


def test_修前记录仍可复用且结果正确(tmp_path, monkeypatch):
    """A19：ed1c644 生成的有效识别记录（任务0保留的副本）修后仍能命中并正确搬运。"""
    旧记录 = (Path(__file__).resolve().parent.parent / "运行" / "验收"
            / "复核五项修复_20261106_013116" / "A19_旧记录原始副本")
    assert 旧记录.is_dir(), "缺少ed1c644旧记录副本，无法验收兼容性"
    运行记录 = tmp_path / "旧记录运行"
    import shutil

    shutil.copytree(旧记录, 运行记录, dirs_exist_ok=True)
    monkeypatch.setenv("审计附注_识别记录目录", str(运行记录))
    monkeypatch.setenv("审计附注_复用", "1")
    输入, 模型 = _标题任务(tmp_path, "旧记录", 记录=运行记录)
    结果 = 跑场景(输入, tmp_path / "任务", 模型, monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "正常", "B1") == 标题新文A
    assert 读Excel格(输出, "正常", "B3") == "本次甲"
    assert 读Excel格(输出, "正常", "C3") == 901
    assert 读Excel格(输出, "正常", "D3") == 102
    调用B = next(x for x in 结果["调用"] if x["侧"] == "B")
    assert 调用B["第三遍"] == 0, 调用B


# —— 问题 3：标题不越界写隐藏格 ——

标题旧文 = "截至2024年12月31日／比较日2023年12月31日"


def _越界簿(路径, 变体):
    """复用已验证的标题样本：受限表（货币资金）＋独立表（存货）；
    按变体把受限表的标题格放进隐藏行、隐藏列或打印区域外。"""
    from 测试.test_组合搬运 import _标题工作簿

    A = _标题工作簿(路径)
    wb = load_workbook(路径)
    wb["正常"].title = "受限"
    wb["冲突"].title = "独立"
    受限 = wb["受限"]
    if 变体 == "隐藏行":
        受限.row_dimensions[1].hidden = True
    elif 变体 == "隐藏列":
        受限["B1"] = None
        受限["E1"] = 标题旧文
        受限.column_dimensions["E"].hidden = True
    elif 变体 == "打印区域":
        受限.print_area = "$B$2:$D$4"
    wb.save(路径)
    wb.close()
    return A


def _越界模型(标题位="B1"):
    def 调用(配置, 指令, 材料, 目录):
        目录 = Path(目录)
        侧 = 目录.name[0]
        序 = next(int(x[1:]) - 1 for x in 目录.parts if x in ("P001", "P002"))
        科目, 表名, 编号们 = (("货币资金", "余额构成", ("N001", "N002")),
                          ("存货", "存货明细", ("N003", "N004")))[序]
        顶 = 2 if 侧 == "A" else 1
        底 = 顶 + 2
        if 指令.startswith("把整份文件"):
            划分 = ([{"行范围": [1, 1], "科目": "非表格"}] if 侧 == "A" else [])
            划分.append({"行范围": [顶, 底], "科目": 科目})
            return json.dumps({"划分": 划分}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表 = {"行范围": [顶, 底], "科目": 科目, "表名": 表名,
                  "表组": "G1", "续表序": 1}
            if 侧 == "A" and 序 == 0:
                表["标题"] = {"位置": 标题位, "原文": 标题旧文,
                              "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"},
                                      {"片段": "2023年12月31日", "期间角色": "上期期末"}]}
            elif 侧 == "A":
                表["标题"] = {"位置": "B1", "原文": "截至2024年12月31日",
                              "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"}]}
            return json.dumps({"表": [表]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"}]
        for 字母, 编号, 角, 日 in zip(("C", "D"), 编号们,
                                  ("本期期末", "上期期末"), ("2025-12-31", "2024-12-31")):
            列.append({"列": 字母, "路径": 角, "角色": "数据列", "期间角色": 角,
                      "期间实际": {"值": 日, "位置": f"{字母}{顶}"}})
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
            {"行": 顶 + i, "名字": f"名单{i}", "角色": "数据", "序号": i,
             "槽位": {"C": 编号们[0], "D": 编号们[1]}} for i in (1, 2)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": f"B{顶}", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def _越界输入(tmp_path, 变体, 关):
    from 测试.test_组合搬运 import _标题工作簿, _标题金

    A = _越界簿(tmp_path / "A.xlsx", 变体)
    B源 = _标题工作簿(tmp_path / "B源.xlsx", 来源=True)
    wb = load_workbook(B源)
    for 表名 in ("来源正常", "来源冲突"):
        wb[表名]["B2"], wb[表名]["C2"], wb[表名]["D2"] = "本次甲", 901, 102
        wb[表名]["B3"], wb[表名]["C3"], wb[表名]["D3"] = "本次乙", 1001, 202
    wb.save(B源)
    wb.close()
    B = tmp_path / "B.xlsx"
    import shutil

    shutil.copy2(B源, B)
    金 = _标题金(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["受限", "独立"]},
            "B": {"路径": str(B), "工作表": ["来源正常", "来源冲突"]},
            "金标准": str(金), **关}
    return A, B, 输入


@pytest.mark.parametrize("变体", ["隐藏行", "隐藏列", "打印区域"])
def test_越界标题不改写且受限组保留(tmp_path, monkeypatch, 变体):
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    标题位 = "E1" if 变体 == "隐藏列" else "B1"
    A, B, 输入 = _越界输入(tmp_path, 变体, 关)
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景(输入, tmp_path / "任务", _越界模型(标题位), monkeypatch)
    assert 结果["状态"]["步骤"] == "部分保留", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "受限", 标题位) == 标题旧文
    assert 读Excel格(输出, "受限", "C3") == 100
    assert 读Excel格(输出, "受限", "D3") == 110
    assert 读Excel格(输出, "受限", "B3") == "正常甲"
    assert 读Excel格(输出, "独立", "B1") == "截至2025年12月31日"
    assert 读Excel格(输出, "独立", "C3") == 901
    assert 读Excel格(输出, "独立", "D3") == 102
    assert 读Excel格(输出, "独立", "B3") == "本次甲"
    wb = load_workbook(输出)
    try:
        受限 = wb["受限"]
        if 变体 == "隐藏行":
            assert 受限.row_dimensions[1].hidden
        elif 变体 == "隐藏列":
            assert 受限.column_dimensions["E"].hidden
        else:
            assert "$B$2:$D$4" in str(受限.print_area)
    finally:
        wb.close()
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "可见范围" in 问题
    assert 前 == (文件哈希(A), 文件哈希(B))


def test_可见表前标题仍更新(tmp_path, monkeypatch):
    """对照：同样的双表样本不加任何范围限制，两个可见标题都按本次日期更新。"""
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A, B, 输入 = _越界输入(tmp_path, "", 关)
    结果 = 跑场景(输入, tmp_path / "任务", _越界模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "受限", "B1") == 标题新文A
    assert 读Excel格(输出, "受限", "C3") == 901
    assert 读Excel格(输出, "受限", "D3") == 102
    assert 读Excel格(输出, "独立", "B1") == "截至2025年12月31日"
    assert 读Excel格(输出, "独立", "C3") == 901
    assert 读Excel格(输出, "独立", "D3") == 102


def test_可见集合缺失不得写入(tmp_path, monkeypatch):
    """A18 负向注入：交付请求边界去掉可见范围记录，真实写入器运行但必须全部保留。"""
    from 核心 import 流程 as 流程模块
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A, B, 输入 = _越界输入(tmp_path, "", 关)
    原交付 = 流程模块.交付Excel

    def 包装(原件, 输出, 变更, A快照, A结构, 公式清单, 标题变更, 取消=None):
        for 表 in A快照.get("sheets") or []:
            表.pop("可见范围", None)      # 注入边界：仅删交付请求里的可见范围记录
        return 原交付(原件, 输出, 变更, A快照, A结构, 公式清单, 标题变更, 取消=取消)

    monkeypatch.setattr(流程模块, "交付Excel", 包装)
    结果 = 跑场景(输入, tmp_path / "任务", _越界模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "部分保留", 结果["状态"]
    输出 = 结果["输出"]
    assert 读Excel格(输出, "受限", "C3") == 100
    assert 读Excel格(输出, "受限", "D3") == 110
    assert 读Excel格(输出, "受限", "B1") == 标题旧文
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "缺少可见范围依据" in 问题


# —— 问题 4：文字名单保持原文 ——

def test_文字名单保持原文与文本类型(tmp_path, monkeypatch):
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    wb = Workbook()
    wb.remove(wb.active)
    资金 = wb.create_sheet("资金")
    资金["B1"] = "项目（单位：人民币元）"
    资金["C1"] = "2025-01-01"
    资金["D1"] = "2025-12-31"
    A名单 = ("甲公司", "乙公司", "丙公司", "丁公司", "戊公司")
    for i, 名 in enumerate(A名单, 2):
        资金.cell(i, 2, 名)
        资金.cell(i, 3, 100 * i - 100).number_format = "0.00"
        资金.cell(i, 4, 100 * i - 90).number_format = "0.00"
    A = tmp_path / "A.xlsx"
    wb.save(A)
    wb.close()
    wb = Workbook()
    wb.remove(wb.active)
    余额 = wb.create_sheet("余额")
    余额["B1"] = "项目（单位：人民币元）"
    余额["C1"] = "2025-01-01"
    余额["D1"] = "2025-12-31"
    B名单 = ("中文名称", "00123", "1-2", "=客户甲", "'=带引号")
    for i, (名, 期初, 期末) in enumerate(
            zip(B名单, (901, 1001, 1101, 1201, 1301), (102, 202, 302, 402, 502)), 2):
        格 = 余额.cell(i, 2, 名)
        if isinstance(名, str) and 名.startswith("="):
            格.data_type = "s"          # 源头就是文字，不是公式
        余额.cell(i, 3, 期初)
        余额.cell(i, 4, 期末)
    B = tmp_path / "B.xlsx"
    wb.save(B)
    wb.close()
    前 = (文件哈希(A), 文件哈希(B))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"]},
            "金标准": str(金), **关}

    def 模型():
        def 调用(配置, 指令, 材料, 目录):
            if 指令.startswith("把整份文件"):
                return json.dumps({"划分": [{"行范围": [1, 6], "科目": "货币资金"}]}, ensure_ascii=False)
            if 指令.startswith("在第一遍"):
                return json.dumps({"表": [{"行范围": [1, 6], "科目": "货币资金", "表名": "余额构成",
                    "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
            列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
                  {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                   "期间实际": {"值": "2025-01-01", "位置": "C1"}},
                  {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                   "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
            行 = [{"行": 1, "名字": "项目", "角色": "表头"}] + [
                {"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
                 "槽位": {"C": "N001", "D": "N001"}} for i in range(2, 7)]
            作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
                {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
                 "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
            return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                              ensure_ascii=False)
        return 调用

    结果 = 跑场景(输入, tmp_path / "任务", 模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert 结果["状态"].get("回读差异", 0) == 0, 结果["状态"]
    wb = load_workbook(结果["输出"], data_only=False)
    try:
        ws = wb["资金"]
        for i, 名 in enumerate(B名单, 2):
            格 = ws.cell(i, 2)
            assert 格.value == 名, (i, 格.value)
            assert 格.data_type == "s", (i, 格.data_type)
        assert ws["C2"].value == 901 and ws["D2"].value == 102
        assert ws["C2"].number_format == "0.00"
    finally:
        wb.close()
    assert 前 == (文件哈希(A), 文件哈希(B))


# —— 问题 5：材料准备失败后同任务续跑 ——

def test_准备失败后续跑只补缺失端(tmp_path, monkeypatch):
    """负向故障注入：仅B端材料准备首次调用注入一次临时读取错误，撤销后续跑。"""
    from 核心 import 流程 as 流程模块
    from 测试.安全目录 import 安全记录目录

    关 = _关复用(monkeypatch, 安全记录目录())
    A = _简单簿(tmp_path / "A.xlsx", "资金",
              [("甲公司", 100, 110), ("乙公司", 200, 210)])
    B = _简单簿(tmp_path / "B.xlsx", "余额",
              [("本次甲", 901, 102), ("本次乙", 1001, 202)])
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["余额"]},
            "金标准": str(金), **关}
    任务 = tmp_path / "任务"
    原准备 = 流程模块.准备一侧
    调用记录 = []

    def 注入(侧, 描述, 目录):
        调用记录.append(侧)
        if 侧 == "B" and len([x for x in 调用记录 if x == "B"]) == 1:
            raise OSError("注入：B 材料准备临时读取失败")
        return 原准备(侧, 描述, 目录)

    monkeypatch.setattr(流程模块, "准备一侧", 注入)
    with pytest.raises(OSError, match="临时读取失败"):
        from 核心.流程 import 运行任务

        运行任务(输入, 任务, {"测试": True})
    索引 = json.loads((任务 / "接续索引.json").read_text(encoding="utf-8-sig"))
    assert sorted((索引.get("材料") or {}).keys()) == ["A"]
    A副本 = 任务 / "A" / "输入副本.xlsx"
    assert A副本.is_file()
    副本哈希 = 文件哈希(A副本)
    原件哈希 = (文件哈希(A), 文件哈希(B))

    结果 = 跑场景(输入, 任务, _简单模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert 读Excel格(结果["输出"], "资金", "C2") == 901
    assert 读Excel格(结果["输出"], "资金", "D3") == 202
    assert 调用记录.count("A") == 1 and 调用记录[-1] == "B", 调用记录
    assert 文件哈希(A副本) == 副本哈希
    assert (文件哈希(A), 文件哈希(B)) == 原件哈希
