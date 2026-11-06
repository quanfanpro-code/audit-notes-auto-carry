# coding: utf-8
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook

from 测试.组合样本 import 造Excel, 造Word, 造金标准, 文件哈希


def _输入(tmp_path, A格式="word", B格式="excel"):
    A = 造Word(tmp_path / "A.docx") if A格式 == "word" else 造Excel(tmp_path / "A.xlsx")
    B = 造Excel(tmp_path / "B.xlsx") if B格式 == "excel" else 造Word(tmp_path / "B.docx", 901, 102)
    金 = 造金标准(tmp_path / "金标准.jsonl")
    return {"A": {"路径": str(A), "工作表": ["余额"] if A格式 == "excel" else []},
            "B": {"路径": str(B), "工作表": ["余额"] if B格式 == "excel" else []},
            "金标准": str(金), "复用": False}


def test_输入规范化_新旧参数与工作表(tmp_path):
    from 核心.文件输入 import 规范化输入

    B = 造Excel(tmp_path / "多表.xlsx", {"资金": [[1]], "往来": [[2]], "隐藏": [[3]]}, 隐藏=("隐藏",))
    输入 = _输入(tmp_path)
    输入["B"] = {"路径": str(B), "工作表": ["往来", "资金"]}
    出 = 规范化输入(输入)
    assert 出["A"]["类型"] == "word" and 出["A"]["工作表"] == []
    assert 出["B"]["类型"] == "excel" and 出["B"]["工作表"] == ["往来", "资金"]
    assert 出["金标准"] == 输入["金标准"]

    旧 = 规范化输入({"word": "原.docx", "b表": str(B), "b工作表": "资金",
                  "金标准": 输入["金标准"], "复用": False})
    assert 旧["A"] == {"路径": "原.docx", "类型": "word", "工作表": []}
    assert 旧["B"]["工作表"] == ["资金"]
    with pytest.raises(ValueError, match="同时提供"):
        规范化输入({**输入, "word": "旧.docx"})
    with pytest.raises(ValueError, match="隐藏"):
        规范化输入({"A": 输入["A"], "B": {"路径": str(B), "工作表": ["隐藏"]},
                  "金标准": 输入["金标准"]})
    with pytest.raises(ValueError, match="至少"):
        规范化输入({"A": 输入["A"], "B": {"路径": str(B), "工作表": []},
                  "金标准": 输入["金标准"]})


def test_输入拒绝伪装格式(tmp_path):
    from 核心.文件输入 import 规范化输入

    假 = tmp_path / "假.xlsx"
    假.write_text("不是工作簿", encoding="utf-8")
    输入 = _输入(tmp_path)
    输入["B"] = {"路径": str(假), "工作表": ["余额"]}
    with pytest.raises(ValueError):
        规范化输入(输入)
    假Word = tmp_path / "假.docx"
    假Word.write_text("不是Word", encoding="utf-8")
    输入 = _输入(tmp_path)
    输入["A"] = {"路径": str(假Word), "工作表": []}
    with pytest.raises(ValueError):
        规范化输入(输入)


def test_双Word材料真实分离(tmp_path):
    from 核心.文件输入 import 准备一侧, 规范化输入

    输入 = 规范化输入(_输入(tmp_path, "word", "word"))
    前 = {侧: 文件哈希(输入[侧]["路径"]) for 侧 in ("A", "B")}
    甲 = 准备一侧("A", 输入["A"], tmp_path / "任务" / "A")
    乙 = 准备一侧("B", 输入["B"], tmp_path / "任务" / "B")
    assert len(甲) == len(乙) == 1
    assert 甲[0]["材料"] != 乙[0]["材料"]
    assert 甲[0]["上下文路径"] != 乙[0]["上下文路径"]
    assert 甲[0]["带连接Word"] != 乙[0]["带连接Word"]
    数值 = []
    for 部分 in (甲[0], 乙[0]):
        wb = load_workbook(部分["材料"], data_only=True)
        try:
            数值.append({c.value for row in wb.active.iter_rows() for c in row
                       if isinstance(c.value, (int, float))})
        finally:
            wb.close()
        assert json.loads(部分["上下文路径"].read_text(encoding="utf-8-sig"))["tables"]
    assert 数值[0] == {100, 110, 200, 210}
    assert 数值[1] == {901, 102, 1001, 202}
    assert {侧: 文件哈希(输入[侧]["路径"]) for 侧 in ("A", "B")} == 前


def _三组任务(tmp_path, 同名=False, monkeypatch=None):
    from 测试.test_真实桥接集成 import _三组样本, _三组金标准
    from 测试.组合样本 import 造三组Excel, 三组多表模型, 跑场景

    A = _三组样本(tmp_path)
    B = 造三组Excel(tmp_path / "B.xlsx")
    金 = _三组金标准(tmp_path)
    输入 = {"A": {"路径": str(A), "工作表": []},
            "B": {"路径": str(B), "工作表": ["资金", "往来", "其他"]},
            "金标准": str(金), "复用": False}
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景(输入, tmp_path / "任务", 三组多表模型(同名), monkeypatch)
    assert 前 == (文件哈希(A), 文件哈希(B))
    return 结果


def test_多表同坐标进入正确Word组(tmp_path, monkeypatch):
    from 测试.组合样本 import 读Word格

    结果 = _三组任务(tmp_path, monkeypatch=monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    预期 = [("901", "910"), ("701", "710"), ("801", "810")]
    for 表, (期初, 期末) in enumerate(预期):
        assert 读Word格(结果["输出"], 表, 1, 0) == "甲公司"
        assert 读Word格(结果["输出"], 表, 1, 1) == 期初
        assert 读Word格(结果["输出"], 表, 1, 2) == 期末
    调用 = [x for x in 结果["调用"] if x["侧"] == "B"]
    assert {x["工作表"] for x in 调用} == {"资金", "往来", "其他"}
    assert all((x["第一遍"], x["第二遍"], x["第三遍"]) == (1, 1, 1) for x in 调用)


def test_同名组保留且独立组更新(tmp_path, monkeypatch):
    from 测试.组合样本 import 读Word格, 读Word段落

    结果 = _三组任务(tmp_path, 同名=True, monkeypatch=monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    for 表, 原值 in ((0, ("100", "110")), (1, ("200", "210"))):
        assert 读Word格(结果["输出"], 表, 1, 0) == "甲公司"
        assert (读Word格(结果["输出"], 表, 1, 1),
                读Word格(结果["输出"], 表, 1, 2)) == 原值
    assert (读Word格(结果["输出"], 2, 1, 1),
            读Word格(结果["输出"], 2, 1, 2)) == ("801", "810")
    段落 = 读Word段落(结果["输出"])
    assert "截至2024年12月31日（第一组）" in 段落
    assert "截至2024年12月31日（第二组）" in 段落
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "同名业务表归属不明" in 问题


def test_Word来源使用本次上下文金额和名单(tmp_path, monkeypatch):
    from 测试.组合样本 import 双Word模型, 读Word格, 跑场景

    A = 造Word(tmp_path / "A.docx")
    B = 造Word(tmp_path / "B.docx", 901, 102, 名单=("客户甲", "客户乙"))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": []}, "B": {"路径": str(B), "工作表": []},
            "金标准": str(金), "复用": False}
    结果 = 跑场景(输入, tmp_path / "任务", 双Word模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert (读Word格(结果["输出"], 0, 1, 0), 读Word格(结果["输出"], 0, 1, 1),
            读Word格(结果["输出"], 0, 1, 2)) == ("客户甲", "901", "102")
    assert (读Word格(结果["输出"], 0, 2, 0), 读Word格(结果["输出"], 0, 2, 1),
            读Word格(结果["输出"], 0, 2, 2)) == ("客户乙", "1001", "202")


def _标题工作簿(路径, 来源=False):
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    if 来源:
        配置 = [("来源正常", None, "2025-12-31", "2024-12-31", (100, 110), (200, 210), "本次"),
                ("来源冲突", None, "2025-12-31", "2025-06-30", (999, 888), (999, 888), "错误")]
    else:
        配置 = [("正常", "截至2024年12月31日／比较日2023年12月31日", "2024-12-31", "2023-12-31", (100, 110), (200, 210), "正常"),
                ("冲突", "截至2024年12月31日", "2024-12-31", "2023-12-31", (300, 310), (400, 410), "冲突")]
    for 名, 标题, 日期1, 日期2, 值1, 值2, 名前 in 配置:
        ws = wb.create_sheet(名)
        顶 = 1 if 来源 else 2
        if 标题:
            ws["B1"] = 标题
        ws.cell(顶, 2, "项目（单位：人民币元）")
        ws.cell(顶, 3, 日期1)
        ws.cell(顶, 4, 日期2)
        for i, 值 in ((1, 值1), (2, 值2)):
            ws.cell(顶 + i, 2, f"{名前}{'甲' if i == 1 else '乙'}")
            ws.cell(顶 + i, 3, 值[0])
            ws.cell(顶 + i, 4, 值[1])
        if 名 == "冲突":
            ws["C3"] = "=100+200"
    wb.save(路径)
    wb.close()
    return 路径


def _标题金(路径):
    行们 = []
    for 编号, 科目, 表名 in (("N001", "货币资金", "余额构成"), ("N002", "货币资金", "余额构成"),
                          ("N003", "存货", "存货明细"), ("N004", "存货", "存货明细")):
        行们.append({"id": 编号, "note": {"name": 科目}, "table": {"name": 表名},
                    "slot": {"name": 编号}, "dimensions": [{"name": "period", "required": True}]})
    Path(路径).write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们), encoding="utf-8")
    return Path(路径)


def _标题模型():
    def 调用(配置, 指令, 材料, 目录):
        目录 = Path(目录)
        侧 = 目录.name[0]
        序 = next(int(x[1:]) - 1 for x in 目录.parts if x in ("P001", "P002"))
        科目, 表名 = (("货币资金", "余额构成"), ("存货", "存货明细"))[序]
        编号们 = (("N001", "N002"), ("N003", "N004"))[序]
        顶 = 2 if 侧 == "A" else 1
        底 = 顶 + 2
        if 指令.startswith("把整份文件"):
            划分 = ([{"行范围": [1, 1], "科目": "非表格"}] if 侧 == "A" else [])
            划分.append({"行范围": [顶, 底], "科目": 科目})
            return json.dumps({"划分": 划分}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表 = {"行范围": [顶, 底], "科目": 科目, "表名": 表名, "表组": "G1", "续表序": 1}
            if 侧 == "A" and 序 == 0:
                表["标题"] = {"位置": "B1", "原文": "截至2024年12月31日／比较日2023年12月31日",
                            "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"},
                                    {"片段": "2023年12月31日", "期间角色": "上期期末"}]}
            elif 侧 == "A":
                表["标题"] = {"位置": "B1", "原文": "截至2024年12月31日",
                            "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"}]}
            return json.dumps({"表": [表]}, ensure_ascii=False)
        日期 = (("2025-12-31", "2024-12-31") if 序 == 0 else ("2025-12-31", "2025-06-30"))
        角色 = ("本期期末", "上期期末") if 序 == 0 else ("本期期末", "本期期末")
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"}]
        for 字母, 编号, 角, 日 in zip(("C", "D"), 编号们, 角色, 日期):
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


def test_Excel标题换年与日期冲突整组保留(tmp_path, monkeypatch):
    from 测试.组合样本 import 读Excel格, 跑场景

    A = _标题工作簿(tmp_path / "A.xlsx")
    B = _标题工作簿(tmp_path / "B.xlsx", 来源=True)
    金 = _标题金(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["正常", "冲突"]},
            "B": {"路径": str(B), "工作表": ["来源正常", "来源冲突"]},
            "金标准": str(金), "复用": False}
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景(输入, tmp_path / "任务", _标题模型(), monkeypatch)
    assert 结果["输出"].is_file(), 结果["状态"]
    assert 读Excel格(结果["输出"], "正常", "B1") == "截至2025年12月31日／比较日2024年12月31日"
    assert 读Excel格(结果["输出"], "正常", "C3") == 100
    assert 读Excel格(结果["输出"], "正常", "B3") == "本次甲"
    assert 读Excel格(结果["输出"], "冲突", "B1") == "截至2024年12月31日"
    assert 读Excel格(结果["输出"], "冲突", "B3") == "冲突甲"
    assert 读Excel格(结果["输出"], "冲突", "C3", 公式=True) == "=100+200"
    assert 读Excel格(结果["输出"], "冲突", "D3") == 310
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "日期" in 问题 and "冲突" in 问题
    assert 前 == (文件哈希(A), 文件哈希(B))


def test_取消后续跑只补剩余工作表(tmp_path, monkeypatch):
    import threading
    from 核心.流程 import 运行任务
    from 测试.test_Excel公式 import _多表, _多表金, _多表模型, _多表输入
    from 测试.组合样本 import 读Excel格, 跑场景

    A, B = _多表(tmp_path / "A.xlsx"), _多表(tmp_path / "B.xlsx", 来源=True)
    金 = _多表金(tmp_path / "金标准.jsonl")
    输入 = _多表输入(A, B, 金)
    取消 = threading.Event()
    基 = _多表模型()

    def 模型(配置, 指令, 材料, 目录):
        if "B" in Path(目录).parts and "P001" in Path(目录).parts and not 指令.startswith(("把整份文件", "在第一遍")):
            取消.set()
        return 基(配置, 指令, 材料, 目录)

    monkeypatch.setattr("核心.识别.调用识别", 模型)
    任务 = tmp_path / "任务"
    状态 = 运行任务(输入, 任务, {"测试": True}, 取消=取消)
    assert 状态["步骤"] == "用户中断" and not 状态.get("待答复")
    assert "识别B:P001" in 状态["已完成范围"]
    assert "识别B:P002" not in 状态["已完成范围"]
    结果 = 跑场景(输入, 任务, _多表模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert 读Excel格(结果["输出"], "资金", "C2") == 901
    assert 读Excel格(结果["输出"], "往来", "C2") == 701
    assert "识别B:P002" in 结果["状态"]["已完成范围"]


def test_续跑输入变化拒绝且已有副本不覆盖(tmp_path, monkeypatch):
    from 核心.流程 import 运行任务
    from 测试.test_Excel公式 import _多表, _多表金, _多表模型, _多表输入
    from 测试.组合样本 import 跑场景

    A, B = _多表(tmp_path / "A.xlsx"), _多表(tmp_path / "B.xlsx", 来源=True)
    金 = _多表金(tmp_path / "金标准.jsonl")
    输入 = _多表输入(A, B, 金)
    任务 = tmp_path / "任务"
    结果 = 跑场景(输入, 任务, _多表模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done"
    副本 = 任务 / "B" / "输入副本.xlsx"
    输出前 = 文件哈希(结果["输出"])
    副本前 = 文件哈希(副本)
    from openpyxl import load_workbook
    wb = load_workbook(B)
    wb["来源一"]["C2"] = 9999
    wb.save(B)
    wb.close()
    with pytest.raises(RuntimeError, match="内容已变"):
        运行任务(输入, 任务, {"测试": True})
    assert 文件哈希(结果["输出"]) == 输出前 and 文件哈希(副本) == 副本前

    改选择 = _多表输入(A, B, 金)
    改选择["B"]["工作表"] = ["来源一"]
    with pytest.raises(RuntimeError, match="工作表不同|文件内容已变"):
        运行任务(改选择, 任务, {"测试": True})
    monkeypatch.setattr("核心.流程.证据模块.规则摘要", lambda: "本轮新规则")
    with pytest.raises(RuntimeError, match="规则版本"):
        运行任务(输入, 任务, {"测试": True})
    assert 文件哈希(结果["输出"]) == 输出前 and 文件哈希(副本) == 副本前


def _可见工作簿(路径, 来源=False):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "余额" if 来源 else "资金"
    值 = (901, 102) if 来源 else (100, 110)
    ws["B1"] = "项目（单位：人民币元）"
    ws["C1"] = "2025-01-01"
    ws["D1"] = "2025-12-31"
    ws["E1"] = "隐藏说明"
    ws["B2"] = "隐藏甲"
    ws["C2"], ws["D2"], ws["E2"] = 999, 888, "隐藏旧"
    ws["B3"] = "打印空档"
    ws["C3"], ws["D3"] = 777, 666
    ws["B4"] = "本次甲" if 来源 else "原甲"
    ws["C4"], ws["D4"] = 值
    ws["E4"] = "隐藏列不动"
    ws.row_dimensions[2].hidden = True
    ws.column_dimensions["E"].hidden = True
    ws.print_area = ["$B$1:$E$2", "$B$4:$E$4"]
    wb.save(路径)
    wb.close()
    return 路径


def _可见模型():
    def 调用(配置, 指令, 材料, 目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,4],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return '{"表":[{"行范围":[1,4],"科目":"货币资金","表名":"余额构成","表组":"G1","续表序":1}]}'
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C1"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
        行 = [{"行": 1, "名字": "项目", "角色": "表头"},
              {"行": 4, "名字": "可见名单", "角色": "数据", "序号": 1,
               "槽位": {"C": "N001", "D": "N001"}}]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def test_隐藏与分离打印区域只搬可见格(tmp_path, monkeypatch):
    from 核心.表格 import 取可见集, 读快照
    from 测试.组合样本 import 读Excel格, 跑场景

    A, B = _可见工作簿(tmp_path / "A.xlsx"), _可见工作簿(tmp_path / "B.xlsx", 来源=True)
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]}, "B": {"路径": str(B), "工作表": ["余额"]},
            "金标准": str(金), "复用": False}
    结果 = 跑场景(输入, tmp_path / "任务", _可见模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert 读Excel格(结果["输出"], "资金", "B4") == "本次甲"
    assert 读Excel格(结果["输出"], "资金", "C4") == 901
    assert 读Excel格(结果["输出"], "资金", "D4") == 102
    assert 读Excel格(结果["输出"], "资金", "C2") == 999
    assert 读Excel格(结果["输出"], "资金", "D2") == 888
    assert 读Excel格(结果["输出"], "资金", "C3") == 777
    assert 读Excel格(结果["输出"], "资金", "E2") == "隐藏旧"
    wb = load_workbook(结果["输出"])
    try:
        区域 = str(wb["资金"].print_area)
        assert "$B$1:$E$2" in 区域 and "$B$4:$E$4" in 区域
    finally:
        wb.close()
    可见 = 取可见集(读快照(B), "余额")
    assert 可见.含(4, 3) and not 可见.含(2, 3) and not 可见.含(3, 3) and not 可见.含(4, 5)


def _续表模型():
    def 调用(配置, 指令, 材料, 目录):
        目录 = Path(目录)
        侧 = 目录.name[0]
        if 指令.startswith("把整份文件"):
            底 = 4 if 侧 == "A" else 8
            return json.dumps({"划分": [{"行范围": [1, 1], "科目": "非表格"},
                {"行范围": [2, 底], "科目": "货币资金"}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表们 = ([{"行范围": [2, 4], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 1}]
                  if 侧 == "A" else
                  [{"行范围": [2, 3], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 1},
                   {"行范围": [7, 8], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 2}])
            return json.dumps({"表": 表们}, ensure_ascii=False)
        if 侧 == "A":
            顶, 数据 = 2, (1, 2)
        else:
            序 = int(目录.name.rsplit("表", 1)[1])
            顶, 数据 = (2, (1,)) if 序 == 1 else (7, (2,))
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
            {"行": 顶 + i, "名字": f"名单{序}", "角色": "数据", "序号": 序,
             "槽位": {"C": "N001", "D": "N001"}} for i, 序 in enumerate(数据, 1)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "上下文", "位置": ["unit_text_candidates", 0], "原文": "金额单位：人民币元",
             "片段": "元", "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def test_Word来源续表保留各自含义(tmp_path, monkeypatch):
    from 测试.组合样本 import 造续表Word, 读Word格, 跑场景

    A = 造Word(tmp_path / "A.docx")
    B = 造续表Word(tmp_path / "B.docx")
    金 = 造金标准(tmp_path / "金标准.jsonl")
    输入 = {"A": {"路径": str(A), "工作表": []}, "B": {"路径": str(B), "工作表": []},
            "金标准": str(金), "复用": False}
    前 = (文件哈希(A), 文件哈希(B))
    结果 = 跑场景(输入, tmp_path / "任务", _续表模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "done", 结果["状态"]
    assert (读Word格(结果["输出"], 0, 1, 0), 读Word格(结果["输出"], 0, 1, 1),
            读Word格(结果["输出"], 0, 1, 2)) == ("客户甲", "901", "102")
    assert (读Word格(结果["输出"], 0, 2, 0), 读Word格(结果["输出"], 0, 2, 1),
            读Word格(结果["输出"], 0, 2, 2)) == ("客户乙", "1001", "202")
    结构 = [json.loads(x) for x in (结果["任务"] / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()]
    B续 = [x for x in 结构 if x.get("侧") == "B"]
    assert [(x["表组"], x["续表序"]) for x in B续] == [("G1", 1), ("G1", 2)]
    assert 前 == (文件哈希(A), 文件哈希(B))
