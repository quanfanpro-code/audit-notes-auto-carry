# coding: utf-8
"""本次标题：日期只取本次 B 对应表组的期间实际，配对后生成；生成不了整组撤回。"""
import json
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from 核心.标题 import 生成本次标题, 解析日期, 格式化
from 核心.识别 import _解析表
from 核心.流程 import 运行任务
from 测试.安全目录 import 安全记录目录

标题句 = "截至2024年12月31日，本公司货币资金余额如下"


def _A结构(原文=标题句, 片段="2024年12月31日", 角色="本期期末", 区域="TA_001",
         科目="货币资金", 表名="余额构成", 标题=True):
    记录 = {"侧": "A", "工作表": "sheet1", "区域": 区域, "范围": "A1:C3", "行范围": [1, 3],
           "科目": 科目, "表名": 表名, "表组": "G1", "续表序": 1, "列": [], "行": [],
           "作用范围": {}, "空名单": [], "例外": [], "完成": True}
    if 标题:
        记录["标题"] = {"原文": 原文,
                     "日期": [{"片段": 片段, "期间角色": 角色}] if 片段 else [],
                     "旧新文": None, "旧新文已忽略": False}
    return 记录


def _B合集(组, 条目):
    return [{"格": f"word附注!{坐标}", "科目": 组[0], "表名": 组[1],
             "维度值": {"期间角色": 角色, "期间实际": 日期}} for 坐标, 角色, 日期 in 条目]


def _上下文(标题=标题句, 区域="TA_001"):
    return {"tables": [{"range_name": 区域, "title": 标题,
                        "unit_text_candidates": ["金额单位：人民币元"]}]}


def test_期初期末两处分别取本次B日期():
    A = [_A结构(原文="期初2024年1月1日，期末2024年12月31日",
              片段="2024年1月1日", 角色="本期期初")]
    A[0]["标题"]["日期"].append({"片段": "2024年12月31日", "期间角色": "本期期末"})
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}]
    B格 = _B合集(("货币资金", "余额构成"),
                [("B2", "本期期初", "2025-01-01"), ("C2", "本期期末", "2025-12-31")])
    更新, 问题, 撤回, 证据 = 生成本次标题(
        A, B, _上下文("期初2024年1月1日，期末2024年12月31日"),
        {"sheet1!B2": 1, "sheet1!C2": 2}, B格)
    assert not 问题 and not 撤回
    assert 更新[0]["text"] == "期初2025年1月1日，期末2025年12月31日"
    assert 证据[0]["日期"][1]["坐标"] == ["word附注!C2"]


def test_同角色日期冲突整组撤回():
    A = [_A结构()]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}]
    B格 = _B合集(("货币资金", "余额构成"),
                [("B2", "本期期末", "2025-12-31"), ("C2", "本期期末", "2025-06-30")])
    更新, 问题, 撤回, _ = 生成本次标题(A, B, _上下文(), {"sheet1!C2": 1}, B格)
    assert not 更新 and 撤回 == {("货币资金", "余额构成")}
    assert "冲突" in 问题[0]["说明"]


def test_日期缺失整组撤回():
    A = [_A结构()]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}]
    更新, 问题, 撤回, _ = 生成本次标题(A, B, _上下文(), {"sheet1!C2": 1}, [])
    assert not 更新 and 撤回 and "没有" in 问题[0]["说明"]


def test_标题原文与本次不一致整组撤回():
    A = [_A结构(原文="截至2023年12月31日", 片段="2023年12月31日")]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}]
    B格 = _B合集(("货币资金", "余额构成"), [("C2", "本期期末", "2025-12-31")])
    更新, 问题, 撤回, _ = 生成本次标题(A, B, _上下文(标题=标题句), {"sheet1!C2": 1}, B格)
    assert not 更新 and 撤回 and "不一致" in 问题[0]["原因"]


def test_日期没变不需要写回():
    A = [_A结构()]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}]
    B格 = _B合集(("货币资金", "余额构成"), [("C2", "本期期末", "2024-12-31")])
    更新, 问题, 撤回, _ = 生成本次标题(A, B, _上下文(), {"sheet1!C2": 1}, B格)
    assert 更新 == [] and not 问题 and not 撤回


def test_另一个科目的日期不会被借用():
    A = [_A结构()]
    另 = _A结构(区域="TA_002", 科目="存货", 表名="存货明细")
    另["范围"] = "A4:C6"
    另["行范围"] = [4, 6]
    A.append(另)
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"},
         {"科目": "存货", "表名": "存货明细", "表组": "G2"}]
    B格 = _B合集(("货币资金", "余额构成"), [("C2", "本期期末", "2025-12-31")]) + \
        _B合集(("存货", "存货明细"), [("C8", "本期期末", "2026-12-31")])
    更新, 问题, 撤回, _ = 生成本次标题(
        A, B, {"tables": [{"range_name": "TA_001", "title": 标题句},
                          {"range_name": "TA_002", "title": 标题句}]},
        {"sheet1!C2": 1, "sheet1!C5": 2}, B格)
    assert {u["range_name"]: u["text"] for u in 更新}["TA_001"] == "截至2025年12月31日，本公司货币资金余额如下"
    assert {u["range_name"]: u["text"] for u in 更新}["TA_002"] == "截至2026年12月31日，本公司货币资金余额如下"


def test_日期写法按原片段保留():
    assert 格式化(解析日期("2025-12-31"), "2024年12月31日") == "2025年12月31日"
    assert 格式化(解析日期("2025-12-31"), "2024-12-31") == "2025-12-31"
    assert 格式化(解析日期("2025-01-05"), "2024/1/5") == "2025/1/5"
    assert 解析日期("截至2024年12月31日") is not None
    assert 解析日期("无日期") is None


def test_旧新文答复不直接写入():
    册 = {"货币资金": ["余额构成"]}
    划分 = [{"行范围": [1, 3], "科目": "货币资金"}]
    答 = ('{"表":[{"行范围":[1,3],"科目":"货币资金","表名":"余额构成","表组":"G1",'
         '"续表序":1,"标题":{"原文":"截至2024年12月31日","新文":"截至2099年12月31日"}}]}')
    表 = _解析表(答, 册, 划分)[0]
    assert 表["标题"]["旧新文已忽略"] is True
    assert 表["标题"]["旧新文"] == "截至2099年12月31日"
    assert 表["标题"]["日期"] == []
    # 没有日期规则就生成不了本次指令，整组撤回，2099 绝不会被写出去
    更新, 问题, 撤回, _ = 生成本次标题(
        [{**_A结构(), "标题": 表["标题"]}],
        [{"科目": "货币资金", "表名": "余额构成", "表组": "G1"}],
        _上下文(), {"sheet1!C2": 1},
        _B合集(("货币资金", "余额构成"), [("C2", "本期期末", "2025-12-31")]))
    assert not 更新 and 撤回


# —— 走完整任务：金额与上年相同，期间已变，标题仍要更新 ——

def _造(tmp_path, 年, 前缀, 金额):
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append(["项目", f"{年}-01-01", f"{年}-12-31"])   # 表头写着期间实际日期
    ws.append(["甲公司", 金额, 金额])
    A = tmp_path / f"{前缀}A.xlsx"
    wb.defined_names.add(DefinedName("TA_001", attr_text="'sheet1'!$A$1:$C$2"))
    wb.save(A)
    wb.close()
    b = Workbook()
    ws2 = b.active
    ws2.title = "word附注"
    ws2.append(["项目", f"{年}-01-01", f"{年}-12-31"])
    ws2.append(["甲公司", 金额, 金额])
    B = tmp_path / f"{前缀}B.xlsx"
    b.save(B)
    b.close()
    上下文 = {"tables": [{"range_name": "TA_001", "title": f"截至{年 - 1}年12月31日",
                       "unit_text_candidates": ["金额单位：人民币元"]}]}
    return A, B, 上下文


def _跑(tmp_path, 名, 年, 金额, monkeypatch, 记录目录=None, 带标题=True):
    金 = tmp_path / "金标准.jsonl"
    金.write_text(json.dumps({"id": "N001", "note": {"name": "货币资金"},
                              "table": {"name": "余额构成"},
                              "slot": {"name": "余额"}, "row_path": ["甲公司"],
                              "column_path": ["期末余额"],
                              "dimensions": [{"name": "period", "required": True}]},
                            ensure_ascii=False), encoding="utf-8")
    A, B, 上下文 = _造(tmp_path, 年, 名, 金额)

    def 假提取(word, excel, 带连接, 路径):
        shutil.copyfile(A, excel)
        Path(带连接).write_text("连接", encoding="utf-8")
        Path(路径).write_text(json.dumps(上下文, ensure_ascii=False), encoding="utf-8")
        return {"ok": True, "warnings": []}

    收到 = []

    def 假更新(word, excel, output, operations, caption_updates):
        收到.extend(caption_updates)
        Path(output).write_text(json.dumps({"excel": str(excel)}), encoding="utf-8")
        return {"ok": True}

    def 假回读(word, excel, linked, 路径):
        shutil.copyfile(json.loads(Path(word).read_text(encoding="utf-8"))["excel"], excel)

    def 模型(配置, 指令, 材料, 目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,2],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            表 = {"行范围": [1, 2], "科目": "货币资金", "表名": "余额构成",
                 "表组": "G1", "续表序": 1}
            if 带标题:
                表["标题"] = {"原文": f"截至{年 - 1}年12月31日",
                           "日期": [{"片段": f"{年 - 1}年12月31日", "期间角色": "本期期末"}]}
            return json.dumps({"表": [表]}, ensure_ascii=False)
        return json.dumps({"列": [{"列": "A", "路径": "项目", "角色": "行名列"},
                                {"列": "B", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                                 "期间实际": {"值": f"{年}-01-01", "位置": "B1"}},
                                {"列": "C", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                                 "期间实际": {"值": f"{年}-12-31", "位置": "C1"}}],
                            "行": [{"行": 1, "角色": "表头"},
                                  {"行": 2, "角色": "数据", "序号": 1, "槽位": {"B": "N001", "C": "N001"}}],
                            "作用范围": {"币种": "人民币", "单位": "元"}, "空名单": [], "例外": []},
                           ensure_ascii=False)

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假回读)
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    入 = {"word": "原件.docx", "b表": str(B), "金标准": str(金), "b工作表": "word附注"}
    if 记录目录:
        入["识别记录目录"] = str(记录目录)
    状态 = 运行任务(入, tmp_path / f"任务{名}", {"测试": True})
    return 状态, 收到


def test_金额全同期间已变标题仍更新(tmp_path, monkeypatch):
    状态, 收到 = _跑(tmp_path, "同额", 2025, 100, monkeypatch)
    assert 状态["步骤"] == "done", 状态.get("步骤")
    assert 收到 and 收到[0]["old_text"] == "截至2024年12月31日"
    assert 收到[0]["text"] == "截至2025年12月31日"       # 来自本次 B，不是上年
    Excel = Path(json.loads(Path(状态["产物"]["Word新副本"]).read_text(encoding="utf-8"))["excel"])
    from openpyxl import load_workbook
    wb = load_workbook(Excel)
    try:
        assert wb["sheet1"]["C2"].value == 100
    finally:
        wb.close()


def test_无需搬运且不改标题的对照组(tmp_path, monkeypatch):
    """识别没有提标题：这一次不生成任何标题指令，也不留标题问题。"""
    状态, 收到 = _跑(tmp_path, "同年", 2024, 100, monkeypatch, 带标题=False)
    assert 状态["步骤"] == "done"
    assert 收到 == []
    问题 = (tmp_path / "任务同年" / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "标题" not in 问题


def test_复用后的第二份仍按本次B日期更新标题(tmp_path, monkeypatch):
    """标题规则随解释复用，片段用本次标题回填，新文仍由本次 B 日期生成。"""
    monkeypatch.delenv("审计附注_复用", raising=False)
    记录目录 = 安全记录目录()
    状态1, 收到1 = _跑(tmp_path, "复一", 2025, 100, monkeypatch, 记录目录=记录目录)
    assert 状态1["步骤"] == "done"
    assert 收到1[0]["text"] == "截至2025年12月31日"
    状态2, 收到2 = _跑(tmp_path, "复二", 2026, 100, monkeypatch, 记录目录=记录目录)
    assert 状态2["步骤"] == "done"
    记录 = json.loads((tmp_path / "任务复二" / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第三遍"] == 0                     # 第三遍整组复用
    assert 收到2 and 收到2[0]["old_text"] == "截至2025年12月31日"
    assert 收到2[0]["text"] == "截至2026年12月31日"        # 新文来自本次 B


def test_只有标题骨架时按本次标题重建规则():
    """第二遍复用只剩骨架、第三遍重认丢了标题结论：仍要按本次标题生成写回指令。"""
    from 核心.标题 import 生成本次标题

    A = [{"区域": "TA_001", "科目": "货币资金", "表名": "余额构成", "表组": "G1",
          "工作表": "sheet1", "行范围": [1, 3], "范围": "A1:D3",
          "标题规则": {"有标题": True, "骨架": "截至{本期期末}", "角色": ["本期期末"]}}]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1",
          "维度值": {"期间角色": "本期期末", "期间实际": "2026-12-31"}, "格": "word附注!D2"}]
    上下文 = {"tables": [{"range_name": "TA_001", "title": "截至2025年12月31日",
                       "unit_text_candidates": ["金额单位：人民币元"]}]}
    更新, 问题, 撤回, _ = 生成本次标题(A, B, 上下文, {"sheet1!C2": 1}, B)
    assert not 问题 and not 撤回, (问题, 撤回)
    assert 更新[0]["old_text"] == "截至2025年12月31日"
    assert 更新[0]["text"] == "截至2026年12月31日"


def test_只有标题骨架时按本次标题重建规则():
    """第二遍复用只剩骨架、第三遍重认丢了标题结论：仍要按本次标题生成写回指令。"""
    from 核心.标题 import 生成本次标题

    A = [{"区域": "TA_001", "科目": "货币资金", "表名": "余额构成", "表组": "G1",
          "工作表": "sheet1", "行范围": [1, 3], "范围": "A1:D3",
          "标题规则": {"有标题": True, "骨架": "截至{本期期末}", "角色": ["本期期末"]}}]
    B = [{"科目": "货币资金", "表名": "余额构成", "表组": "G1",
          "维度值": {"期间角色": "本期期末", "期间实际": "2026-12-31"}, "格": "word附注!D2"}]
    上下文 = {"tables": [{"range_name": "TA_001", "title": "截至2025年12月31日",
                       "unit_text_candidates": ["金额单位：人民币元"]}]}
    更新, 问题, 撤回, _ = 生成本次标题(A, B, 上下文, {"sheet1!C2": 1}, B)
    assert not 问题 and not 撤回, (问题, 撤回)
    assert 更新[0]["old_text"] == "截至2025年12月31日"
    assert 更新[0]["text"] == "截至2026年12月31日"
