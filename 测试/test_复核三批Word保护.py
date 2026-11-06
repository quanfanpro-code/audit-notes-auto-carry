# coding: utf-8
"""复核三批：真实Word主表、续表各带标题，任一张标题核对失败时整组保留的保护用例。"""
import copy
import json
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from 测试.组合样本 import 造Word, 文件哈希, 读Word格, 读Word段落, 跑场景


def _标题金(路径):
    行们 = []
    for 编号, 科目, 表名 in (("N001", "货币资金", "余额构成"), ("N002", "货币资金", "余额构成"),
                          ("N003", "存货", "存货明细"), ("N004", "存货", "存货明细")):
        行们.append({"id": 编号, "note": {"name": 科目}, "table": {"name": 表名},
                    "slot": {"name": 编号}, "dimensions": [{"name": "period", "required": True}]})
    Path(路径).write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们),
                          encoding="utf-8")
    return Path(路径)


@pytest.mark.parametrize("坏哪张", ["续表", "主表"])
def test_Word续表标题失败整组保留独立组正常(tmp_path, monkeypatch, 坏哪张):
    """真实Word主表、续表各带标题，一张的标题原文与文件不符时，
    本组两条标题、金额、名单全部原样；同文件独立业务组照常更新；状态部分保留。"""
    from lxml import etree as E

    from 测试.组合样本 import W

    记录 = tmp_path / "识别记录"
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    monkeypatch.setenv("审计附注_复用", "0")

    主表题, 续表题, 独立题 = ("截至2024年12月31日（主表）",
                          "截至2024年12月31日（续表）",
                          "截至2024年12月31日（独立）")
    文内续表题 = f"注：{续表题}" if 坏哪张 == "续表" else 续表题
    文内主表题 = f"注：{主表题}" if 坏哪张 == "主表" else 主表题
    基 = 造Word(tmp_path / "主表样本.docx", 标题=文内主表题, 表内单位=True, 表前单位=False)
    A = tmp_path / "A.docx"
    with zipfile.ZipFile(基) as zin, zipfile.ZipFile(A, "w") as zout:
        for 条 in zin.infolist():
            数据 = zin.read(条.filename)
            if 条.filename == "word/document.xml":
                树 = E.fromstring(数据)
                body = 树.find(W + "body")
                表原型 = body.find(W + "tbl")
                for 题 in (文内续表题, 独立题):
                    段 = copy.deepcopy(body.find(W + "p"))
                    段.find(".//" + W + "t").text = 题
                    表 = copy.deepcopy(表原型)
                    body.append(段)
                    body.append(表)
                数据 = E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
            zout.writestr(条, 数据)
    wb = Workbook()
    wb.remove(wb.active)
    for 表名, 块顶们, 起序 in (("余额", (1, 4), 1), ("货源", (1,), 5)):
        ws = wb.create_sheet(表名)
        for n, 顶 in enumerate(块顶们):
            ws.cell(顶, 2, "项目（单位：人民币元）")
            ws.cell(顶, 3, "2025-01-01")
            ws.cell(顶, 4, "2025-12-31")
            for j in (1, 2):
                序 = 起序 + n * 2 + j - 1
                ws.cell(顶 + j, 2, f"本次{序}")
                ws.cell(顶 + j, 3, 900 + 序)
                ws.cell(顶 + j, 4, 100 + 序)
    B = tmp_path / "B.xlsx"
    wb.save(B)
    wb.close()
    金 = _标题金(tmp_path / "金标准.jsonl")
    前 = (文件哈希(A), 文件哈希(B))
    输入 = {"A": {"路径": str(A), "工作表": []},
            "B": {"路径": str(B), "工作表": ["余额", "货源"]},
            "金标准": str(金), "识别记录目录": str(记录), "复用": False}

    def 模型():
        def 调用(配置, 指令, 材料, 目录):
            目录 = Path(目录)
            侧 = 目录.name[0]
            if 侧 == "A":
                顶们 = (2, 8, 14)
                科目们 = (("货币资金", "余额构成", "N001"), ("货币资金", "余额构成", "N001"),
                       ("存货", "存货明细", "N003"))
            elif "P001" in 目录.parts:
                顶们 = (1, 4)
                科目们 = (("货币资金", "余额构成", "N001"),) * 2
            else:
                顶们 = (1,)
                科目们 = (("存货", "存货明细", "N003"),)
            if 指令.startswith("把整份文件"):
                if 侧 == "A":
                    划分 = [{"行范围": [1, 1], "科目": "非表格"},
                            {"行范围": [2, 6], "科目": "货币资金"},
                            {"行范围": [7, 7], "科目": "非表格"},
                            {"行范围": [8, 12], "科目": "货币资金"},
                            {"行范围": [13, 13], "科目": "非表格"},
                            {"行范围": [14, 16], "科目": "存货"}]
                elif "P001" in 目录.parts:
                    划分 = [{"行范围": [1, 6], "科目": "货币资金"}]
                else:
                    划分 = [{"行范围": [1, 3], "科目": "存货"}]
                return json.dumps({"划分": 划分}, ensure_ascii=False)
            if 指令.startswith("在第一遍"):
                if 侧 == "A":
                    表们 = [{"行范围": [2, 4], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 1},
                            {"行范围": [8, 10], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 2},
                            {"行范围": [14, 16], "科目": "存货", "表名": "存货明细", "表组": "G2", "续表序": 1}]
                    题们 = (主表题, 续表题, 独立题)
                    for i, 表 in enumerate(表们):
                        表["标题"] = {"原文": 题们[i],
                                      "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"}]}
                elif "P001" in 目录.parts:
                    表们 = [{"行范围": [1, 3], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 1},
                            {"行范围": [4, 6], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 2}]
                else:
                    表们 = [{"行范围": [1, 3], "科目": "存货", "表名": "存货明细", "表组": "G2", "续表序": 1}]
                return json.dumps({"表": 表们}, ensure_ascii=False)
            序 = int(目录.name.rsplit("表", 1)[1]) - 1
            顶, (科目, 表名, 编号) = 顶们[序], 科目们[序]
            名单基 = {0: 0, 1: 2, 2: 0}[序] if 侧 == "A" else ((0, 2)[序] if "P001" in 目录.parts else 0)
            列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
                  {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                   "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
                  {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                   "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
            行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
                {"行": 顶 + i, "名字": f"名单{名单基 + i}", "角色": "数据", "序号": 名单基 + i,
                 "槽位": {"C": 编号, "D": 编号}} for i in (1, 2)]
            作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
                {"来源": "表内", "位置": f"B{顶}", "原文": "项目（单位：人民币元）", "片段": "元",
                 "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
            return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                              ensure_ascii=False)
        return 调用

    结果 = 跑场景(输入, tmp_path / "任务", 模型(), monkeypatch)
    assert 结果["状态"]["步骤"] == "部分保留", 结果["状态"]
    输出 = 结果["输出"]
    段落们 = [x for x in 读Word段落(输出) if x.strip()]
    assert 文内主表题 in 段落们, 段落们          # 主表标题保持原样（含失败方向的坏前缀）
    assert 文内续表题 in 段落们, 段落们          # 续表标题保持原样
    assert "截至2025年12月31日（独立）" in 段落们, 段落们   # 独立组标题按本次更新
    assert [读Word格(输出, 0, 1, c) for c in (0, 1, 2)] == ["甲公司", "100", "110"]
    assert [读Word格(输出, 1, 1, c) for c in (0, 1, 2)] == ["甲公司", "100", "110"]
    assert [读Word格(输出, 2, 1, c) for c in (0, 1, 2)] == ["本次5", "905", "105"]
    问题 = (结果["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "本次标题原文与识别给出的原文不一致" in 问题, 问题
    assert "相关组的标题与数值一并保持原样" in 问题, 问题
    assert 前 == (文件哈希(A), 文件哈希(B))
