# coding: utf-8
"""三遍跨任务复用（走完整任务入口）：连续三份、跳过后补认、删中间科目、冲突。

调用次数按侧、按遍记录在 复用记录_{侧}.json；断言同时核对实际模型调用次数。
"""
import json
import re
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from 核心.流程 import 运行任务
from 核心.共用记录簿 import 记录簿, 发布事件
from 核心.复用证据 import 证据键, 解释键, 版本号

科目表 = [("货币资金", "余额构成"), ("应收账款", "按账龄披露"), ("存货", "存货明细")]


def _金标准(tmp_path):
    行们 = []
    for i, (科, 表) in enumerate(科目表, 1):
        行们.append({"id": f"N{i:03d}", "note": {"name": 科, "display_order": i},
                   "table": {"name": 表, "display_order": 1},
                   "slot": {"name": f"{科}／余额"}, "row_path": ["金额"],
                   "column_path": ["余额"],
                   "dimensions": [{"name": "period", "required": True}]})
    路径 = tmp_path / "金标准.jsonl"
    路径.write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们), encoding="utf-8")
    return 路径


def _造(tmp_path, 名, 科目们, 年):
    """A：每个科目一张命名区域，3 行；B：同布局的连续区块（统一底表）。"""
    wb = Workbook()
    wa = wb.active
    wa.title = "sheet1"
    底 = Workbook()
    wb2 = 底.active
    wb2.title = "word附注"
    顶 = 1
    表们 = []
    for i, (科, 表) in enumerate(科目们, 1):
        for r in range(顶, 顶 + 3):
            名2 = "项目" if r == 顶 else ("甲公司" if r == 顶 + 1 else "乙公司")
            期初 = f"{年}-01-01" if r == 顶 else 100 + r
            期末 = f"{年}-12-31" if r == 顶 else 200 + r
            wa.cell(r, 1, 名2); wa.cell(r, 2, 期初); wa.cell(r, 3, 期末)
            wb2.cell(r, 1, 名2); wb2.cell(r, 2, 期初); wb2.cell(r, 3, 期末)
        wa.cell(顶, 4, f"{科}（单位：元）")     # 每张表的固定说明文字，也是它唯一的身份锚点
        wb2.cell(顶, 4, f"{科}（单位：元）")
        区域 = f"TA_{i:03d}"
        wa.defined_names.add(DefinedName(区域, attr_text=f"'sheet1'!$A${顶}:$D${顶 + 2}"))
        表们.append({"range_name": 区域, "title": f"{科}（附注）", "sheet": "sheet1",
                   "first_row": 顶, "first_column": 1, "row_count": 3, "column_count": 4})
        顶 += 3
    A = tmp_path / f"A_{名}.xlsx"
    B = tmp_path / f"B_{名}.xlsx"
    wb.save(A); wb.close()
    底.save(B); 底.close()
    return A, {"tables": 表们}, B


def _模型(年, 侧科目):
    """固定答复；按调用目录前缀区分 A／B。返回（函数, 记录）。"""
    记录 = []

    def 调用(配置, 指令, 材料, 目录):
        侧 = "A" if Path(目录).name.startswith("A_") else "B"
        科目们 = 侧科目[侧]
        记录.append(f"{侧}:{Path(目录).name}")
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [3 * (i - 1) + 1, 3 * i], "科目": 科}
                                    for i, (科, _) in enumerate(科目们, 1)]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [3 * (i - 1) + 1, 3 * i], "科目": 科,
                                    "表名": 表, "表组": f"G{i}", "续表序": 1}
                                   for i, (科, 表) in enumerate(科目们, 1)]}, ensure_ascii=False)
        科目 = re.search(r"科目：(.+)", 材料).group(1).strip()
        表名 = re.search(r"表名：(.+)", 材料).group(1).strip()
        # 编号跟着科目走，行号跟着本侧顺序走：少了科目时行号会变、编号不变
        编号 = f"N{next(i for i, (科, _) in enumerate(科目表, 1) if 科 == 科目):03d}"
        序 = next(i for i, (科, t) in enumerate(科目们, 1) if 科 == 科目 and t == 表名)
        顶 = 3 * (序 - 1) + 1
        return json.dumps({
            "列": [{"列": "A", "路径": "项目", "角色": "行名列", "期间角色": None},
                  {"列": "B", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                   "期间实际": {"值": f"{年}-01-01", "位置": f"B{顶}"}},
                  {"列": "C", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                   "期间实际": {"值": f"{年}-12-31", "位置": f"C{顶}"}},
                  {"列": "D", "路径": "单位说明", "角色": "说明列", "期间角色": None}],
            "行": [{"行": 顶, "名字": "项目", "角色": "表头"},
                  {"行": 顶 + 1, "名字": "甲公司", "角色": "数据", "序号": 1,
                   "槽位": {"B": 编号, "C": 编号}},
                  {"行": 顶 + 2, "名字": "乙公司", "角色": "数据", "序号": 2,
                   "槽位": {"B": 编号, "C": 编号}}],
            "作用范围": {"币种": "人民币", "单位": "元", "期间跨度": "时点"},
            "空名单": [], "例外": []}, ensure_ascii=False)
    return 调用, 记录


def _跑(tmp_path, 名, 科目A, 科目B, 年, 记录目录, 金, monkeypatch):
    A, 上下文, B = _造(tmp_path, f"{名}A", 科目A, 年)
    _, _, B = _造(tmp_path, f"{名}B", 科目B, 年)

    def 假提取(word, excel, 带连接, 上下文路径):
        shutil.copyfile(A, excel)
        Path(带连接).write_text("连接体", encoding="utf-8")
        Path(上下文路径).write_text(json.dumps(上下文, ensure_ascii=False), encoding="utf-8")
        return {"ok": True, "warnings": []}

    def 假更新(word, excel, output, operations, caption_updates):
        Path(output).write_text(json.dumps({"excel": str(excel)}, ensure_ascii=False), encoding="utf-8")
        return {"ok": True}

    def 假回读(word, excel, linked, 上下文路径):
        shutil.copyfile(json.loads(Path(word).read_text(encoding="utf-8"))["excel"], excel)

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假回读)
    模型, 记录 = _模型(年, {"A": 科目A, "B": 科目B})
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    任务 = tmp_path / f"任务_{名}"
    状态 = 运行任务({"word": "原件.docx", "b表": str(B), "金标准": str(金),
                 "b工作表": "word附注", "识别记录目录": str(记录目录)},
                任务, {"测试": True})
    assert 状态["步骤"] in ("done",), 状态.get("步骤")
    return 任务, 记录


def _复用记录(任务, 侧):
    return json.loads((任务 / f"复用记录_{侧}.json").read_text(encoding="utf-8"))


def _调用(任务, 侧):
    return _复用记录(任务, 侧)["调用"]


def _开复用(monkeypatch):
    monkeypatch.delenv("审计附注_复用", raising=False)


def test_连续三份同一布局后两份全复用(tmp_path, monkeypatch):
    _开复用(monkeypatch)
    金 = _金标准(tmp_path)
    记录目录 = tmp_path / "记录"
    任务1, 记1 = _跑(tmp_path, "一", 科目表, 科目表, 2025, 记录目录, 金, monkeypatch)
    assert _调用(任务1, "A") == {"第一遍": 1, "第二遍": 1, "第三遍": 3, "重试": 0}
    assert _调用(任务1, "B") == {"第一遍": 1, "第二遍": 1, "第三遍": 3, "重试": 0}
    assert not _复用记录(任务1, "A")["复用"]
    assert len(记1) == 10

    任务2, 记2 = _跑(tmp_path, "二", 科目表, 科目表, 2026, 记录目录, 金, monkeypatch)
    assert _调用(任务2, "A") == {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0}
    assert _调用(任务2, "B") == {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0}
    assert len(记2) == 0                                   # 实际一次模型都没发
    assert len(_复用记录(任务2, "A")["复用"]) == 5          # 分科目、分表、三个表组

    任务3, 记3 = _跑(tmp_path, "三", 科目表, 科目表, 2027, 记录目录, 金, monkeypatch)
    assert _调用(任务3, "A") == {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0}
    assert _调用(任务3, "B") == {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0}
    assert len(记3) == 0

    合集 = [json.loads(x) for x in (任务3 / "A合集.jsonl").read_text(encoding="utf-8").splitlines()]
    值 = {x["原值"] for x in 合集 if x["角色"] == "数据格"}
    assert {102, 103, 202, 203, 108, 109, 208, 209} <= 值   # 值来自本次文件，不是旧记录
    结构 = [json.loads(x) for x in (任务3 / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()]
    assert all(x["完整性"]["第三遍逐行"] for x in 结构)


def test_首份B因A缺科目跳过第二份补认一组(tmp_path, monkeypatch):
    _开复用(monkeypatch)
    金 = _金标准(tmp_path)
    记录目录 = tmp_path / "记录"
    任务1, _ = _跑(tmp_path, "缺", 科目表[:2], 科目表, 2025, 记录目录, 金, monkeypatch)
    assert _调用(任务1, "B")["第三遍"] == 2                # 存货组被粗比跳过
    结构 = [json.loads(x) for x in (任务1 / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()]
    assert [x.get("跳过") for x in 结构 if x["侧"] == "B"].count("粗比：科目不在A") == 1

    任务2, 记2 = _跑(tmp_path, "补", 科目表, 科目表, 2026, 记录目录, 金, monkeypatch)
    b = _调用(任务2, "B")
    assert (b["第一遍"], b["第二遍"], b["第三遍"]) == (0, 0, 1)   # 只补认这一个单物理表组
    assert len([x for x in 记2 if x.startswith("B:")]) == 1
    assert _复用记录(任务2, "B")["复用"][0]["步骤"] == "第一遍"


def test_A删除中间科目后面的表移位仍复用(tmp_path, monkeypatch):
    _开复用(monkeypatch)
    金 = _金标准(tmp_path)
    记录目录 = tmp_path / "记录"
    任务1, _ = _跑(tmp_path, "全", 科目表, 科目表, 2025, 记录目录, 金, monkeypatch)
    assert _调用(任务1, "A")["第三遍"] == 3
    任务2, 记2 = _跑(tmp_path, "少", [科目表[0], 科目表[2]], 科目表, 2026,
                  记录目录, 金, monkeypatch)
    a = _调用(任务2, "A")
    assert (a["第一遍"], a["第二遍"]) == (0, 0)              # 表题锚点重定位成功
    assert a["第三遍"] == 0
    assert len([x for x in 记2 if x.startswith("A:")]) == 0
    划分 = [json.loads(x) for x in (任务2 / "科目划分_A.jsonl").read_text(encoding="utf-8").splitlines()]
    assert [(x["行范围"], x["科目"]) for x in 划分] == [([1, 3], "货币资金"), ([4, 6], "存货")]
    合集 = [json.loads(x) for x in (任务2 / "A合集.jsonl").read_text(encoding="utf-8").splitlines()]
    值 = {x["原值"] for x in 合集 if x["角色"] == "数据格"}
    assert {105, 205, 106, 206} <= 值                       # 存货已落到本次位置


def test_同一结构两种解释不自动选用(tmp_path, monkeypatch):
    _开复用(monkeypatch)
    金 = _金标准(tmp_path)
    记录目录 = tmp_path / "记录"
    _跑(tmp_path, "起", 科目表, 科目表, 2025, 记录目录, 金, monkeypatch)
    索引 = 记录簿(记录目录).快照()
    项0 = next(x for x in 索引["解释"].values() if x["侧"] == "A" and x["步骤"] == "第三遍")
    冲突解释 = json.loads(json.dumps(项0["解释"], ensure_ascii=False))
    冲突解释["行"][1]["槽位"] = {"B": "N002", "C": "N002"}
    键 = 证据键(项0["证据"])
    解 = 解释键(冲突解释)
    v2 = 版本号(键, 解, 项0["金标准版本"], 项0["规则版本"])
    记录簿(记录目录).提交([发布事件("A", "第三遍", 项0["证据"], 冲突解释, 键, 解, v2,
                          项0["金标准版本"], 项0["规则版本"],
                          {"任务": "另一份", "侧": "A", "表组": "G1"}, 形状键=None)], "冲突注入")
    任务2, 记2 = _跑(tmp_path, "撞", 科目表, 科目表, 2026, 记录目录, 金, monkeypatch)
    a = _调用(任务2, "A")
    assert (a["第一遍"], a["第二遍"]) == (0, 0)
    assert a["第三遍"] >= 1 and len([x for x in 记2 if x.startswith("A:")]) >= 1
    assert any("冲突" in (x.get("原因") or "") for x in _复用记录(任务2, "A")["回退"])
