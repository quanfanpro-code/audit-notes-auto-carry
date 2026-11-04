# coding: utf-8
"""只补测试：复核97ea504的错误路径、旧任务接续及文本一致性。"""
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from 测试.test_第五阶段独立验收 import 金原, _造场景, _两侧识别, _写证据
from 测试.test_第五阶段独立验收 import test_回读区分文本金额和空白 as _验回读
from 核心.流程 import 运行任务
from 核心.识别 import 识别一侧
from 核心.比对 import 建立对应
from 核心.搬运 import 生成更新表, 保存更新表

根 = Path(__file__).resolve().parent.parent


def _桥接替身(monkeypatch, 源表, 上下文表们, 首次失败=None):
    收到 = []
    def 提取(word, excel, linked, context):
        shutil.copyfile(源表, excel)
        Path(linked).write_text("模拟连接", encoding="utf-8")
        Path(context).write_text(json.dumps({"tables": 上下文表们}, ensure_ascii=False), encoding="utf-8")
        return {"ok": True}
    def 更新(word, excel, output, operations, caption_updates):
        wb = load_workbook(excel)
        try:
            格 = {c.coordinate: c.value for row in wb.active for c in row}
        finally:
            wb.close()
        收到.append({"标题": list(caption_updates), "Excel": 格})
        if 首次失败 and len(收到) == 1:
            raise ValueError("标题原文无法唯一定位：" + 首次失败)
        Path(output).write_text(json.dumps({"excel": str(excel)}), encoding="utf-8")
        return {"ok": True}
    def 回读提取(word, excel, linked, context):
        src = json.loads(Path(word).read_text(encoding="utf-8"))["excel"]
        shutil.copyfile(src, excel)
    monkeypatch.setattr("核心.提取.调用桥接提取", 提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 回读提取)
    return 收到


@pytest.mark.parametrize("冲突", [False, True])
def test_整表口径冲突时标题与金额同进同退(tmp_path, monkeypatch, 金原, 冲突):
    文件, 金, 原模型 = _造场景(tmp_path, monkeypatch, 金原,
                          [(101, 111), (201, 211), (303, 313)], 标题=True)
    if 冲突:
        wb = load_workbook(文件["B"])
        wb.active["A1"] = "项目（单位：万元）"
        wb.save(文件["B"])
        wb.close()
    def 模型(配置, 指令, 材料, 调用目录):
        答 = json.loads(原模型(配置, 指令, 材料, 调用目录))
        if 冲突 and Path(调用目录).name.startswith("B_") and "作用范围" in 答:
            答["作用范围"]["单位"] = "万元"
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    收到 = _桥接替身(monkeypatch, 文件["A"],
                   [{"range_name": "TA_001", "title": "截至2024年12月31日"}])
    原件 = tmp_path / "模拟原件.docx"
    原件.write_text("模拟桥接边界", encoding="utf-8")
    入 = {"word": str(原件), "b表": str(文件["B"]), "金标准": str(金), "b工作表": "word附注"}
    前 = {k: Path(入[k]).read_bytes() for k in ("word", "b表", "金标准")}
    状态 = 运行任务(入, tmp_path / "任务", {"测试": True})
    _写证据(tmp_path, 口径冲突=冲突, 步骤=状态["步骤"], 桥接收到=收到,
          问题清单=(tmp_path / "任务" / "问题清单.jsonl").read_text(encoding="utf-8"))
    assert all(Path(入[k]).read_bytes() == v for k, v in 前.items())
    if 冲突:
        assert 收到[-1]["Excel"]["B2"] == 100
        assert 收到[-1]["标题"] == [], "整表金额未换，日期标题却仍发出了更新"
    else:
        assert 收到[-1]["Excel"]["B2"] == 101
        assert 收到[-1]["标题"][0]["text"] == "截至2025年12月31日"


@pytest.mark.parametrize("有零项", [False, True])
def test_原序号超出A容量的项目必须有新增项目提示(tmp_path, monkeypatch, 金原, 有零项):
    _两侧识别(tmp_path, monkeypatch, 金原,
          [(101, 111), (0, 0) if 有零项 else (201, 211), (303, 313), (404, 414)])
    证据 = json.loads((tmp_path / "实际结果.json").read_text(encoding="utf-8"))
    提示 = [x.get("原因", "") for x in 证据["对应"] if str(x.get("原因") or "").startswith("B多出项")]
    assert any("404" in x and "414" in x for x in 提示), 提示


@pytest.mark.parametrize("期望,实际", [("甲公司 ", "甲公司"), ("00123", "00123 "), ("甲公司\n", "甲公司")])
def test_文字逐字回读不能去掉前后字符(tmp_path, monkeypatch, 期望, 实际):
    _验回读(tmp_path, monkeypatch, 期望, 实际, False)


def test_桥接真实子进程超时而不是只检查参数(tmp_path, monkeypatch):
    import 核心.桥接 as 模
    脚本 = tmp_path / "等待桥接.py"
    脚本.write_text("import time; time.sleep(5); print('{\"ok\":true}')", encoding="utf-8")
    monkeypatch.setattr(模, "桥接超时秒", 0.15)
    起 = time.monotonic()
    with pytest.raises(RuntimeError, match="没有返回"):
        模.调用桥接({}, 工作目录=tmp_path, exe=脚本, 前缀=sys.executable)
    秒 = time.monotonic() - 起
    _写证据(tmp_path, 测试超时秒=0.15, 实际等待秒=秒)
    assert 秒 < 2


def _旧号任务(tmp_path, monkeypatch, 金原, 用旧索引, 仅建旧结构=False):
    旧 = next(r for r in 金原 if r["id"] == "C-N054-T001-S002")
    新 = next(r for r in 金原 if r["id"] == "C-N054-T001-S022")
    assert 旧["status"] == "superseded" and 新["status"] == "active"
    字段 = ("id", "note", "table", "slot", "row_path", "column_path", "dimensions", "value_type", "status", "superseded_by")
    金 = tmp_path / "金标准.jsonl"
    金.write_text("\n".join(json.dumps({k: r.get(k) for k in 字段}, ensure_ascii=False)
                           for r in (旧, 新)), encoding="utf-8")
    文件 = {}
    for 侧, 表名, 值 in (("A", "sheet1", 100), ("B", "word附注", 999)):
        wb = Workbook()
        wb.active.title = 表名
        wb.active["A1"] = 值
        if 侧 == "A":
            wb.defined_names.add(DefinedName("TA_001", attr_text="'sheet1'!$A$1:$A$1"))
        文件[侧] = tmp_path / (侧 + ".xlsx")
        wb.save(文件[侧])
        wb.close()
    def 模型(配置, 指令, 材料, 目录):
        if 指令.startswith("把整份文件"):
            答 = {"划分": [{"行范围": [1, 1], "科目": 旧["note"]["name"]}]}
        elif 指令.startswith("在第一遍"):
            答 = {"表": [{"行范围": [1, 1], "科目": 旧["note"]["name"],
                       "表名": 旧["table"]["name"], "表组": "G", "续表序": 1}]}
        else:
            期间 = "本期期初" if Path(目录).name.startswith("A_") else "本期期末"
            答 = {"列": [{"列": "A", "角色": "数据列", "期间角色": 期间}],
                 "行": [{"行": 1, "角色": "数据", "槽位": {"A": 旧["id"]}}]}
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    任务 = tmp_path / "任务"
    def 两侧():
        识别一侧({"测试": True}, 文件["A"], "A", 任务, 金)
        识别一侧({"测试": True}, 文件["B"], "B", 任务, 金, 只跑表=["word附注"])
        读 = lambda 名: [json.loads(x) for x in (任务 / 名).read_text(encoding="utf-8").splitlines()]
        a, b, 结构 = 读("A合集.jsonl"), 读("B合集.jsonl"), 读("结构清单.jsonl")
        对 = 建立对应(a, b, [x for x in 结构 if x["侧"] == "A"], [x for x in 结构 if x["侧"] == "B"])
        改, 问题 = 生成更新表({}, 对, b, a)
        return {"A": a, "B": b, "对应": 对, "变更": 改, "问题": 问题}
    前 = None
    if 用旧索引:
        # 仅在内存中模拟旧版不筛状态的索引，原金标准文件仍保存停用状态。
        from 核心.金标准 import 建索引
        def 旧索引(条目):
            return 建索引([{**r, "status": "active"} for r in 条目])
        with monkeypatch.context() as mp:
            mp.setattr("核心.识别.建索引", 旧索引)
            前 = 两侧()
            assert 前["变更"] == {}, 前
        if 仅建旧结构:
            return 前
    后 = 两侧()
    保存更新表(文件["A"], 后["变更"], tmp_path / "实际更新表.xlsx")
    _写证据(tmp_path, 使用旧完成结构=用旧索引, 旧索引下=前, 新索引下=后)
    return 后


def test_新任务拒绝已停用号_对照(tmp_path, monkeypatch, 金原):
    结果 = _旧号任务(tmp_path, monkeypatch, 金原, False)
    assert 结果["变更"] == {}
    assert not any(r.get("槽位") == "C-N054-T001-S002" for r in 结果["A"])


def test_旧完成结构不能绕过新停用号校验并串期(tmp_path, monkeypatch, 金原):
    _旧号任务(tmp_path, monkeypatch, 金原, True, 仅建旧结构=True)
    旧文件 = {p: p.read_bytes() for p in (tmp_path / "任务").rglob("*") if p.is_file()}
    with pytest.raises(RuntimeError, match="失效.*新建任务"):
        识别一侧({"测试": True}, tmp_path / "A.xlsx", "A", tmp_path / "任务", tmp_path / "金标准.jsonl")
    assert all(p.read_bytes() == b for p, b in 旧文件.items())


def test_续表一处失败后整组数字与标题一起撤回(tmp_path, monkeypatch):
    编号 = "S1"
    金 = tmp_path / "金标准.jsonl"
    金.write_text(json.dumps({"id": 编号, "note": {"name": "现金"},
                             "table": {"name": "余额"}, "slot": {"name": "金额"},
                             "dimensions": [{"name": "period", "required": True}]}), encoding="utf-8")
    源 = tmp_path / "A续表.xlsx"
    wb = Workbook()
    s = wb.active
    s.title = "sheet1"
    for 顶, 号, 值 in ((1, 1, 100), (4, 2, 200)):
        s.cell(顶, 1, "项目")
        s.cell(顶, 2, "期末余额")
        s.cell(顶+1, 1, "旧" + str(号))
        s.cell(顶+1, 2, 值)
        wb.defined_names.add(DefinedName("TA_00" + str(号), attr_text=f"'sheet1'!$A${顶}:$B${顶+1}"))
    wb.save(源)
    wb.close()
    B = tmp_path / "B.xlsx"
    wb = Workbook()
    wb.active.title = "word附注"
    for row in (["项目", "期末余额"], ["新1", 101], ["新2", 202]):
        wb.active.append(row)
    wb.save(B)
    wb.close()
    def 模型(配置, 指令, 材料, 目录):
        A侧 = Path(目录).name.startswith("A_")
        if 指令.startswith("把整份文件"):
            答 = {"划分": [{"行范围": [1, 5 if A侧 else 3], "科目": "现金"}]}
        elif 指令.startswith("在第一遍"):
            答 = {"表": [{"行范围": [顶, 底], "科目": "现金", "表名": "余额",
                       "表组": "G", "续表序": i+1,
                       **({"标题": {"原文": f"旧日期{i+1}", "新文": f"新日期{i+1}"}} if A侧 else {})}
                      for i, (顶, 底) in enumerate([(1, 2), (4, 5)] if A侧 else [(1, 3)])]}
        else:
            if A侧:
                第二表 = Path(目录).name.endswith("002")
                顶, 数据行, 序 = (4, 5, 2) if 第二表 else (1, 2, 1)
                行 = [{"行": 顶, "角色": "表头"}, {"行": 数据行, "角色": "数据",
                       "序号": 序, "槽位": {"B": 编号}}]
            else:
                行 = [{"行": 1, "角色": "表头"},
                     {"行": 2, "角色": "数据", "序号": 1, "槽位": {"B": 编号}},
                     {"行": 3, "角色": "数据", "序号": 2, "槽位": {"B": 编号}}]
            答 = {"列": [{"列": "A", "角色": "行名列"},
                        {"列": "B", "角色": "数据列", "期间角色": "本期期末"}],
                 "行": 行, "作用范围": {}}
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    收到 = _桥接替身(monkeypatch, 源, [
        {"range_name": "TA_001", "title": "旧日期1"},
        {"range_name": "TA_002", "title": "旧日期2"}], 首次失败="TA_001")
    Word = tmp_path / "模拟.docx"
    Word.write_text("桥接替身", encoding="utf-8")
    状态 = 运行任务({"word": str(Word), "b表": str(B), "金标准": str(金), "b工作表": "word附注"},
               tmp_path / "任务", {"测试": True})
    _写证据(tmp_path, 步骤=状态["步骤"], 撤回=状态["撤回表"], 桥接收到=收到)
    assert 收到[-1]["Excel"]["B2"] == 100 and 收到[-1]["Excel"]["B5"] == 200
    assert 收到[-1]["标题"] == [], "整组数字撤回了，另一页的标题却仍更新"


def test_完整续跑不能把旧号的期末金额写进期初(tmp_path, monkeypatch, 金原):
    from 核心.流程 import _输入指纹
    from 核心.接续 import 读索引, 写索引
    _旧号任务(tmp_path, monkeypatch, 金原, True, 仅建旧结构=True)
    Word = tmp_path / "模拟原件.docx"
    Word.write_text("模拟原件", encoding="utf-8")
    上下文 = tmp_path / "旧任务上下文.json"
    上下文.write_text(json.dumps({"tables": [{"range_name": "TA_001"}]}), encoding="utf-8")
    任务 = tmp_path / "任务"
    入 = {"word": str(Word), "b表": str(tmp_path / "B.xlsx"),
         "金标准": str(tmp_path / "金标准.jsonl"), "b工作表": "word附注"}
    索引 = 读索引(任务)
    索引["输入指纹"] = _输入指纹(入)
    索引["步骤"] = "done"                # 真实旧任务可能已经完成，拒绝续跑不能仍显示完成
    索引["产物"].update({"表A": str(tmp_path / "A.xlsx"), "带连接Word": str(Word),
                      "上下文": str(上下文)})
    写索引(任务, 索引)
    收到 = _桥接替身(monkeypatch, tmp_path / "A.xlsx", [{"range_name": "TA_001"}])
    with pytest.raises(RuntimeError, match="失效.*新建任务"):
        运行任务(入, 任务, {"测试": True})
    assert 收到 == []
    assert not (任务 / "更新后附注表格.xlsx").exists()
    assert 读索引(任务)["步骤"] == "识别失败"
