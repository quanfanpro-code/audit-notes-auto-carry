# coding: utf-8
"""全链集成测试（stub 模型/桥接）：断言实际写入值、操作记录、原件哈希不变。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import hashlib, json
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from 核心.流程 import 运行任务


def md5(路径):
    return hashlib.md5(open(路径, "rb").read()).hexdigest()




def _造金标准(tmp_path):
    路径 = tmp_path / "金标准.jsonl"
    路径.write_text(json.dumps({"note": {"name": "货币资金", "display_order": 1},
                     "table": {"name": "余额构成", "display_order": 1},
                     "slot": {"name": "库存现金 / 期末余额"}, "id": "C-N006-T001-S002",
                     "row_path": ["库存现金"], "column_path": ["期末余额"]},
                    ensure_ascii=False) + "\n", encoding="utf-8")
    return 路径




def _假提取(A值, word, excel, 带连接word, 上下文):
    wb = Workbook(); wb.active.title = "sheet1"; wb.active["A1"] = A值
    wb.active["B1"] = "单位：人民币元"; wb.save(excel)
    Path(带连接word).write_text("连接体", encoding="utf-8")
    Path(上下文).write_text(json.dumps({"tables": []}), encoding="utf-8")
    return {"ok": True, "table_count": 1, "warnings": []}




def test_全链stub(tmp_path, monkeypatch):
    # 真实小样本：B 是真实 xlsx（工作表 word附注，值为 9）；A 由假提取产出（sheet1，值为 100）
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径); b前 = md5(b路径)
    金路径 = _造金标准(tmp_path)

    def 假提取(word, excel, 带连接word, 上下文):
        return _假提取(9 if "新副本" in str(word) else 100, word, excel, 带连接word, 上下文)

    def 假模型(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,1],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,1],"科目":"货币资金","表名":"余额构成",'
                    '"表组":"G1","续表序":1}]}')
        return ('{"列":[{"列":"A","路径":"期末余额","角色":"数据列","期间角色":"本期期末"},{"列":"B","路径":"单位","角色":"说明列"}],'
                '"行":[{"行":1,"名字":"库存现金","角色":"数据","序号":null,"槽位":{"A":"C-N006-T001-S002"}}],'
                '"作用范围":{"币种":"人民币","单位":"元","单位依据":[{"来源":"表内","位置":"B1","原文":"单位：人民币元","片段":"人民币元","单位":"元","币种":"人民币","列":["A"]}]},"空名单":[],"例外":[]}')

    更新入参 = {}

    def 假更新(word, excel, output, operations, caption_updates):
        更新入参.update(excel=excel, operations=operations, caption_updates=caption_updates)
        Path(output).write_text("新Word", encoding="utf-8")
        return {"ok": True, "written_cells": 1, "errors": []}

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.识别.调用识别", 假模型)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假提取)

    运行目录 = tmp_path / "任务1"
    状态 = 运行任务({"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)},
                运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    assert 状态["步骤"] == "done"
    新 = load_workbook(运行目录 / "更新后附注表格.xlsx")
    v = 新["sheet1"]["A1"].value
    assert v == 9 and isinstance(v, (int, float))   # 实际写入 B 的数字（文件读取，非模型）
    assert 更新入参["operations"] == []              # 结构未变 → 无结构操作
    assert (运行目录 / "回读记录.jsonl").read_text(encoding="utf-8").strip() == ""   # 无回读差异
    assert md5(b路径) == b前                         # 原件不动
    对应 = (运行目录 / "对应清单.jsonl").read_text(encoding="utf-8").splitlines()
    assert len(对应) == 1 and json.loads(对应[0])["状态"] == "对应"   # 两侧挂到同一编号




def test_无配置时停在待答复并写出材料(tmp_path, monkeypatch):
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    monkeypatch.setattr("核心.提取.调用桥接提取",
                    lambda word, excel, 带连接word, 上下文: _假提取(100, word, excel, 带连接word, 上下文))
    状态 = 运行任务({"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))},
                tmp_path / "任务2", {})
    assert 状态["步骤"] == "identify_a"          # 停在第一遍，不硬往下走
    材料 = Path(状态["待答复"][0]) / "材料.txt"
    assert 材料.is_file() and "整份文件的表格文字" in 材料.read_text(encoding="utf-8")




# —— R7：标题改不了整表不换数、回读差异不算完成、人工改过的产物不被覆盖 ——



def _读值(路径, 表="sheet1", 格="A1"):
    wb = load_workbook(路径)
    try:
        return wb[表][格].value
    finally:
        wb.close()




def _装stub(monkeypatch, 上下文段="本表按账龄披露，无日期"):
    """A、B 两遍识别都能跑通；A 的期间是 2024 年、B 是 2025 年，用来触发标题检查。"""
    def 假提取(word, excel, 带连接word, 上下文):
        wb = Workbook(); wb.active.title = "sheet1"
        wb.active["A1"] = 9 if "新副本" in str(word) else 100
        wb.defined_names["TA_001"] = DefinedName("TA_001", attr_text="sheet1!$A$1:$B$1")
        wb.active["B1"] = "单位：人民币元"; wb.save(excel)
        Path(带连接word).write_text("连接体", encoding="utf-8")
        Path(上下文).write_text(json.dumps(
            {"tables": [{"range_name": "TA_001", "chapter_context": 上下文段}]},
            ensure_ascii=False), encoding="utf-8")
        return {"ok": True, "table_count": 1, "warnings": []}

    def 假模型(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,1],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,1],"科目":"货币资金","表名":"余额构成",'
                    '"表组":"G1","续表序":1}]}')
        日期 = "2025-12-31" if "word附注" in 材料 else "2024-12-31"
        return ('{"列":[{"列":"A","路径":"期末余额","角色":"数据列","期间角色":"本期期末",'
                '"期间实际":"%s"},{"列":"B","路径":"单位","角色":"说明列"}],'
                '"行":[{"行":1,"名字":"库存现金","角色":"数据","序号":null,'
                '"槽位":{"A":"C-N006-T001-S002"}}],'
                '"作用范围":{"币种":"人民币","单位":"元","单位依据":[{"来源":"表内","位置":"B1","原文":"单位：人民币元","片段":"人民币元","单位":"元","币种":"人民币","列":["A"]}]},"空名单":[],"例外":[]}' % 日期)

    def 假更新(word, excel, output, operations, caption_updates):
        Path(output).write_text("新Word", encoding="utf-8")
        return {"ok": True, "written_cells": 1, "errors": []}

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.识别.调用识别", 假模型)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假提取)







def test_回读差异不算完成(tmp_path, monkeypatch):
    _装stub(monkeypatch, "本表期末指2024年12月31日")
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    monkeypatch.setattr("核心.流程.回读验证",
                    lambda *a, **k: [{"格": "sheet1!A1", "期望": 9, "实际": 100}])
    状态 = 运行任务({"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))},
                tmp_path / "任务R7c", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    assert 状态["步骤"] == "回读不一致"           # 有差异不许标完成
    assert 状态["回读差异"] == 1
    差异 = (tmp_path / "任务R7c" / "回读记录.jsonl").read_text(encoding="utf-8").strip()
    assert "sheet1!A1" in 差异                    # 差异文件保留，不删证据




def test_人工改过的产物不被覆盖(tmp_path, monkeypatch):
    _装stub(monkeypatch, "本表期末指2024年12月31日")
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    运行目录 = tmp_path / "任务R7d"
    输入 = {"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))}
    运行任务(输入, 运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})

    产物 = 运行目录 / "更新后附注表格.xlsx"
    改后 = load_workbook(产物)
    try:
        改后["sheet1"]["A1"] = 999
        改后.save(产物)
    finally:
        改后.close()

    运行任务(输入, 运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    assert _读值(产物) == 999                          # 人工成果没被覆盖
    assert list(运行目录.glob("更新后附注表格_人工改过_*.xlsx"))  # 另存了新文件






# —— 独立复核第三轮：中途失败、续跑核验、按表失败兜底 ——

def test_中途失败后人工改Excel再续跑不覆盖(tmp_path, monkeypatch):
    _装stub(monkeypatch, "本表期末指2024年12月31日")

    def 先失败(word, excel, output, operations, caption_updates):
        raise ValueError("桥接中途失败")
    monkeypatch.setattr("核心.写回.调用桥接更新", 先失败)

    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    运行目录 = tmp_path / "任务失败后续跑"
    输入 = {"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))}
    try:
        运行任务(输入, 运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    except Exception:
        pass

    产物 = 运行目录 / "更新后附注表格.xlsx"
    assert 产物.is_file()                                  # Excel 已经生成、Word 没写成
    改后 = load_workbook(产物)
    try:
        改后["sheet1"]["A1"] = "人工修改999"
        改后.save(产物)
    finally:
        改后.close()

    def 再成功(word, excel, output, operations, caption_updates):
        Path(output).write_text("新Word", encoding="utf-8")
        return {"ok": True}
    monkeypatch.setattr("核心.写回.调用桥接更新", 再成功)

    运行任务(输入, 运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    assert _读值(产物) == "人工修改999"                    # 人工成果没被覆盖
    assert list(运行目录.glob("更新后附注表格_人工改过_*.xlsx"))   # 另存了新文件


def test_续跑换了B表就拒绝(tmp_path, monkeypatch):
    _装stub(monkeypatch, "本表期末指2024年12月31日")
    b1 = tmp_path / "b1.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"
    wb.save(b1)
    b2 = tmp_path / "b2.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 7
    wb.active["B1"] = "单位：人民币元"
    wb.save(b2)
    运行目录 = tmp_path / "任务换输入"
    运行任务({"word": "原.docx", "b表": str(b1), "金标准": str(_造金标准(tmp_path))},
            运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    try:
        运行任务({"word": "原.docx", "b表": str(b2), "金标准": str(_造金标准(tmp_path))},
                运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
        assert False, "换了 B 表还沿用旧 A，必须拒绝"
    except RuntimeError as e:
        assert "输入与原任务不一致" in str(e)


def test_一张表写失败就恢复原值再交付(tmp_path, monkeypatch):
    _装stub(monkeypatch, "本表期末指2024年12月31日")
    次数 = {"n": 0}

    def 一次失败(word, excel, output, operations, caption_updates):
        次数["n"] += 1
        if 次数["n"] == 1:
            raise ValueError("所属表前标题原文不存在或不唯一：TA_001")
        Path(output).write_text("新Word", encoding="utf-8")
        return {"ok": True}
    monkeypatch.setattr("核心.写回.调用桥接更新", 一次失败)

    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    运行目录 = tmp_path / "任务按表失败"
    状态 = 运行任务({"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))},
                运行目录, {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    assert 次数["n"] == 2                                  # 撤回失败表后重试了一次
    assert 状态["步骤"] == "部分保留" and 状态["撤回表"] == ["TA_001"]
    assert _读值(运行目录 / "更新后附注表格.xlsx") == 100    # 该表恢复原值，不写半截
    问题 = (运行目录 / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "已保持原值" in 问题
