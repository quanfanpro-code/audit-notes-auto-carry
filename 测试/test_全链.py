# coding: utf-8
"""全链集成测试（stub 模型/桥接）：断言实际写入值、操作记录、原件哈希不变。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import hashlib, json
from pathlib import Path
from openpyxl import Workbook, load_workbook
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
    wb.save(excel)
    Path(带连接word).write_text("连接体", encoding="utf-8")
    Path(上下文).write_text(json.dumps({"tables": []}), encoding="utf-8")
    return {"ok": True, "table_count": 1, "warnings": []}


def test_全链stub(tmp_path, monkeypatch):
    # 真实小样本：B 是真实 xlsx（工作表 word附注，值为 9）；A 由假提取产出（sheet1，值为 100）
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.save(b路径); b前 = md5(b路径)
    金路径 = _造金标准(tmp_path)

    def 假提取(word, excel, 带连接word, 上下文):
        return _假提取(9 if "新副本" in str(word) else 100, word, excel, 带连接word, 上下文)

    def 假模型(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,1],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,1],"科目":"货币资金","表名":"余额构成",'
                    '"表组":"G1","续表序":1}]}')
        return ('{"列":[{"列":"A","路径":"期末余额","角色":"数据列","期间角色":"本期期末"}],'
                '"行":[{"行":1,"名字":"库存现金","角色":"数据","序号":null,"槽位":{"A":"C-N006-T001-S002"}}],'
                '"作用范围":{"币种":"人民币","单位":"元"},"空名单":[],"例外":[]}')

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
    wb.save(b路径)
    monkeypatch.setattr("核心.提取.调用桥接提取",
                    lambda word, excel, 带连接word, 上下文: _假提取(100, word, excel, 带连接word, 上下文))
    状态 = 运行任务({"word": "原.docx", "b表": str(b路径), "金标准": str(_造金标准(tmp_path))},
                tmp_path / "任务2", {})
    assert 状态["步骤"] == "identify_a"          # 停在第一遍，不硬往下走
    材料 = Path(状态["待答复"][0]) / "材料.txt"
    assert 材料.is_file() and "整份文件的表格文字" in 材料.read_text(encoding="utf-8")

