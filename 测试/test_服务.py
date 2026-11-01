# coding: utf-8
"""Web 服务测试：首页可达、设置往返不回传密钥、任务真跑到完成、待答复能看见。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import json
import time
from openpyxl import Workbook
from fastapi.testclient import TestClient
from 服务 import web as 服务模块
from 服务.web import 应用


def _重置(配置路径=None, 运行根=None):
    应用.state.状态 = {"线程": None, "运行目录": None, "日志": []}
    应用.state.配置路径 = 配置路径 or "不存在的配置.json"


def _样本(tmp_path):
    b路径 = tmp_path / "b.xlsx"
    wb = Workbook(); wb.active.title = "word附注"; wb.active["A1"] = 9
    wb.save(b路径)
    金路径 = tmp_path / "金标准.jsonl"
    金路径.write_text(json.dumps({"note": {"name": "货币资金", "display_order": 1},
                     "table": {"name": "余额构成", "display_order": 1},
                     "slot": {"name": "库存现金 / 期末余额"}, "id": "C-N006-T001-S002",
                     "row_path": ["库存现金"], "column_path": ["期末余额"]},
                    ensure_ascii=False) + "\n", encoding="utf-8")
    return b路径, 金路径


def _打桩(monkeypatch):
    def 假提取(word, excel, 带连接word, 上下文):
        wb = Workbook(); wb.active.title = "sheet1"; wb.active["A1"] = 100
        wb.save(excel)
        pathlib.Path(带连接word).write_text("连接体", encoding="utf-8")
        pathlib.Path(上下文).write_text(json.dumps({"tables": []}), encoding="utf-8")
        return {"ok": True, "table_count": 1, "warnings": []}

    def 假模型(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,1],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return '{"表":[{"行范围":[1,1],"科目":"货币资金","表名":"余额构成"}]}'
        return ('{"列":[{"列":"A","路径":"期末余额","角色":"数据列","期间":"期末"}],'
                '"行":[{"行":1,"名字":"库存现金","角色":"数据","槽位":{"A":"C-N006-T001-S002"}}],'
                '"作用范围":{"币种":"人民币","单位":"元"},"例外":[]}')

    def 假更新(word, excel, output, operations, caption_updates):
        pathlib.Path(output).write_text("新Word", encoding="utf-8")
        return {"ok": True, "written_cells": 1, "errors": []}

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.识别.调用识别", 假模型)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)


def _等(c, 步骤=None, 次数=200):
    最后 = {}
    for _ in range(次数):
        最后 = c.get("/api/进度").json()
        if 步骤 is None or 最后.get("步骤") == 步骤:
            return 最后
        time.sleep(0.05)
    return 最后


def test_页面与设置往返(tmp_path):
    _重置(tmp_path / "模型配置.json")
    c = TestClient(应用)
    assert c.get("/").status_code == 200
    r = c.post("/api/设置", json={"url": "https://u", "api_key": "k", "model": "m", "timeout": 60})
    assert r.status_code == 200 and r.json()["ok"] is True
    cfg = c.get("/api/设置").json()
    assert cfg["model"] == "m" and "api_key" not in cfg   # 不回传密钥明文


def test_任务缺字段返回400(tmp_path):
    _重置(tmp_path / "配置.json")
    c = TestClient(应用)
    r = c.post("/api/任务", json={"word": "只给一个"})
    assert r.status_code == 400


def test_任务真跑到完成并能取结果(tmp_path, monkeypatch):
    配置路径 = tmp_path / "配置.json"
    配置路径.write_text("{}", encoding="utf-8")      # 有配置文件才会去读模型配置
    _重置(配置路径, tmp_path / "运行")
    服务模块.运行根 = tmp_path / "运行"
    monkeypatch.setattr(服务模块, "读取模型配置",
                    lambda p: {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    _打桩(monkeypatch)
    b路径, 金路径 = _样本(tmp_path)
    c = TestClient(应用)
    r = c.post("/api/任务", json={"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)})
    assert r.status_code == 200 and pathlib.Path(r.json()["运行目录"]).is_dir()
    p = _等(c, "done")
    assert p["步骤"] == "done"
    assert p["待答复"] == []                         # 跑完就没有待答复
    结果 = c.get("/api/结果").json()
    assert len(结果["对应"]) == 1 and 结果["对应"][0]["状态"] == "对应"


def test_待答复在进度里看得见(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json", tmp_path / "运行")
    服务模块.运行根 = tmp_path / "运行"
    monkeypatch.setattr(服务模块, "读取模型配置", lambda p: {})     # 没配模型 → 只组装材料
    _打桩(monkeypatch)
    b路径, 金路径 = _样本(tmp_path)
    c = TestClient(应用)
    c.post("/api/任务", json={"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)})
    p = _等(c, "identify_a")
    assert p["步骤"] == "identify_a"
    assert p["待答复"] and pathlib.Path(p["待答复"][0], "材料.txt").is_file()   # 卡在哪份材料要说得出

