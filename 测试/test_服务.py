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
    wb.active["B1"] = "单位：人民币元"; wb.save(b路径)
    金路径 = tmp_path / "金标准.jsonl"
    金路径.write_text(json.dumps({"note": {"name": "货币资金", "display_order": 1},
                     "table": {"name": "余额构成", "display_order": 1},
                     "slot": {"name": "库存现金 / 期末余额"}, "id": "C-N006-T001-S002",
                     "row_path": ["库存现金"], "column_path": ["期末余额"]},
                    ensure_ascii=False) + "\n", encoding="utf-8")
    return b路径, 金路径




def _打桩(monkeypatch):
    def 假提取(word, excel, 带连接word, 上下文):
        wb = Workbook(); wb.active.title = "sheet1"
        wb.active["A1"] = 9 if "新副本" in str(word) else 100
        wb.active["B1"] = "单位：人民币元"; wb.save(excel)
        pathlib.Path(带连接word).write_text("连接体", encoding="utf-8")
        pathlib.Path(上下文).write_text(json.dumps({"tables": []}), encoding="utf-8")
        return {"ok": True, "table_count": 1, "warnings": []}

    def 假模型(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,1],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,1],"科目":"货币资金","表名":"余额构成",'
                    '"表组":"G1","续表序":1}]}')
        return ('{"列":[{"列":"A","路径":"期末余额","角色":"数据列","期间角色":"本期期末"},{"列":"B","路径":"单位","角色":"说明列"}],'
                '"行":[{"行":1,"名字":"库存现金","角色":"数据","序号":null,"槽位":{"A":"C-N006-T001-S002"}}],'
                '"作用范围":{"币种":"人民币","单位":"元","单位依据":[{"来源":"表内","位置":"B1","原文":"单位：人民币元","片段":"人民币元","单位":"元","币种":"人民币","列":["A"]}]},"空名单":[],"例外":[]}')

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
    from 核心.设置 import 保存模型配置
    保存模型配置(tmp_path / "配置.json", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
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
    for _ in range(40):                    # 待答复与步骤同一次落盘；读端可能撞上替换瞬间，给它留出重读
        if p.get("待答复"):
            break
        time.sleep(0.05)
        p = c.get("/api/进度").json()
    assert p["步骤"] == "identify_a"
    assert p["待答复"] and pathlib.Path(p["待答复"][0], "材料.txt").is_file()   # 卡在哪份材料要说得出




# —— R8：续跑沿用原目录、人工回写字段、401 不算连通、端口顺延 ——



def test_续跑沿用原目录(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json")
    服务模块.运行根 = tmp_path / "运行"
    from 核心.设置 import 保存模型配置
    保存模型配置(tmp_path / "配置.json", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    _打桩(monkeypatch)
    b路径, 金路径 = _样本(tmp_path)
    c = TestClient(应用)
    入参 = {"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)}
    原读取 = 服务模块.读取模型配置
    monkeypatch.setattr(服务模块, "读取模型配置", lambda p: {})
    第一次 = c.post("/api/任务", json=入参).json()["运行目录"]
    线程 = 应用.state.状态["线程"]
    线程.join(10)
    assert not 线程.is_alive()
    等待 = c.get("/api/进度").json()
    assert 等待["步骤"] == "identify_a" and 等待["待答复"]
    monkeypatch.setattr(服务模块, "读取模型配置", 原读取)
    续跑 = c.post("/api/任务", json={**入参, "续跑": True})
    assert 续跑.status_code == 200
    assert 续跑.json()["运行目录"] == 第一次       # 等待答复的任务继续用原目录
    assert _等(c, "done")["步骤"] == "done"
    assert c.post("/api/任务", json={**入参, "续跑": True}).status_code == 409




def test_人工回写认输出字段(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json")
    monkeypatch.setattr(服务模块, "运行根", tmp_path / "运行")
    收到 = {}

    def 假更新(word, excel, output, operations, caption_updates):
        收到["output"] = output
        pathlib.Path(output).write_bytes(b"argument-test-placeholder")
        return {"ok": True, "written_cells": 1, "errors": []}

    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    c = TestClient(应用)
    输出 = str(tmp_path / "新.docx")
    r = c.post("/api/人工回写", json={"带连接word": "w.docx", "excel": "e.xlsx", "输出": 输出})
    assert r.json()["ok"] is True and 收到["output"] == 输出




def test_设置测试401不算连通(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json")
    c = TestClient(应用)
    c.post("/api/设置", json={"url": "https://u", "api_key": "k", "model": "m", "timeout": 60})

    class 假响应:
        status_code = 401

    monkeypatch.setattr("服务.web.httpx.get", lambda *a, **k: 假响应())
    r = c.post("/api/设置/测试", json={}).json()
    assert r["ok"] is False and "密钥" in r["说明"]     # 401 必须如实报密钥问题




def test_端口被占自动顺延():
    import socket
    from 服务.启动 import 选端口
    占位 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    占位.bind(("127.0.0.1", 0))
    占位.listen(1)
    端口 = 占位.getsockname()[1]
    try:
        选 = 选端口(端口, 5)
        assert 选 != 端口 and 端口 < 选 < 端口 + 5
    finally:
        占位.close()




def test_结果页给出摘要与产物(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json")
    服务模块.运行根 = tmp_path / "运行"
    from 核心.设置 import 保存模型配置
    保存模型配置(tmp_path / "配置.json", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    _打桩(monkeypatch)
    b路径, 金路径 = _样本(tmp_path)
    c = TestClient(应用)
    c.post("/api/任务", json={"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)})
    _等(c, "done")
    r = c.get("/api/结果").json()
    assert r["摘要"]["对应条数"] >= 1
    assert r["产物"]["运行目录"] and r["产物"]["更新后附注表格"]




# —— R8：续跑沿用原目录、人工回写字段、401 不算连通、端口顺延 ——


# —— 独立复核第三轮：无控制台日志、桥接提示 ——

def test_无控制台也能完成启动日志配置(tmp_path, monkeypatch):
    import uvicorn
    from 服务 import 启动
    monkeypatch.setattr(启动, "日志路径", tmp_path / "启动日志.txt")
    monkeypatch.setattr(sys, "stdout", None)
    monkeypatch.setattr(sys, "stderr", None)
    配置 = uvicorn.Config("服务.web:应用", host="127.0.0.1", port=8765,
                        log_level="info", log_config=启动.日志配置())
    配置.configure_logging()               # 修复前这一句会抛 ValueError
    assert 启动.日志配置()["handlers"]["文件"]["filename"].endswith("服务日志.txt")


def test_桥接提示能到结果接口(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json")
    服务模块.运行根 = tmp_path / "运行"
    from 核心.设置 import 保存模型配置
    保存模型配置(tmp_path / "配置.json", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    _打桩(monkeypatch)

    import 核心.提取 as 提取模块
    原 = 提取模块.调用桥接提取

    def 带提示(word, excel, 带连接word, 上下文):
        结果 = 原(word, excel, 带连接word, 上下文)
        return {**结果, "warnings": ["第 3 张表有合并单元格，已按左上角处理"]}
    monkeypatch.setattr("核心.提取.调用桥接提取", 带提示)

    b路径, 金路径 = _样本(tmp_path)
    c = TestClient(应用)
    c.post("/api/任务", json={"word": "原.docx", "b表": str(b路径), "金标准": str(金路径)})
    _等(c, "done")
    r = c.get("/api/结果").json()
    assert any("合并单元格" in str(x) for x in r["桥接提示"])   # 桥接的话要能到页面


def test_工作表接口返回实际名称与可见性(tmp_path):
    from 测试.组合样本 import 造Excel

    _重置(tmp_path / "配置.json")
    路径 = 造Excel(tmp_path / "多表.xlsx", {"资金": [[1]], "往来": [[2]], "隐藏": [[3]]}, 隐藏=("隐藏",))
    c = TestClient(应用)
    r = c.post("/api/工作表", json={"路径": str(路径)})
    assert r.status_code == 200
    assert r.json()["工作表"] == [{"名称": "资金", "可见": True},
                                {"名称": "往来", "可见": True}, {"名称": "隐藏", "可见": False}]
    坏 = tmp_path / "坏.xlsx"
    坏.write_text("不是工作簿", encoding="utf-8")
    assert c.post("/api/工作表", json={"路径": str(坏)}).status_code == 400


def test_新任务接口返回真实Excel和公式明细(tmp_path, monkeypatch):
    from 测试.test_Excel公式 import _造A, _造B, _模型
    from 测试.组合样本 import 读Excel格, 造金标准

    _重置(tmp_path / "配置.json", tmp_path / "运行")
    服务模块.运行根 = tmp_path / "运行"
    from 核心.设置 import 保存模型配置
    保存模型配置(tmp_path / "配置.json", {"url": "u", "api_key": "k", "model": "m", "timeout": 1})
    monkeypatch.setattr("核心.识别.调用识别", _模型())
    A, B = _造A(tmp_path / "A.xlsx"), _造B(tmp_path / "B.xlsx")
    c = TestClient(应用)
    r = c.post("/api/任务", json={"A": {"路径": str(A), "工作表": ["资金"]},
                                 "B": {"路径": str(B), "工作表": ["余额"]},
                                 "金标准": str(造金标准(tmp_path / "金标准.jsonl")), "复用": False})
    assert r.status_code == 200, r.text
    p = _等(c, "done", 次数=400)
    assert p["步骤"] == "done", p
    结果 = c.get("/api/结果").json()
    输出 = 结果["产物"]["A端新副本"]
    assert 结果["产物"]["Excel新副本"] == 输出
    assert 读Excel格(输出, "资金", "C2") == 901
    assert 结果["公式待核数"] == 0
    记录 = next(x for x in 结果["公式核对"] if x["a格"] == "资金!C4")
    assert 记录["A结果"] == 1902 and 记录["B金额"] == 1902
    assert 结果["复用统计"]


def test_后台异常写入失败状态而不是停在写回(tmp_path, monkeypatch):
    _重置(tmp_path / "配置.json", tmp_path / "运行")
    服务模块.运行根 = tmp_path / "运行"
    monkeypatch.setattr(服务模块, "运行任务", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("测试故障")))
    c = TestClient(应用)
    r = c.post("/api/任务", json={"word": "原.docx", "b表": "B.xlsx", "金标准": "金.jsonl"})
    assert r.status_code == 200
    p = _等(c, "任务失败", 次数=100)
    assert p["步骤"] == "任务失败"
    assert any("测试故障" in x for x in p["日志"])
