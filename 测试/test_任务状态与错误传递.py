# coding: utf-8
"""结果状态回归：固定模型答复，真实文件、桥接及Excel工作进程照常执行。"""
import copy
import json
import os
import threading
from pathlib import Path
from zipfile import ZipFile

import pytest
from fastapi.testclient import TestClient
from lxml import etree as E
from openpyxl import Workbook, load_workbook
from 服务 import web
from 核心.接续 import 读索引
from 测试.组合样本 import W, 造Word, 造Excel, 文件哈希, 读Word格, 读Word段落, 三组多表模型
from 测试.test_Excel公式 import _多表, _多表金, _多表模型, _多表输入

def _写簿(路径, 行们):
    wb = Workbook()
    ws = wb.active
    ws.title = "资金"
    for 行号, 行 in enumerate(行们, 1):
        for 列号, 值 in enumerate(行, 2):
            ws.cell(行号, 列号, 值)
    wb.save(路径)
    wb.close()


def _重叠输入(tmp_path):
    A = tmp_path / "A.xlsx"
    B = tmp_path / "B.xlsx"
    _写簿(A, [
        ["项目（单位：人民币元）", "2025-01-01", "2025-12-31"],
        ["甲公司（单位：人民币元）", 100, 110],
        ["乙公司", 200, 210],
        ["丙公司", 300, 310],
    ])
    _写簿(B, [
        ["项目（单位：人民币元）", "2025-01-01", "2025-12-31"],
        ["本次甲", 901, 102],
        ["本次一（单位：人民币元）", 1001, 202],
        ["本次二", 2001, 302],
        ["本次三", 3001, 402],
    ])
    金 = tmp_path / "金标准.jsonl"
    行们 = []
    for 编号, 表名 in (("N001", "余额构成"), ("N002", "其他构成")):
        行们.append({"id": 编号, "note": {"name": "货币资金"}, "table": {"name": 表名},
                   "slot": {"name": 编号}, "dimensions": [{"name": "period", "required": True}]})
    金.write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们), encoding="utf-8")
    输入 = {"A": {"路径": str(A), "工作表": ["资金"]},
            "B": {"路径": str(B), "工作表": ["资金"]},
            "金标准": str(金), "识别记录目录": str(tmp_path / "独立识别记录"), "复用": False}
    return A, B, 输入


def _重叠模型(配置, 指令, 材料, 目录):
    目录 = Path(目录)
    侧 = 目录.name[0]
    底 = 4 if 侧 == "A" else 5
    if 指令.startswith("把整份文件"):
        return json.dumps({"划分": [{"行范围": [1, 底], "科目": "货币资金"}]}, ensure_ascii=False)
    if 指令.startswith("在第一遍"):
        二起 = 2 if 侧 == "A" else 3
        return json.dumps({"表": [
            {"行范围": [1, 2], "科目": "货币资金", "表名": "余额构成", "表组": "G1", "续表序": 1},
            {"行范围": [二起, 底], "科目": "货币资金", "表名": "其他构成", "表组": "G2", "续表序": 1},
        ]}, ensure_ascii=False)
    第二 = 目录.name.endswith("表002")
    编号 = "N002" if 第二 else "N001"
    顶 = (2 if 侧 == "A" else 3) if 第二 else 1
    底 = (4 if 侧 == "A" else 5) if 第二 else 2
    列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
          {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
           "期间实际": {"值": "2025-01-01", "位置": None}},
          {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
           "期间实际": {"值": "2025-12-31", "位置": None}}]
    行 = [{"行": r, "名字": f"第{r}行", "角色": "数据", "序号": r - 顶 + 1,
           "槽位": {"C": 编号, "D": 编号}} for r in range(顶, 底 + 1)]
    单位格 = f"B{顶}"
    单位原文 = (("甲公司" if 侧 == "A" else "本次一") + "（单位：人民币元）") if 第二 else "项目（单位：人民币元）"
    作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
        {"来源": "表内", "位置": 单位格, "原文": 单位原文, "片段": "元",
         "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
    return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []}, ensure_ascii=False)



def 留证(名字, 值):
    (证据 / (名字 + ".json")).write_text(json.dumps(值, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

@pytest.fixture(autouse=True)
def 隔离(monkeypatch, tmp_path):
    global 证据
    证据 = Path(os.environ.get("状态修复证据目录") or tmp_path)
    证据.mkdir(parents=True, exist_ok=True)
    monkeypatch.setenv("审计附注_复用", "0")
    monkeypatch.setenv("审计附注_识别记录目录", str(tmp_path / "独立识别记录"))
    monkeypatch.setattr(web, "运行根", tmp_path / "服务任务")
    monkeypatch.setattr(web, "读取模型配置", lambda *_: {"测试": True})
    web.应用.state.状态 = {"线程": None, "运行目录": None, "日志": []}
    web.应用.state.配置路径 = str(tmp_path / "不用模型配置.json")
    yield

def 开始(客户端, 输入):
    回 = 客户端.post("/api/任务", json=输入)
    assert 回.status_code == 200, 回.text
    assert "运行目录" in 回.json(), 回.json()
    return Path(回.json()["运行目录"])

def 等完():
    线程 = web.应用.state.状态["线程"]
    线程.join(120)
    assert not 线程.is_alive(), "探针等待超时，不能作为业务结论"

def 读簿(路径):
    wb = load_workbook(路径, data_only=True)
    try:
        return {ws.title: [[ws.cell(r,c).value for c in (2,3,4)] for r in (2,3)] for ws in wb}
    finally:
        wb.close()

def Word场景(tmp_path, 标题含链接):
    原型 = 造Word(tmp_path / "原型.docx", 标题="截至2024年12月31日（资金）", 表内单位=True, 表前单位=False)
    标题们 = ["截至2024年12月31日（资金）", "截至2024年12月31日（往来）", "截至2024年12月31日（存货）"]
    A = tmp_path / "A.docx"
    with ZipFile(原型) as zin, ZipFile(A, "w") as zout:
        for 项 in zin.infolist():
            数据 = zin.read(项.filename)
            if 项.filename == "word/document.xml":
                树 = E.fromstring(数据)
                body = 树.find(W + "body")
                段, 表 = body.find(W + "p"), body.find(W + "tbl")
                for 标题 in 标题们[1:]:
                    新段 = copy.deepcopy(段)
                    新段.find(".//" + W + "t").text = 标题
                    body.append(新段)
                    body.append(copy.deepcopy(表))
                if 标题含链接:
                    run = 段.find(W + "r")
                    段.remove(run)
                    link = E.SubElement(段, W + "hyperlink")
                    link.set(W + "anchor", "_测试标题书签")
                    link.append(run)
                数据 = E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
            zout.writestr(项, 数据)
    业务 = [("资金","货币资金","余额构成",901,102),
            ("往来","应收账款","按账龄披露",701,710),
            ("其他","存货","存货明细",801,810)]
    B = 造Excel(tmp_path / "B.xlsx", {名: [
        ["项目（单位：人民币元）","2025-01-01","2025-12-31"],
        ["本次甲",初,末],["本次乙",初+1,末+1]] for 名,科,表,初,末 in 业务})
    金 = tmp_path / "金标准.jsonl"
    金.write_text("\n".join(json.dumps({"id":f"N{i:03d}","note":{"name":科},"table":{"name":表},
        "slot":{"name":"余额"},"dimensions":[{"name":"period","required":True}]},ensure_ascii=False)
        for i,(_,科,表,_,_) in enumerate(业务,1)),encoding="utf-8")
    原模型 = 三组多表模型()
    def 模型(配置, 指令, 材料, 目录):
        数据 = json.loads(原模型(配置, 指令, 材料, 目录))
        if Path(目录).name.startswith("A_") and 指令.startswith("在第一遍"):
            for 表,题 in zip(数据["表"], 标题们):
                表["标题"] = {"原文":题,"日期":[{"片段":"2024年12月31日","期间角色":"本期期末"}]}
        return json.dumps(数据,ensure_ascii=False)
    return A,B,{"A":{"路径":str(A),"工作表":[]},"B":{"路径":str(B),"工作表":["资金","往来","其他"]},
               "金标准":str(金),"复用":False},模型

@pytest.mark.parametrize("含链接",[True,False],ids=["真实Word拒绝标题","普通标题对照"])
def test_Word桥接撤回与最终状态(tmp_path, monkeypatch, 含链接):
    A,B,输入,模型=Word场景(tmp_path,含链接)
    原哈希 = [文件哈希(A),文件哈希(B)]
    monkeypatch.setattr("核心.识别.调用识别",模型)
    with TestClient(web.应用) as 客户端:
        任务=开始(客户端,输入)
        等完()
        索引=读索引(任务)
        进度=客户端.get("/api/进度").json()
        结果=客户端.get("/api/结果").json()
    输出=Path(索引.get("产物",{}).get("Word新副本",""))
    assert 输出.is_file(), 索引
    实际=[[读Word格(输出,t,1,c) for c in (0,1,2)] for t in range(3)]
    标题=读Word段落(输出)
    留证("Word桥接撤回" if 含链接 else "Word正常对照",{
        "输入":输入,"任务":任务,"输出":输出,"实际首行":实际,"实际标题":标题,"索引":索引,
        "网页进度":进度,"网页结果":结果,"原件哈希不变":原哈希==[文件哈希(A),文件哈希(B)]})
    assert 原哈希==[文件哈希(A),文件哈希(B)]
    assert 实际[1:]==[["本次甲","701","710"],["本次甲","801","810"]],实际
    if 含链接:
        assert 实际[0]==["甲公司","100","110"]
        assert "截至2024年12月31日（资金）" in 标题
        assert 索引["撤回表"]==["TA_001"],索引
        assert "标题包含域、链接或其他对象" in (任务/"问题清单.jsonl").read_text(encoding="utf-8")
        assert 索引["步骤"]==进度["步骤"]==结果["摘要"]["步骤"]=="部分保留"
    else:
        assert 实际[0]==["本次甲","901","102"]
        assert all("2025年12月31日" in x for x in 标题 if x.strip())
        assert 索引["步骤"]=="done"

@pytest.mark.parametrize("取消部分",["P001","P002"],ids=["第一张B表后停止","最后一张B表后停止"])
def test_服务取消后真实文件与续跑(tmp_path, monkeypatch, 取消部分):
    A,B=_多表(tmp_path/"A.xlsx"),_多表(tmp_path/"B.xlsx",来源=True)
    金=_多表金(tmp_path/"金标准.jsonl")
    输入=_多表输入(A,B,金)
    前=[文件哈希(A),文件哈希(B)]
    到达,放行=threading.Event(),threading.Event()
    原模型=_多表模型()
    调用=[]
    def 模型(配置,指令,材料,目录):
        目录=Path(目录)
        调用.append(str(目录))
        答复=原模型(配置,指令,材料,目录)
        if "B" in 目录.parts and 取消部分 in 目录.parts and "_表" in 目录.name:
            到达.set()
            assert 放行.wait(30), "探针没能按顺序提交停止"
        return 答复
    monkeypatch.setattr("核心.识别.调用识别",模型)
    with TestClient(web.应用) as 客户端:
        任务=开始(客户端,输入)
        assert 到达.wait(30), "没有进入预期精读点"
        中断响应=客户端.post("/api/中断",json={}).json()
        放行.set()
        等完()
        中断索引=读索引(任务)
        中断进度=客户端.get("/api/进度").json()
        停止时交付请求 = list(任务.glob("Excel请求_*.json"))
        中断原件不变=前==[文件哈希(A),文件哈希(B)]
        monkeypatch.setattr("核心.识别.调用识别",原模型)
        续跑目录=开始(客户端,{**输入,"续跑":True})
        等完()
        续跑索引=读索引(任务)
        输出=Path(续跑索引.get("产物",{}).get("Excel新副本",""))
        实际=读簿(输出) if 输出.is_file() else {}
    留证("停止_"+取消部分,{"任务":任务,"停止响应":中断响应,"停止索引":中断索引,
        "停止网页进度":中断进度,"实际调用目录":调用,"续跑任务":续跑目录,"续跑索引":续跑索引,
        "最终Excel":输出,"实际金额名单":实际,"停止时原件不变":中断原件不变,
        "续跑后原件不变":前==[文件哈希(A),文件哈希(B)]})
    assert 中断原件不变 and 前==[文件哈希(A),文件哈希(B)]
    assert 续跑目录==任务
    assert 续跑索引["步骤"]=="done",续跑索引
    assert 实际["资金"][0]==["资金甲",901,102] and 实际["往来"][0]==["往来甲",701,710]
    assert 中断索引["步骤"]=="用户中断",中断索引["步骤"]
    assert not 中断索引.get("待答复"),中断索引
    assert 中断进度["步骤"] == "用户中断"
    assert not 停止时交付请求
    if 取消部分 == "P001":
        assert not any("B" in Path(p).parts and "P002" in Path(p).parts for p in 调用)

def test_Excel真实错误要传到任务及页面(tmp_path,monkeypatch):
    A,B,输入=_重叠输入(tmp_path)
    前=[文件哈希(A),文件哈希(B)]
    monkeypatch.setattr("核心.识别.调用识别",_重叠模型)
    import 核心.流程 as 流程
    原写=流程.写索引
    状态轨迹=[]
    def 记状态(目录,索引):
        状态轨迹.append({"步骤":索引.get("步骤"),"Excel错误":索引.get("Excel错误")})
        return 原写(目录,索引)
    monkeypatch.setattr(流程,"写索引",记状态)
    with TestClient(web.应用) as 客户端:
        任务=开始(客户端,输入)
        等完()
        索引=读索引(任务)
        进度=客户端.get("/api/进度").json()
    工作结果=json.loads((任务/"Excel结果.json").read_text(encoding="utf-8"))
    候选=Path(索引["产物"]["Excel候选副本"])
    实际=读簿(候选)
    留证("Excel错误传递",{"任务":任务,"工作进程结果":工作结果,"内部状态轨迹":状态轨迹,
        "根任务索引":索引,"网页进度":进度,"候选副本实际值":实际,
        "原件哈希不变":前==[文件哈希(A),文件哈希(B)]})
    assert "拟写格不能唯一归入业务表组：资金!B2" in 工作结果["错误"]
    assert 实际["资金"][0]==["甲公司（单位：人民币元）",100,110]
    assert 前==[文件哈希(A),文件哈希(B)]
    assert any(x["步骤"]=="Excel交付失败" for x in 状态轨迹)
    assert "资金!B2" in 索引.get("错误","") and 索引["步骤"]=="Excel交付失败",索引

@pytest.mark.parametrize("有答复",[True,False],ids=["无效模型答复","没有配置等待答复"])
def test_识别失败与待答复对照(tmp_path,monkeypatch,有答复):
    A,B=_多表(tmp_path/"A.xlsx"),_多表(tmp_path/"B.xlsx",来源=True)
    输入=_多表输入(A,B,_多表金(tmp_path/"金标准.jsonl"))
    前=[文件哈希(A),文件哈希(B)]
    if 有答复:
        monkeypatch.setattr("核心.识别.调用识别",lambda *_: "这不是JSON")
    else:
        monkeypatch.setattr(web,"读取模型配置",lambda *_: {})
    with TestClient(web.应用) as 客户端:
        任务=开始(客户端,输入)
        等完()
        索引=读索引(任务)
        进度=客户端.get("/api/进度").json()
    留证("识别异常" if 有答复 else "真正待答复",{"任务":任务,"索引":索引,"网页进度":进度,
        "原件哈希不变":前==[文件哈希(A),文件哈希(B)]})
    assert 前==[文件哈希(A),文件哈希(B)]
    if 有答复:
        assert 索引.get("识别错误")
        assert 索引["步骤"]=="识别失败",索引
    else:
        assert 索引["步骤"]=="identify_a"
        assert all(Path(p,"材料.txt").is_file() for p in 索引["待答复"])


def test_续跑的新故障不能沿用旧失败类别(tmp_path, monkeypatch):
    A, B = _多表(tmp_path / "A.xlsx"), _多表(tmp_path / "B.xlsx", 来源=True)
    金 = _多表金(tmp_path / "金标准.jsonl")
    输入 = _多表输入(A, B, 金)
    monkeypatch.setattr("核心.识别.调用识别", lambda *_: "这不是JSON")
    with TestClient(web.应用) as 客户端:
        任务 = 开始(客户端, 输入)
        等完()
        assert 读索引(任务)["步骤"] == "识别失败"
        # 输入不一致现在在接受前拒绝；这里继续验证真正运行后的新故障类别。
        def 新故障(*args, **kwargs):
            raise RuntimeError("第二次运行读取材料失败")
        monkeypatch.setattr(web, "运行任务", 新故障)
        开始(客户端, {**输入, "续跑": True})
        等完()
        索引 = 读索引(任务)
        assert 索引["步骤"] == "任务失败"
        assert "第二次运行读取材料失败" in 索引["错误"]
        assert "识别错误" not in 索引


def test_真实Excel运行中停止后可继续(tmp_path, monkeypatch):
    import time
    import ctypes

    A, B = _多表(tmp_path / "A.xlsx"), _多表(tmp_path / "B.xlsx", 来源=True)
    输入 = _多表输入(A, B, _多表金(tmp_path / "金标准.jsonl"))
    前 = [文件哈希(A), 文件哈希(B)]
    monkeypatch.setattr("核心.识别.调用识别", _多表模型())
    with TestClient(web.应用) as 客户端:
        任务 = 开始(客户端, 输入)
        状态文件 = 任务 / "Excel实例.json"
        截止 = time.monotonic() + 30
        实例 = {}
        while time.monotonic() < 截止:
            try:
                实例 = json.loads(状态文件.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                pass
            if 实例.get("excel_hwnd"):
                break
            time.sleep(0.005)
        assert 实例.get("excel_hwnd"), "未观察到真实Excel启动，不能冒充验收通过"
        assert ctypes.windll.user32.IsWindow(实例["excel_hwnd"])
        客户端.post("/api/中断", json={})
        等完()
        中断索引 = 读索引(任务)
        assert 中断索引["步骤"] == "用户中断", 中断索引
        assert not ctypes.windll.user32.IsWindow(实例["excel_hwnd"])
        assert 前 == [文件哈希(A), 文件哈希(B)]
        开始(客户端, {**输入, "续跑": True})
        等完()
        索引 = 读索引(任务)
        输出 = Path(索引["产物"]["Excel新副本"])
        实际 = 读簿(输出)
        assert 索引["步骤"] == "done"
        assert 实际["资金"][0] == ["资金甲", 901, 102]
        assert 实际["往来"][0] == ["往来甲", 701, 710]
        assert 前 == [文件哈希(A), 文件哈希(B)]
    留证("真实Excel停止", {"任务": 任务, "中断索引": 中断索引, "续跑索引": 索引,
                         "最终Excel": 输出, "实际金额名单": 实际, "原件哈希不变": True})


@pytest.mark.parametrize("等待侧", ["A", "B"])
def test_等待人工答复时中断保留材料并可继续(tmp_path, monkeypatch, 等待侧):
    import 核心.识别 as 识别

    A, B = _多表(tmp_path / "A.xlsx"), _多表(tmp_path / "B.xlsx", 来源=True)
    输入 = _多表输入(A, B, _多表金(tmp_path / "金标准.jsonl"))
    原件哈希 = [文件哈希(A), 文件哈希(B)]
    原取答复, 模型 = 识别._取答复, _多表模型()

    def 等待答复(目录, 指令, 材料, 配置, 只用已有答复, **参数):
        return 原取答复(目录, 指令, 材料, 配置,
                       只用已有答复 or Path(目录).name.startswith(等待侧 + "_"), **参数)

    monkeypatch.setattr(识别, "_取答复", 等待答复)
    monkeypatch.setattr(识别, "调用识别", 模型)
    with TestClient(web.应用) as 客户端:
        assert 客户端.post("/api/中断", json={}).json() == {"ok": True}
        任务 = 开始(客户端, 输入)
        等完()
        等待索引 = 读索引(任务)
        assert 等待索引["步骤"] == ("identify_a" if 等待侧 == "A" else "identify_b")
        目录们 = [Path(p) for p in 等待索引["待答复"]]
        assert 目录们 and all((p / "材料.txt").is_file() for p in 目录们)
        原材料 = {p: p.read_bytes() for d in 目录们 for p in d.glob("*.txt")}
        assert 客户端.post("/api/中断", json={}).json() == {"ok": True}
        中断索引 = 读索引(任务)
        进度 = 客户端.get("/api/进度").json()
        摘要 = 客户端.get("/api/结果").json()["摘要"]
        assert 中断索引["步骤"] == 进度["步骤"] == 摘要["步骤"] == "用户中断"
        assert not 进度["待答复"]
        assert all(p.read_bytes() == 内容 for p, 内容 in 原材料.items())
        前索引 = (任务 / "接续索引.json").read_bytes()
        客户端.post("/api/中断", json={})
        assert (任务 / "接续索引.json").read_bytes() == 前索引

        # 把答复放进原材料目录，再继续同一任务，证明中断没有销毁等待成果。
        for d in 目录们:
            答复 = 模型({"测试": True}, (d / "指令.txt").read_text(encoding="utf-8"),
                        (d / "材料.txt").read_text(encoding="utf-8"), d)
            (d / "答复.txt").write_text(答复, encoding="utf-8")
        monkeypatch.setattr(识别, "_取答复", 原取答复)
        assert 开始(客户端, {**输入, "续跑": True}) == 任务
        等完()
        续跑索引 = 读索引(任务)
        输出 = Path(续跑索引["产物"]["Excel新副本"])
        实际 = 读簿(输出)
        assert 续跑索引["步骤"] == "done"
        assert 实际["资金"][0] == ["资金甲", 901, 102]
        assert 实际["往来"][0] == ["往来甲", 701, 710]
        assert all(p.read_bytes() == 内容 for p, 内容 in 原材料.items())
        assert 原件哈希 == [文件哈希(A), 文件哈希(B)]
        完成索引 = (任务 / "接续索引.json").read_bytes()
        完成文件哈希 = 文件哈希(输出)
        客户端.post("/api/中断", json={})
        assert (任务 / "接续索引.json").read_bytes() == 完成索引
        assert 文件哈希(输出) == 完成文件哈希
    留证("等待答复中断_" + 等待侧, {"任务": 任务, "中断索引": 中断索引,
         "网页进度": 进度, "结果摘要": 摘要, "原材料目录": 目录们,
         "续跑索引": 续跑索引, "输出": 输出, "实际名单金额": 实际,
         "原件哈希不变": True, "原材料内容不变": True})
