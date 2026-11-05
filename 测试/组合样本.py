# coding: utf-8
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from openpyxl import Workbook, load_workbook

样本 = Path(__file__).resolve().parent / "样本" / "01_普通表格.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def 文件哈希(路径):
    return hashlib.sha256(Path(路径).read_bytes()).hexdigest()


def 造Word(路径, 期初=100, 期末=110, 名单=("甲公司", "乙公司"), 标题=None, 单位="元",
         表内单位=False, 表前单位=True):
    路径 = Path(路径)
    with ZipFile(样本) as z:
        树 = E.fromstring(z.read("word/document.xml"))
    body = 树.find(W + "body")
    for 子 in list(body):
        body.remove(子)

    def 段落(文本):
        p = E.Element(W + "p")
        r = E.SubElement(p, W + "r")
        E.SubElement(r, W + "t").text = str(文本)
        return p

    if 表前单位:
        body.append(段落(f"金额单位：人民币{单位}"))
    if 标题:
        body.append(段落(标题))
    表头 = f"项目（单位：人民币{单位}）" if 表内单位 else "项目"
    行们 = [[表头, "2025-01-01", "2025-12-31"]] + [
        [名, 期初 + i * 100, 期末 + i * 100] for i, 名 in enumerate(名单)]
    tbl = E.Element(W + "tbl")
    pr = E.SubElement(tbl, W + "tblPr")
    w = E.SubElement(pr, W + "tblW")
    w.set(W + "w", "0")
    w.set(W + "type", "auto")
    grid = E.SubElement(tbl, W + "tblGrid")
    for _ in range(3):
        E.SubElement(grid, W + "gridCol").set(W + "w", "2400")
    for 行 in 行们:
        tr = E.SubElement(tbl, W + "tr")
        for 值 in 行:
            tc = E.SubElement(tr, W + "tc")
            tcPr = E.SubElement(tc, W + "tcPr")
            E.SubElement(tcPr, W + "tcW").set(W + "w", "2400")
            p = E.SubElement(tc, W + "p")
            r = E.SubElement(p, W + "r")
            E.SubElement(r, W + "t").text = str(值)
    body.append(tbl)
    with ZipFile(样本) as zin, ZipFile(路径, "w") as zout:
        for part in zin.infolist():
            内容 = (E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
                  if part.filename == "word/document.xml" else zin.read(part.filename))
            zout.writestr(part, 内容)
    return 路径


def 造Excel(路径, 表数据=None, 隐藏=(), 起列=1):
    路径 = Path(路径)
    wb = Workbook()
    wb.remove(wb.active)
    for 名, 行们 in (表数据 or {"余额": [["项目", "2025-01-01", "2025-12-31", "单位：元"],
                                      ["甲公司", 901, 102, ""],
                                      ["乙公司", 1001, 202, ""]]}).items():
        ws = wb.create_sheet(名)
        for 行号, 行 in enumerate(行们, 1):
            for 列号, 值 in enumerate(行, 起列):
                ws.cell(行号, 列号, 值)
        if 名 in 隐藏:
            ws.sheet_state = "hidden"
    wb.save(路径)
    wb.close()
    return 路径


def 造金标准(路径, 科目="货币资金", 表名="余额构成", 编号="N001"):
    行 = {"id": 编号, "note": {"name": 科目}, "table": {"name": 表名},
         "slot": {"name": "余额"}, "row_path": ["金额"], "column_path": ["余额"],
         "dimensions": [{"name": "period", "required": True}]}
    Path(路径).write_text(json.dumps(行, ensure_ascii=False) + "\n", encoding="utf-8")
    return Path(路径)


def 读Excel格(输出, 工作表, 坐标, 公式=False):
    wb = load_workbook(输出, data_only=not 公式)
    try:
        return wb[工作表][坐标].value
    finally:
        wb.close()


def _Word正文(路径):
    with ZipFile(路径) as z:
        return E.fromstring(z.read("word/document.xml")).find(W + "body")


def 读Word格(输出, 表序, 行序, 列序):
    表们 = _Word正文(输出).findall(W + "tbl")
    行们 = 表们[表序].findall(W + "tr")
    格们 = 行们[行序].findall(W + "tc")
    return "".join(格们[列序].itertext())


def 读Word段落(输出):
    return ["".join(p.itertext()).strip() for p in _Word正文(输出).findall(W + "p")]


def 造三组Excel(路径):
    return 造Excel(路径, {
        "资金": [["项目（单位：人民币元）", "期初余额", "期末余额"],
                ["甲公司", 901, 910], ["乙公司", 902, 911]],
        "往来": [["项目（单位：人民币元）", "期初余额", "期末余额"],
                ["甲公司", 701, 710], ["乙公司", 702, 711]],
        "其他": [["项目（单位：人民币元）", "期初余额", "期末余额"],
                ["甲公司", 801, 810], ["乙公司", 802, 811]],
    })


def 三组多表模型(同名=False):
    from 测试.单位样本 import 单位作用

    科目们 = ("货币资金", "货币资金" if 同名 else "应收账款", "存货")
    表名们 = ("余额构成", "余额构成" if 同名 else "按账龄披露", "存货明细")
    部分科目 = {"P001": 0, "P002": 1, "P003": 2}

    def 调用(配置, 指令, 材料, 目录):
        from openpyxl.utils import get_column_letter
        目录 = Path(目录)
        侧 = 目录.name[0]
        if 侧 == "A":
            if 指令.startswith("把整份文件"):
                return json.dumps({"划分": [{"行范围": 段, "科目": 科} for 段, 科 in
                    zip([[1, 6], [7, 12], [13, 16]], 科目们)]}, ensure_ascii=False)
            if 指令.startswith("在第一遍"):
                return json.dumps({"表": [{"行范围": [顶, 顶 + 2], "科目": 科, "表名": 表,
                    "表组": f"G{i}", "续表序": 1} for i, (顶, 科, 表) in enumerate(
                    zip((2, 8, 14), 科目们, 表名们), 1)]}, ensure_ascii=False)
            序 = int(目录.name.rsplit("表", 1)[1]) - 1
            顶, 列号 = (2, 8, 14)[序], 2
        else:
            序 = 部分科目[next(x for x in 目录.parts if x in 部分科目)]
            if 指令.startswith("把整份文件"):
                return json.dumps({"划分": [{"行范围": [1, 3], "科目": 科目们[序]}]}, ensure_ascii=False)
            if 指令.startswith("在第一遍"):
                return json.dumps({"表": [{"行范围": [1, 3], "科目": 科目们[序],
                    "表名": 表名们[序], "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
            顶, 列号 = 1, 1
        列字 = [get_column_letter(列号 + i) for i in range(3)]
        编号 = "N002B" if 同名 and 序 == 1 else f"N{序 + 1:03d}"
        列 = [{"列": 列字[0], "路径": "项目", "角色": "行名列"},
              {"列": 列字[1], "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": f"{列字[1]}{顶}"}},
              {"列": 列字[2], "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": f"{列字[2]}{顶}"}}]
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
            {"行": 顶 + i, "名字": 名, "角色": "数据", "序号": i,
             "槽位": {列字[1]: 编号, 列字[2]: 编号}}
            for i, 名 in enumerate(("甲公司", "乙公司"), 1)]
        return json.dumps({"列": 列, "行": 行, "作用范围": 单位作用(
            f"{列字[0]}{顶}", "项目（单位：人民币元）", 列字[1:]), "空名单": [], "例外": []},
            ensure_ascii=False)
    return 调用


def 跑场景(输入, 任务, 模型, monkeypatch):
    from 核心.流程 import 运行任务

    monkeypatch.setattr("核心.识别.调用识别", 模型)
    状态 = 运行任务(输入, 任务, {"测试": True})
    调用 = []
    for 侧 in ("A", "B"):
        路径 = Path(任务) / f"复用记录_{侧}.json"
        if not 路径.is_file():
            continue
        记录 = json.loads(路径.read_text(encoding="utf-8"))
        for 部分 in 记录.get("部分") or []:
            调用.append({"侧": 侧, "读取部分": 部分.get("读取部分"),
                       "工作表": 部分.get("工作表"), **(部分.get("调用") or {})})
    产物 = 状态.get("产物") or {}
    输出 = 产物.get("A端新副本") or 产物.get("Word新副本") or 产物.get("Excel新副本")
    return {"任务": Path(任务), "状态": 状态, "输出": Path(输出) if 输出 else None,
            "调用": 调用}


def 双Word模型():
    def 调用(配置, 指令, 材料, 目录):
        侧 = Path(目录).name[0]
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[2,4],"科目":"货币资金"}]}'
        if 指令.startswith("在第一遍"):
            return '{"表":[{"行范围":[2,4],"科目":"货币资金","表名":"余额构成","表组":"G1","续表序":1}]}'
        名单 = ("客户甲", "客户乙") if 侧 == "B" else ("甲公司", "乙公司")
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C2"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D2"}}]
        行 = [{"行": 2, "名字": "项目", "角色": "表头"}] + [
            {"行": 2 + i, "名字": 名, "角色": "数据", "序号": i,
             "槽位": {"C": "N001", "D": "N001"}} for i, 名 in enumerate(名单, 1)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "上下文", "位置": ["unit_text_candidates", 0],
             "原文": "金额单位：人民币元", "片段": "元", "币种": "人民币", "单位": "元",
             "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def 换端Word模型():
    def 调用(配置, 指令, 材料, 目录):
        三项 = "丙公司" in 材料 or "客户丙" in 材料
        底 = 5 if 三项 else 4
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [1, 1], "科目": "非表格"},
                {"行范围": [2, 底], "科目": "货币资金"}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [2, 底], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        if "客户" in 材料:
            名单 = ["客户甲", "客户乙", "客户丙"][:底 - 2]
        else:
            名单 = ["甲公司", "乙公司", "丙公司"][:底 - 2]
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C2"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D2"}}]
        行 = [{"行": 2, "名字": "项目", "角色": "表头"}] + [
            {"行": 2 + i, "名字": 名, "角色": "数据", "序号": i,
             "槽位": {"C": "N001", "D": "N001"}} for i, 名 in enumerate(名单, 1)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "上下文", "位置": ["unit_text_candidates", 0],
             "原文": "金额单位：人民币元", "片段": "元", "币种": "人民币", "单位": "元",
             "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def 跨格式模型():
    def 调用(配置, 指令, 材料, 目录):
        侧 = Path(目录).name[0]
        实际 = 材料.split("【二、本表身份】", 1)[0] if "【二、本表身份】" in 材料 else 材料
        顶 = 2 if ("以下为原文件二维网格" in 实际
                or "金额单位：人民币元" in 实际) else 1
        底 = 顶 + 2
        划分 = []
        if 顶 == 2:
            划分.append({"行范围": [1, 1], "科目": "非表格"})
        划分.append({"行范围": [顶, 底], "科目": "货币资金"})
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": 划分}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [顶, 底], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": f"C{顶}"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": f"D{顶}"}}]
        行 = [{"行": 顶, "名字": "项目", "角色": "表头"}] + [
            {"行": 顶 + i, "名字": f"名单{i}", "角色": "数据", "序号": i,
             "槽位": {"C": "N001", "D": "N001"}} for i in (1, 2)]
        if "金额单位：人民币元" in 实际:
            依据 = [{"来源": "上下文", "位置": ["unit_text_candidates", 0],
                    "原文": "金额单位：人民币元", "片段": "元", "币种": "人民币", "单位": "元",
                    "列": ["C", "D"]}]
        elif "项目（单位：人民币元）" in 实际:
            依据 = [{"来源": "表内", "位置": f"B{顶}", "原文": "项目（单位：人民币元）",
                    "片段": "元", "币种": "人民币", "单位": "元", "列": ["C", "D"]}]
        else:
            依据 = []
        作用 = {"币种": "人民币" if 依据 else "未知", "单位": "元" if 依据 else "未知",
                "期间跨度": "时点", "单位依据": 依据}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def 造续表Word(路径):
    import copy

    路径 = 造Word(路径, 901, 102, 名单=("客户甲", "客户乙"), 表内单位=True)
    with ZipFile(路径) as z:
        树 = E.fromstring(z.read("word/document.xml"))
    body = 树.find(W + "body")
    表 = body.find(W + "tbl")
    行们 = 表.findall(W + "tr")
    续 = copy.deepcopy(表)
    表.remove(行们[2])
    续.remove(续.findall(W + "tr")[1])
    body.insert(list(body).index(表) + 1, 续)
    with ZipFile(路径) as zin:
        内容们 = {p.filename: (E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
                            if p.filename == "word/document.xml" else zin.read(p.filename))
                 for p in zin.infolist()}
    with ZipFile(路径, "w") as zout:
        for 名, 内容 in 内容们.items():
            zout.writestr(名, 内容)
    return 路径
