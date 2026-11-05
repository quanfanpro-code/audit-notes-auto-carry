# coding: utf-8
import shutil
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook

from .提取 import 提取任务
from .表格 import file_hash


def 列出工作表(路径):
    try:
        wb = load_workbook(路径, read_only=True, data_only=False)
    except Exception as e:
        raise ValueError(f"不是可读取的xlsx工作簿：{路径}；{e}") from e
    try:
        return [{"名称": ws.title, "可见": ws.sheet_state == "visible"} for ws in wb.worksheets]
    finally:
        wb.close()


def _核验Word(路径):
    p = Path(路径)
    if not p.exists():
        return
    try:
        with ZipFile(p) as z:
            if "word/document.xml" not in z.namelist():
                raise ValueError("缺少word/document.xml")
    except (BadZipFile, OSError, ValueError) as e:
        raise ValueError(f"不是可读取的docx文档：{p}；{e}") from e


def _规范化一侧(侧, 描述, 旧调用):
    if not isinstance(描述, dict) or not 描述.get("路径"):
        raise ValueError(f"缺少{侧}端文件路径")
    路径 = str(描述["路径"])
    if not 旧调用 and not Path(路径).is_absolute():
        raise ValueError(f"{侧}端须使用文件选择窗口提供的绝对路径")
    后缀 = Path(路径).suffix.lower()
    类型 = {".docx": "word", ".xlsx": "excel"}.get(后缀)
    if 类型 is None:
        raise ValueError(f"{侧}端只支持docx和xlsx：{路径}")
    工作表 = 描述.get("工作表") or []
    if isinstance(工作表, str):
        工作表 = [工作表]
    if not isinstance(工作表, (list, tuple)) or any(not isinstance(x, str) or not x for x in 工作表):
        raise ValueError(f"{侧}端工作表选择无效")
    工作表 = list(工作表)
    if len(set(工作表)) != len(工作表):
        raise ValueError(f"{侧}端工作表重复选择")
    if 类型 == "word":
        if 工作表:
            raise ValueError("Word不需要选择工作表")
        if not 旧调用:
            _核验Word(路径)
    else:
        可省略 = 旧调用 and 侧 == "B" and not 工作表
        if not 工作表 and not 可省略:
            raise ValueError(f"{侧}端Excel至少选择一张工作表")
        实际 = {x["名称"]: x["可见"] for x in 列出工作表(路径)}
        for 名 in 工作表:
            if 名 not in 实际:
                raise ValueError(f"{侧}端工作表不存在：{名}")
            if not 实际[名]:
                raise ValueError(f"{侧}端不能选择隐藏工作表：{名}")
    return {"路径": 路径, "类型": 类型, "工作表": 工作表}


def 规范化输入(输入: dict) -> dict:
    if not isinstance(输入, dict):
        raise ValueError("任务输入必须是对象")
    新 = "A" in 输入 or "B" in 输入
    旧 = any(k in 输入 for k in ("word", "b表", "b工作表"))
    if 新 and 旧:
        raise ValueError("新旧A/B参数不能同时提供")
    if not 新 and not 旧:
        raise ValueError("缺少A端和B端文件")
    if 新 and not all(输入.get(k) for k in ("A", "B")):
        raise ValueError("新输入须同时提供A和B")
    出 = {k: v for k, v in 输入.items() if k not in ("word", "b表", "b工作表", "A", "B")}
    if 新:
        甲, 乙 = 输入["A"], 输入["B"]
    else:
        甲 = {"路径": 输入.get("word"), "工作表": []}
        乙 = {"路径": 输入.get("b表"), "工作表": 输入.get("b工作表") or []}
    出["A"] = _规范化一侧("A", 甲, 旧)
    出["B"] = _规范化一侧("B", 乙, 旧)
    return 出


def _准备Word(描述, 目录):
    目录.mkdir(parents=True, exist_ok=True)
    前 = file_hash(描述["路径"]) if Path(描述["路径"]).is_file() else ""
    产物 = 提取任务(描述["路径"], 目录)
    if 前 and file_hash(描述["路径"]) != 前:
        raise RuntimeError("Word原件在提取期间发生变化")
    上下文路径 = Path(产物["上下文"])
    import json
    上下文 = json.loads(上下文路径.read_text(encoding="utf-8-sig"))
    return [{"标识": "P001", "类型": "word", "材料": Path(产物["表A"]),
             "工作表": None, "上下文": 上下文.get("tables", []),
             "上下文路径": 上下文路径, "带连接Word": Path(产物["带连接Word"]), "目录": 目录}]


def _准备Excel(描述, 目录):
    目录.mkdir(parents=True, exist_ok=True)
    原件 = Path(描述["路径"])
    前 = file_hash(原件)
    副本 = 目录 / "输入副本.xlsx"
    if 副本.exists():
        if file_hash(副本) != 前:
            raise FileExistsError(f"输入副本已被改动，拒绝覆盖：{副本}")
    else:
        shutil.copy2(原件, 副本)
    if file_hash(原件) != 前 or file_hash(副本) != 前:
        raise RuntimeError("Excel原件在复制期间发生变化")
    工作表 = 描述["工作表"]
    if not 工作表:
        工作表 = [x["名称"] for x in 列出工作表(副本) if x["可见"]]
        if len(工作表) != 1:
            raise ValueError("旧输入省略工作表时，Excel必须只有一张可见工作表")
    出 = []
    for 序, 名 in enumerate(工作表, 1):
        部分目录 = 目录 / f"P{序:03d}"
        部分目录.mkdir(parents=True, exist_ok=True)
        出.append({"标识": f"P{序:03d}", "类型": "excel", "材料": 副本,
                   "工作表": 名, "上下文": [], "上下文路径": None,
                   "带连接Word": None, "目录": 部分目录})
    return 出


def 准备一侧(侧: str, 描述: dict, 目录) -> list[dict]:
    目录 = Path(目录)
    if 描述.get("类型") == "word":
        return _准备Word(描述, 目录 / "P001")
    if 描述.get("类型") == "excel":
        return _准备Excel(描述, 目录)
    raise ValueError(f"{侧}端文件类型未规范化")
