# coding: utf-8
"""从任务留证整理展示数据；不重判语义、不执行搬运、不改历史文件。"""
import hashlib
import json
import os
from collections import Counter, OrderedDict
from decimal import Decimal
from pathlib import Path

from 核心.接续 import 读索引
from .任务服务 import 任务目录, 读说明, 拒绝, 输入线索
from fastapi import HTTPException

_缓存 = OrderedDict()
_文件缓存 = OrderedDict()
_资料 = ("接续索引.json", "任务说明.json", "交付明细.jsonl", "A合集.jsonl", "B合集.jsonl",
         "结构清单.jsonl", "问题清单.jsonl", "回读记录.jsonl", "公式核对记录.jsonl",
         "复用记录_A.json", "复用记录_B.json")


def 值对象(值, 状态="值", 已取=True, 来源="任务记录"):
    if not 已取 or 状态 in (None, "读取失败", "公式无保存结果", "不可见"):
        return {"类型": 状态 or "未记录", "原样文本": "", "是否取得": False, "来源": 来源}
    类型 = ("空白" if 值 is None or 值 == "" else "布尔" if isinstance(值, bool)
          else "数字" if isinstance(值, (int, float, Decimal)) else "文字")
    return {"类型": 类型, "原样文本": "" if 值 is None else str(值), "是否取得": True, "来源": 来源}


def 位置(格):
    if not 格 or "!" not in 格:
        return {"工作表": "", "格": 格 or ""}
    表, 坐标 = 格.rsplit("!", 1)
    return {"工作表": 表, "格": 坐标}


def _标识(*项):
    return hashlib.sha256(json.dumps(项, ensure_ascii=False).encode()).hexdigest()[:24]


def _组键(目录, 条):
    return _标识(目录.name, 条.get("侧", "A"), 条.get("读取部分"), 条.get("表组"))


def _版本(目录):
    return _标识(*[(名, (p.stat().st_size, p.stat().st_mtime_ns) if p.is_file() else None)
                 for 名 in _资料 for p in [目录 / 名]])


def _读(目录, 名, 提示):
    p = 目录 / 名
    if not p.is_file():
        return [] if p.suffix == ".jsonl" else {}
    try:
        文 = p.read_text(encoding="utf-8-sig")
        return ([json.loads(x, parse_float=Decimal) for x in 文.splitlines() if x.strip()]
                if p.suffix == ".jsonl" else json.loads(文))
    except (OSError, ValueError, TypeError) as e:
        提示.append(f"{名}无法读取：{e}")
        return [] if p.suffix == ".jsonl" else {}


def 整理结果(目录):
    目录 = Path(目录)
    版本 = _版本(目录)
    键 = (str(目录), 版本)
    if 键 in _缓存:
        _缓存.move_to_end(键)
        return _缓存[键]
    提示 = []
    索引 = _读(目录, "接续索引.json", 提示)
    说明 = _读(目录, "任务说明.json", 提示)
    次数 = 说明.get("执行次数", 0)
    事实 = _读(目录, "交付明细.jsonl", 提示) if (索引.get("产物") or {}).get("交付明细") else []
    明细哈希 = (索引.get("产物哈希") or {}).get("交付明细")
    明细文件 = 目录 / "交付明细.jsonl"
    明细变化 = bool(明细哈希 and (not 明细文件.is_file() or hashlib.sha256(明细文件.read_bytes()).hexdigest() != 明细哈希))
    if 明细变化:
        事实 = []
        提示.append("交付明细已变化，无法确认最终动作")
    事实 = [x for x in 事实 if x.get("执行次数", 0) == 次数]
    完整性 = ("交付明细已变化，无法确认最终动作" if 明细变化 else 索引.get("展示错误")
              or ("最终事实已记录" if 事实 else "旧任务未记录最终动作" if not 说明 else "本次尚无最终明细"))
    if 索引.get("步骤") == "done" and 索引.get("撤回表"):
        提示.append("旧记录状态不一致")
    A = _读(目录, "A合集.jsonl", 提示)
    结构 = _读(目录, "结构清单.jsonl", 提示)
    问题 = _读(目录, "问题清单.jsonl", 提示)
    格记录 = [x for x in 事实 if x.get("类型") == "格子"]
    if not 事实:
        格记录 = [{**x, "a格": x.get("格"), "侧": "A", "最终动作": "未记录"} for x in A]
    表们 = OrderedDict()
    for x in 结构:
        组键 = _组键(目录, x)
        组 = 表们.setdefault(组键, {"id": 组键, "组键": 组键, "表组": x.get("表组"),
                  "侧": x.get("侧"), "读取部分": x.get("读取部分"), "科目": x.get("科目"),
                  "表名": x.get("表名"), "物理表": [], "计数": {}, "问题数": 0})
        组["物理表"].append({k: x.get(k) for k in ("工作表", "范围", "区域", "续表序", "作用范围")})
    明细, 格到组, 问题到组 = [], {}, {}
    for i, x in enumerate(格记录):
        组键 = _组键(目录, x)
        格到组[x.get("a格")] = 组键
        if 组键 not in 表们:
            表们[组键] = {"id": 组键, "组键": 组键, "表组": x.get("表组"), "侧": "A",
                         "读取部分": x.get("读取部分"), "科目": x.get("科目"), "表名": x.get("表名"),
                         "物理表": [], "计数": {}, "问题数": 0}
        计数 = 表们[组键]["计数"]
        动作 = x.get("最终动作", "未记录")
        计数[动作] = 计数.get(动作, 0) + 1
        原因 = [问题[n] for n in x.get("原因引用", []) if isinstance(n, int) and 0 <= n < len(问题)]
        for n in x.get("原因引用", []):
            if isinstance(n, int) and 0 <= n < len(问题):
                问题到组.setdefault(n, set()).add(组键)
        明细.append({"id": _标识(目录.name, x.get("a格"), i), "组键": 组键,
                    "科目": x.get("科目"), "表名": x.get("表名"), "角色": x.get("角色"),
                    "项目": " / ".join(x.get("行路径") or []), "期间": " / ".join(x.get("列路径") or []),
                    "维度值": x.get("维度值") or {}, "A位置": 位置(x.get("a格")), "B位置": 位置(x.get("b格")),
                    "A原值": 值对象(x.get("原值"), x.get("原值状态")),
                    "B本次值": 值对象(x.get("来源值"), x.get("来源值状态")),
                    "A实际值": 值对象(x.get("实际值"), 已取=x.get("实际值已取得", False), 来源="最终回读"),
                    "最终动作": 动作, "原因": 原因, "公式核对": x.get("公式核对")})
    问题们 = []
    for i, x in enumerate(问题):
        组键 = 格到组.get(x.get("a格"))
        if not 组键 and x.get("表组"):
            候选 = [k for k, v in 表们.items() if v["表组"] == x["表组"]
                    and (not x.get("侧") or v["侧"] == x["侧"])]
            组键 = 候选[0] if len(候选) == 1 else None
        关联 = set(问题到组.get(i, ()))
        if 组键:
            关联.add(组键)
        条 = {**x, "id": f"问题{i}", "组键": 组键, "关联组键": sorted(关联)}
        问题们.append(条)
        for k in 关联:
            if k in 表们:
                表们[k]["问题数"] += 1
    标题 = [{**x, "id": f"标题{i}", "组键": _组键(目录, x)} for i, x in enumerate(事实) if x["类型"] == "标题"]
    公式 = [{**x, "id": f"公式{i}", "组键": 格到组.get(x.get("a格")),
             "A计算值": 值对象(x.get("A结果")), "B核对值": 值对象(x.get("B金额")),
             "差额值": 值对象(x.get("差额"), 已取=x.get("差额") is not None)}
            for i, x in enumerate(_读(目录, "公式核对记录.jsonl", 提示))]
    复用 = [{"侧": 侧, **x} for 侧 in ("A", "B")
          for x in (_读(目录, f"复用记录_{侧}.json", 提示).get("部分") or [])]
    for 表 in 表们.values():
        表["需要查看"] = bool(表["问题数"] or any(表["计数"].get(k) for k in ("原值保留", "回读有差异", "未记录"))
                            or any(x.get("组键") == 表["组键"] and x.get("结论") != "本次金额一致" for x in 公式))
    出 = {"task_id": 目录.name, "执行次数": 次数, "结果版本": 版本, "步骤": 索引.get("步骤"),
          "明细完整性": 完整性, "提示": 提示, "计数": dict(Counter(x["最终动作"] for x in 明细)),
          "需查看表数": sum(x["侧"] == "A" and x["需要查看"] for x in 表们.values()),
          "问题条数": len(问题们), "视图": {"表": list(表们.values()), "明细": 明细,
                                      "问题": 问题们, "标题": 标题, "公式": 公式, "复用": 复用}}
    # Decimal只在数值转换成原样文本前使用；其余证据也按字符串传递，避免前端改精度。
    出 = json.loads(json.dumps(出, ensure_ascii=False, default=str))
    _缓存[键] = 出
    while len(_缓存) > 8:
        _缓存.popitem(last=False)
    return 出


def 文件列表(目录):
    索引 = 读索引(目录) or {}
    出 = []
    登记 = dict(索引.get("产物") or {})
    try:
        输入 = 输入线索(目录)
    except HTTPException:
        输入 = {}
    for 侧, 旧键, 人工键 in (("A", "word", "带连接word"), ("B", "b表", "excel")):
        原件 = (输入.get(侧) or {}).get("路径") or 输入.get(旧键) or 输入.get(人工键)
        if 原件:
            登记[侧 + "输入原件"] = 原件
    for 名, 路径 in 登记.items():
        if not isinstance(路径, str):
            continue
        p = Path(路径)
        存在 = p.is_file()
        状况 = "文件缺失"
        if 存在:
            预期 = (索引.get("产物哈希") or {}).get(名)
            if 名.endswith("输入原件"):
                指纹 = 索引.get("输入指纹") or {}
                指纹项 = 指纹.get(名[0]) or 指纹.get("word" if 名[0] == "A" else "b表") or {}
                预期 = 指纹项.get("哈希")
            if 名 == "A端新副本" and not 预期:
                预期 = next((v for k, v in (索引.get("产物哈希") or {}).items()
                            if (索引.get("产物") or {}).get(k) == 路径), None)
            键 = (str(p), p.stat().st_size, p.stat().st_mtime_ns)
            if 预期:
                if 键 not in _文件缓存:
                    _文件缓存[键] = hashlib.sha256(p.read_bytes()).hexdigest()
                状况 = "未改动" if _文件缓存[键] == 预期 else "文件已被改动"
            else:
                状况 = "未记录原哈希"
        出.append({"file_id": _标识(名, 路径), "名称": 名, "路径": 路径,
                  "正式": 名 in ("A端新副本", "Excel新副本", "Word新副本"),
                  "候选": "候选" in 名, "存在": 存在, "文件状况": 状况})
    while len(_文件缓存) > 128:
        _文件缓存.popitem(last=False)
    return 出


def 挂接结果接口(应用, 宿主):
    @应用.get("/api/桌面/任务/{task_id}/结果")
    def 查询(task_id: str, view: str = "表", 组: str = "", 搜索: str = "", 状态: str = "",
           offset: int = 0, limit: int = 50, 结果版本: str = "", 侧: str = "", 仅需查看: bool = False):
        数据 = 整理结果(任务目录(宿主, task_id))
        if 结果版本 and 结果版本 != 数据["结果版本"]:
            拒绝(409, "本任务结果已变化，请刷新后查看")
        if view not in 数据["视图"]:
            拒绝(400, "没有这个结果视图")
        行们 = 数据["视图"][view]
        行们 = [x for x in 行们 if (not 组 or x.get("组键") == 组 or 组 in x.get("关联组键", []))
               and (not 侧 or x.get("侧") == 侧) and (not 仅需查看 or x.get("需要查看"))
               and (not 状态 or x.get("最终动作") == 状态)
               and (not 搜索 or 搜索.casefold() in json.dumps(x, ensure_ascii=False).casefold())]
        return {**{k: v for k, v in 数据.items() if k != "视图"}, "总数": len(行们),
                "条目": 行们[max(0, offset):max(0, offset) + min(100, max(1, limit))]}

    @应用.get("/api/桌面/任务/{task_id}/文件")
    def 文件(task_id: str):
        return {"task_id": task_id, "条目": 文件列表(任务目录(宿主, task_id))}

    @应用.post("/api/桌面/任务/{task_id}/文件/{file_id}/打开")
    def 打开(task_id: str, file_id: str, 请求: dict):
        目录 = 任务目录(宿主, task_id)
        if file_id == "任务目录":
            p = 目录
        else:
            条 = next((x for x in 文件列表(目录) if x["file_id"] == file_id), None)
            if not 条:
                拒绝(404, "没有这份已登记文件")
            p = Path(条["路径"])
            if 请求.get("打开目录"):
                p = p.parent
        if not p.exists():
            拒绝(404, "文件或目录已不在原位置")
        try:
            os.startfile(str(p))
        except OSError as e:
            拒绝(400, str(e), "打开失败")
        return {"ok": True}
