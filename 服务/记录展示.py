# coding: utf-8
"""记录管理界面适配；发布、停用、冲突判断和依赖追溯全部沿用原函数。"""
import json
from datetime import datetime
from pathlib import Path
from 核心.共用记录簿 import 取记录簿
from 核心.复用接入 import 版本可用
from 核心.复用证据 import 内容摘要, 规则摘要
from 核心 import 复用管理
from .任务服务 import 拒绝
from .界面设置 import 读取默认


def 当前记录簿(宿主):
    return 取记录簿(读取默认().get("识别记录目录") or None)


def 当前上下文():
    默认 = 读取默认()
    p = Path(默认.get("金标准") or "")
    return {"金标准": str(p), "金标准版本": 内容摘要(p) if p.is_file() else None, "规则版本": 规则摘要()}


def _摘要(项, 索引, 上下文):
    可用, 原因 = 版本可用(项, 上下文["金标准版本"], 上下文["规则版本"])
    if not 上下文["金标准版本"] or not 上下文["规则版本"]:
        可用, 原因 = False, "当前金标准或规则依据不完整"
    冲突们 = (索引.get("冲突") or {}).get(项.get("证据键"), [])
    if 冲突们:
        可用, 原因 = False, "同一证据有冲突解释"
    解释 = 项.get("解释") or {}
    return {"版本": 项["版本"], "侧": 项.get("侧"), "步骤": 项.get("步骤"),
            "科目": 解释.get("科目"), "表名": 解释.get("表名"), "存档状态": 项.get("状态"),
            "当前适用": "可自动复用" if 可用 else 原因, "冲突": bool(冲突们),
            "证据组": 项.get("证据键"), "来源数": len(项.get("来源") or [])}


def 任务影响(宿主, 目录):
    索引 = 当前记录簿(宿主).快照()
    出 = []
    for 标识, 项 in (索引.get("请求") or {}).items():
        事件 = 项.get("事件") or {}
        if 事件.get("类型") not in ("停用", "裁决"):
            continue
        结果 = 事件.get("结果") or {}
        if str(目录) in (结果.get("待复核") or []):
            出.append({"请求": 标识, "原因": 事件.get("原因"), "依据": 事件.get("依据"),
                      "表": [x for 类 in ("直接", "间接") for x in (结果.get(类) or {}).get("任务表组", []) if x.get("任务") == str(目录)]})
    return 出


def 挂接记录接口(应用, 宿主):
    @应用.get("/api/桌面/记录")
    def 列表(搜索: str = "", 状态: str = "", offset: int = 0, limit: int = 50):
        簿 = 当前记录簿(宿主)
        索引 = 簿.快照()
        上下文 = 当前上下文()
        行们 = [_摘要(x, 索引, 上下文) for x in (索引.get("解释") or {}).values()]
        行们 = [x for x in 行们 if (not 搜索 or 搜索.casefold() in json.dumps(x, ensure_ascii=False).casefold())
              and (not 状态 or (状态 == "冲突" and x["冲突"]) or x["存档状态"] == 状态 or x["当前适用"] == 状态)]
        return {"目录": str(簿.目录 or ""), "可用": 索引.get("可用"), "原因": 索引.get("原因"),
                "代次": 索引.get("代次"), "上下文": 上下文, "总数": len(行们),
                "条目": 行们[max(0, offset):max(0, offset) + min(100, max(1, limit))]}

    @应用.get("/api/桌面/记录/请求/{request_id}")
    def 已提交(request_id: str):
        try:
            return 复用管理.查询(当前记录簿(宿主), 请求标识=request_id)
        except ValueError as e:
            拒绝(404, str(e))

    @应用.get("/api/桌面/记录/{version}")
    def 详情(version: str):
        簿 = 当前记录簿(宿主)
        索引 = 簿.快照()
        项 = (索引.get("解释") or {}).get(version)
        if not 项:
            拒绝(404, "没有这个解释版本")
        冲突们 = (索引.get("冲突") or {}).get(项.get("证据键"), [])
        return {**_摘要(项, 索引, 当前上下文()), "记录": 项,
                "冲突解释": [索引["解释"][v] for v in 冲突们],
                "影响": 复用管理.查询(簿, 解释版本=version),
                "历史操作": [x["事件"] for x in 索引.get("请求", {}).values()
                           if version in ((x.get("事件") or {}).get("目标", []) + (x.get("事件") or {}).get("失效", []))]}

    @应用.post("/api/桌面/记录/影响预览")
    def 预览(请求: dict):
        索引 = 当前记录簿(宿主).快照()
        版本 = 请求.get("版本") or []
        if not 版本 or not all(v in (索引.get("解释") or {}) for v in 版本):
            拒绝(400, "请选择存在的解释版本")
        return {"代次": 索引.get("代次"), "时点": datetime.now().isoformat(timespec="seconds"),
                **复用管理._影响(索引, 版本)}

    @应用.post("/api/桌面/记录/停用")
    def 停用(请求: dict):
        try:
            结果 = 复用管理.停用(当前记录簿(宿主), 请求.get("请求标识"), 请求.get("版本"), 请求.get("原因"), 请求.get("依据"))
            return {**结果, "预览后变化": 请求.get("预览代次") != 结果.get("代次", 0) - 1}
        except (ValueError, RuntimeError) as e:
            拒绝(409, str(e), "停用未接受")

    @应用.post("/api/桌面/记录/裁决")
    def 裁决(请求: dict):
        try:
            return 复用管理.裁决(当前记录簿(宿主), 请求.get("请求标识"), 请求.get("证据组"),
                              请求.get("预期冲突版本"), 保留=请求.get("保留"), 停用们=请求.get("停用"),
                              原因=请求.get("原因"), 依据=请求.get("依据"))
        except (ValueError, RuntimeError) as e:
            拒绝(409, str(e), "冲突处理未接受")
