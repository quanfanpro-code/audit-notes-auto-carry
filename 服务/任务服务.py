# coding: utf-8
"""任务目录仍是事实来源；只统一接受操作、身份和历史查询。"""
import copy
import hashlib
import json
import os
import threading
import uuid
from datetime import datetime
from pathlib import Path
from 运行位置 import 检查数据位置

from fastapi import HTTPException
from 核心.接续 import 读索引, 写索引
from 核心.文件输入 import 规范化输入
from 核心.流程 import _输入指纹, _指纹差异

执行锁 = threading.RLock()
结果步骤 = {"done", "部分保留", "公式待核", "回读不一致"}
终态 = {"done", "部分保留", "公式待核", "回读不一致", "识别失败", "Excel交付失败", "任务失败", "用户中断"}
已知步骤 = 终态 | {"extract", "identify_a", "identify_b", "carry", "writeback", "readback"}


def 拒绝(码, 说明, 类别="操作未接受"):
    raise HTTPException(码, {"类别": 类别, "说明": 说明})


def 读说明(目录):
    p = Path(目录) / "任务说明.json"
    return json.loads(p.read_text(encoding="utf-8-sig")) if p.is_file() else {}


def 原子JSON(路径, 数据):
    路径 = Path(路径)
    临时 = 路径.with_name(路径.name + "." + uuid.uuid4().hex + ".tmp")
    临时.write_text(json.dumps(数据, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(临时, 路径)


def 目录列表(宿主):
    根 = Path(宿主.运行根).resolve()
    return sorted((p for p in 根.glob("任务_*") if p.is_dir() and p.resolve().parent == 根
                   and ((p / "接续索引.json").is_file() or (p / "任务说明.json").is_file())), reverse=True)


def 任务目录(宿主, 标识):
    if not 标识 or Path(标识).name != 标识 or not 标识.startswith("任务_"):
        拒绝(404, "没有找到这份任务")
    p = Path(宿主.运行根) / 标识
    if p not in 目录列表(宿主):
        拒绝(404, "没有找到这份任务")
    return p


def 正在执行(宿主):
    线程 = 宿主.应用.state.状态.get("线程")
    return bool(线程 and 线程.is_alive())


def 输入线索(目录):
    说明 = 读说明(目录)
    if "原始输入" in 说明:
        return copy.deepcopy(说明["原始输入"])
    指纹 = (读索引(目录) or {}).get("输入指纹") or {}
    if not 指纹.get("金标准", {}).get("路径"):
        拒绝(409, "旧任务没有可恢复的原输入，请新建任务")
    出 = {"金标准": 指纹["金标准"]["路径"]}
    if "A" in 指纹 and "B" in 指纹:
        for 侧 in ("A", "B"):
            出[侧] = {k: copy.deepcopy(v) for k, v in 指纹[侧].items() if k in ("路径", "类型", "工作表")}
    elif "word" in 指纹 and "b表" in 指纹:
        出.update({k: 指纹[k]["路径"] for k in ("word", "b表")})
        出["b工作表"] = 指纹.get("b工作表")
    else:
        拒绝(409, "旧任务缺少两端输入，不能猜测续跑")
    return 出


def 原始输入(目录):
    if (读索引(目录) or {}).get("步骤") not in 已知步骤:
        拒绝(409, "未知任务步骤，只能查看已有记录")
    if "原始输入" not in 读说明(目录):
        拒绝(409, "旧任务未完整保存原始选项，请选用原文件新建任务")
    return 输入线索(目录)


def 任务进度(宿主, 目录):
    索引 = 读索引(目录) or {}
    说明 = 读说明(目录)
    状态 = 宿主.应用.state.状态
    活动 = 正在执行(宿主) and Path(状态.get("运行目录") or "") == Path(目录)
    等待 = bool(索引.get("待答复")) and 索引.get("步骤") in ("identify_a", "identify_b") and not 活动
    类型 = 说明.get("类型", "自动搬运")
    步骤 = 索引.get("步骤")
    显示 = ("未知任务步骤：" + str(步骤) if 步骤 and 步骤 not in 已知步骤 else
            "正在中断" if 活动 and 状态.get("取消") and 状态["取消"].is_set() else
            "等待答复" if 等待 else "上次运行未正常结束" if not 活动 and 步骤 and 步骤 not in 终态 else 步骤 or "准备文件")
    原因 = ""
    try:
        原始输入(目录)
    except HTTPException as e:
        原因 = e.detail["说明"]
    if not 原因 and 步骤 in 结果步骤:
        原因 = "本任务已处理结束，请用原文件新建任务"
    可继续 = 类型 == "自动搬运" and not 正在执行(宿主) and not 原因
    return {"task_id": Path(目录).name, "运行目录": str(目录), "类型": 类型,
            "执行次数": 说明.get("执行次数", 0), "活动": 活动, "等待": 等待,
            "步骤": 步骤, "显示状态": 显示, "待答复": 索引.get("待答复", []),
            "已完成范围": len(索引.get("已完成范围") or []),
            "错误": 索引.get("错误"), "允许继续": 可继续, "不能继续原因": 原因 if 类型 == "自动搬运" else "人工回写请另存新任务",
            "允许中断": 类型 == "自动搬运" and (活动 or 等待),
            "复用统计": 宿主._复用统计(目录),
            "日志": 状态.get("日志", [])[-20:] if Path(状态.get("运行目录") or "") == Path(目录) else []}


def 任务摘要(宿主, 目录):
    进度 = 任务进度(宿主, 目录)
    说明 = 读说明(目录)
    try:
        输入 = 输入线索(目录)
    except HTTPException:
        输入 = {}
    A = 输入.get("A") or {"路径": 输入.get("word") or 输入.get("带连接word", ""), "工作表": []}
    B = 输入.get("B") or {"路径": 输入.get("b表") or 输入.get("excel", ""), "工作表": 输入.get("b工作表") or []}
    from .记录展示 import 任务影响
    try:
        影响 = 任务影响(宿主, 目录)
    except (OSError, ValueError, RuntimeError):
        影响 = None
    return {**进度, "名称": Path(A.get("路径") or "人工回写").name,
            "创建时间": 说明.get("创建时间") or datetime.fromtimestamp(Path(目录).stat().st_mtime).isoformat(timespec="seconds"),
            "时间来源": "任务记录" if 说明.get("创建时间") else "目录时间",
            "A": A, "B": B, "原始输入": 输入, "模型": 说明.get("模型", {}), "待复核影响": 影响}


def _校验输入(输入, 类型):
    try:
        if 输入.get("识别记录目录"):
            检查数据位置(输入["识别记录目录"])
        if 类型 == "人工回写" and (输入.get("输出") or 输入.get("output")):
            检查数据位置(输入.get("输出") or 输入["output"])
    except ValueError as e:
        拒绝(400, str(e))
    if 类型 == "人工回写":
        if not all(输入.get(k) for k in ("带连接word", "excel")) or not (输入.get("输出") or 输入.get("output")):
            拒绝(400, "请选择带连接Word、已修改Excel和新副本位置")
        输出 = Path(输入.get("输出") or 输入["output"]).resolve()
        if 输出 in {Path(输入[k]).resolve() for k in ("带连接word", "excel")}:
            拒绝(400, "新副本不能覆盖输入原件")
        return
    if 类型 != "自动搬运":
        拒绝(400, "不支持的任务类型")
    缺 = [k for k in (("A", "B", "金标准") if "A" in 输入 or "B" in 输入 else ("word", "b表", "金标准")) if not 输入.get(k)]
    if 缺:
        拒绝(400, "缺少字段：" + "、".join(缺))
    if "A" in 输入 or "B" in 输入:
        try:
            规范化输入(输入)
        except (ValueError, OSError) as e:
            拒绝(400, str(e))


def 接受任务(宿主, 输入, 请求标识=None, 类型="自动搬运", 续跑目录=None, 后台=True):
    输入 = copy.deepcopy(输入)
    输入.pop("续跑", None)
    if any(k in 输入 for k in ("api_key", "api_key_enc", "密钥")):
        拒绝(400, "密钥只能在模型设置中保存")
    请求标识 = str(请求标识 or uuid.uuid4().hex)
    if len(请求标识) > 128:
        拒绝(400, "请求标识过长")
    内容键 = hashlib.sha256(json.dumps([类型, 输入, str(续跑目录 or "")], sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    with 执行锁:
        if getattr(宿主.应用.state, "退出中", False):
            拒绝(409, "程序正在安全退出，请等待结束")
        for p in 目录列表(宿主):
            for 旧 in 读说明(p).get("请求", []):
                if 旧["标识"] == 请求标识:
                    if 旧["内容键"] != 内容键:
                        拒绝(409, "同一请求标识不能用于不同输入")
                    return 旧["返回"]
        if 正在执行(宿主):
            拒绝(409, "已有任务在运行，请等待其安全结束")
        _校验输入(输入, 类型)
        if 续跑目录:
            索引 = 读索引(续跑目录) or {}
            if 索引.get("步骤") in 结果步骤:
                拒绝(409, "本任务已处理结束，请用原文件新建任务")
            差 = _指纹差异(索引.get("输入指纹"), _输入指纹(输入))
            if 差:
                拒绝(409, "续跑输入与原任务不一致：" + "；".join(差))
            目录 = Path(续跑目录)
        else:
            目录 = Path(宿主.运行根) / ("任务_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        目录.mkdir(parents=True, exist_ok=True)
        说明 = 读说明(目录)
        try:
            配置 = 宿主.读取模型配置(宿主.应用.state.配置路径)
        except Exception:
            配置 = {}
        次数 = 说明.get("执行次数", 0) + 1
        返回 = {"task_id": 目录.name, "运行目录": str(目录), "执行次数": 次数}
        说明.update({"类型": 类型, "原始输入": 输入, "执行次数": 次数,
                   "创建时间": 说明.get("创建时间") or datetime.now().isoformat(timespec="seconds"),
                   "模型": {k: 配置[k] for k in ("url", "model", "timeout") if k in 配置}})
        说明.setdefault("请求", []).append({"标识": 请求标识, "内容键": 内容键, "返回": 返回})
        原子JSON(目录 / "任务说明.json", 说明)
        取消 = threading.Event()
        def 跑():
            try:
                if 类型 == "人工回写":
                    写索引(目录, {"步骤": "writeback", "产物": {}})
                    结果 = 宿主.人工回写(输入)
                    输出 = 输入.get("输出") or 输入.get("output")
                    摘要 = hashlib.sha256(Path(输出).read_bytes()).hexdigest()
                    写索引(目录, {"步骤": "done", "产物": {"Word新副本": 输出, "A端新副本": 输出},
                                 "产物哈希": {"Word新副本": 摘要, "A端新副本": 摘要}, "结果": 结果})
                    返回["结果"] = 结果
                else:
                    宿主.运行任务(输入, 目录, 配置, 取消=取消)
            except Exception as e:
                宿主.应用.state.状态["日志"].append("失败：" + str(e))
                索引 = 读索引(目录) or {"已完成范围": [], "产物": {}}
                步骤 = getattr(e, "任务步骤", "任务失败")
                if 步骤 == "任务失败":
                    for k in ("识别错误", "Excel错误", "待答复"):
                        索引.pop(k, None)
                索引.update({"步骤": 步骤, "错误": str(e)})
                写索引(目录, 索引)
                返回["错误"] = str(e)
        线程 = threading.Thread(target=跑, daemon=True)
        宿主.应用.state.状态.update({"线程": 线程, "运行目录": 目录, "日志": [], "取消": 取消})
        线程.start()
    if not 后台:
        线程.join()
    return 返回


def 中断任务(宿主, 目录):
    with 执行锁:
        状态 = 宿主.应用.state.状态
        if 正在执行(宿主) and Path(状态.get("运行目录") or "") != Path(目录):
            拒绝(409, "正在运行的是另一份任务，本次操作未中断它")
        if 读说明(目录).get("类型") == "人工回写" and 正在执行(宿主):
            拒绝(409, "人工回写请等待当前写回安全结束")
        if 正在执行(宿主):
            状态["取消"].set()
        else:
            索引 = 读索引(目录) or {}
            if 索引.get("步骤") in ("identify_a", "identify_b") and 索引.get("待答复"):
                索引["步骤"] = "用户中断"
                索引.pop("待答复", None)
                写索引(目录, 索引)
        return {"ok": True, **任务进度(宿主, 目录)}


def 挂接任务接口(应用, 宿主):
    @应用.get("/api/桌面/概况")
    def 概况():
        状态 = 宿主.应用.state.状态
        目录 = 状态.get("运行目录")
        from .界面设置 import 概况补充
        return {"活动task_id": Path(目录).name if 目录 and 正在执行(宿主) else None,
                "服务就绪": True, **概况补充(宿主)}

    @应用.get("/api/桌面/任务")
    def 列表(搜索: str = "", 状态: str = "", offset: int = 0, limit: int = 50):
        出 = []
        for p in 目录列表(宿主):
            try:
                条目 = 任务摘要(宿主, p)
            except (OSError, ValueError, TypeError):
                条目 = {"task_id": p.name, "名称": p.name, "显示状态": "记录不可读", "创建时间": ""}
            if 搜索 and 搜索.casefold() not in json.dumps(条目, ensure_ascii=False).casefold():
                continue
            if 状态 and 条目.get("显示状态") != 状态:
                continue
            出.append(条目)
        本页 = 出[max(0, offset):max(0, offset) + min(100, max(1, limit))]
        from .结果展示 import 整理结果, 文件列表
        for 条 in 本页:
            try:
                p = Path(宿主.运行根) / 条["task_id"]
                数据 = 整理结果(p)
                条["需查看表数"] = 数据["需查看表数"]
                文件 = [x for x in 文件列表(p) if x["正式"]]
                条["文件状况"] = "、".join(dict.fromkeys(x["文件状况"] for x in 文件)) or "未登记正式副本"
            except (OSError, ValueError, TypeError):
                条["需查看表数"], 条["文件状况"] = None, "资料不完整"
        return {"条目": 本页, "总数": len(出)}

    @应用.post("/api/桌面/任务")
    def 开始(请求: dict):
        return 接受任务(宿主, 请求.get("原始输入") or {}, 请求.get("请求标识"), 请求.get("类型", "自动搬运"))

    @应用.get("/api/桌面/任务/{task_id}/进度")
    def 进度(task_id: str):
        return 任务进度(宿主, 任务目录(宿主, task_id))

    @应用.get("/api/桌面/任务/{task_id}")
    def 详情(task_id: str):
        return 任务摘要(宿主, 任务目录(宿主, task_id))

    @应用.post("/api/桌面/任务/{task_id}/继续")
    def 继续(task_id: str, 请求: dict):
        p = 任务目录(宿主, task_id)
        if 读说明(p).get("类型") == "人工回写":
            拒绝(409, "人工回写不支持续跑，请重新另存")
        return 接受任务(宿主, 原始输入(p), 请求.get("请求标识"), 续跑目录=p)

    @应用.post("/api/桌面/任务/{task_id}/中断")
    def 中断(task_id: str):
        return 中断任务(宿主, 任务目录(宿主, task_id))
