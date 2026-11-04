# coding: utf-8
"""跨任务复用的接入层：三遍调用点前的候选核对、使用登记与发布。

本模块只做机械核对：证据来自首次识别的标注，金额不参与比较，位置一律用本次文件重定位。
任何不确定都返回“未命中／冲突”及原因，让调用方走现有完整识别，不在任务外另建流程。
"""
from __future__ import annotations

import json
import time

from openpyxl.utils import range_boundaries, column_index_from_string
from openpyxl.utils.cell import coordinate_to_tuple

from . import 复用证据 as 证据模块
from .共用记录簿 import 使用事件, 发布事件
from .标题 import 本次标题, 解析日期


def 区域标题们(上下文):
    """A 侧当前各物理表的表题与行范围；只用作定位锚点，不参与科目判断。"""
    if isinstance(上下文, dict):
        表们 = 上下文.get("tables") or []
    elif isinstance(上下文, (list, tuple)):
        表们 = list(上下文)
    else:
        return []
    出 = []
    for 项 in 表们:
        if not isinstance(项, dict):
            continue
        名, 题 = 项.get("range_name"), (项.get("title") or "").strip()
        顶, 数 = 项.get("first_row"), 项.get("row_count")
        if not 名 or not 题 or not isinstance(顶, int) or not isinstance(数, int) or 数 <= 0:
            continue
        出.append({"区域": 名, "标题": 题, "行范围": [顶, 顶 + 数 - 1]})
    return 出


def 取复用(簿):
    """任务开始时取一次已提交快照；索引建立失败即本次禁用复用。"""
    if 簿 is None:
        return {"可用": False, "原因": "未启用跨任务复用", "索引": {}, "秒": 0.0}
    起 = time.perf_counter()
    索引 = 簿.快照()
    return {"可用": bool(索引.get("可用")), "原因": 索引.get("原因") or "",
            "索引": 索引, "秒": time.perf_counter() - 起}


def 版本可用(项, 金标准版本, 规则版本):
    if not isinstance(项, dict):
        return False, "记录缺失或损坏"
    if 项.get("状态") != "有效":
        return False, "记录已停用"
    if 项.get("金标准版本") != 金标准版本:
        return False, "金标准版本已变"
    if 项.get("规则版本") != 规则版本:
        return False, "识别规则版本已变"
    if (项.get("证据") or {}).get("证据版本") != 证据模块.证据版本:
        return False, "证据版本不兼容"
    return True, ""


def 候选版本(阅, 侧, 步骤, 形状键=None):
    if not 阅.get("可用"):
        return []
    槽 = f"{侧}|{步骤}|{形状键 or '-'}"
    return [(v, (阅["索引"].get("解释") or {}).get(v))
            for v in (阅["索引"].get("按步") or {}).get(槽, [])]


def 登记使用(簿, 版本, 任务, 侧, 表组, 代次):
    """依赖旧解释生成本次结果之前先登记；登记失败就放弃这次命中。"""
    if 簿 is None:
        return False, "记录簿不可用"
    事件 = 使用事件(版本, 任务, 侧, 表组, 代次)
    try:
        簿.提交([事件], 请求标识=事件["请求"])
        return True, ""
    except Exception as e:
        return False, f"使用登记失败：{e}"


def 发布记录(簿, 侧, 步骤, 证据, 解释, 金标准版本, 规则版本, 来源, 任务,
         形状键=None, 依赖=None):
    if 簿 is None:
        return False, "未启用跨任务复用"
    证据键值 = 证据模块.证据键(证据)
    解释键值 = 证据模块.解释键(解释)
    版本 = 证据模块.版本号(证据键值, 解释键值, 金标准版本, 规则版本)
    事件 = 发布事件(侧, 步骤, 证据, 解释, 证据键值, 解释键值, 版本,
                金标准版本, 规则版本, 来源, 形状键=形状键, 依赖=依赖)
    try:
        簿.提交([事件], 请求标识=f"发布:{任务}:{侧}:{步骤}:{版本}")
        return True, 版本
    except Exception as e:
        return False, f"发布失败：{e}"


def 试复用科目(阅, 侧, 快照, 用可见范围, 工作表, 标题们, 金标版, 规则版):
    候选 = 候选版本(阅, 侧, "第一遍")
    if not 候选:
        return {"结果": "未命中", "原因": "没有同侧分科目记录" if 阅.get("可用")
                else (阅.get("原因") or "记录簿不可用")}
    有效, 原因们 = [], []
    for 版本, 项 in 候选:
        可用, 因 = 版本可用(项, 金标版, 规则版)
        if not 可用:
            原因们.append(因)
            continue
        出 = 证据模块.核对分科目(项["证据"], 快照, 用可见范围, 工作表, 标题们)
        if 出["结果"] == "唯一":
            有效.append((版本, 项, 出["划分"]))
        else:
            原因们.append(出["原因"])
    不同 = {json.dumps(x[2], sort_keys=True, ensure_ascii=False) for x in 有效}
    if len(不同) > 1:
        return {"结果": "冲突", "原因": "同一份材料对应多种科目划分，不自动选用"}
    if len(有效) == 1:
        return {"结果": "唯一", "原因": "", "数据": 有效[0][2], "版本": 有效[0][0],
                "项": 有效[0][1]}
    return {"结果": "未命中", "原因": "；".join(dict.fromkeys(原因们))[:300] or "没有可用记录"}


def 试复用分表(阅, 侧, 快照, 划分, 用可见范围, 工作表, 标题们, 金标版, 规则版):
    候选 = 候选版本(阅, 侧, "第二遍")
    if not 候选:
        return {"结果": "未命中", "原因": "没有同侧分表记录" if 阅.get("可用")
                else (阅.get("原因") or "记录簿不可用")}
    有效, 原因们 = [], []
    for 版本, 项 in 候选:
        可用, 因 = 版本可用(项, 金标版, 规则版)
        if not 可用:
            原因们.append(因)
            continue
        出 = 证据模块.核对分表(项["证据"], 快照, 划分, 用可见范围, 工作表, 标题们)
        if 出["结果"] == "唯一":
            有效.append((版本, 项, 出["表"]))
        else:
            原因们.append(出["原因"])
    不同 = {json.dumps(x[2], sort_keys=True, ensure_ascii=False) for x in 有效}
    if len(不同) > 1:
        return {"结果": "冲突", "原因": "同一份材料对应多种分表结论，不自动选用"}
    if len(有效) == 1:
        return {"结果": "唯一", "原因": "", "数据": 有效[0][2], "版本": 有效[0][0],
                "项": 有效[0][1]}
    return {"结果": "未命中", "原因": "；".join(dict.fromkeys(原因们))[:300] or "没有可用记录"}


def 组形状(快照, 记录们, 用可见范围):
    记录们 = sorted(记录们, key=lambda x: x.get("续表序") or 1)
    顶 = min(x["行范围"][0] for x in 记录们)
    底 = max(x["行范围"][1] for x in 记录们)
    左 = min(range_boundaries(str(x["范围"]))[0] for x in 记录们)
    右 = max(range_boundaries(str(x["范围"]))[2] for x in 记录们)
    视图 = 证据模块.表视图(快照, 记录们[0]["工作表"], (左, 顶, 右, 底), 用可见范围)
    分表 = [[x["行范围"][0] - 顶, x["行范围"][1] - 顶] for x in 记录们]
    return 视图, 分表, 顶


def 试复用精读(阅, 侧, 快照, 记录们, 索引, 用可见范围, 金标版, 规则版):
    """按整组形状找候选，再逐项核对固定证据；返回本次坐标下的解释。"""
    if not 阅.get("可用"):
        return {"结果": "未命中", "原因": 阅.get("原因") or "记录簿不可用"}
    视图, 分表, 顶 = 组形状(快照, 记录们, 用可见范围)
    形状键 = 证据模块.形状摘要(视图.形状(分表))
    候选 = 候选版本(阅, 侧, "第三遍", 形状键)
    if not 候选:
        return {"结果": "未命中", "原因": "没有同结构记录"}
    有效, 原因们 = [], []
    for 版本, 项 in 候选:
        可用, 因 = 版本可用(项, 金标版, 规则版)
        if not 可用:
            原因们.append(因)
            continue
        if (阅["索引"].get("冲突") or {}).get(项.get("证据键")):
            原因们.append("同一证据存在冲突解释，未经裁决不自动选用")
            continue
        出 = 证据模块.核对精读(项["证据"], 视图)
        if 出["结果"] != "唯一":
            原因们.append(出["原因"])
            continue
        问题 = 解释可用(出["解释"], 索引, 记录们) or 补本次日期(出["解释"], 视图)
        if 问题:
            原因们.append(问题)
            continue
        有效.append((版本, 项, 出["解释"]))
    不同 = {证据模块.解释键(x[2]) for x in 有效}
    if len(不同) > 1:
        return {"结果": "冲突", "原因": "同一结构对应多种解释，不按新旧或次数选用"}
    if len(有效) == 1:
        return {"结果": "唯一", "原因": "", "数据": 有效[0][2], "版本": 有效[0][0],
                "项": 有效[0][1], "顶": 顶, "形状键": 形状键}
    return {"结果": "未命中", "原因": "；".join(dict.fromkeys(原因们))[:300] or "没有可用记录"}


def 补本次日期(解释, 视图):
    """复用只带解释；日期要按标注位置从本次文件重读，绝不沿用上一份的值。

    位置是唯一能重读的依据。位置缺失却带着旧值继续，就会把上一份的日期写进本次标题，
    所以这种情况判为不可复用；位置本来就空、也没有旧值的项不受影响。
    """
    for 项 in list(解释.get("列") or []) + list(解释.get("行") or []):
        位置 = 项.get("期间实际位置")
        if not 位置:
            if 项.get("期间实际"):
                return f"记录只有日期值没有位置，无法按本次文件重读：{项.get('期间实际')}"
            continue
        try:
            行, 列 = coordinate_to_tuple(str(位置).upper())
        except (TypeError, ValueError):
            return f"日期位置不是有效单元格：{位置}"
        文 = 视图.文本(行, 列)
        if not 文:
            return f"日期位置本次没有内容：{位置}"
        日期 = 解析日期(文)
        if 日期 is None:
            return f"日期位置本次不是日期：{位置}“{文}”"
        项["期间实际"] = 日期.isoformat()
    return ""


def 解释可用(解释, 索引, 记录们):
    """换成本次坐标后再过一遍必备检查：编号、期间、槽位与表名。"""
    from .识别 import _校验格期间

    if (解释.get("科目"), 解释.get("表名")) != (记录们[0].get("科目"), 记录们[0].get("表名")):
        return "记录的表名与本次分表结论不一致"
    编号们 = {编号 for 行 in 解释.get("行") or []
            for 编号 in (行.get("槽位") or {}).values() if 编号}
    失效 = sorted(编号们 - 索引["编号"])
    if 失效:
        return "槽位编号已失效：" + "、".join(失效[:3])
    try:
        _校验格期间(解释.get("列") or [], 解释.get("行") or [], 编号们)
    except RuntimeError as e:
        return str(e)
    for 行 in 解释.get("行") or []:
        if 行.get("角色") in ("数据", "小计", "合计") and not (行.get("槽位") or 行.get("不搬原因")):
            return f"第 {行.get('行')} 行缺少槽位映射"
    return ""


def 应用精读(记录们, 解释, 版本, 快照, 上下文=None):
    """把整组解释按本次物理表分表落到每张表；位置已是本次绝对坐标。

    标题只复用“骨架＋期间角色”，片段用本次标题回填；对不上就不设标题，
    由配对后的标题生成阶段记组级失败，不会拿旧新文硬写。
    """
    from .识别 import _合并们

    记录们 = sorted(记录们, key=lambda x: x.get("续表序") or 1)
    for 记录 in 记录们:
        起, 止 = 记录["行范围"]
        左, _, 右, _ = range_boundaries(str(记录["范围"]))
        列们 = [dict(c) for c in 解释.get("列") or []
              if 左 <= column_index_from_string(str(c.get("列")).upper()) <= 右]
        行们 = [dict(r) for r in 解释.get("行") or [] if 起 <= r.get("行") <= 止]
        例外 = []
        for e in 解释.get("例外") or []:
            try:
                左2, 顶2, 右2, 底2 = range_boundaries(str(e.get("范围") or "").replace("$", ""))
            except (ValueError, TypeError):
                continue
            if 顶2 <= 止 and 底2 >= 起:
                例外.append(dict(e))
        记录.update({"列": 列们, "行": 行们, "作用范围": dict(解释.get("作用范围") or {}),
                    "空名单": list(解释.get("空名单") or []) if 记录 is 记录们[0] else [],
                    "例外": 例外, "完成": True, "复用版本": 版本,
                    "原始结构": {"合并": _合并们(快照, 记录["工作表"], 记录["范围"])}})
    规则 = 解释.get("标题规则") or {}
    if 规则.get("有标题") and 记录们:
        当前 = 本次标题(上下文, 记录们[0].get("区域"))
        片段们 = 证据模块.按骨架取片段(规则.get("骨架") or "", 当前)
        记录们[0]["标题"] = ({"原文": 当前,
                          "日期": [{"片段": 片段, "期间角色": 角色} for 角色, 片段 in 片段们],
                          "旧新文": None, "旧新文已忽略": False, "复用还原": True}
                         if 片段们 else None)
    return 记录们
