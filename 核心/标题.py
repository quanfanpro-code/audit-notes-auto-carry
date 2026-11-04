# coding: utf-8
"""本次标题：配对之后用本次 B 的期间日期生成，不取上一份任务的写回指令。

A 的第二遍只标明标题定位、日期片段与期间角色；日期一律来自本次 B 对应业务表组里
已确认的“期间实际”和它的证据坐标。取不到、同角色冲突或标题原文对不上，就整组撤回，
标题与数值一起保持原样，其他组继续。
"""
from __future__ import annotations

import re
from datetime import date

_日期 = re.compile(r"(\d{4})\s*([年\-/\.])\s*(\d{1,2})\s*[月\-/\.]\s*(\d{1,2})")


def 解析日期(文本):
    """只认明确写出年月日的文本；解析不了就返回 None，不按系统日期补。"""
    m = _日期.search(str(文本 or ""))
    if not m:
        return None
    年, 月, 日 = int(m.group(1)), int(m.group(3)), int(m.group(4))
    try:
        return date(年, 月, 日)
    except ValueError:
        return None


def 格式化(日期值: date, 片段: str):
    """按原片段的写法换年月日：中式保留“年月日”，横线／斜杠保留分隔与补零习惯。"""
    片段 = str(片段 or "")
    补零 = bool(re.search(r"[-/\.]0\d", 片段))
    if "年" in 片段:
        return f"{日期值.year}年{日期值.month}月{日期值.day}日"
    if "/" in 片段:
        return (f"{日期值.year}/{日期值.month:02d}/{日期值.day:02d}" if 补零
                else f"{日期值.year}/{日期值.month}/{日期值.day}")
    if "-" in 片段:
        return (f"{日期值.year}-{日期值.month:02d}-{日期值.day:02d}" if 补零
                else f"{日期值.year}-{日期值.month}-{日期值.day}")
    if "." in 片段:
        return (f"{日期值.year}.{日期值.month:02d}.{日期值.day:02d}" if 补零
                else f"{日期值.year}.{日期值.month}.{日期值.day}")
    if 片段.isdigit() and len(片段) == 8:
        return f"{日期值.year:04d}{日期值.month:02d}{日期值.day:02d}"
    return None


def _组(记录):
    return (str(记录.get("科目") or ""), str(记录.get("表名") or ""))


def 本次标题(上下文, 区域):
    """本次文件里该物理表的表题；上下文既可以是 {"tables": [...]} 也可以是列表。"""
    表们 = ((上下文 or {}).get("tables") or []) if isinstance(上下文, dict) else list(上下文 or [])
    for 项 in 表们:
        if isinstance(项, dict) and 项.get("range_name") == 区域:
            return str(项.get("title") or "").strip()
    return ""


def 生成本次标题(A结构, B结构, 上下文, 变更, B合集):
    """返回（标题更新, 问题清单, 需整组撤回的组键集合, 证据清单）。"""
    a按组 = {}
    for 记录 in A结构 or []:
        a按组.setdefault(_组(记录), []).append(记录)
    b按组 = {_组(x): x for x in B结构 or []}
    b格 = {}
    for r in B合集 or []:
        b格.setdefault((r.get("科目"), r.get("表名")), []).append(r)
    范围表 = []
    for 记录 in A结构 or []:
        左, 顶, 右, 底 = _范围(记录.get("范围"))
        if 右:
            范围表.append((str(记录.get("工作表") or ""), 左, 顶, 右, 底, _组(记录)))
    范围表.sort(key=lambda x: (x[3] - x[1] + 1) * (x[4] - x[2] + 1))
    a格到组 = {}
    for r in (变更 or {}):
        表名, 行, 列 = _坐标(r)
        for 名, 左, 顶, 右, 底, 组 in 范围表:
            if 名 == 表名 and 左 <= 列 <= 右 and 顶 <= 行 <= 底:
                a格到组[r] = 组
                break
    待搬 = set(a格到组.values())
    更新, 问题, 撤回, 证据 = [], [], set(), []
    for 组 in sorted(待搬):
        记录们 = a按组.get(组) or []
        规则 = next((x.get("标题") for x in 记录们 if isinstance(x.get("标题"), dict)), None)
        if not 规则:
            continue
        原文 = str(规则.get("原文") or "").strip()
        片段们 = 规则.get("日期") or []
        候选 = b格.get(组) or []
        if not 片段们:
            撤回.add(组)
            问题.append({"原因": "标题缺少日期片段与期间角色，不能生成写回指令；"
                            "该组标题与数值保持原样", "科目": 组[0], "表名": 组[1]})
            continue
        新文, 失败 = 原文, ""
        证据组 = []
        for 项 in 片段们:
            角色 = str(项.get("期间角色") or "")
            片段 = str(项.get("片段") or "")
            同角色 = [r for r in 候选
                   if str((r.get("维度值") or {}).get("期间角色") or "") == 角色]
            日期们 = {str((r.get("维度值") or {}).get("期间实际")) for r in 同角色
                   if str((r.get("维度值") or {}).get("期间实际") or "") not in ("", "未知")}
            if len(日期们) != 1:
                失败 = (f"B 侧“{角色}”日期缺失或冲突（{sorted(日期们)}）" if 日期们
                      else f"B 侧没有“{角色}”的期间实际日期")
                break
            日期值 = 解析日期(next(iter(日期们)))
            if 日期值 is None:
                失败 = f"B 侧“{角色}”日期无法解析：{next(iter(日期们))}"
                break
            新片段 = 格式化(日期值, 片段)
            if not 新片段:
                失败 = f"无法按原写法替换日期片段：{片段}"
                break
            if 片段 not in 新文:
                失败 = f"标题原文里找不到日期片段：{片段}"
                break
            新文 = 新文.replace(片段, 新片段, 1)
            证据组.append({"期间角色": 角色, "片段": 片段, "新片段": 新片段,
                        "日期": str(日期值), "坐标": sorted({r.get("格") for r in 同角色})})
        if 失败:
            撤回.add(组)
            问题.append({"原因": "标题未能按本次 B 日期生成，该组标题与数值保持原样",
                        "科目": 组[0], "表名": 组[1], "说明": 失败})
            continue
        if 新文 == 原文:
            continue                         # 日期没变，不需要动标题
        区域 = next((str(x.get("区域") or "") for x in 记录们 if x.get("区域")), "")
        现标题 = 本次标题(上下文, 区域)
        if not 区域:
            撤回.add(组)
            问题.append({"原因": "有标题结论但缺区域，不能定位段落；该组标题与数值保持原样",
                        "科目": 组[0], "表名": 组[1]})
            continue
        if 现标题 and 现标题 != 原文:
            撤回.add(组)
            问题.append({"原因": "本次标题原文与识别给出的原文不一致，不能写回；"
                            "该组标题与数值保持原样",
                        "科目": 组[0], "表名": 组[1],
                        "说明": f"本次“{现标题}”，识别“{原文}”"})
            continue
        # 桥接只认这三个字段：多一个都会拒绝，组信息留在标题证据里
        更新.append({"range_name": 区域, "old_text": 原文, "text": 新文})
        证据.append({"科目": 组[0], "表名": 组[1], "区域": 区域,
                  "原文": 原文, "新文": 新文, "日期": 证据组})
    return 更新, 问题, 撤回, 证据


def _范围(范围):
    from openpyxl.utils import range_boundaries
    try:
        return range_boundaries(str(范围 or "").replace("$", ""))
    except (ValueError, TypeError):
        return (0, 0, 0, 0)


def _坐标(格):
    from openpyxl.utils.cell import coordinate_to_tuple
    try:
        表名, 地址 = str(格).split("!", 1)
        行, 列 = coordinate_to_tuple(地址)
        return 表名, 行, 列
    except (IndexError, TypeError, ValueError):
        return "", 0, 0
