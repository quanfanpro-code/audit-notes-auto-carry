# coding: utf-8
"""搬运：整项同进同退，项标签随项走，清尾整项写空；只改格子，保格式与命名区域。

“一项”由识别给的项方式决定：名单在行上就是同一明细行，名单在列上就是同一列。
物理行号只用来找要写哪一格，不再决定哪些格属于同一项——横排竖排都能搬。
一项里任何一格不能搬，这一项的名字和数字都不动——不许出现“名字换了、数没换”。
清尾只在整项都判了清尾时才生效，绝不盖过未对应、缺编号或读取失败。
"""
from pathlib import Path

from openpyxl import load_workbook

_读不出 = ("读取失败", "公式无保存结果")


def _项键(记录) -> tuple:
    return (str(记录.get("表组") or ""), str(记录.get("项方式") or "行"),
            str(记录.get("项位置")))


def _逐格(对应清单: list, b按格: dict) -> tuple:
    """没有 A 合集时的退化处理：每格自成一项，只保证单格政策。"""
    变更, 问题 = {}, []
    for 条 in 对应清单:
        if 条.get("状态") != "对应":
            continue
        乙 = b按格.get(条.get("b格"))
        if 乙 is None:
            问题.append({"a格": 条.get("a格"), "b格": 条.get("b格"), "原因": "读取失败"})
            continue
        状态 = 乙.get("原值状态")
        if 状态 in _读不出:
            问题.append({"a格": 条.get("a格"), "b格": 条.get("b格"), "原因": 状态})
            continue
        变更[条["a格"]] = "" if 状态 == "真实空白" else 乙.get("原值")
    return 变更, 问题


def 生成更新表(表A数据: dict, 对应清单: list, B合集: list, A合集=None):
    """返回（变更清单 {格: 新值}, 问题清单）。

    有 A 合集时按“一项”处理：整项可搬才写，含项标签；清尾项整行／整列写空。
    0 照搬；真实空白清空；读取失败或未对应则整项保留原值并列问题。
    """
    b按格 = {r.get("格"): r for r in B合集}
    b标签 = {_项键(r): r for r in B合集 if r.get("角色") == "项标签"}
    if A合集 is None:
        return _逐格(对应清单, b按格)

    a按格 = {r.get("格"): r for r in A合集}
    条按a格 = {条.get("a格"): 条 for 条 in 对应清单 if 条.get("a格")}

    # A 的哪一项对上 B 的哪一项：从已配上的格反查，不靠物理行号
    a项到b项 = {}
    for 条 in 对应清单:
        if 条.get("状态") != "对应":
            continue
        甲 = a按格.get(条.get("a格"))
        乙 = b按格.get(条.get("b格"))
        if 甲 is None or 乙 is None:
            continue
        a项到b项.setdefault(_项键(甲), _项键(乙))

    项态 = {}
    for 条 in 对应清单:
        甲 = a按格.get(条.get("a格"))
        if 甲 is None:
            continue
        态 = 项态.setdefault(_项键(甲), {"失败": None, "清尾": 0})
        if 甲.get("角色") == "项标签" and not (甲.get("不搬") and 甲.get("序号") is not None):
            continue                                  # 项标签跟随所在项，见下
        if 条.get("状态") == "清尾":
            态["清尾"] += 1
            continue
        if 条.get("状态") != "对应":
            态["失败"] = 态.get("失败") or (条.get("原因") or "未对应")
            continue
        乙 = b按格.get(条.get("b格"))
        if 乙 is None:
            态["失败"] = 态.get("失败") or "读取失败"
            continue
        if 乙.get("原值状态") in _读不出:
            态["失败"] = 态.get("失败") or str(乙.get("原值状态"))

    # 名单项：名字和数字必须一起换；B 侧标签取不到或读不出，整项不动
    标签来源 = {}
    for 甲 in a按格.values():
        if 甲.get("角色") != "项标签" or 甲.get("序号") is None:
            continue
        键 = _项键(甲)
        态 = 项态.setdefault(键, {"失败": None, "清尾": 0})
        if 态.get("失败") or 态.get("清尾"):
            continue
        b项 = a项到b项.get(键)
        乙 = b标签.get(b项) if b项 else None
        if 乙 is None:
            态["失败"] = "项标签取不到"
        elif 乙.get("不搬"):
            态["失败"] = "识别不搬，交人工"               # B 的列标题标了不搬，名字不可靠
        elif 乙.get("原值状态") in _读不出:
            态["失败"] = "项标签" + str(乙.get("原值状态"))
        else:
            标签来源[键] = 乙

    变更, 问题 = {}, []
    已记项 = set()
    清尾已记 = set()
    for 格, 甲 in a按格.items():
        键 = _项键(甲)
        态 = 项态.get(键) or {}
        if 态.get("失败"):
            if 甲.get("角色") != "项标签":
                问题.append({"a格": 格, "b格": None, "原因": 态.get("失败")})
            elif 键 not in 已记项:
                已记项.add(键)
                问题.append({"a格": 格, "b格": None, "原因": 态.get("失败"),
                           "说明": "整项不动"})
            continue
        if 态.get("清尾"):
            变更[格] = ""                             # 清尾：标签与数字一起空
            if 键 not in 清尾已记:                     # 每一项都要在问题清单里交代一次
                清尾已记.add(键)
                问题.append({"a格": 格, "b格": None,
                           "原因": "B无此项，已清空，请人工确认是否删行或删列"})
            continue
        if 甲.get("角色") == "项标签":
            if 甲.get("序号") is None:
                continue                              # 固定项标签是身份，不搬
            乙 = 标签来源.get(键)
            if 乙 is None:
                continue
            变更[格] = "" if 乙.get("原值状态") == "真实空白" else 乙.get("原值")
            continue
        条 = 条按a格.get(格) or {}
        乙 = b按格.get(条.get("b格"))
        if 乙 is None:
            问题.append({"a格": 格, "b格": 条.get("b格"), "原因": "读取失败"})
            continue
        变更[格] = "" if 乙.get("原值状态") == "真实空白" else 乙.get("原值")
    return 变更, 问题


def 保存更新表(提取表A路径, 变更清单: dict, 输出路径) -> None:
    源 = Path(提取表A路径)
    出 = Path(输出路径)
    if 源.resolve() == 出.resolve():
        raise ValueError("输出不能覆盖提取产物")
    wb = load_workbook(源)
    try:
        for 格, 值 in 变更清单.items():
            表名, 坐标 = 格.split("!", 1)
            wb[表名][坐标] = 值
        wb.save(出)
    finally:
        wb.close()
