# coding: utf-8
"""搬运：整项同进同退，项标签随项走，清尾整行写空；只改格子，保格式与命名区域。

“一项”= 同一张表里的同一明细行（科目＋表名＋行号）。
一项里任何一格不能搬，这一项的名称和数字都不动——不许出现“名字换了、数没换”。
"""
from pathlib import Path

from openpyxl import load_workbook


def _行号(格) -> int:
    数字 = "".join(ch for ch in str(格 or "").split("!")[-1] if ch.isdigit())
    return int(数字) if 数字 else -1


def _项键(记录, 格) -> tuple:
    return (记录.get("科目") or "", 记录.get("表名") or "",
            记录.get("行号", _行号(格)))


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
        if 状态 in ("读取失败", "公式无保存结果"):
            问题.append({"a格": 条.get("a格"), "b格": 条.get("b格"), "原因": 状态})
            continue
        变更[条["a格"]] = "" if 状态 == "真实空白" else 乙.get("原值")
    return 变更, 问题


def 生成更新表(表A数据: dict, 对应清单: list, B合集: list, A合集=None):
    """返回（变更清单 {格: 新值}, 问题清单）。

    有 A 合集时按“一项”处理：整项可搬才写，含项标签；清尾项整行写空。
    0 照搬；真实空白清空；读取失败或未对应则整项保留原值并列问题。
    """
    b按格 = {r.get("格"): r for r in B合集}
    if A合集 is None:
        return _逐格(对应清单, b按格)

    a按格 = {r.get("格"): r for r in A合集}
    条按a格 = {条.get("a格"): 条 for 条 in 对应清单 if 条.get("a格")}

    项态 = {}
    for 条 in 对应清单:
        甲 = a按格.get(条.get("a格")) or {}
        键 = _项键(甲, 条.get("a格"))
        态 = 项态.setdefault(键, {"可搬": True, "清尾": False, "原因": None})
        if 条.get("状态") == "清尾":
            态["清尾"] = True
            continue
        if 甲.get("角色") == "项标签":
            continue                                  # 项标签跟随所在项，不单独决定
        if 条.get("状态") != "对应":
            态["可搬"] = False
            态["原因"] = 态["原因"] or 条.get("原因") or "未对应"
            continue
        乙 = b按格.get(条.get("b格"))
        if 乙 is None or 乙.get("原值状态") in ("读取失败", "公式无保存结果"):
            态["可搬"] = False
            态["原因"] = 态["原因"] or (乙 or {}).get("原值状态") or "读取失败"

    # A 的某一项对应到 B 的哪一行：项标签要从 B 同项取文字
    a项到b行 = {}
    for 条 in 对应清单:
        if 条.get("状态") != "对应":
            continue
        甲 = a按格.get(条.get("a格")) or {}
        乙 = b按格.get(条.get("b格")) or {}
        if 甲.get("角色") != "项标签" and 乙.get("行号") is not None:
            a项到b行[_项键(甲, 条.get("a格"))] = 乙["行号"]

    变更, 问题 = {}, []
    for 格, 甲 in a按格.items():
        键 = _项键(甲, 格)
        态 = 项态.get(键) or {}
        if 态.get("清尾"):
            变更[格] = ""                             # 清尾：标签与数字一起空
            continue
        if not 态.get("可搬"):
            if 甲.get("角色") != "项标签":
                问题.append({"a格": 格, "b格": None, "原因": 态.get("原因") or "未对应"})
            continue
        if 甲.get("角色") == "项标签":
            if 甲.get("序号") is None:
                continue                              # 固定项标签是身份，不搬
            b行 = a项到b行.get(键)
            if b行 is None:
                continue
            乙 = next((r for r in B合集
                       if (r.get("科目"), r.get("表名")) == 键[:2]
                       and r.get("行号") == b行 and r.get("角色") == "项标签"), None)
            if 乙 is None:
                continue
            变更[格] = "" if 乙.get("原值状态") == "真实空白" else 乙.get("原值")
            continue
        乙 = b按格.get((条按a格.get(格) or {}).get("b格"))
        if 乙 is None:
            问题.append({"a格": 格, "b格": None, "原因": "读取失败"})
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
