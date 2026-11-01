# coding: utf-8
"""搬运：生成变更（0/真实空白/失败政策）并基于提取A副本只改格保存（保格式与命名区域）。"""
from pathlib import Path

from openpyxl import load_workbook


def 生成更新表(表A数据: dict, 对应清单: list, B合集: list):
    """返回（变更清单 {格: 新值}, 问题清单）。0 照搬；真实空白清空；读取失败/公式无保存结果保留原值。"""
    b按格 = {r.get("格"): r for r in B合集}
    变更 = {}
    问题 = []
    for 项 in 对应清单:
        if 项.get("状态") != "对应":
            continue
        b = b按格.get(项.get("b格"))
        if b is None:
            问题.append({"a格": 项.get("a格"), "b格": 项.get("b格"), "原因": "读取失败"})
            continue
        状态 = b.get("原值状态")
        if 状态 in ("读取失败", "公式无保存结果"):
            问题.append({"a格": 项.get("a格"), "b格": 项.get("b格"), "原因": 状态})
            continue
        变更[项["a格"]] = "" if 状态 == "真实空白" else b.get("原值")
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
