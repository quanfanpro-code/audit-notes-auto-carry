# coding: utf-8
"""结构操作：只执行给定的操作清单。

插行、删行、换列由识别或操作人给出结论，程序只校验字段、原样交给桥接。
自己比较两份文件去推断“这里插了一行”“这一列换到了别处”属于判断，已整体移除。
"""
import json
from pathlib import Path

_轴 = {"row", "column"}
_动作 = {"insert", "delete"}


def 读结构操作(来源) -> list:
    """来源可以是 json 文件路径，也可以是已解析的列表；格式不对直接报错，不猜。"""
    原始 = (json.loads(Path(来源).read_text(encoding="utf-8"))
            if isinstance(来源, (str, Path)) else 来源)
    if not isinstance(原始, list):
        raise RuntimeError("结构操作必须是数组")
    出 = []
    for 项 in 原始:
        if not isinstance(项, dict):
            raise RuntimeError(f"结构操作不是对象：{项}")
        区域 = str(项.get("range_name") or "").strip()
        轴 = str(项.get("axis") or "").strip()
        动作 = str(项.get("action") or "").strip()
        位置, 数 = 项.get("index"), 项.get("count")
        if not 区域:
            raise RuntimeError(f"结构操作缺区域：{项}")
        if 轴 not in _轴:
            raise RuntimeError(f"结构操作轴向不认识：{项}")
        if 动作 not in _动作:
            raise RuntimeError(f"结构操作动作不认识：{项}")
        if not isinstance(位置, int) or 位置 < 0:
            raise RuntimeError(f"结构操作位置不合规：{项}")
        if not isinstance(数, int) or 数 < 1:
            raise RuntimeError(f"结构操作行数不合规：{项}")
        记 = {"range_name": 区域, "axis": 轴, "action": 动作, "index": 位置, "count": 数}
        if isinstance(项.get("template_index"), int):
            记["template_index"] = 项["template_index"]
        出.append(记)
    return 出
