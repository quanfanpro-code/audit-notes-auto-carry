# coding: utf-8
"""金标准三层索引：科目 → 表 → 槽位。

识别材料不放金标准条目；只在第三遍给该表候选槽位，比对按编号配对。
程序不做语义判断、不做字符串替换、不据行列路径反推槽位名（实测 1.3% 不等）。
"""
import json
from pathlib import Path


def 加载金标准(主文件路径) -> list[dict]:
    lines = Path(主文件路径).read_text(encoding="utf-8-sig").splitlines()
    return [json.loads(line) for line in lines if line.strip()]


def 建索引(金标准: list) -> dict:
    """三层嵌套：{科目: {"序", "表": {表名: {"序", "槽位": [槽位记录]}}}}，另存全部编号集合。"""
    索引 = {"科目": {}, "编号": set()}
    for 记录 in 金标准:
        note = 记录.get("note") or {}
        表 = 记录.get("table") or {}
        科目名, 表名 = note.get("name"), 表.get("name")
        if not 科目名 or not 表名:
            continue
        科 = 索引["科目"].setdefault(科目名, {"序": note.get("display_order") or 0, "表": {}})
        表项 = 科["表"].setdefault(表名, {"序": 表.get("display_order") or 0, "槽位": []})
        编号 = 记录.get("id")
        if 编号:
            索引["编号"].add(编号)
        表项["槽位"].append({"id": 编号,
                          "槽位名": (记录.get("slot") or {}).get("name"),
                          "行路径": list(记录.get("row_path") or []),
                          "列路径": list(记录.get("column_path") or []),
                          "别名": 记录.get("aliases")})
    return 索引


def 科目清单(索引: dict) -> list:
    科目们 = 索引["科目"]
    return sorted(科目们, key=lambda 名: (科目们[名]["序"], 名))


def 表清单(索引: dict, 科目: str) -> list:
    表们 = (索引["科目"].get(科目) or {}).get("表") or {}
    return sorted(表们, key=lambda 名: (表们[名]["序"], 名))


def 候选槽位(索引: dict, 科目: str, 表名: str) -> list:
    表项 = ((索引["科目"].get(科目) or {}).get("表") or {}).get(表名) or {}
    return [dict(槽) for 槽 in 表项.get("槽位", [])]


def 编号存在(索引: dict, 编号) -> bool:
    return bool(编号) and 编号 in 索引["编号"]
