# coding: utf-8
"""写回：期间标题更新构造（完整原标题＋所属表）、桥接 update 封装。"""
from .桥接 import 更新 as 调用桥接更新


def _中文日期(文本):
    if not 文本 or "-" not in str(文本):
        return str(文本) if 文本 else ""
    y, m, d = str(文本).split("-")[:3]
    return f"{y}年{int(m)}月{int(d)}日"


def 构造标题更新(对应清单: list, A合集: list, 上下文: dict) -> list:
    表们 = (上下文 or {}).get("tables", [])
    a按格 = {r.get("格"): r for r in A合集}
    更新 = []
    for 项 in 对应清单:
        if 项.get("状态") != "对应":
            continue
        记录 = a按格.get(项.get("a格"))
        if not 记录 or "标题" not in str(记录.get("槽位名", "")):
            continue
        旧日期 = _中文日期((记录.get("维度值") or {}).get("期间实际"))
        新日期 = _中文日期(项.get("采用来源日期"))
        if not 旧日期 or not 新日期 or 旧日期 == 新日期:
            continue
        表 = next((t for t in 表们 if 旧日期 in str(t.get("chapter_context", ""))), None)
        if not 表:
            continue
        段 = next((p for p in str(表["chapter_context"]).splitlines() if 旧日期 in p),
                 str(表["chapter_context"]))
        更新.append({"range_name": 表.get("range_name"), "old_text": 段,
                   "text": 段.replace(旧日期, 新日期)})
    return 更新


def 写回Word(带连接word, 更新表xlsx, output, operations=None, caption_updates=None) -> dict:
    return 调用桥接更新(str(带连接word), str(更新表xlsx), str(output),
                    operations or [], caption_updates or [])
