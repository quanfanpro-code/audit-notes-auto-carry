# coding: utf-8
"""写回：期间标题更新构造（按该表真实原句）、桥接 update 封装。"""
from .桥接 import 更新 as 调用桥接更新


def _中文日期(文本):
    if not 文本 or "-" not in str(文本):
        return str(文本) if 文本 else ""
    y, m, d = str(文本).split("-")[:3]
    return f"{y}年{int(m)}月{int(d)}日"


def _表日期(合集: list) -> dict:
    """按「科目＋表名」收集该表用到的期间日期（未知不算）。"""
    出 = {}
    for r in 合集:
        d = str((r.get("维度值") or {}).get("期间实际") or "").strip()
        if not d or d == "未知":
            continue
        出.setdefault((r.get("科目") or "", r.get("表名") or ""), set()).add(d)
    return 出


def 构造标题更新(对应清单: list, A合集: list, 上下文: dict) -> tuple:
    """返回（标题更新列表, 定位失败的表清单）。

    需不需要改标题，只看这张表两侧的期间日期是否不同，不看槽位名里有没有“标题”两个字。
    该表上下文里找不到唯一一句旧日期时，这张表列为失败——调用方必须整表不换数，
    不许留下“旧标题 + 新数字”。
    """
    a日期 = _表日期(A合集)
    b日期 = {}
    a按格 = {r.get("格"): r for r in A合集}
    for 项 in 对应清单:
        if 项.get("状态") != "对应" or not 项.get("采用来源日期"):
            continue
        甲 = a按格.get(项.get("a格")) or {}
        键 = (甲.get("科目") or "", 甲.get("表名") or "")
        b日期.setdefault(键, set()).add(str(项["采用来源日期"]).strip())
    表们 = (上下文 or {}).get("tables", [])
    更新, 失败 = [], []
    for 键 in a日期:
        旧组, 新组 = a日期.get(键, set()), b日期.get(键, set())
        if len(旧组) != 1 or len(新组) != 1:
            continue                                   # 日期不唯一，判断不了，不动标题
        旧, 新 = next(iter(旧组)), next(iter(新组))
        if 旧 == 新:
            continue                                   # 两边同一期，标题不用改
        旧文, 新文 = _中文日期(旧), _中文日期(新)
        候选 = [t for t in 表们 if 旧文 in str(t.get("chapter_context", ""))]
        if len(候选) != 1:
            失败.append({"科目": 键[0], "表名": 键[1],
                       "原因": "原句找不到" if not 候选 else "原句不唯一",
                       "旧日期": 旧, "新日期": 新})
            continue
        表 = 候选[0]
        章 = str(表.get("chapter_context", ""))
        段 = next((p for p in 章.splitlines() if 旧文 in p), 章)
        更新.append({"range_name": 表.get("range_name"), "old_text": 段,
                   "text": 段.replace(旧文, 新文)})
    return 更新, 失败


def 写回Word(带连接word, 更新表xlsx, output, operations=None, caption_updates=None) -> dict:
    return 调用桥接更新(str(带连接word), str(更新表xlsx), str(output),
                    operations or [], caption_updates or [])
