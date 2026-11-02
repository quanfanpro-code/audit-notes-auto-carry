# coding: utf-8
"""写回：期间标题更新构造（按该表真实原句与自己的命名区域）、桥接 update 封装。"""
import re

from .桥接 import 更新 as 调用桥接更新

_日期式 = re.compile(r"\d{4}年\d{1,2}月\d{1,2}日")


def _中文日期(文本):
    if not 文本 or "-" not in str(文本):
        return str(文本) if 文本 else ""
    y, m, d = str(文本).split("-")[:3]
    return f"{y}年{int(m)}月{int(d)}日"


def _换日期(原: str, 替换: dict) -> str:
    """按原句一次换完：已换出来的新日期不再参加下一轮，免得 2023→2024 又被改成 2025。"""
    映射 = {_中文日期(旧): _中文日期(新) for 旧, 新 in 替换.items()}
    return _日期式.sub(lambda m: 映射.get(m.group(0), m.group(0)), 原)


def _段落们(表: dict) -> list:
    """把上下文里这段表的原文摊成逐段文字。

    真实提取产出的 chapter_context 是段落列表，不是一整个字符串；
    直接 str(列表) 会带出方括号和引号，桥接就找不到原句。
    同一句标题常同时挂在 title、preceding_blocks、chapter_context 上，那是一个段落被
    多处引用，只算一段；同一字段里真的重复两次的，仍然算两段。
    """
    if 表.get("title"):
        标题 = [str(表["title"])]
    else:
        标题 = []
    前文 = []
    for b in 表.get("preceding_blocks") or []:
        if isinstance(b, dict) and b.get("text"):
            前文.append(str(b["text"]))
        elif isinstance(b, str) and b:
            前文.append(b)
    单位 = [str(u) for u in (表.get("unit_text_candidates") or []) if u]
    章 = 表.get("chapter_context")
    if isinstance(章, (list, tuple)):
        章节 = [str(x) for x in 章 if x]
    elif 章:
        章节 = [x for x in str(章).splitlines() if x]
    else:
        章节 = []
    出, 见过 = [], set()
    for 字段 in (标题, 前文, 单位, 章节):
        本字段 = set()
        for 文 in 字段:
            if not 文 or 文 == "None":
                continue
            if 文 not in 见过 or 文 in 本字段:
                出.append(文)
            见过.add(文)
            本字段.add(文)
    return 出


def _换期映射(对应清单: list, A合集: list) -> dict:
    """按「科目＋表名」收集（旧日期→新日期），期间角色各自对应。

    同一角色出现两个不同日期，说明这张表本身有多套期间且识别没分清，不猜。
    """
    a按格 = {r.get("格"): r for r in A合集}
    出 = {}
    for 项 in 对应清单:
        if 项.get("状态") != "对应" or not 项.get("采用来源日期"):
            continue
        甲 = a按格.get(项.get("a格")) or {}
        维度 = 甲.get("维度值") or {}
        旧 = str(维度.get("期间实际") or "").strip()
        新 = str(项.get("采用来源日期") or "").strip()
        角色 = str(维度.get("期间角色") or "未知")
        if not 旧 or 旧 == "未知" or not 新:
            continue
        键 = (甲.get("科目") or "", 甲.get("表名") or "")
        出.setdefault(键, {}).setdefault(角色, set()).add((旧, 新))
    return 出


def 构造标题更新(对应清单: list, A合集: list, 上下文: dict, A结构=None) -> tuple:
    """返回（标题更新列表, 定位失败的表清单）。

    需不需要改标题，看这张表两侧的期间日期是否不同，不看槽位名里有没有“标题”两个字。
    定位顺序：先按这张表自己的命名区域找上下文段落，找不到唯一原句就列为失败——
    调用方必须整表不换数，不许留下“旧标题 + 新数字”。
    """
    换期 = _换期映射(对应清单, A合集)
    区域表 = {(t.get("科目") or "", t.get("表名") or ""): t.get("区域")
           for t in (A结构 or []) if t.get("区域")}
    表们 = (上下文 or {}).get("tables", [])
    更新, 失败 = [], []
    for 键, 角色表 in 换期.items():
        替换, 坏 = {}, None
        for 角色, 对 in 角色表.items():
            if len(对) != 1:
                坏 = f"期间「{角色}」的日期不唯一"
                break
            旧, 新 = next(iter(对))
            if 旧 == 新:
                continue
            if 旧 in 替换 and 替换[旧] != 新:
                坏 = "同一旧日期对应多个新日期"
                break
            替换[旧] = 新
        if 坏:
            失败.append({"科目": 键[0], "表名": 键[1], "原因": 坏})
            continue
        if not 替换:
            continue                                   # 两边同一期，标题不用改
        区域 = 区域表.get(键)
        表 = next((t for t in 表们 if t.get("range_name") == 区域), None) if 区域 else None
        if 表 is None:
            候选 = [t for t in 表们
                  if any(_中文日期(旧) in 段 for 旧 in 替换 for 段 in _段落们(t))]
            if len(候选) != 1:
                失败.append({"科目": 键[0], "表名": 键[1],
                           "原因": "原句找不到" if not 候选 else "原句不唯一"})
                continue
            表 = 候选[0]
        段们 = _段落们(表)
        命中 = [段 for 段 in 段们 if any(_中文日期(旧) in 段 for 旧 in 替换)]
        if not 命中:
            失败.append({"科目": 键[0], "表名": 键[1], "原因": "原句找不到"})
            continue
        if len(命中) > 1:
            失败.append({"科目": 键[0], "表名": 键[1], "原因": "原句不唯一"})
            continue
        原 = 命中[0]
        新文 = _换日期(原, 替换)
        if 新文 == 原:
            失败.append({"科目": 键[0], "表名": 键[1], "原因": "原句没变"})
            continue
        更新.append({"range_name": 表.get("range_name"), "old_text": 原, "text": 新文})
    return 更新, 失败


def 写回Word(带连接word, 更新表xlsx, output, operations=None, caption_updates=None) -> dict:
    return 调用桥接更新(str(带连接word), str(更新表xlsx), str(output),
                    operations or [], caption_updates or [])
