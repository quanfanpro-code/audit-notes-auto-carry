# coding: utf-8
"""写回：只执行给定的标题更新与桥接 update。

改哪一段、改成什么，由识别给出；程序不再自己找段落、自己换年份。
“没给结论”就是不动标题，不推断、不补救。
"""
from .桥接 import 更新 as 调用桥接更新


def 读标题更新(结构记录们: list) -> list:
    """把识别给的「标题：{原文, 新文}」转成桥接要的形状；没给就不动。"""
    出 = []
    for 表 in 结构记录们 or []:
        标题 = 表.get("标题")
        if not 标题:
            continue
        if not isinstance(标题, dict):
            raise RuntimeError(f"标题结论格式不对：{表.get('科目')}／{表.get('表名')}")
        原 = str(标题.get("原文") or "").strip()
        新 = str(标题.get("新文") or "").strip()
        if not 原 or not 新:
            raise RuntimeError(f"标题结论缺原文或新文：{表.get('科目')}／{表.get('表名')}")
        if 原 == 新:
            continue
        区域 = str(表.get("区域") or "").strip()
        if not 区域:
            raise RuntimeError(f"给了标题结论但缺区域：{表.get('科目')}／{表.get('表名')}")
        出.append({"range_name": 区域, "old_text": 原, "text": 新})
    return 出


def 写回Word(带连接word, 更新表xlsx, output, operations=None, caption_updates=None) -> dict:
    return 调用桥接更新(str(带连接word), str(更新表xlsx), str(output),
                    operations or [], caption_updates or [])
