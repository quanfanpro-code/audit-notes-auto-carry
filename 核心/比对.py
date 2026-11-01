# coding: utf-8
"""比对引擎：只认金标准槽位编号。写法差异在第三遍挂编号时已吸收，金额不参与配对。"""


def _合法(槽位) -> bool:
    return isinstance(槽位, str) and 槽位.startswith("C-N")


def _项(a格, b格, 槽位, 状态="未对应", 原因=None, 来源日期=None) -> dict:
    return {"a格": a格, "b格": b格, "槽位": 槽位, "状态": 状态,
            "原因": 原因, "采用来源日期": 来源日期}


def 建立对应(A合集: list, B合集: list) -> list:
    a组, b组 = {}, {}
    for 序, 记录 in enumerate(A合集):
        if _合法(记录.get("槽位")):
            a组.setdefault(记录["槽位"], []).append(序)
    for 序, 记录 in enumerate(B合集):
        if _合法(记录.get("槽位")):
            b组.setdefault(记录["槽位"], []).append(序)
    清单, 已用B = [], set()
    for 序, 记录 in enumerate(A合集):
        槽位 = 记录.get("槽位")
        if not _合法(槽位):
            清单.append(_项(记录.get("格"), None, 槽位, 原因="槽位编号缺失"))
            continue
        a数, b位 = len(a组[槽位]), b组.get(槽位, [])
        if a数 > 1:
            清单.append(_项(记录.get("格"), None, 槽位, 原因="A侧重复"))
        elif len(b位) > 1:
            清单.append(_项(记录.get("格"), None, 槽位, 原因="B侧重复"))
        elif b位:
            乙 = B合集[b位[0]]
            已用B.add(b位[0])
            清单.append(_项(记录.get("格"), 乙.get("格"), 槽位, "对应",
                          None, (乙.get("维度值") or {}).get("期间实际")))
        else:
            清单.append(_项(记录.get("格"), None, 槽位, 原因="仅A侧有"))
    for 序, 记录 in enumerate(B合集):
        if 序 in 已用B:
            continue
        槽位 = 记录.get("槽位")
        if not _合法(槽位):
            清单.append(_项(None, 记录.get("格"), 槽位, 原因="槽位编号缺失"))
            continue
        a数 = len(a组.get(槽位, []))
        原因 = "仅B侧有" if a数 == 0 else ("A侧重复" if a数 > 1 else "B侧重复")
        清单.append(_项(None, 记录.get("格"), 槽位, 原因=原因))
    return 清单

