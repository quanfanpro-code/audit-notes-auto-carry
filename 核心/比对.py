# coding: utf-8
"""比对引擎：按“编号（＋序号）（＋期间角色）”配对，清尾只对名单且先证明 B 完整。

程序只机械执行识别给的判断，不推断性质、不猜方向：
- 认表用「科目＋表名」，两侧各自的表号（TA_008 / word附注!120）不参与比较。
- 期间角色只在金标准要求期间时才进配对键；不要求期间时，期间不同也照配，原值保留。
- 序号只对名单项有意义；固定项没有序号，按编号配，顺序不同也能对上。
- 汇总格再加分组，两个分组的小计不串组。
- 判重使用同一套完整键：不同编号各一格不算重复。
"""
_口径键 = ("币种", "单位", "期间跨度")
_完整性键 = ("第一遍覆盖", "第二遍归段", "第三遍逐行", "空名单已声明")


def _项(甲, 乙, 状态="未对应", 原因=None, 来源日期=None) -> dict:
    源 = 甲 or 乙 or {}
    return {"a格": (甲 or {}).get("格"), "b格": (乙 or {}).get("格"),
            "表组": 源.get("表组"), "槽位": 源.get("槽位"),
            "序号": 源.get("序号"), "状态": 状态, "原因": 原因,
            "采用来源日期": 来源日期}


def _组(记录) -> tuple:
    return (str(记录.get("科目") or ""), str(记录.get("表名") or ""))


def _结构索引(结构) -> dict:
    出 = {}
    for t in 结构 or []:
        出[_组(t)] = t
    return 出


def _行号(记录):
    行号 = 记录.get("行号")
    if isinstance(行号, int):
        return 行号
    坐标 = str(记录.get("格") or "")
    数字 = "".join(ch for ch in 坐标.split("!")[-1] if ch.isdigit())
    return int(数字) if 数字 else -1


def _列字母(记录) -> str:
    坐标 = str(记录.get("格") or "")
    字母 = "".join(ch for ch in 坐标.split("!")[-1] if ch.isalpha())
    return 字母.upper()


def _键(记录) -> tuple:
    """完整配对键。期间只在金标准要求时才进；汇总格再加分组。"""
    维度 = 记录.get("维度值") or {}
    键 = [str(记录.get("槽位") or "")]
    if 记录.get("序号") is not None:
        键.append(("序号", 记录.get("序号")))
    if 记录.get("期间必填"):
        键.append(("期间", str(维度.get("期间角色") or "未知")))
    if 记录.get("角色") == "汇总格":
        键.append(("分组", str(记录.get("分组") or "")))
    return tuple(键)


def _口径冲突(甲: list, 乙: list) -> bool:
    for 名 in _口径键:
        a = {str((r.get("维度值") or {}).get(名) or "").strip() for r in 甲}
        b = {str((r.get("维度值") or {}).get(名) or "").strip() for r in 乙}
        a = {v for v in a if v and v != "未知"}
        b = {v for v in b if v and v != "未知"}
        if a and b and not (a & b):
            return True                        # 两侧都有值且没有交集：口径根本不是一个
    return False


def _完整性好(表记录) -> bool:
    if not 表记录:
        return False
    完 = 表记录.get("完整性") or {}
    return all(bool(完.get(k)) for k in _完整性键)


def _需加项(甲: list, 乙: list) -> bool:
    """B 的行或列比 A 多，A 装不下：整表停，交人工加行加列。"""
    def 量(记录们):
        有效 = [r for r in 记录们 if r.get("角色") != "项标签"]
        return ({_行号(r) for r in 有效}, {_列字母(r) for r in 有效})
    a行, a列 = 量(甲)
    b行, b列 = 量(乙)
    return len(b行) > len(a行) or len(b列) > len(a列)


def _期间未知(记录) -> bool:
    """该编号要求期间、识别却写了“未知”：两侧都写未知也不能凑成一对。"""
    if not 记录.get("期间必填"):
        return False
    维度 = 记录.get("维度值") or {}
    return str(维度.get("期间角色") or "未知") == "未知"


def _清尾可用(记录, 乙: list) -> bool:
    """A 多出来的名单项能不能清：序号必须排在 B 同编号所有项之后。"""
    if 记录.get("序号") is None:
        return False
    同 = [x.get("序号") for x in 乙
         if x.get("槽位") == 记录.get("槽位") and x.get("序号") is not None]
    return bool(同) and 记录.get("序号") > max(同)


def _一组配对(甲: list, 乙: list, 清尾可: bool) -> list:
    清单 = []
    a桶, b桶 = {}, {}
    for r in 甲:
        if r.get("角色") == "项标签":
            清单.append(_项(r, None, "不参与", "项标签"))
            continue
        if not r.get("槽位"):
            清单.append(_项(r, None, "未对应", "槽位编号缺失"))
            continue
        if _期间未知(r):
            清单.append(_项(r, None, "未对应", "期间角色未知"))
            continue
        a桶.setdefault(_键(r), []).append(r)
    for r in 乙:
        if r.get("角色") == "项标签":
            continue                           # B 侧项标签由 A 侧那条带出清单，不重复记
        if not r.get("槽位"):
            清单.append(_项(None, r, "未对应", "槽位编号缺失"))
            continue
        if _期间未知(r):
            清单.append(_项(None, r, "未对应", "期间角色未知"))
            continue
        b桶.setdefault(_键(r), []).append(r)
    已用 = set()
    for 键, 记录们 in a桶.items():
        if len(记录们) > 1:
            for r in 记录们:
                清单.append(_项(r, None, "未对应", "同侧重复"))
            continue
        对手们 = b桶.get(键, [])
        if len(对手们) > 1:
            清单.append(_项(记录们[0], None, "未对应", "同侧重复"))
            continue
        甲记 = 记录们[0]
        if not 对手们:
            if 清尾可 and _清尾可用(甲记, 乙):
                清单.append(_项(甲记, None, "清尾", None))
            else:
                清单.append(_项(甲记, None, "未对应", "仅A侧有"))
            continue
        乙记 = 对手们[0]
        已用.add(id(乙记))
        清单.append(_项(甲记, 乙记, "对应", None,
                      (乙记.get("维度值") or {}).get("期间实际")))
    for 键, 记录们 in b桶.items():
        a条数 = len(a桶.get(键, []))
        for r in 记录们:
            if id(r) in 已用:
                continue
            if len(记录们) > 1 or a条数 > 1:
                清单.append(_项(None, r, "未对应", "同侧重复"))   # 整组判不了，交人工
            else:
                清单.append(_项(None, r, "未对应", "仅B侧有"))
    return 清单


def _整组未对应(甲: list, 乙: list, 原因: str) -> list:
    清单 = [_项(r, None, "未对应", 原因) for r in 甲 if r.get("角色") != "项标签"]
    清单 += [_项(None, r, "未对应", 原因) for r in 乙 if r.get("角色") != "项标签"]
    return 清单


def 建立对应(A合集: list, B合集: list, A结构=None, B结构=None) -> list:
    A结, B结 = _结构索引(A结构), _结构索引(B结构)
    a组, b组 = {}, {}
    for r in A合集:
        a组.setdefault(_组(r), []).append(r)
    for r in B合集:
        b组.setdefault(_组(r), []).append(r)
    清单 = []
    for 组键 in list(dict.fromkeys(list(a组) + list(b组) + list(A结) + list(B结))):
        甲, 乙 = a组.get(组键, []), b组.get(组键, [])
        if _口径冲突(甲, 乙):
            清单 += _整组未对应(甲, 乙, "口径冲突")
            continue
        if not 甲 or not 乙:
            清单 += _整组未对应(甲, 乙, "仅A侧有" if 甲 else "仅B侧有")
            continue
        if _需加项(甲, 乙):
            清单 += _整组未对应(甲, 乙, "需加项")
            continue
        清单 += _一组配对(甲, 乙, _完整性好(B结.get(组键)))
    return 清单
