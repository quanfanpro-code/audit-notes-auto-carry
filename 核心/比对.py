# coding: utf-8
"""比对引擎：按“编号（＋序号）（＋期间角色）”配对，清尾只对名单且先证明 B 完整。

程序只机械执行识别给的判断，不推断性质、不猜方向：
- 认表用「科目＋表名」，两侧各自的表号（TA_008 / word附注!120）不参与比较。
- 期间角色只在金标准要求期间时才进配对键；不要求期间时，期间不同也照配，原值保留。
- 序号只对名单项有意义；固定项没有序号，按编号配，顺序不同也能对上。
- 汇总格再加分组，两个分组的小计不串组。
- 项数按完整名单清单算（含识别失败的占位项），不按物理行列数：固定项、合计各按各的规则办。
- 清尾只清名单项，且只清尾部；尾巴按识别给的序号认，不按列字母或行号重排。
- 判重使用同一套完整键，同侧撞键就整表停：不同编号各一格不算重复。
"""
from openpyxl.utils import column_index_from_string

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


def _项键(记录) -> tuple:
    """一项 = 同一张业务表里的同一条明细行，或者同一列（横向名单）。"""
    return (str(记录.get("表组") or ""), str(记录.get("项方式") or "行"),
            str(记录.get("项位置")))


def _按项(记录们: list) -> dict:
    出 = {}
    for r in 记录们:
        出.setdefault(_项键(r), []).append(r)
    return 出


def _有序项(项表: dict) -> list:
    """按识别给的序号排序；没有序号的按物理位置。

    列上成名单时位置是列字母，不能拿字符串比大小——"AA" 会排到 "Z" 前面，
    尾巴就认错了项；这里一律换算成列序号。
    """
    def 排序键(键):
        序号 = next((r.get("序号") for r in 项表[键] if r.get("序号") is not None), None)
        if 序号 is not None:
            return (键[1], 0, int(序号), 0)
        位置 = str(键[2])
        if 位置.isdigit():
            return (键[1], 1, int(位置), 0)
        try:
            return (键[1], 1, 0, column_index_from_string(位置.upper()))
        except ValueError:
            return (键[1], 1, 0, 0)
    return sorted(项表, key=排序键)


def _是名单项(记录们: list) -> bool:
    return any(r.get("序号") is not None for r in 记录们)


def _序号(项表: dict, 键) -> int | None:
    return next((r.get("序号") for r in 项表[键] if r.get("序号") is not None), None)


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


def _期间未知(记录) -> bool:
    """该编号要求期间、识别却写了“未知”：两侧都写未知也不能凑成一对。"""
    if not 记录.get("期间必填"):
        return False
    维度 = 记录.get("维度值") or {}
    return str(维度.get("期间角色") or "未知") == "未知"


def _整组未对应(甲: list, 乙: list, 原因: str) -> list:
    """整组判不了时也要逐格进清单；项标签另有标记，不算“漏了”。"""
    清单 = []
    for r in 甲:
        if r.get("角色") == "项标签":
            清单.append(_项(r, None, "不参与", "项标签"))
        else:
            清单.append(_项(r, None, "未对应", 原因))
    for r in 乙:
        if r.get("角色") != "项标签":
            清单.append(_项(None, r, "未对应", 原因))
    return 清单


def _重复键(记录们: list) -> bool:
    """同一侧是否出现两组以上的完整配对键——撞键就整表停，不按顺序硬配。

    清尾项也要查：尾项撞键同样是这张表没法判，不能因为“反正要清掉”就放过。
    """
    见 = set()
    for r in 记录们:
        if r.get("角色") == "项标签" or r.get("不搬"):
            continue
        if not r.get("槽位") or _期间未知(r):
            continue                            # 缺编号、期间未知另有专门处理，不算这里
        键 = _键(r)
        if 键 in 见:
            return True
        见.add(键)
    return False


def _项有数据(记录们: list) -> bool:
    """仅确认至少两期均为数值0时不搬；空白、失败和缺号仍保留项位。"""
    数据 = [r for r in 记录们 if r.get("角色") != "项标签"]
    if not 数据 or any(r.get("不搬") or not r.get("槽位") or _期间未知(r)
                     or r.get("原值状态") != "值" for r in 数据):
        return True
    期间 = {str((r.get("维度值") or {}).get("期间角色") or "未知") for r in 数据}
    if len(期间) < 2 or "未知" in 期间:
        return True
    return not all(isinstance(r.get("原值"), (int, float))
                   and not isinstance(r.get("原值"), bool) and r["原值"] == 0 for r in 数据)


def _空项(表: dict, 记录们: list) -> dict:
    """确认两期全零的项只记账，不搬，仍保留原序号。"""
    名 = next((r.get("原值") for r in 记录们 if r.get("角色") == "项标签"), None)
    return _项(None, 记录们[0] if 记录们 else None, "未对应",
               f"B空项：{名 if 名 not in (None, '') else '（无名称）'}；"
               "已确认两期均为数值0，保留原序号、不搬")


def _超额项(表: dict, 记录们: list) -> dict:
    """B 有数据、但 A 已经没有行可放的一项。

    以 B 为准：搬得下的按序先搬，多出来的不硬塞、也不整表停，而是在这里记账，
    把名称和金额写进原因，由人工在 A 侧加行后再跑。
    """
    名 = next((r.get("原值") for r in 记录们 if r.get("角色") == "项标签"), None)
    金额 = [f"{'/'.join(r.get('列路径') or [])}={r.get('原值')}"
         for r in 记录们
         if r.get("角色") != "项标签" and r.get("原值") not in (None, "")]
    说明 = (f"B多出项：{名 if 名 not in (None, '') else '（无名称）'}；"
            + ("；".join(金额[:4]) if 金额 else "金额全空")
            + "；A 无对应行，需在 A 加行，交人工")
    return _项(None, 记录们[0] if 记录们 else None, "未对应", 说明)


def _项有失败(记录们: list) -> bool:
    """这一项里有格已经交人工（不搬、缺编号、期间未知）：整项保留，不能拿来清尾。"""
    for r in 记录们:
        if r.get("不搬"):
            return True
        if r.get("角色") == "项标签":
            continue                            # 项标签本来就没有编号
        if not r.get("槽位") or _期间未知(r):
            return True
    return False


def _配对(甲: list, 乙: list, 清尾项: set) -> list:
    清单 = []
    a桶, b桶 = {}, {}
    for r in 甲:
        if _项键(r) in 清尾项:
            清单.append(_项(r, None, "清尾", None))   # 整项清空：标签和数字一起
            continue
        if r.get("角色") == "项标签":
            if r.get("不搬") and r.get("序号") is not None:
                清单.append(_项(r, None, "未对应", "识别不搬，交人工"))   # 名字不可靠，数也不搬
            else:
                清单.append(_项(r, None, "不参与", "项标签"))
            continue
        if r.get("不搬"):
            清单.append(_项(r, None, "未对应", "识别不搬，交人工"))
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
            if r.get("不搬") and r.get("序号") is not None:
                清单.append(_项(None, r, "未对应", "识别不搬，交人工"))
            continue                           # B 侧项标签由 A 侧那条带出清单，不重复记
        if r.get("不搬"):
            清单.append(_项(None, r, "未对应", "识别不搬，交人工"))
            continue
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
        A该表, B该表 = A结.get(组键) or {}, B结.get(组键) or {}
        if _口径冲突(甲, 乙):
            清单 += _整组未对应(甲, 乙, "口径冲突")
            continue
        if A该表.get("例外未执行") or B该表.get("例外未执行"):
            清单 += _整组未对应(甲, 乙, "例外未执行")     # 例外执行不了，整表不许自动搬
            continue
        if A该表.get("项标签缺失") or B该表.get("项标签缺失"):
            清单 += _整组未对应(甲, 乙, "列标签缺失")     # 名字写不进去，就不搬数
            continue
        if not 甲:
            清单 += _整组未对应(甲, 乙, "仅B侧有")
            continue
        if not 乙 and not _完整性好(B该表):
            清单 += _整组未对应(甲, 乙, "仅A侧有")
            continue

        a项, b项 = _按项(甲), _按项(乙)
        a名单 = [k for k in _有序项(a项) if _是名单项(a项[k])]
        b全部 = [k for k in _有序项(b项) if _是名单项(b项[k])]
        b名单 = [k for k in b全部 if _项有数据(b项[k])]
        b空项 = [k for k in b全部 if k not in set(b名单)]
        if _重复键(甲) or _重复键(乙):
            清单 += _整组未对应(甲, 乙, "同侧重复")        # 撞键：整张表交人工，一项都不许动
            continue
        a末 = max((_序号(a项, k) for k in a名单), default=0)
        # 全零项仍占原序号；与清尾采用同一位置口径，不用过滤后的数量判断容量。
        超额项 = [k for k in b名单 if _序号(b项, k) > a末]
        踢出 = set(b空项) | set(超额项)
        if 踢出:
            乙 = [r for r in 乙 if _项键(r) not in 踢出]
        清尾项 = set()
        if _完整性好(B该表):
            明确零项 = bool(b名单) or bool(B该表.get("空名单"))
            if 明确零项:
                # 尾巴按 B 的最大序号认，不按项数：B 中间有全零项时项数会少，
                # 按数量截会把后面仍在 B 里的非零项误清（全零项也占着序号位）
                b末 = max((_序号(b项, k) for k in b全部), default=None)
                清尾项 = {k for k in a名单
                        if (b末 is None or _序号(a项, k) > b末) and not _项有失败(a项[k])}
        清单 += _配对(甲, 乙, 清尾项)
        for 键 in 超额项:
            清单.append(_超额项(B该表, b项[键]))
        for 键 in b空项:
            清单.append(_空项(B该表, b项[键]))
    return 清单
