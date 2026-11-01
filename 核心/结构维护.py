# coding: utf-8
"""结构维护：先定位变更落在哪个命名区域，再算该区域内的相对行号（桥接语义）。

一张工作表里可能装着几十张表，所以不能用工作表首行当基准。
区域外、跨区域、多处增删、合并结构变化一律交人工，不猜。
"""
from openpyxl.utils import range_boundaries

from .表格 import 读快照


def _行签名(表):
    rows = {}
    for 坐标, 格 in 表["cells"].items():
        值 = 格.get("value")
        if 值 in (None, ""):
            continue                       # 空行不算结构行（删行后的残留空样式行不参与比对）
        rows.setdefault(格["row"], {})[格["column"]] = 值
    行号 = sorted(rows)
    return ([tuple(str(rows[r][c]) for c in sorted(rows[r])) for r in 行号], 行号)


def _合并集(表):
    """合并区按（左列、上行、右列、下行）存集合，便于按行映射归一后比较。"""
    出 = set()
    for m in 表.get("merges", []):
        try:
            出.add(range_boundaries(str(m)))
        except ValueError:
            continue
    return 出


def _前后缀(甲, 乙):
    """返回（相同前缀长度，相同后缀长度）。单处增删时两侧余段必有一侧为空。"""
    前 = 0
    while 前 < len(甲) and 前 < len(乙) and 甲[前] == 乙[前]:
        前 += 1
    后 = 0
    while (后 < len(甲) - 前 and 后 < len(乙) - 前
           and 甲[len(甲) - 1 - 后] == 乙[len(乙) - 1 - 后]):
        后 += 1
    return 前, 后


def _区域们(快照, 工作表):
    for 名, 位置们 in (快照.get("names") or {}).items():
        for 位 in 位置们:
            if 位.get("sheet") != 工作表:
                continue
            try:
                左, 顶, 右, 底 = range_boundaries(str(位["range"]))
            except ValueError:
                continue
            yield 名, 顶, 底


def 区域归属(快照, 工作表: str, 起始行: int, 行数: int):
    """返回完整覆盖该行段的命名区域名；没有返回 None。最小覆盖区优先。"""
    末行 = 起始行 + 行数 - 1
    候选 = [(底 - 顶, 名) for 名, 顶, 底 in _区域们(快照, 工作表) if 顶 <= 起始行 and 末行 <= 底]
    return min(候选)[1] if 候选 else None


def _区域范围(快照, 工作表: str, 区域: str):
    for 名, 顶, 底 in _区域们(快照, 工作表):
        if 名 == 区域:
            return 顶, 底
    raise RuntimeError(f"命名区域 {区域} 在工作表 {工作表} 上找不到")


def _区内合并(合并集, 顶: int, 底: int) -> set:
    """只取完整落在该行段内的合并，行号相对区首行——整表位移不影响比较。"""
    return {(l, t - 顶, r, b - 顶) for l, t, r, b in 合并集 if 顶 <= t and b <= 底}


def _连续(序列):
    runs, run = [], [序列[0]] if 序列 else []
    for x in 序列[1:]:
        if x == run[-1] + 1:
            run.append(x)
        else:
            runs.append(run); run = [x]
    if run:
        runs.append(run)
    return runs


def 生成结构操作(基准表A路径, 修改后excel路径) -> list:
    """返回桥接 operations 数组；结构未变返回 []。落不到区域就 raise，交人工。"""
    甲 = 读快照(基准表A路径)
    乙 = 读快照(修改后excel路径)
    if [s["name"] for s in 甲["sheets"]] != [s["name"] for s in 乙["sheets"]]:
        raise RuntimeError("工作表结构变化，无法锚定")
    ops = []
    for 表甲, 表乙 in zip(甲["sheets"], 乙["sheets"]):
        名 = 表甲["name"]
        甲行, 甲号 = _行签名(表甲)
        乙行 = _行签名(表乙)[0]
        甲合并, 乙合并 = _合并集(表甲), _合并集(表乙)
        if len(甲行) == len(乙行) and 甲合并 == 乙合并:
            continue                             # 只改名等：结构没动，直接放行
        前, 后 = _前后缀(甲行, 乙行)
        甲余 = list(range(前, len(甲行) - 后))
        乙余 = list(range(前, len(乙行) - 后))
        if not 甲余 and not 乙余:
            raise RuntimeError(f"工作表 {名} 行结构变化无法判断增删，交人工")
        if 甲余 and 乙余:
            raise RuntimeError(f"工作表 {名} 同时存在增行与删行，暂不支持，交人工")

        if 甲余:
            段们 = _连续(甲余)
            if len(段们) > 1:
                raise RuntimeError(f"工作表 {名} 多处删行，暂不支持，交人工")
            段 = 段们[0]
            首行 = 甲号[段[0]]
            区域 = 区域归属(甲, 名, 首行, len(段))
            if 区域 is None:
                raise RuntimeError(f"工作表 {名} 删除行段不在任何命名区域内，交人工")
            顶, 底 = _区域范围(甲, 名, 区域)
            if _区内合并(甲合并, 顶, 底) != _区内合并(乙合并, 顶, 底 - len(段)):
                raise RuntimeError(f"工作表 {名} 区域 {区域} 内合并结构变化，暂不支持，交人工")
            ops.append({"range_name": 区域, "axis": "row", "action": "delete",
                        "index": 首行 - 顶, "count": len(段)})
        else:
            段们 = _连续(乙余)
            if len(段们) > 1:
                raise RuntimeError(f"工作表 {名} 多处增行，暂不支持，交人工")
            段 = 段们[0]
            插入行 = 甲号[0] if 前 == 0 else (甲号[-1] + 1 if 前 >= len(甲号) else 甲号[前])
            区域 = 区域归属(甲, 名, 插入行, 1) or 区域归属(甲, 名, max(1, 插入行 - 1), 1)
            if 区域 is None:
                raise RuntimeError(f"工作表 {名} 插入位置不在任何命名区域内，交人工")
            顶, 底 = _区域范围(甲, 名, 区域)
            if _区内合并(甲合并, 顶, 底) != _区内合并(乙合并, 顶, 底 + len(段)):
                raise RuntimeError(f"工作表 {名} 区域 {区域} 内合并结构变化，暂不支持，交人工")
            序 = 插入行 - 顶
            ops.append({"range_name": 区域, "axis": "row", "action": "insert",
                        "index": 序, "count": len(段), "template_index": max(0, 序 - 1)})
    return ops

