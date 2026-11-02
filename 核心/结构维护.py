# coding: utf-8
"""结构维护：先定位变更落在哪个命名区域，再算该区域内的相对行号（桥接语义）。

判结构不靠“有内容才算行”，也不靠“有特殊样式才算行”：
- 只清空内容、行还在（命名区域仍罩着这一行）→ 不算删行，不生成任何操作。
- 中间插入空白行 → 按“后面的内容整体下移”认出来，生成插行操作。
- 只有“后面的内容整体上移”或“行已不在任何命名区域里”才算真删行。
- 列数或列位置变化（含两列对调）→ 明确报不支持，交人工。
一张工作表里可能装着几十张表，所以区域外、跨区域、多处增删一律交人工，不猜。
"""
from openpyxl.utils import range_boundaries

from .表格 import 读快照


def _值(格):
    v = 格.get("value")
    return None if v is None or v == "" else v


def _占用尾(表) -> int:
    """这张工作表里最后一个有值的行号；全空返回 0。"""
    行们 = [c["row"] for c in (表.get("cells") or {}).values() if _值(c) is not None]
    return max(行们) if 行们 else 0


def _行序列(表, 尾: int) -> list:
    """第 1 到 尾 行，每行一个内容元组；空行是 ()——空行本身也是结构证据。"""
    每行 = {}
    for c in (表.get("cells") or {}).values():
        if 1 <= c["row"] <= 尾 and _值(c) is not None:
            每行.setdefault(c["row"], []).append((c["column"], str(_值(c))))
    return [tuple(v for _, v in sorted(每行.get(r, []))) for r in range(1, 尾 + 1)]


def _合并集(表):
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
            yield 名, 顶, 底, 左, 右


def 区域归属(快照, 工作表: str, 起始行: int, 行数: int):
    """返回完整覆盖该行段的命名区域名；没有返回 None。最小覆盖区优先。"""
    末行 = 起始行 + 行数 - 1
    候选 = [(底 - 顶, 名) for 名, 顶, 底, _左, _右 in _区域们(快照, 工作表)
          if 顶 <= 起始行 and 末行 <= 底]
    return min(候选)[1] if 候选 else None


def _区域范围(快照, 工作表: str, 区域: str):
    for 名, 顶, 底, 左, 右 in _区域们(快照, 工作表):
        if 名 == 区域:
            return 顶, 底, 左, 右
    raise RuntimeError(f"命名区域 {区域} 在工作表 {工作表} 上找不到")


def _行还在区域里(快照, 工作表: str, 行号: int) -> bool:
    return any(顶 <= 行号 <= 底 for _名, 顶, 底, _左, _右 in _区域们(快照, 工作表))


def _列值(表, 顶: int, 底: int) -> dict:
    """区域内每一列的值集合（可重复计数）。列被对调时，一列的值会整体跑到另一列。"""
    from collections import Counter
    每列 = {}
    for c in (表.get("cells") or {}).values():
        值 = _值(c)
        if 值 is None or not (顶 <= c["row"] <= 底):
            continue
        每列.setdefault(c["column"], Counter())[str(值)] += 1
    return 每列


def _查列变化(表甲, 表乙, 名: str, 快照甲, 快照乙) -> None:
    """列数、列位置（含两列对调）变化都明确拒绝，不许返回空操作。

    只比“列还在不在、值有没有整体换到别的列”，不比类型——行整体上移下移时，
    同一列里会临时混进表头文字，用类型比会误判成列结构变了。
    """
    甲列 = {c["column"] for c in (表甲.get("cells") or {}).values() if _值(c) is not None}
    乙列 = {c["column"] for c in (表乙.get("cells") or {}).values() if _值(c) is not None}
    if 甲列 and 乙列 and (max(乙列) > max(甲列) or min(乙列) < min(甲列)):
        raise RuntimeError("列结构变化，本期不支持，交人工")      # 真加了一列或删了一列
    区域甲 = {名2: (顶, 底, 左, 右) for 名2, 顶, 底, 左, 右 in _区域们(快照甲, 名)}
    区域乙 = {名2: (顶, 底, 左, 右) for 名2, 顶, 底, 左, 右 in _区域们(快照乙, 名)}
    for 区域, (顶, 底, 左, 右) in 区域甲.items():
        if 区域 not in 区域乙:
            continue
        顶2, 底2, 左2, 右2 = 区域乙[区域]
        if (左, 右) != (左2, 右2):
            raise RuntimeError("列结构变化，本期不支持，交人工")  # 区域左右边界变了
        甲值 = _列值(表甲, 顶, min(底, 底2))
        乙值 = _列值(表乙, 顶2, min(底, 底2))
        for 列, 值们 in 甲值.items():
            if 列 not in 乙值 or 乙值[列] == 值们:
                continue
            if any(另列 != 列 and 另值 == 值们 for 另列, 另值 in 乙值.items()):
                raise RuntimeError("列结构变化，本期不支持，交人工")   # 这一列整体跑到了别的列


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


def _查合并(表甲, 表乙, 名: str, 区域: str, 顶: int, 底甲: int, 底乙: int) -> None:
    if _区内合并(_合并集(表甲), 顶, 底甲) != _区内合并(_合并集(表乙), 顶, 底乙):
        raise RuntimeError(f"工作表 {名} 区域 {区域} 内合并结构变化，暂不支持，交人工")


def 生成结构操作(基准表A路径, 修改后excel路径) -> list:
    """返回桥接 operations 数组；结构未变返回 []。落不到区域就 raise，交人工。"""
    甲 = 读快照(基准表A路径)
    乙 = 读快照(修改后excel路径)
    if [s["name"] for s in 甲["sheets"]] != [s["name"] for s in 乙["sheets"]]:
        raise RuntimeError("工作表结构变化，无法锚定")
    ops = []
    for 表甲, 表乙 in zip(甲["sheets"], 乙["sheets"]):
        名 = 表甲["name"]
        _查列变化(表甲, 表乙, 名, 甲, 乙)
        甲尾, 乙尾 = _占用尾(表甲), _占用尾(表乙)
        if not 甲尾 and not 乙尾:
            continue
        甲序, 乙序 = _行序列(表甲, 甲尾), _行序列(表乙, 乙尾)
        if 甲序 == 乙序:
            continue                             # 结构没动（可能只是改名）

        前, 后 = _前后缀(甲序, 乙序)
        甲余 = list(range(前, len(甲序) - 后))
        乙余 = list(range(前, len(乙序) - 后))
        if 甲余 and 乙余 and len(甲余) == len(乙余) and all(not 乙序[i] for i in 乙余):
            continue                             # 两侧同位置都空：只是把内容清空了，行还在
        if 甲余 and 乙余 and len(甲余) == len(乙余):
            continue                             # 两侧都改了内容、行数没变：内容变化，不动结构
        if not 甲余 and 乙余:
            段 = _连续(乙余)
            if len(段) > 1:
                raise RuntimeError(f"工作表 {名} 多处增行，暂不支持，交人工")
            插入行 = 段[0][0] + 1
            区域 = 区域归属(甲, 名, 插入行, 1) or 区域归属(甲, 名, max(1, 插入行 - 1), 1)
            if 区域 is None:
                raise RuntimeError(f"工作表 {名} 插入位置不在任何命名区域内，交人工")
            顶, 底, _l, _r = _区域范围(甲, 名, 区域)
            _查合并(表甲, 表乙, 名, 区域, 顶, 底, 底 + len(段[0]))
            ops.append({"range_name": 区域, "axis": "row", "action": "insert",
                        "index": 插入行 - 顶, "count": len(段[0]),
                        "template_index": max(0, 插入行 - 顶 - 1)})
            continue
        if 甲余 and not 乙余:
            段 = _连续(甲余)
            if len(段) > 1:
                raise RuntimeError(f"工作表 {名} 多处删行，暂不支持，交人工")
            首行 = 段[0][0] + 1
            if 后 == 0 and _行还在区域里(乙, 名, 首行):
                continue                         # 尾部只是清空（区域还罩着这一行），不算删行
            区域 = 区域归属(甲, 名, 首行, len(段[0]))
            if 区域 is None:
                raise RuntimeError(f"工作表 {名} 删除行段不在任何命名区域内，交人工")
            顶, 底, _l, _r = _区域范围(甲, 名, 区域)
            _查合并(表甲, 表乙, 名, 区域, 顶, 底, 底 - len(段[0]))
            ops.append({"range_name": 区域, "axis": "row", "action": "delete",
                        "index": 首行 - 顶, "count": len(段[0])})
            continue
        raise RuntimeError(f"工作表 {名} 同时存在增行与删行或改动交错，暂不支持，交人工")
    return ops
