# coding: utf-8
"""识别证据与可复用解释：只核对首次识别已标明用途的原文、结构和位置。

本模块不判断语义：固定文字必须逐字相同，金额位置不比较金额，名单项只比较位置与项数，
日期位置只要求仍是日期。任何一条对不上就返回“未命中”及具体原因，交回现有识别流程。

证据分三个步骤保存：
- 第一遍：整侧科目段的首行/末行原文与行数，按段锚点重定位；
- 第二遍：每个业务表（含续表组）的首行原文、行数与组内关系；
- 第三遍：整组物理表的固定文字、合并形状、名单标签位置、日期位置与行列解释。

解释里的位置一律相对本组保存，绝对坐标只在本次重新定位后生成。
"""
from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_to_tuple

from .表格 import 取可见集

证据版本 = 1
匹配版本 = 1

_日期像 = re.compile(r"(\d{4})\s*[年\-/\.]\s*(\d{1,2})\s*[月\-/\.]\s*(\d{1,2})")
_数值像 = re.compile(r"^[（(]?[-+]?[\d,，\s]*\.?\d+[)）]?%?$")
_空白标记 = {"", "-", "—", "－", "–", "/", "不适用", "N/A", "n/a", "无"}


# —— 摘要与版本 ——

def _摘要(对象) -> str:
    return hashlib.sha256(json.dumps(对象, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode("utf-8")).hexdigest()


def 内容摘要(路径) -> str:
    """金标准按实际内容取版本；文件名或更新时间不构成版本。"""
    p = Path(路径)
    if not p.is_file():
        return ""
    h = hashlib.sha256()
    with p.open("rb") as f:
        for 块 in iter(lambda: f.read(1024 * 1024), b""):
            h.update(块)
    return h.hexdigest()


def 规则摘要() -> str:
    """影响认表、范围、证据或解释的规则版本。

    指令与格式原文、可见范围规则、匹配规则、证据版本任一变化都会改变摘要，
    旧记录随之退出自动复用；不按文件时间或版本号字符串判断。
    """
    from . import 识别

    规则 = {
        "第一遍指令": 识别._第一遍指令, "第一遍格式": 识别._第一遍格式,
        "第二遍指令": 识别._第二遍指令, "第二遍格式": 识别._第二遍格式,
        "第三遍指令": 识别._第三遍指令, "第三遍格式": 识别._第三遍格式,
        "期间角色": sorted(识别._期间角色集),
        "证据版本": 证据版本, "匹配版本": 匹配版本,
        "可见范围": "打印区域并集去重＋扣隐藏行列；无打印区域取可见内容；合并跨界只记结构",
        "匹配": "固定文字逐字、金额不比、名单比位置与项数、日期只认日期",
    }
    return _摘要(规则)


def 形状摘要(形状) -> str:
    return _摘要(形状 or {})


def 证据键(证据: dict) -> str:
    """固定证据摘要：同一结构、同一固定原文才算同一个证据组。"""
    return _摘要({"侧": 证据.get("侧"), "步骤": 证据.get("步骤"),
                  "证据版本": 证据.get("证据版本"), "形状": 证据.get("形状"),
                  "固定": 证据.get("固定"), "段": 证据.get("段")})


def 解释键(解释: dict) -> str:
    """解释摘要：不含旧格值、旧标题新文和绝对坐标。"""
    return _摘要({"科目": 解释.get("科目"), "表名": 解释.get("表名"),
                  "表组序": 解释.get("表组序"), "续表": 解释.get("分表"),
                  "列": 解释.get("列"), "行": 解释.get("行"),
                  "作用范围": 解释.get("作用范围"), "空名单": 解释.get("空名单"),
                  "例外": 解释.get("例外"), "标题规则": 解释.get("标题规则"),
                  "划分": 解释.get("划分"), "表清单": 解释.get("表清单")})


def 版本号(键: str, 解释: str, 金标准版本: str, 规则版本: str) -> str:
    return "V" + _摘要({"证据": 键, "解释": 解释, "金标准": 金标准版本,
                        "规则": 规则版本, "证据版本": 证据版本})[:24]


# —— 视图：从快照读一个矩形的文字、合并与形状 ——

class 表视图:
    """某工作表某矩形范围内的单元格视图；B 侧只含可见集合里的位置。"""

    def __init__(self, 快照, 工作表, 矩形, 用可见范围=False):
        self.工作表 = 工作表
        self.左, self.顶, self.右, self.底 = 矩形
        self.用可见范围 = 用可见范围
        表 = next((s for s in 快照.get("sheets") or [] if s.get("name") == 工作表), None)
        if 表 is None:
            raise ValueError(f"快照里没有工作表：{工作表}")
        self._格 = 表.get("cells") or {}
        self._合并 = []
        for 合并 in 表.get("merges") or []:
            左, 顶, 右, 底 = range_boundaries(合并)
            if 右 < self.左 or 左 > self.右 or 底 < self.顶 or 顶 > self.底:
                continue
            self._合并.append((max(顶, self.顶), max(左, self.左),
                             min(底, self.底), min(右, self.右)))
        self._可见 = 取可见集(快照, 工作表) if 用可见范围 else None

    def 在范围(self, 行, 列) -> bool:
        return self.顶 <= 行 <= self.底 and self.左 <= 列 <= self.右

    def 可见(self, 行, 列) -> bool:
        if not self.在范围(行, 列):
            return False
        return True if self._可见 is None else self._可见.含(行, 列)

    def 文本(self, 行, 列):
        """返回文字；不可见或没有记录返回 None，空白格返回空串。"""
        if not self.在范围(行, 列):
            return None
        if self._可见 is not None and not self._可见.含(行, 列):
            return None
        格 = self._格.get(f"{get_column_letter(列)}{行}")
        if 格 is None:
            return ""
        值 = 格.get("cached_value") if 格.get("formula") else 格.get("value")
        if 值 is None:
            return ""
        return str(值).strip()

    def 数值(self, 行, 列) -> bool:
        格 = self._格.get(f"{get_column_letter(列)}{行}")
        if 格 is None:
            return False
        值 = 格.get("cached_value") if 格.get("formula") else 格.get("value")
        return isinstance(值, (int, float)) and not isinstance(值, bool)

    def 行首文本(self, 行) -> str:
        for 列 in range(self.左, self.右 + 1):
            文 = self.文本(行, 列)
            if 文:
                return 文
        return ""

    def 行尾文本(self, 行) -> str:
        for 列 in range(self.右, self.左 - 1, -1):
            文 = self.文本(行, 列)
            if 文:
                return 文
        return ""

    def 有内容行(self) -> list:
        """有文字或数值的行；隐藏行与打印区域外的行不算。"""
        行们 = []
        for 行 in range(self.顶, self.底 + 1):
            for 列 in range(self.左, self.右 + 1):
                if not self.可见(行, 列):
                    continue
                文 = self.文本(行, 列)
                if (文 not in (None, "")) or self.数值(行, 列):
                    行们.append(行)
                    break
        return 行们

    def 形状(self, 分表=None) -> dict:
        return {"行数": self.底 - self.顶 + 1, "列数": self.右 - self.左 + 1,
                "合并": sorted([[顶 - self.顶, 左 - self.左, 底 - self.顶, 右 - self.左]
                              for 顶, 左, 底, 右 in self._合并]),
                "分表": 分表 or []}

    def 合并覆盖(self) -> dict:
        """合并从属格 → 锚点（相对本视图左上角）。"""
        覆盖 = {}
        for 顶, 左, 底, 右 in self._合并:
            for 行 in range(顶, 底 + 1):
                for 列 in range(左, 右 + 1):
                    if (行, 列) != (顶, 左):
                        覆盖[(行, 列)] = (顶, 左)
        return 覆盖


# —— 证据构造 ——

def _段证据(视图, 切段) -> list:
    出 = []
    for 序, (起, 止, 科目) in enumerate(切段):
        出.append({"序": 序, "科目": 科目, "行数": 止 - 起 + 1,
                  "首行": 视图.行首文本(起), "末行": 视图.行尾文本(止)})
    return 出


def _内容矩形(快照, 工作表, 用可见范围):
    """本次材料的内容矩形：B 侧取可见集合，A 侧取已提取的全部格。"""
    if 用可见范围:
        可见 = 取可见集(快照, 工作表)
        if 可见 is None or 可见.空():
            return None
        return (min(可见.列们()), min(可见.行们()), max(可见.列们()), max(可见.行们()))
    表 = next((s for s in 快照.get("sheets") or [] if s.get("name") == 工作表), None)
    格们 = (表 or {}).get("cells") or {}
    if not 格们:
        return None
    return (min(c["column"] for c in 格们.values()), min(c["row"] for c in 格们.values()),
            max(c["column"] for c in 格们.values()), max(c["row"] for c in 格们.values()))


def 建证据_分科目(侧, 划分: list, 快照, 用可见范围=False, 工作表=None,
              本侧=None, 标题们=None) -> dict:
    """整侧科目段证据：每段首表标题、首末行原文与行数；不存金额、不存旧原文全篇。

    A 侧的表题（如“3、货币资金”）是章节锚点，比重复的表头文字可靠；B 侧没有表题时
    退回行首原文。两者都只作定位，不参与科目含义判断。
    """
    工作表 = 工作表 or 快照["sheets"][0]["name"]
    矩形 = _内容矩形(快照, 工作表, 用可见范围)
    if 矩形 is None:
        raise ValueError(f"工作表 {工作表} 没有内容，不能建立科目段证据")
    视图 = 表视图(快照, 工作表, 矩形, 用可见范围)
    切段 = [(x["行范围"][0], x["行范围"][1], x["科目"]) for x in 划分]
    段们 = _段证据(视图, 切段)
    for 段, (起, 止, _) in zip(段们, 切段):
        首个 = next((x for x in (本侧 or []) if 起 <= x["行范围"][0] <= 止), None)
        段["标题"] = (标题们 or {}).get((首个 or {}).get("区域")) or ""
    return {"侧": 侧, "步骤": "第一遍", "证据版本": 证据版本,
            "工作表": 工作表, "形状": {"段数": len(切段)}, "段": 段们}


def 建证据_分表(侧, 本侧: list, 快照, 用可见范围=False, 工作表=None, 标题们=None) -> dict:
    """每个业务表（续表按其组整体）的首行原文、行数与组内顺序。"""
    工作表 = 工作表 or (本侧[0]["工作表"] if 本侧 else 快照["sheets"][0]["name"])
    表们 = []
    for 序, 记录 in enumerate(本侧):
        起, 止 = 记录.get("行范围") or [0, 0]
        矩形 = _内容矩形(快照, 记录.get("工作表") or 工作表, 用可见范围) or (1, 起, 1, 止)
        视图 = 表视图(快照, 记录.get("工作表") or 工作表,
                    (矩形[0], 起, 矩形[2], 止), 用可见范围)
        表们.append({"序": 序, "科目": 记录.get("科目"), "表名": 记录.get("表名"),
                    "表组": 记录.get("表组"), "续表序": 记录.get("续表序"),
                    "标题": (标题们 or {}).get(记录.get("区域")) or "",
                    "行数": 止 - 起 + 1, "首行": 视图.行首文本(起), "末行": 视图.行尾文本(止)})
    return {"侧": 侧, "步骤": "第二遍", "证据版本": 证据版本,
            "形状": {"表数": len(表们)}, "表": 表们}


def _格类别(记录们, 行, 列, 字母, 文本, 日期位置) -> str:
    数据列 = {str(c.get("列") or "").upper() for c in (记录.get("列") or [])
            if c.get("角色") == "数据列"} if 记录 else set()
    if (行, 列) in 日期位置:
        return "日期"
    if 字母 in 数据列:
        return "数值"
    return "固定" if 文本 else "空"


def 建证据_精读(侧, 记录们: list, 快照, 用可见范围=False) -> dict:
    """一组业务表（主表＋续表）的固定证据与可复用解释。

    金额格记为“数值”，不复存金额；名单标签记为“名单”，名称可变而位置与项数不变；
    首次标明用途的日期格记为“日期”，只要求仍是日期。
    """
    记录们 = sorted(记录们, key=lambda x: x.get("续表序") or 1)
    工作表 = 记录们[0]["工作表"]
    顶 = min(x["行范围"][0] for x in 记录们)
    底 = max(x["行范围"][1] for x in 记录们)
    左 = min(range_boundaries(str(x["范围"]))[0] for x in 记录们 if x.get("范围"))
    右 = max(range_boundaries(str(x["范围"]))[2] for x in 记录们 if x.get("范围"))
    视图 = 表视图(快照, 工作表, (左, 顶, 右, 底), 用可见范围)

    日期位置 = _日期位置(记录们, 顶, 左)
    名单格, 固定格, 数值格, 空格 = [], [], [], []
    表头行 = {r.get("行") for r in (记录们[0].get("行") or []) if r.get("角色") == "表头"}
    行名列 = next((str(c.get("列") or "").upper() for c in (记录们[0].get("列") or [])
                if c.get("角色") == "行名列"), None)
    数据列 = {str(c.get("列") or "").upper() for c in (记录们[0].get("列") or [])
            if c.get("角色") == "数据列"}
    行序号 = {r.get("行"): r.get("序号") for r in (记录们[0].get("行") or [])}
    列序号 = {str(c.get("列") or "").upper(): c.get("序号") for c in (记录们[0].get("列") or [])}
    标签行 = {str(c.get("列") or "").upper(): c.get("标签行") for c in (记录们[0].get("列") or [])}

    for 行 in range(顶, 底 + 1):
        for 列 in range(左, 右 + 1):
            字母 = get_column_letter(列)
            文本 = 视图.文本(行, 列)
            if 文本 is None:
                continue                       # 不可见：不进入证据
            if (行, 列) in 日期位置:
                名单格.append({"相对": [行 - 顶, 列 - 左], "类别": "日期", "文本": ""})
                continue
            if 字母 in 数据列:
                数值格.append([行 - 顶, 列 - 左])
                continue
            if 字母 == 行名列 and 行 in 行序号:
                类 = "名单" if 行序号[行] is not None else "固定"
                (名单格 if 类 == "名单" else 固定格).append(
                    {"相对": [行 - 顶, 列 - 左], "类别": 类, "文本": 文本})
                continue
            if 列序号.get(字母) is not None and 标签行.get(字母) == 行:
                名单格.append({"相对": [行 - 顶, 列 - 左], "类别": "名单", "文本": 文本})
                continue
            if 行 in 表头行 or 文本:
                固定格.append({"相对": [行 - 顶, 列 - 左], "类别": "固定", "文本": 文本})
            else:
                空格.append([行 - 顶, 列 - 左])

    分表 = [[x["行范围"][0] - 顶, x["行范围"][1] - 顶] for x in 记录们]
    形状 = 视图.形状(分表)
    return {"侧": 侧, "步骤": "第三遍", "证据版本": 证据版本, "工作表": 工作表,
            "形状": 形状,
            "固定": {"格": 固定格, "名单": 名单格, "数值": 数值格, "空": 空格},
            "解释": _抽解释(记录们, 顶, 左)}


def _日期位置(记录们, 顶, 左) -> set:
    位置 = set()
    for 记录 in 记录们:
        for 项 in list(记录.get("列") or []) + list(记录.get("行") or []):
            坐标 = 项.get("期间实际位置")
            if not 坐标:
                continue
            try:
                行, 列 = coordinate_to_tuple(str(坐标).upper())
            except (TypeError, ValueError):
                continue
            位置.add((行, 列))
    return 位置


def _相对坐标(坐标, 顶, 左):
    try:
        行, 列 = coordinate_to_tuple(str(坐标).upper())
    except (TypeError, ValueError):
        return None
    return [行 - 顶, 列 - 左]


def _抽解释(记录们, 顶, 左) -> dict:
    """把绝对坐标换成本组相对位置，剔除旧格值、旧标题新文与旧表组名。"""
    主 = 记录们[0]
    列们 = []
    for c in 主.get("列") or []:
        列们.append({"相对列": ord(str(c.get("列")).upper()) - 64, "路径": c.get("路径"),
                   "角色": c.get("角色"), "期间角色": c.get("期间角色"),
                   "序号": c.get("序号"), "标签行": (c.get("标签行") - 顶) if c.get("标签行") else None,
                   "期间实际": c.get("期间实际"),
                   "期间实际位置": _相对坐标(c.get("期间实际位置"), 顶, 左) if c.get("期间实际位置") else None})
    行们 = []
    for r in 主.get("行") or []:
        槽 = {}
        for 字母, 编号 in (r.get("槽位") or {}).items():
            槽[str(ord(str(字母).upper()) - 64)] = 编号
        按列 = {}
        for 字母, 角色 in (r.get("期间角色按列") or {}).items():
            按列[str(ord(str(字母).upper()) - 64)] = 角色
        行们.append({"相对行": r.get("行") - 顶, "名字": r.get("名字"), "角色": r.get("角色"),
                   "序号": r.get("序号"), "期间角色": r.get("期间角色"),
                   "期间角色按列": 按列, "期间实际": r.get("期间实际"),
                   "期间实际位置": _相对坐标(r.get("期间实际位置"), 顶, 左) if r.get("期间实际位置") else None,
                   "分组": r.get("分组"), "槽位": 槽, "不搬原因": r.get("不搬原因")})
    例外 = []
    for e in 主.get("例外") or []:
        try:
            左2, 顶2, 右2, 底2 = range_boundaries(str(e.get("范围") or "").replace("$", ""))
        except (ValueError, TypeError):
            continue
        例外.append({"相对": [顶2 - 顶, 左2 - 左, 底2 - 顶, 右2 - 左],
                   "处理": e.get("处理"), "原因": e.get("原因")})
    return {"科目": 主.get("科目"), "表名": 主.get("表名"),
            "表组序": 主.get("表组序") or 0,
            "分表": [[x["行范围"][0] - 顶, x["行范围"][1] - 顶] for x in 记录们],
            "列": 列们, "行": 行们,
            "作用范围": dict(主.get("作用范围") or {}),
            "空名单": list(主.get("空名单") or []),
            "例外": 例外,
            "标题规则": _标题规则(主, 顶)}


def _标题规则(记录, 顶) -> dict:
    """把标题存成“骨架＋角色”：原文换年也能按同一条规则重新取本次片段。

    旧“新文”一律不存；骨架里的 {角色} 占位会在复用时用本次标题的日期片段回填，
    这样既不需要旧标题原文，也不需要模型在第二份重新说明改哪里。
    """
    标题 = 记录.get("标题")
    if not isinstance(标题, dict):
        return {}
    原文 = str(标题.get("原文") or "")
    骨架, 角色们 = 原文, []
    for 项 in 标题.get("日期") or []:
        片段, 角色 = str(项.get("片段") or ""), str(项.get("期间角色") or "")
        if not 片段 or not 角色 or 片段 not in 骨架:
            return {}
        骨架 = 骨架.replace(片段, "{" + 角色 + "}", 1)
        角色们.append(角色)
    if not 角色们:
        return {}
    return {"有标题": True, "骨架": 骨架, "角色": 角色们}


def 按骨架取片段(骨架: str, 当前: str):
    """用骨架匹配本次标题，取出本次的日期片段；对不上就返回 None（不猜）。"""
    if not 骨架 or not 当前:
        return None
    import re as _re
    日期式 = r"(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})"
    模式, 角色们 = "", []
    余下 = 骨架
    while True:
        起 = 余下.find("{")
        止 = 余下.find("}", 起 + 1) if 起 >= 0 else -1
        if 起 < 0 or 止 < 0:
            模式 += _re.escape(余下)
            break
        模式 += _re.escape(余下[:起]) + "(" + 日期式 + ")"
        角色们.append(余下[起 + 1:止])
        余下 = 余下[止 + 1:]
    m = _re.fullmatch(模式, 当前.strip())
    if not m:
        return None
    return [(角色, m.group(i + 1)) for i, 角色 in enumerate(角色们)]


# —— 核对 ——

_标题日期 = re.compile(r"\d{4}\s*[年\-/\.]\s*\d{1,2}\s*[月\-/\.]\s*\d{1,2}\s*日?")


def _标题键(文本) -> str:
    """表题作锚点时把日期遮掉：正常换年不应让整章定位失败。"""
    return _标题日期.sub("#日", str(文本 or "").strip())


_标题日期 = re.compile(r"\d{4}\s*[年\-/\.]\s*\d{1,2}\s*[月\-/\.]\s*\d{1,2}\s*日?")


def _标题键(文本) -> str:
    """表题作锚点时把日期遮掉：正常换年不应让整章定位失败。"""
    return _标题日期.sub("#日", str(文本 or "").strip())


def _像日期(文本) -> bool:
    return bool(_日期像.search(str(文本 or "")))


def _像金额(文本) -> bool:
    文 = str(文本 or "").strip()
    if 文 in _空白标记:
        return True
    return bool(_数值像.match(文.replace(",", "").replace("，", "")))


def 核对精读(证据: dict, 视图: 表视图) -> dict:
    """逐项核对固定证据；全部对上才返回唯一解释，否则给出未命中原因。"""
    if 证据.get("证据版本") != 证据版本:
        return {"结果": "损坏", "原因": "证据版本不兼容"}
    形状 = 证据.get("形状") or {}
    if 视图.形状(形状.get("分表")) != 形状:
        return {"结果": "未命中", "原因": "表组结构、分表或合并关系已变"}
    固定 = 证据.get("固定") or {}
    for 项 in 固定.get("格") or []:
        行, 列 = 视图.顶 + 项["相对"][0], 视图.左 + 项["相对"][1]
        文 = 视图.文本(行, 列)
        if 文 is None or 文 != 项.get("文本"):
            return {"结果": "未命中",
                    "原因": f"固定文字不符：{get_column_letter(列)}{行} 应为“{项.get('文本')}”，实为“{文 or ''}”"}
    for 项 in 固定.get("名单") or []:
        行, 列 = 视图.顶 + 项["相对"][0], 视图.左 + 项["相对"][1]
        文 = 视图.文本(行, 列)
        if not 文:
            return {"结果": "未命中",
                    "原因": f"名单项位置没有名称：{get_column_letter(列)}{行}"}
        if 项.get("类别") == "日期" and not _像日期(文):
            return {"结果": "未命中",
                    "原因": f"日期位置已不是日期：{get_column_letter(列)}{行}“{文}”"}
    for 相对 in 固定.get("数值") or []:
        行, 列 = 视图.顶 + 相对[0], 视图.左 + 相对[1]
        文 = 视图.文本(行, 列)
        if 文 is None:
            return {"结果": "未命中", "原因": "金额位置落在本次不可见范围内"}
        if 文 and not _像金额(文) and not 视图.数值(行, 列):
            return {"结果": "未命中",
                    "原因": f"金额位置出现非金额文字：{get_column_letter(列)}{行}“{文}”"}
    for 相对 in 固定.get("空") or []:
        行, 列 = 视图.顶 + 相对[0], 视图.左 + 相对[1]
        文 = 视图.文本(行, 列)
        if 文:
            return {"结果": "未命中",
                    "原因": f"原空白位置出现文字：{get_column_letter(列)}{行}“{文}”"}
    return {"结果": "唯一", "原因": "", "解释": 重定位(证据.get("解释") or {}, 视图)}


def 重定位(解释: dict, 视图: 表视图) -> dict:
    """相对位置 → 本次绝对位置；行列角色、槽位、例外范围一并换算。"""
    顶, 左 = 视图.顶, 视图.左

    def 字母(序):
        return get_column_letter(左 + int(序) - 1)

    def 地址(相对):
        return f"{字母(相对[1] + 1)}{顶 + 相对[0]}"

    列们 = []
    for c in 解释.get("列") or []:
        列们.append({"列": 字母(c["相对列"]), "路径": c.get("路径"), "角色": c.get("角色"),
                   "期间角色": c.get("期间角色"), "序号": c.get("序号"),
                   "标签行": (顶 + c["标签行"]) if c.get("标签行") is not None else None,
                   "期间实际": c.get("期间实际"),
                   "期间实际位置": 地址(c["期间实际位置"]) if c.get("期间实际位置") else None})
    行们 = []
    for r in 解释.get("行") or []:
        槽 = {字母(k): v for k, v in (r.get("槽位") or {}).items()}
        按列 = {字母(k): v for k, v in (r.get("期间角色按列") or {}).items()}
        行们.append({"行": 顶 + r["相对行"], "名字": r.get("名字"), "角色": r.get("角色"),
                   "序号": r.get("序号"), "期间角色": r.get("期间角色"),
                   "期间角色按列": 按列, "期间实际": r.get("期间实际"),
                   "期间实际位置": 地址(r["期间实际位置"]) if r.get("期间实际位置") else None,
                   "分组": r.get("分组"), "槽位": 槽, "不搬原因": r.get("不搬原因") or ""})
    例外 = []
    for e in 解释.get("例外") or []:
        顶2, 左2, 底2, 右2 = e["相对"]
        例外.append({"范围": f"{get_column_letter(左 + 左2)}{顶 + 顶2}:"
                          f"{get_column_letter(左 + 右2)}{顶 + 底2}",
                   "处理": e.get("处理"), "原因": e.get("原因")})
    return {"科目": 解释.get("科目"), "表名": 解释.get("表名"),
            "表组序": 解释.get("表组序") or 0,
            "列": 列们, "行": 行们, "作用范围": dict(解释.get("作用范围") or {}),
            "空名单": list(解释.get("空名单") or []), "例外": 例外,
            "标题规则": dict(解释.get("标题规则") or {})}


def 核对分科目(证据: dict, 快照, 用可见范围=False, 工作表=None, 区域标题们=None) -> dict:
    """按段锚点把旧科目划分放回本次材料：每段首行必须唯一出现，覆盖必须完整。"""
    if 证据.get("证据版本") != 证据版本:
        return {"结果": "损坏", "原因": "证据版本不兼容"}
    工作表 = 证据.get("工作表") or 工作表
    if 用可见范围:
        可见 = 取可见集(快照, 工作表)
        if 可见 is None or 可见.空():
            return {"结果": "未命中", "原因": "本次没有可见内容，无法重定位科目段"}
        矩形 = (min(可见.列们()), min(可见.行们()), max(可见.列们()), max(可见.行们()))
    else:
        格们 = (快照["sheets"][0].get("cells") or {})
        if not 格们:
            return {"结果": "未命中", "原因": "本次没有内容，无法重定位科目段"}
        行们 = [c["row"] for c in 格们.values()]
        列们 = [c["column"] for c in 格们.values()]
        矩形 = (min(列们), min(行们), max(列们), max(行们))
    视图 = 表视图(快照, 工作表, 矩形, 用可见范围)
    内容行 = 视图.有内容行()
    if not 内容行:
        return {"结果": "未命中", "原因": "本次没有可见内容行"}
    段们 = 证据.get("段") or []
    if (证据.get("侧") or "") == "B":
        # 整张底表按原位核对布局：行数、段首段末原文都对上才沿用，不靠表题猜测
        位 = 内容行[0]
        划分 = []
        for 段 in 段们:
            止 = 位 + int(段.get("行数") or 0) - 1
            if 止 > 内容行[-1] or 视图.行首文本(位) != (段.get("首行") or "")                     or 视图.行尾文本(止) != (段.get("末行") or ""):
                return {"结果": "未命中", "原因": f"底表布局与记录不一致（第 {位} 行起的科目段）"}
            划分.append({"行范围": [位, 止], "科目": 段.get("科目")})
            位 = 止 + 1
        if 位 - 1 != 内容行[-1]:
            return {"结果": "未命中", "原因": "底表行数与记录不一致，可能有插行或删行"}
        return {"结果": "唯一", "原因": "", "划分": 划分}
    锚点 = []
    for 段 in 段们:
        标题 = (段.get("标题") or "").strip()
        if 标题:
            同题 = [x for x in (区域标题们 or []) if _标题键(x.get("标题")) == _标题键(标题)]
            if len(同题) > 1:
                return {"结果": "未命中", "原因": f"表题“{标题}”在本次出现 {len(同题)} 次，无法唯一定位"}
            if 同题:
                锚点.append((同题[0]["行范围"][0], 段.get("科目"), 标题, 段))
                continue
            continue                     # 有表题却找不到：这一段本次没有，不拿重复表头去猜
        文 = 段.get("首行") or ""
        if not 文:
            return {"结果": "未命中", "原因": f"第 {段.get('序', 0) + 1} 段缺少可核对的首行原文"}
        命中 = [r for r in 内容行 if 视图.行首文本(r) == 文]
        if len(命中) > 1:
            return {"结果": "未命中", "原因": f"段首原文“{文}”在本次材料中出现 {len(命中)} 次，无法唯一定位"}
        if 命中:
            锚点.append((命中[0], 段.get("科目"), 文, 段))
    锚点.sort()
    if not 锚点:
        return {"结果": "未命中", "原因": "旧科目段一个也定位不到"}
    划分 = []
    for i, (起, 科目, _文, 段) in enumerate(锚点):
        止 = (锚点[i + 1][0] - 1) if i + 1 < len(锚点) else 内容行[-1]
        if 止 - 起 + 1 != int(段.get("行数") or 0):
            return {"结果": "未命中",
                    "原因": f"第 {起} 行起的科目段长度与记录不一致（记录 {段.get('行数')} 行，本次 {止 - 起 + 1} 行）"}
        划分.append({"行范围": [起, 止], "科目": 科目})
    if 划分[0]["行范围"][0] > 内容行[0] or 划分[-1]["行范围"][1] < 内容行[-1]:
        return {"结果": "未命中", "原因": "本次内容超出旧科目段锚点覆盖的范围"}
    未覆盖 = [r for r in 内容行 if not any(a <= r <= b for a, b in (x["行范围"] for x in 划分))]
    if 未覆盖:
        return {"结果": "未命中", "原因": f"有 {len(未覆盖)} 行内容没有科目归属（首个 {未覆盖[0]} 行）"}
    return {"结果": "唯一", "原因": "", "划分": 划分}


def 核对分表(证据: dict, 快照, 划分: list, 用可见范围=False, 工作表=None,
         区域标题们=None) -> dict:
    """把旧分表结论放回本次科目段：每个表组按首行原文唯一定位，覆盖必须完整。"""
    if 证据.get("证据版本") != 证据版本:
        return {"结果": "损坏", "原因": "证据版本不兼容"}
    工作表 = 证据.get("工作表") or 工作表 or 快照["sheets"][0]["name"]
    矩形 = _内容矩形(快照, 工作表, 用可见范围)
    if 矩形 is None:
        return {"结果": "未命中", "原因": "本次没有内容，无法重定位业务表"}
    视图 = 表视图(快照, 工作表, 矩形, 用可见范围)
    内容行 = 视图.有内容行()
    组们 = []
    for t in 证据.get("表") or []:
        键 = t.get("表组") or f"{t.get('科目')}／{t.get('表名')}"
        if 组们 and 组们[-1]["键"] == 键:
            组们[-1]["表"].append(t)
        else:
            组们.append({"键": 键, "表": [t], "科目": t.get("科目")})
    出 = []
    for 段 in 划分:
        起, 止 = 段["行范围"]
        科目 = 段["科目"]
        段内 = [g for g in 组们 if g.get("科目") == 科目]
        if not 段内:
            if any(起 <= r <= 止 for r in 内容行):
                return {"结果": "未命中",
                        "原因": f"科目“{科目}”在旧记录里没有对应业务表，本次不能靠旧划分覆盖"}
            continue
        内容 = [r for r in 内容行 if 起 <= r <= 止]
        if (证据.get("侧") or "") == "B":
            位 = 起
            for g in 段内:
                for t in sorted(g["表"], key=lambda x: x.get("续表序") or 1):
                    末 = 位 + int(t.get("行数") or 0) - 1
                    if 末 > 止 or 视图.行首文本(位) != (t.get("首行") or ""):
                        return {"结果": "未命中",
                                "原因": f"底表布局与记录不一致（第 {位} 行起的表）"}
                    出.append({"行范围": [位, 末], "科目": 科目, "表名": t.get("表名"),
                              "表组": g["键"], "续表序": t.get("续表序")})
                    位 = 末 + 1
            if 位 - 1 != 止:
                return {"结果": "未命中", "原因": "科目段内行数与记录不一致"}
            continue
        锚点 = []
        for g in 段内:
            标题 = (g["表"][0].get("标题") or "").strip()
            if 标题:
                同题 = [x for x in (区域标题们 or []) if _标题键(x.get("标题")) == _标题键(标题)]
                if len(同题) > 1:
                    return {"结果": "未命中",
                            "原因": f"表题“{标题}”在本次出现 {len(同题)} 次，无法唯一定位"}
                if 同题:
                    锚点.append((同题[0]["行范围"][0], g))
                    continue
            首 = (g["表"][0].get("首行") or "").strip()
            if not 首:
                return {"结果": "未命中", "原因": f"表“{g['表'][0].get('表名')}”没有可核对的首行原文"}
            命中 = [r for r in 内容 if 视图.行首文本(r) == 首]
            if len(命中) != 1:
                return {"结果": "未命中",
                        "原因": f"表首行原文“{首}”在本次出现 {len(命中)} 次，无法唯一定位"}
            锚点.append((命中[0], g))
        锚点.sort(key=lambda x: x[0])
        for i, (顶, g) in enumerate(锚点):
            表们 = sorted(g["表"], key=lambda x: x.get("续表序") or 1)
            总 = sum(int(t.get("行数") or 0) for t in 表们)
            底 = 顶 + 总 - 1
            if 底 > 止:
                return {"结果": "未命中", "原因": f"表“{表们[0].get('表名')}”超出本次科目段"}
            if i + 1 < len(锚点) and 底 >= 锚点[i + 1][0]:
                return {"结果": "未命中", "原因": "两组业务表的范围重叠，续表归属不明"}
            偏移 = 顶
            for t in 表们:
                末 = 偏移 + int(t.get("行数") or 0) - 1
                出.append({"行范围": [偏移, 末], "科目": 科目, "表名": t.get("表名"),
                          "表组": g["键"], "续表序": t.get("续表序")})
                偏移 = 末 + 1
        表段 = [(x["行范围"][0], x["行范围"][1]) for x in 出 if 起 <= x["行范围"][0] <= 止]
        未覆盖 = [r for r in 内容 if not any(a <= r <= b for a, b in 表段)]
        if 未覆盖:
            return {"结果": "未命中",
                    "原因": f"科目段内有 {len(未覆盖)} 行未归入任何业务表（首个 {未覆盖[0]} 行）"}
    全未覆盖 = [r for r in 内容行
             if not any(x["行范围"][0] <= r <= x["行范围"][1] for x in 出)]
    if 全未覆盖:
        return {"结果": "未命中",
                "原因": f"本次有 {len(全未覆盖)} 行内容没有旧分表依据（首个 {全未覆盖[0]} 行）"}
    return {"结果": "唯一", "原因": "", "表": 出}


_完整性键 = ("第一遍覆盖", "第二遍归段", "第三遍逐行", "空名单已声明")


def 可发布(记录们: list, 完整性=None, 跳过原因=None) -> tuple:
    """入库门槛：按业务表组逐条判定，任何一条不成立都不发布，（False, 原因）。"""
    原因 = []
    if 跳过原因:
        原因.append(f"本组已跳过：{跳过原因}")
    if not 记录们:
        原因.append("没有可发布的表组记录")
    if 完整性 is not None:
        缺 = [k for k in _完整性键 if not 完整性.get(k)]
        if 缺:
            原因.append("本次整侧完整性未通过：" + "、".join(缺))
    for 记录 in 记录们:
        if not 记录.get("完成"):
            原因.append("第三遍未完成")
        if 记录.get("例外未执行"):
            原因.append("存在未执行的例外")
        if 记录.get("越界解释"):
            原因.append("解释越出本次可见范围")
        if 记录.get("项标签缺失"):
            原因.append("名单标签位置缺失")
        数据行 = [x for x in 记录.get("行") or [] if x.get("角色") in ("数据", "小计", "合计")]
        for x in 数据行:
            if not (x.get("槽位") or x.get("不搬原因")):
                原因.append(f"第 {x.get('行')} 行既没有槽位映射也没有不搬原因")
        if not 记录.get("列") and 记录.get("完成"):
            原因.append("缺少列清单")
    去重 = list(dict.fromkeys(原因))
    return (not 去重, "；".join(去重))
