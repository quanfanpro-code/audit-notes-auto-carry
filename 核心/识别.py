# coding: utf-8
"""识别编排：三遍线性——先分科目，再分表，最后逐表精读挂槽位编号。

串行依赖：第二遍读第一遍的科目划分，第三遍读前两遍的科目与表名，全部顺序执行、不并发。
第一、二遍不给任何金标准条目；第三遍只给该表候选槽位。
模型只交结构判断；原值与原值状态由程序读原文件取得，答复里的值字段一律丢弃。
"""
import json
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries

from .接续 import 读索引, 写索引
from .表格 import 读快照
from .表格文字 import 表格转文字
from .看表规范 import 分科目规范, 分表规范, 精读规范
from .金标准 import 加载金标准, 建索引, 科目清单, 表清单, 候选槽位
from .模型调用 import 调用识别, 校验指令

_答复名 = "答复.txt"
_数据行角色 = {"数据", "小计", "合计"}


# —— 三遍的指令与材料 ——

_第一遍指令 = ("把整份文件按科目切开。只交哪些行范围属于哪个科目，不交表名、不交行列、不交任何编号。"
           "严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_第一遍格式 = ("严格输出一个 JSON 对象：\n"
           '{"划分":[{"行范围":[起始行,结束行],"科目":"科目原名或非表格"}, ...]}\n'
           "行范围用原文件行号，含首含尾，按行号从小到大；相邻两段首尾相接，覆盖全部有内容的行，不留空档、不重叠。\n"
           "是表但清单里没有对应科目的，写「无对应科目：」再接这张表的原文标题，标题照抄、不改写、不概括，"
           "例如「无对应科目：母公司」「无对应科目：应收关联方款项」；只写「无对应科目」不给标题算不合规。")

_第二遍指令 = ("在第一遍的科目划分内，把每段切成表。只交行范围、科目、表名，不交行列结构、不交任何编号。"
           "严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_第二遍格式 = ("严格输出一个 JSON 对象：\n"
           '{"表":[{"行范围":[起始行,结束行],"科目":"科目原名","表名":"表名原名"}, ...]}\n'
           "表名只能取自该科目的表名清单；行范围不得跨出第一遍给该科目的行范围。")

_第三遍指令 = ("只读这一张表：交列清单、行清单、作用范围、例外，并给每个数据格挑候选槽位编号。"
           "不交任何数值、不改原文。严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_第三遍格式 = """严格输出一个 JSON 对象：
{
  "列": [{"列": "B", "路径": "项目", "角色": "行名列", "期间": null},
         {"列": "C", "路径": "期末余额", "角色": "数据列", "期间": "期末"}],
  "行": [{"行": 56, "名字": "项目", "角色": "表头"},
         {"行": 57, "名字": "库存现金", "角色": "数据", "槽位": {"C": "C-N006-T001-S002"}}],
  "作用范围": {"币种": "人民币", "单位": "元"},
  "例外": []
}
要求：列字母、行号都用原文件的位置；路径多级用斜杠连接，不许拍平；数据行、小计行、合计行的槽位映射必须逐列给全，挑不到就写 ""；编号只能从候选里挑，不许编；"期间"只写原文有的（期末/期初/本期/上期），确知日期时另给 "期间实际"（YYYY-MM-DD）。"""


def 第一遍材料(文件文字: str, 科目名们: list) -> tuple:
    材料 = "\n\n".join([
        "【一、整份文件的表格文字】（不裁剪）\n" + 文件文字,
        "【二、科目候选清单】只能照抄这里的原字，不得改写；说明段、目录、纯文字一律写「非表格」。\n"
        + "\n".join("- " + str(名) for 名 in 科目名们),
        "【三、看表规范】\n" + 分科目规范(),
        "【四、输出格式】\n" + _第一遍格式,
        "【五、示例】名字是假的，只示意形状\n"
        '{"划分":[{"行范围":[1,6],"科目":"示例科目甲"},{"行范围":[7,12],"科目":"非表格"}]}',
    ])
    return _第一遍指令, 材料


def 第二遍材料(文件文字: str, 科目划分: list, 表名册: dict) -> tuple:
    划段 = "\n".join(
        f"- 第 {x['行范围'][0]}–{x['行范围'][1]} 行：{x['科目']}" for x in 科目划分)
    册段 = "\n".join(
        "- " + 科 + "：" + "、".join(str(名) for 名 in 表们)
        for 科, 表们 in 表名册.items())
    材料 = "\n\n".join([
        "【一、整份文件的表格文字】（不裁剪）\n" + 文件文字,
        "【二、第一遍的科目划分】\n" + 划段,
        "【三、表名清单】表名只能从中照抄\n" + (册段 or "（无）"),
        "【四、看表规范】\n" + 分表规范(),
        "【五、输出格式】\n" + _第二遍格式,
        "【六、示例】名字是假的，只示意形状\n"
        '{"表":[{"行范围":[2,9],"科目":"示例科目甲","表名":"示例表名"}]}',
    ])
    return _第二遍指令, 材料


def 第三遍材料(表文字: str, 科目: str, 表名: str, 候选: list) -> tuple:
    行们 = [" | ".join([str(c.get("id") or ""), str(c.get("槽位名") or ""),
                      "行:" + "/".join(map(str, c.get("行路径") or [])),
                      "列:" + "/".join(map(str, c.get("列路径") or []))]) for c in 候选]
    名单 = "\n".join(行们) if 行们 else "（本表没有候选槽位：这一遍槽位一律留空）"
    材料 = "\n\n".join([
        "【一、本表的表格文字】（只这一张表，不裁剪）\n" + 表文字,
        "【二、本表身份】科目：" + 科目 + "\n表名：" + 表名,
        "【三、本表候选槽位】编号 | 槽位名 | 行路径 | 列路径\n" + 名单,
        "【四、看表规范】\n" + 精读规范(),
        "【五、输出格式】\n" + _第三遍格式,
        "【六、示例】名字与编号都是假的\n"
        '{"列":[{"列":"B","路径":"项目","角色":"行名列","期间":null},'
        '{"列":"C","路径":"期末余额","角色":"数据列","期间":"期末"}],'
        '"行":[{"行":2,"名字":"项目","角色":"表头"},'
        '{"行":3,"名字":"库存现金","角色":"数据","槽位":{"C":"C-N999-T999-S999"}}],'
        '"作用范围":{"币种":"人民币","单位":"元"},"例外":[]}',
    ])
    return _第三遍指令, 材料


# —— 文件与答复 ——

def _读jsonl(路径) -> list:
    p = Path(路径)
    if not p.is_file():
        return []
    return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]


def _写jsonl(路径, 记录们) -> None:
    p = Path(路径)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        for r in 记录们:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def _去围栏(文本: str) -> str:
    行们 = [x for x in (文本 or "").strip().splitlines() if not x.strip().startswith("```")]
    return "\n".join(行们).strip()


def _解析对象(答复: str, 说明: str) -> dict:
    try:
        对象 = json.loads(_去围栏(答复))
    except json.JSONDecodeError as e:
        raise RuntimeError(f"{说明}答复不是合法 JSON：{e}") from e
    if not isinstance(对象, dict):
        raise RuntimeError(f"{说明}答复应是 JSON 对象")
    return 对象


def _取答复(目录: Path, 指令: str, 材料: str, 配置, 只用已有答复: bool):
    """有答复就导入；没有就写出材料，允许调用则调用；仍没有返回 None（留待答复）。"""
    目录.mkdir(parents=True, exist_ok=True)
    答复路径 = 目录 / _答复名
    校验指令(指令)                                  # 简短指令硬校验，超 2000 字符直接拒
    if 答复路径.is_file():
        return 答复路径.read_text(encoding="utf-8")
    (目录 / "指令.txt").write_text(指令, encoding="utf-8")
    (目录 / "材料.txt").write_text(材料, encoding="utf-8")
    if 只用已有答复 or not 配置:
        return None
    文本 = 调用识别(配置, 指令, 材料, 目录)
    答复路径.write_text(文本, encoding="utf-8")
    return 文本


# —— 答复校验（程序只校验格式与清单归属，不替模型改字） ——

_无对应前缀 = "无对应科目："


def _合法科目(科目, 科目集) -> bool:
    if not isinstance(科目, str):
        return False
    if 科目 in 科目集 or 科目 == "非表格":
        return True
    return 科目.startswith(_无对应前缀) and len(科目) > len(_无对应前缀)


def 待表科目(科目划分: list, 索引: dict) -> dict:
    """只给真科目配表名清单；无对应科目与非表格都不进第二遍。"""
    科目集 = set(科目清单(索引))
    return {科: 表清单(索引, 科)
            for 科 in sorted({x["科目"] for x in 科目划分}) if 科 in 科目集}

def _解析划分(答复: str, 科目集: set) -> list:
    划分 = _解析对象(答复, "第一遍").get("划分")
    if not isinstance(划分, list) or not 划分:
        raise RuntimeError("第一遍答复缺「划分」")
    出, 上末 = [], 0
    for 段 in 划分:
        范围 = 段.get("行范围") or []
        科目 = 段.get("科目")
        if not (isinstance(范围, list) and len(范围) == 2 and all(isinstance(x, int) for x in 范围)):
            raise RuntimeError(f"第一遍行范围不合规：{段}")
        起, 止 = 范围
        if 起 > 止 or 起 <= 上末:
            raise RuntimeError(f"第一遍行范围重叠或倒序：{段}")
        if not _合法科目(科目, 科目集):
            raise RuntimeError(f"第一遍科目不在清单里：{科目}")
        出.append({"行范围": [起, 止], "科目": 科目})
        上末 = 止
    return 出


def _解析表(答复: str, 表名册: dict, 划分: list) -> list:
    表们 = _解析对象(答复, "第二遍").get("表")
    if not isinstance(表们, list) or not 表们:
        raise RuntimeError("第二遍答复缺「表」")
    出 = []
    for 表 in 表们:
        范围 = 表.get("行范围") or []
        科目, 表名 = 表.get("科目"), 表.get("表名")
        if not (isinstance(范围, list) and len(范围) == 2 and all(isinstance(x, int) for x in 范围)):
            raise RuntimeError(f"第二遍行范围不合规：{表}")
        if 科目 not in 表名册:
            raise RuntimeError(f"第二遍科目不在清单里：{科目}")
        if 表名 not in 表名册[科目]:
            raise RuntimeError(f"第二遍表名不在该科目清单里：{科目}／{表名}")
        段 = [x["行范围"] for x in 划分 if x["科目"] == 科目]
        if not any(起 <= 范围[0] and 范围[1] <= 止 for 起, 止 in 段):
            raise RuntimeError(f"第二遍行范围跨出该科目：{表}")
        出.append({"行范围": list(范围), "科目": 科目, "表名": 表名})
    return 出


def _解析精读(答复: str, 编号集: set) -> dict:
    对象 = _解析对象(答复, "第三遍")
    列们, 行们 = 对象.get("列") or [], 对象.get("行") or []
    if not 列们 or not 行们:
        raise RuntimeError("第三遍答复缺列清单或行清单")
    列出 = []
    for c in 列们:
        字母 = str(c.get("列") or "").strip().upper()
        if not 字母:
            raise RuntimeError(f"第三遍列缺字母：{c}")
        列出.append({"列": 字母, "路径": str(c.get("路径") or ""),
                   "角色": str(c.get("角色") or ""), "期间": c.get("期间") or None,
                   "期间实际": c.get("期间实际") or None})
    行出 = []
    for r in 行们:
        try:
            行号 = int(r.get("行"))
        except (TypeError, ValueError):
            raise RuntimeError(f"第三遍行号不合规：{r}") from None
        槽 = r.get("槽位") or {}
        干净 = {}
        for 字母, 编号 in 槽.items():
            干净[str(字母).strip().upper()] = 编号 if 编号 in 编号集 else ""   # 查不到就留空，不猜
        行出.append({"行": 行号, "名字": str(r.get("名字") or ""),
                   "角色": str(r.get("角色") or ""), "槽位": 干净})
    return {"列": 列出, "行": 行出,
            "作用范围": 对象.get("作用范围") or {}, "例外": 对象.get("例外") or []}


# —— 位置与文字 ——

def _整表范围(表) -> str:
    格们 = 表.get("cells") or {}
    if not 格们:
        return ""
    行们 = [c["row"] for c in 格们.values()]
    列们 = [c["column"] for c in 格们.values()]
    return f"{get_column_letter(min(列们))}{min(行们)}:{get_column_letter(max(列们))}{max(行们)}"


def 全文文字(快照) -> str:
    """整个文件：有命名区域就逐区域出，没有就把整张工作表出。"""
    段 = []
    for 表 in 快照["sheets"]:
        区域们 = [(名, 位["range"]) for 名, 位置们 in (快照.get("names") or {}).items()
                for 位 in 位置们 if 位.get("sheet") == 表["name"]]
        if 区域们:
            段 += [表格转文字(快照, {"sheet": 表["name"], "range": str(范围)}) for _, 范围 in 区域们]
        else:
            范围 = _整表范围(表)
            if 范围:
                段.append(表格转文字(快照, {"sheet": 表["name"], "range": 范围}))
    return "\n\n".join(段)


def _定位(快照, 行范围: list) -> tuple:
    """返回（命名区域名或 None, 工作表, 精确范围）。最小覆盖该行段的区域优先。"""
    起, 止 = 行范围
    最优 = None
    for 名, 位置们 in (快照.get("names") or {}).items():
        for 位 in 位置们:
            try:
                _, 顶, _, 底 = range_boundaries(str(位["range"]))
            except ValueError:
                continue
            if 顶 <= 起 and 底 >= 止 and (最优 is None or 底 - 顶 < 最优[3]):
                最优 = (名, 位["sheet"], str(位["range"]), 底 - 顶)
    if 最优:
        return 最优[0], 最优[1], 最优[2]
    工作表们 = 快照["sheets"]
    if len(工作表们) > 1:
        raise RuntimeError("没有命名区域且不止一张工作表，无法判断表在哪张工作表上")
    工作表 = 工作表们[0]["name"]
    表 = 工作表们[0]
    列们 = [c["column"] for c in (表.get("cells") or {}).values() if 起 <= c["row"] <= 止]
    左 = get_column_letter(min(列们)) if 列们 else "A"
    右 = get_column_letter(max(列们)) if 列们 else "A"
    return None, 工作表, f"{左}{起}:{右}{止}"


def _合并们(快照, 工作表: str, 范围: str) -> list:
    左, 顶, 右, 底 = range_boundaries(范围)
    表 = next(s for s in 快照["sheets"] if s["name"] == 工作表)
    出 = []
    for 合并 in 表.get("merges", []):
        try:
            左2, 顶2, 右2, 底2 = range_boundaries(合并)
        except ValueError:
            continue
        if 左2 <= 右 and 右2 >= 左 and 顶2 <= 底 and 底2 >= 顶:
            锚 = (表.get("cells") or {}).get(f"{get_column_letter(左2)}{顶2}") or {}
            出.append({"范围": 合并, "锚点": 锚.get("value")})
    return 出


# —— 铺格子与补原值（纯机械，不做语义判断） ——

def 铺格子(结构记录: dict) -> list:
    工作表 = 结构记录.get("工作表") or "sheet1"
    作用 = 结构记录.get("作用范围") or {}
    数据列 = [c for c in (结构记录.get("列") or []) if c.get("角色") == "数据列"]
    记 = []
    for 行 in 结构记录.get("行") or []:
        if 行.get("角色") not in _数据行角色:
            continue
        行路径 = [x for x in str(行.get("名字") or "").split("/") if x]
        槽表 = 行.get("槽位") or {}
        for 列 in 数据列:
            字母 = str(列.get("列") or "").upper()
            列路径 = [x for x in str(列.get("路径") or "").split("/") if x]
            记.append({
                "格": f"{工作表}!{字母}{行.get('行')}",
                "科目": 结构记录.get("科目") or "",
                "表名": 结构记录.get("表名") or "",
                "行路径": 行路径, "列路径": 列路径,
                "维度值": {"期间角色": 列.get("期间") or "未知",
                        "期间实际": 列.get("期间实际") or "未知",
                        "选取范围": 行.get("名字") or "未知",
                        "币种": 作用.get("币种") or "未知",
                        "主体": 作用.get("主体") or "未知",
                        "报表范围": 作用.get("报表范围") or "未知",
                        "单位": 作用.get("单位") or "未知"},
                "槽位": 槽表.get(字母, ""), "槽位名": "",
                "原值": None, "原值状态": "待读",
            })
    return 记


def 补原值(记录们: list, 快照) -> list:
    索引 = {f"{表['name']}!{坐标}": 格
          for 表 in 快照["sheets"] for 坐标, 格 in (表.get("cells") or {}).items()}
    出 = []
    for 记录 in 记录们:
        格 = 索引.get(str(记录.get("格")))
        if 格 is None:
            出.append({**记录, "原值": None, "原值状态": "读取失败"})
        elif 格.get("formula"):
            缓存 = 格.get("cached_value")
            出.append({**记录, "原值": 缓存,
                     "原值状态": "值" if 缓存 is not None else "公式无保存结果"})
        else:
            值 = 格.get("value")
            空白 = 值 is None or 值 == ""
            出.append({**记录, "原值": None if 空白 else 值,
                     "原值状态": "真实空白" if 空白 else "值"})
    return 出


# —— 一侧的完整三遍 ——

def 识别一侧(配置, 表文件路径, 侧名, 运行目录, 金标准主文件, 上下文=None,
         只用已有答复=False, 取消=None) -> dict:
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    快照 = 读快照(表文件路径, 上下文)
    索引 = 建索引(加载金标准(金标准主文件)) if 金标准主文件 else {"科目": {}, "编号": set()}
    留证 = 运行目录 / "调用留证"
    统计 = {"科目划分": 0, "表数": 0, "格子": 0}
    文字 = None

    def 共用文字():
        nonlocal 文字
        if 文字 is None:
            文字 = 全文文字(快照)
        return 文字

    # 第一遍：分科目
    划分路径 = 运行目录 / f"科目划分_{侧名}.jsonl"
    划分 = [r for r in _读jsonl(划分路径) if r.get("侧") == 侧名]
    if not 划分:
        指令, 材料 = 第一遍材料(共用文字(), 科目清单(索引))
        目录 = 留证 / f"{侧名}_第一遍"
        答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复)
        if 答复 is None:
            return {"结构清单": None, "合集": None, "统计": 统计, "待答复": [str(目录)]}
        划分 = [{"侧": 侧名, **x} for x in _解析划分(答复, set(科目清单(索引)))]
        _写jsonl(划分路径, 划分)
    统计["科目划分"] = len(划分)
    if 取消 is not None and 取消.is_set():
        return {"结构清单": None, "合集": None, "统计": 统计, "待答复": ["已取消"]}

    # 第二遍：分表
    结构路径 = 运行目录 / "结构清单.jsonl"
    全部 = _读jsonl(结构路径)
    本侧 = [r for r in 全部 if r.get("侧") == 侧名]
    if not 本侧:
        表名册 = 待表科目(划分, 索引)
        指令, 材料 = 第二遍材料(共用文字(), 划分, 表名册)
        目录 = 留证 / f"{侧名}_第二遍"
        答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复)
        if 答复 is None:
            return {"结构清单": None, "合集": None, "统计": 统计, "待答复": [str(目录)]}
        本侧 = []
        for 表 in _解析表(答复, 表名册, 划分):
            区域, 工作表, 范围 = _定位(快照, 表["行范围"])
            本侧.append({"侧": 侧名, "区域": 区域, "工作表": 工作表, "范围": 范围,
                       "行范围": 表["行范围"], "科目": 表["科目"], "表名": 表["表名"],
                       "列": [], "行": [], "作用范围": {}, "原始结构": {}, "例外": [],
                       "完成": False})
        _写jsonl(结构路径, [r for r in 全部 if r.get("侧") != 侧名] + 本侧)
    统计["表数"] = len(本侧)

    # 第三遍：逐表精读（按表走，一张表一次，顺序执行）
    待答复 = []
    for 序号, 记录 in enumerate(本侧, 1):
        if 记录.get("完成"):
            continue
        if 取消 is not None and 取消.is_set():
            待答复.append("已取消")
            break
        表文字 = 表格转文字(快照, {"sheet": 记录["工作表"], "range": 记录["范围"]})
        候选 = 候选槽位(索引, 记录["科目"], 记录["表名"])
        指令, 材料 = 第三遍材料(表文字, 记录["科目"], 记录["表名"], 候选)
        目录 = 留证 / f"{侧名}_表{序号:03d}"
        答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复)
        if 答复 is None:
            待答复.append(str(目录))
            continue
        记录.update(_解析精读(答复, {c["id"] for c in 候选 if c.get("id")}))
        记录["原始结构"] = {"合并": _合并们(快照, 记录["工作表"], 记录["范围"])}
        记录["完成"] = True
        _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)

    # 合集：铺格子 + 程序读原值
    名表 = {}
    for 记录 in 本侧:
        for c in 候选槽位(索引, 记录["科目"], 记录["表名"]):
            if c.get("id"):
                名表[c["id"]] = c.get("槽位名")
    格们 = []
    for 记录 in 本侧:
        if not 记录.get("列"):
            continue
        for 格 in 铺格子(记录):
            格["槽位名"] = 名表.get(格["槽位"], "")
            格们.append(格)
    格们 = 补原值(格们, 快照)
    合集路径 = 运行目录 / f"{侧名}合集.jsonl"
    _写jsonl(合集路径, 格们)
    统计["格子"] = len(格们)

    索引状态 = 读索引(运行目录) or {"已完成范围": [], "产物": {}}
    索引状态.setdefault("已完成范围", []).append(f"识别{侧名}")
    索引状态["已完成范围"] = list(dict.fromkeys(索引状态["已完成范围"]))
    索引状态.setdefault("产物", {})[f"{侧名}合集"] = str(合集路径)
    索引状态["产物"]["结构清单"] = str(结构路径)
    写索引(运行目录, 索引状态)
    return {"结构清单": 结构路径, "合集": 合集路径, "统计": 统计, "待答复": 待答复}



# _旧实现已删除：按行切块、拼全部章节上下文、关键词挑金标准片段三处，均于本轮改为三遍线性。


# _按行切块与整章上下文拼接已删除：表按命名区域／行范围整张读，不切断表头。


# _旧入口（识别侧／准备材料／_块材料／_块清单）已由 识别一侧 取代，随三遍流程一并删除。
