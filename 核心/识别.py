# coding: utf-8
"""识别编排：三遍线性——先分科目，再分表，最后逐表精读挂槽位编号。

串行依赖：第二遍读第一遍的科目划分，第三遍读前两遍的科目与表名，全部顺序执行、不并发。
第一、二遍不给任何金标准条目；第三遍只给该表候选槽位。
模型只交结构判断；原值与原值状态由程序读原文件取得，答复里的值字段一律丢弃。
"""
import json
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries, column_index_from_string

from .接续 import 读索引, 写索引
from .表格 import 读快照
from .表格文字 import 表格转文字
from .看表规范 import 分科目规范, 分表规范, 精读规范
from .金标准 import 加载金标准, 建索引, 科目清单, 表清单, 候选槽位, 期间必填
from .模型调用 import 调用识别, 校验指令

_答复名 = "答复.txt"
_数据行角色 = {"数据", "小计", "合计"}


# —— 三遍的指令与材料 ——

_第一遍指令 = ("把整份文件按科目切开。只交哪些行范围属于哪个科目，不交表名、不交行列、不交任何编号。"
           "严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_第一遍格式 = ("严格输出一个 JSON 对象：\n"
           '{"划分":[{"行范围":[起始行,结束行],"科目":"科目原名或非表格"}, ...]}\n'
           "行范围用原文件行号，含首含尾，按行号从小到大；相邻两段首尾相接，覆盖全部有内容的行，不留空档、不重叠。\n"
           "会计政策、会计估计及其变更这一章的表格（折旧年限、摊销年限、组合划分依据、计提比例、税率等），"
           "整段写「会计政策」；这一章不做识别、不搬运。\n"
           "是表但清单里没有对应科目的，写「无对应科目：」再接这张表的原文标题，标题照抄、不改写、不概括，"
           "例如「无对应科目：母公司」「无对应科目：应收关联方款项」；只写「无对应科目」不给标题算不合规。")

_第二遍指令 = ("在第一遍的科目划分内，把每段切成表，并把分页拆开的续表归到同一组。"
           "只交行范围、科目、表名、表组、续表序，不交行列结构、不交任何编号。"
           "严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_第二遍格式 = ("严格输出一个 JSON 对象：\n"
           '{"表":[{"行范围":[起始行,结束行],"科目":"科目原名","表名":"表名原名",'
           '"表组":"同组标识","续表序":1}, ...]}\n'
           "表名只能取自该科目的表名清单；行范围不得跨出第一遍给该科目的行范围。\n"
           "同一张业务表被分页拆成续表时，各段用同一个「表组」，续表序从 1 递增；"
           "互不相干的两张表即使同名，也必须用不同「表组」。")

_第三遍指令 = ("只读这一张表：交列清单、行清单、作用范围、空名单、例外，并给每个数据格挑候选槽位编号。"
           "不交任何数值、不改原文。严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_期间角色集 = {"本期期初", "本期期末", "上期期初", "上期期末", "本期发生", "上期发生", "未知"}
# 期间分两层写时的半截说法：列写本期／上期、行写期初／期末，两处拼成上面七种之一
_期间部件 = {"本期": ("本期", None), "上期": ("上期", None),
          "期初": (None, "期初"), "期末": (None, "期末")}

_第三遍格式 = """严格输出一个 JSON 对象：
{
  "列": [{"列": "B", "路径": "项目", "角色": "行名列", "期间角色": null, "序号": null},
         {"列": "C", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末", "序号": null,
          "标签行": null}],
  "行": [{"行": 56, "名字": "项目", "角色": "表头", "期间角色": null, "分组": null},
         {"行": 57, "名字": "库存现金", "角色": "数据", "序号": null, "期间角色": null,
          "期间实际": null, "分组": null, "槽位": {"C": "C-N087-T001-S001"}}],
  "作用范围": {"币种": "人民币", "单位": "元", "期间角色": "本期", "期间实际": "2025-12-31", "期间跨度": "时点"},
  "空名单": [],
  "例外": [{"范围": "B2", "处理": "不搬", "原因": "合并说明文字，不是数据"}]
}
要求：列字母、行号都用原文件的位置；路径多级用斜杠连接，不许拍平；数据行、小计行、合计行的槽位映射必须逐列给全，挑不到就写 ""；整行都不搬时写 "不搬原因"。
编号只能从候选里挑，不许编。
「期间角色」最终只有七种：本期期初、本期期末、上期期初、上期期末、本期发生、上期发生、未知。
期间要按这一格真实读法给：期间体现在列上就写在「列」里，体现在行上就写在「行」里；行和列都能看出期间时，两处都要写且必须一致。
表里期间分两层写（例如列是“本期金额／上期金额”，行是“期初余额／期末余额”），两处各写自己那一半即可：
列写「本期」「上期」，行写「期初」「期末」，程序按两处拼成七种之一。两处各写完整角色时必须一致，冲突交人工。
只有一处能看出期间时，一律写七种之一（看不出期初还是期末的发生额写「本期发生」「上期发生」），
不许只写「本期」「上期」——那样拼不出时点，只能记「未知」。
「期间实际」写这一格对应的资产负债表日真实日期（如 2025-12-31）；行上列上哪一处看得到就写哪一处，都看不到留空。
「序号」只给名单项：同一编号在本表出现多行（或多列）时，按出现先后给 1、2、3；单值项留空。
名单排在列上时，序号写在「列」里，行上不要给；排在行上时写在「行」里，列上不要给。识别失败的项也要给序号，后面的项不前移。
名单排在列上时，「列」里还要给「标签行」：这一列的类别名称写在第几行。
表头上下两层（一层写类别名称、一层写单位）时，写名称所在那一行，不许写单位行；单层表头就写表头那一行。
小计属于哪个分组，在「行」里用「分组」写明（如"国内""国外"）；同一张表有多个同类小计时必须写。
「例外」只认两种处理：写 "不搬" 表示这一格确定不搬、交人工；写 "合并" 表示范围是从左上角起的一片合并格，只按左上角那一格算。
范围用 Excel 坐标（如 B2、C2:D2），原因写清为什么。不写处理方式、或范围看不懂的，程序会停下来交人工，不会当成普通格硬搬。
整张表本期没有任何数据行时，「空名单」里逐项写明 {"槽位": "编号", "说明": "本期无行"}；有数据行就留空数组。"""


def 第一遍材料(文件文字: str, 科目名们: list) -> tuple:
    材料 = "\n\n".join([
        "【一、整份文件的表格文字】（不裁剪）\n" + 文件文字,
        "【二、科目候选清单】只能照抄这里的原字，不得改写；说明段、目录、纯文字一律写「非表格」；"
        "会计政策、会计估计那一章的表格整段写「会计政策」。\n"
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


def 第三遍材料(表文字: str, 科目: str, 表名: str, 候选: list, 上下文文本: str = "") -> tuple:
    行们 = []
    for c in 候选:
        维度 = c.get("维度") or {}
        段 = [str(c.get("id") or ""), str(c.get("槽位名") or ""),
              "行:" + "/".join(map(str, c.get("行路径") or [])),
              "列:" + "/".join(map(str, c.get("列路径") or [])),
              "类型:" + str(c.get("值类型") or "未知"),
              "期间:" + ("要求" if 维度.get("period") else "不要求")]
        别名 = c.get("别名") or {}
        别名文 = "；".join(str(k) + "=" + "、".join(map(str, v))
                          for k, v in 别名.items() if v)
        if 别名文:
            段.append("别名:" + 别名文)                  # 别名只为看懂，程序不按别名配对
        行们.append(" | ".join(段))
    名单 = "\n".join(行们) if 行们 else "（本表没有候选槽位：这一遍槽位一律留空）"
    材料 = "\n\n".join([
        "【一、本表的表格文字】（只这一张表，不裁剪）\n" + 表文字,
        "【二、本表身份】科目：" + 科目 + "\n表名：" + 表名,
        "【三、本表候选槽位】编号 | 槽位名 | 行路径 | 列路径 | 值类型 | 期间要求 | 别名\n" + 名单,
        "【四、本表上下文】（原文整段，不裁剪；显示“无”表示提取时没关联到本表）\n"
        + (上下文文本 or "（无）"),
        "【五、看表规范】\n" + 精读规范(),
        "【六、输出格式】\n" + _第三遍格式,
        "【七、示例】名字与编号都是假的\n"
        '{"列":[{"列":"B","路径":"项目","角色":"行名列","期间角色":null},'
        '{"列":"C","路径":"期末余额","角色":"数据列","期间角色":"本期期末"}],'
        '"行":[{"行":2,"名字":"项目","角色":"表头"},'
        '{"行":3,"名字":"库存现金","角色":"数据","序号":null,"槽位":{"C":"C-N999-T999-S999"}}],'
        '"作用范围":{"币种":"人民币","单位":"元"},"空名单":[],"例外":[]}',
    ])
    return _第三遍指令, 材料


def _表上下文(上下文, 区域) -> str:
    """按命名区域把本表的上下文整段拼出来：标题、表前原文、单位与日期口径、章节尾段。

    不按关键词筛（“单位”“系指”这类词会漏掉“以下金额均以人民币万元列示”），
    也不拿别的表的上下文来凑；找不到归属就返回空串，由材料里如实写“无”。
    """
    if not 上下文 or not 区域:
        return ""
    if isinstance(上下文, dict):
        表们 = 上下文.get("tables") or []
    elif isinstance(上下文, (list, tuple)):
        表们 = list(上下文)                  # 调用方直接传表格清单时也认
    else:
        return ""
    for 项 in 表们:
        if not isinstance(项, dict):
            continue
        if 项.get("range_name") != 区域:
            continue
        段 = []
        if 项.get("title"):
            段.append(str(项["title"]))
        for b in (项.get("preceding_blocks") or []):
            if isinstance(b, dict) and b.get("text"):
                段.append(str(b["text"]))
        for u in (项.get("unit_text_candidates") or []):
            if u:
                段.append(str(u))
        章 = 项.get("chapter_context")
        if isinstance(章, list):
            段 += [str(x) for x in 章 if x]
        elif 章:
            段.append(str(章))
        return "\n".join(段)
    return ""


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
    if 科目 in 科目集 or 科目 in ("非表格", "会计政策"):
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
        组 = str(表.get("表组") or "").strip()
        if not 组:
            raise RuntimeError(f"第二遍表缺「表组」，无法判断是不是续表：{表}")
        try:
            续 = int(表.get("续表序"))
        except (TypeError, ValueError):
            raise RuntimeError(f"第二遍续表序不合规：{表}") from None
        if 续 < 1:
            raise RuntimeError(f"第二遍续表序必须从 1 起：{表}")
        出.append({"行范围": list(范围), "科目": 科目, "表名": 表名,
                  "表组": 组, "续表序": 续})
    _校验表组(出)
    return 出


def _校验表组(表们: list) -> None:
    """续表必须同组、续表序从 1 连续；同名表分属两组说明归属不明，交人工。"""
    组内 = {}
    for x in 表们:
        组内.setdefault(x["表组"], []).append(x)
    for 组名, 项们 in 组内.items():
        键 = {(x["科目"], x["表名"]) for x in 项们}
        if len(键) > 1:
            raise RuntimeError(f"表组 {组名} 混了不同的表：{sorted(键)}，交人工")
        序们 = sorted(x["续表序"] for x in 项们)
        if 序们 != list(range(1, len(序们) + 1)):
            raise RuntimeError(f"表组 {组名} 的续表序不是从 1 连续：{序们}")
    同名 = {}
    for x in 表们:
        同名.setdefault((x["科目"], x["表名"]), set()).add(x["表组"])
    for (科目, 表名), 组们 in 同名.items():
        if len(组们) > 1:
            raise RuntimeError(f"同科目下同名表分属多个表组，无法判断是不是续表，交人工："
                               f"{科目}／{表名} {sorted(组们)}")


def _序号值(原值, 说明: str):
    if 原值 is None or 原值 == "":
        return None
    try:
        值 = int(原值)
    except (TypeError, ValueError):
        raise RuntimeError(f"第三遍{说明}序号不合规：{原值}") from None
    if 值 < 1:
        raise RuntimeError(f"第三遍{说明}序号必须从 1 起：{值}")
    return 值


def _期间值(原值, 说明: str):
    """期间角色认七种整写法，也认「本期／上期」「期初／期末」这种等着拼的半截写法。"""
    if 原值 in (None, ""):
        return None
    文 = str(原值)
    if 文 not in _期间角色集 and 文 not in _期间部件:
        raise RuntimeError(f"第三遍{说明}期间角色不合规：{原值}"
                           f"（只认七种写法：{'、'.join(sorted(_期间角色集))}；"
                           f"或分两层时的一半：{'、'.join(sorted(_期间部件))}）")
    return 文


def _标签行值(原值, 说明: str, 行范围):
    """列名单的类别名称写在第几行。越界直接判失败，不许写到表外。"""
    if 原值 in (None, ""):
        return None
    try:
        值 = int(原值)
    except (TypeError, ValueError):
        raise RuntimeError(f"第三遍{说明}标签行不合规：{原值}") from None
    if 行范围 is not None and not (行范围[0] <= 值 <= 行范围[1]):
        raise RuntimeError(f"第三遍{说明}标签行越界：{值} 不在本表 {行范围} 内")
    return 值


def _解析精读(答复: str, 编号集: set, 行范围=None, 列范围=None, 续表序: int = 1) -> dict:
    """解析第三遍答复并做本表范围内的校验；任一不过就抛错，不标完成。

    行范围/列范围给定时，坐标越界直接判失败，不许跑到表外；不给则只做格式校验。
    列序号、行期间、小计分组都要原样留住——少了它们，横向名单和期初期末都配不准。
    """
    对象 = _解析对象(答复, "第三遍")
    列们, 行们 = 对象.get("列") or [], 对象.get("行") or []
    if not 列们 or not 行们:
        raise RuntimeError("第三遍答复缺列清单或行清单")
    列范围集 = set(列范围) if 列范围 is not None else None
    列出, 见字母 = [], set()
    for c in 列们:
        字母 = str(c.get("列") or "").strip().upper()
        if not 字母:
            raise RuntimeError(f"第三遍列缺字母：{c}")
        if 字母 in 见字母:
            raise RuntimeError(f"第三遍列字母重复：{字母}")
        if 列范围集 is not None and 字母 not in 列范围集:
            raise RuntimeError(f"第三遍列 {字母} 不在本表列范围 {sorted(列范围集)} 内")
        见字母.add(字母)
        角色 = str(c.get("角色") or "")
        期间角色 = _期间值(c.get("期间角色"), f"列 {字母} 的")
        if 角色 == "数据列" and 期间角色 is None and c.get("期间角色") is not None:
            raise RuntimeError(f"第三遍数据列 {字母} 的期间角色为空")
        列出.append({"列": 字母, "路径": str(c.get("路径") or ""), "角色": 角色,
                   "期间角色": 期间角色, "期间实际": c.get("期间实际") or None,
                   "序号": _序号值(c.get("序号"), f"列 {字母} 的"),
                   "标签行": _标签行值(c.get("标签行"), f"列 {字母} 的", 行范围)})
    行出, 上行 = [], None
    for r in 行们:
        try:
            行号 = int(r.get("行"))
        except (TypeError, ValueError):
            raise RuntimeError(f"第三遍行号不合规：{r}") from None
        if 行范围 is not None and not (行范围[0] <= 行号 <= 行范围[1]):
            raise RuntimeError(f"第三遍行号越界：{行号} 不在本表 {行范围} 内")
        if 上行 is not None and 行号 <= 上行:
            raise RuntimeError(f"第三遍行号重叠或倒序：{行号} 出现在 {上行} 之后")
        上行 = 行号
        槽 = r.get("槽位")
        if not isinstance(槽, dict):
            槽 = {}
        干净 = {}
        for 字母, 编号 in 槽.items():
            字母 = str(字母).strip().upper()
            if 列范围集 is not None and 字母 not in 列范围集:
                raise RuntimeError(f"第三遍坐标越界：列 {字母} 不在本表列范围 {sorted(列范围集)} 内")
            干净[字母] = 编号 if 编号 in 编号集 else ""   # 查不到就留空，不猜
        行出.append({"行": 行号, "名字": str(r.get("名字") or ""),
                   "角色": str(r.get("角色") or ""),
                   "序号": _序号值(r.get("序号"), f"第 {行号} 行的"),
                   "期间角色": _期间值(r.get("期间角色"), f"第 {行号} 行的"),
                   "期间实际": r.get("期间实际") or None,
                   "分组": (str(r.get("分组")).strip() or None) if r.get("分组") not in (None, "") else None,
                   "槽位": 干净, "不搬原因": str(r.get("不搬原因") or "")})
    数据行 = [r for r in 行出 if r["角色"] in _数据行角色]
    for r in 数据行:
        if not r["槽位"] and not r["不搬原因"]:
            raise RuntimeError(f"第三遍第 {r['行']} 行没有交代：既没给槽位映射，也没写不搬原因")
    空名单 = 对象.get("空名单") or []
    if not 数据行 and not 空名单:
        raise RuntimeError("第三遍只交了表头，也没在「空名单」里声明本期无行")
    _校验序号(数据行, 列出, 续表序=int(续表序 or 1))
    return {"列": 列出, "行": 行出,
            "作用范围": 对象.get("作用范围") or {},
            "空名单": 空名单, "例外": 对象.get("例外") or []}



def _校验序号(数据行: list, 列出: list, 续表序: int = 1) -> None:
    """序号只对名单项有意义，按“项”校验，不按某个编号出现几次。

    - 行上有序号 → 行是项；列上有序号 → 列是项；两边都有，本期交人工。
    - 同一侧不能重号，且连续；续表允许不从 1 起（跨页连号由业务表组统一校验）。
    - 汇总格（小计／合计）不是名单项，不参与。
    - 识别失败的项照样有号，序号不会因为缺编号而前移。
    - 名单已经排在列上（列有号）时，行只是期初／期末，不再另要一套行序号。
    """
    明细 = [r for r in 数据行 if r["角色"] == "数据"]
    行序号们 = [r["序号"] for r in 明细 if r.get("序号") is not None]
    列序号们 = [c["序号"] for c in 列出 if c.get("序号") is not None]
    if 行序号们 and 列序号们:
        raise RuntimeError("第三遍同一张表里行和列都给了序号：本期不支持行列同时成名单，交人工")
    if not 列序号们:
        # 行是项：同一编号落在两行、期间又分不开时，光靠编号必混，必须给行序号。
        # 期初、期末各写一行时期间本来就分得开，那两行是一项的两格，不是两个名单项。
        按编号 = {}
        for r in 明细:
            for 编号 in {v for v in (r.get("槽位") or {}).values() if v}:
                按编号.setdefault(编号, []).append(r)
        for 编号, 行们 in 按编号.items():
            角色们 = {str(r.get("期间角色") or "未知") for r in 行们}
            if (len(行们) >= 2 and len(角色们) < len(行们)
                    and any(r.get("序号") is None for r in 行们)):
                raise RuntimeError(f"第三遍名单项缺序号：{编号}（同编号、同期间出现 {len(行们)} 行）")
    序号们 = 行序号们 or 列序号们
    if not 序号们:
        return
    if len(set(序号们)) != len(序号们):
        raise RuntimeError(f"第三遍序号重号：{sorted(序号们)}")
    有序 = sorted(序号们)
    if 续表序 <= 1 and 有序[0] != 1:
        raise RuntimeError(f"第三遍序号必须从 1 起：{有序}")
    if 有序 != list(range(有序[0], 有序[0] + len(有序))):
        raise RuntimeError(f"第三遍序号跳号：{有序}")


# —— 位置与文字 ——

def _校验组序号(本侧: list) -> None:
    """一张业务表可能分页成几张物理表，序号要按整组连号，不能逐张从 1 起。

    占位项（识别失败但有号）算作一项；汇总格不算名单项。
    """
    组 = {}
    for 记录 in 本侧:
        if not 记录.get("列"):
            continue
        键 = 记录.get("表组") or f"{记录.get('科目')}／{记录.get('表名')}"
        组.setdefault(键, []).append(记录)
    for 键, 记录们 in 组.items():
        记录们.sort(key=lambda x: x.get("续表序") or 1)
        序号 = []
        for 记录 in 记录们:
            序号 += [r["序号"] for r in (记录.get("行") or [])
                   if r.get("角色") == "数据" and r.get("序号") is not None]
            序号 += [c["序号"] for c in (记录.get("列") or []) if c.get("序号") is not None]
        if not 序号:
            continue
        if len(set(序号)) != len(序号) or sorted(序号) != list(range(1, len(序号) + 1)):
            raise RuntimeError(f"业务表组「{键}」的序号不连续：{sorted(序号)}")

def _整表范围(表) -> str:
    格们 = 表.get("cells") or {}
    if not 格们:
        return ""
    行们 = [c["row"] for c in 格们.values()]
    列们 = [c["column"] for c in 格们.values()]
    return f"{get_column_letter(min(列们))}{min(行们)}:{get_column_letter(max(列们))}{max(行们)}"


def 全文文字(快照, 只跑表=None) -> str:
    """整个文件：有命名区域就逐区域出，没有就把该工作表整张出。

    只跑表 给定时只读这些工作表；工作表多于一张又没指定时报错，不默认取第一张。
    """
    表们 = 快照["sheets"]
    if 只跑表 is not None:
        要读 = [s for s in 表们 if s["name"] in set(只跑表)]
        if not 要读:
            raise RuntimeError(f"指定的工作表不存在：{'、'.join(map(str, 只跑表))}")
    elif len(表们) > 1:
        raise RuntimeError("有多张工作表，必须先指定要读哪一张")
    else:
        要读 = 表们
    段 = []
    for 表 in 要读:
        区域们 = [] if 只跑表 is not None else [
            (名, 位["range"]) for 名, 位置们 in (快照.get("names") or {}).items()
            for 位 in 位置们 if 位.get("sheet") == 表["name"]]
        if 区域们:
            段 += [表格转文字(快照, {"sheet": 表["name"], "range": str(范围)}) for _, 范围 in 区域们]
        else:
            范围 = _整表范围(表)
            if 范围:
                段.append(表格转文字(快照, {"sheet": 表["name"], "range": 范围}))
    return "\n\n".join(段)


def _定位(快照, 行范围: list, 只跑表=None) -> tuple:
    """返回（命名区域名或 None, 工作表, 精确范围）。最小覆盖该行段的区域优先。

    指定了工作表就不认命名区域——那类文件里的命名多半是书签，不是表边界。
    """
    起, 止 = 行范围
    最优 = None
    for 名, 位置们 in ([] if 只跑表 is not None else (快照.get("names") or {}).items()):
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
    if 只跑表 is not None:
        候选 = [s for s in 工作表们 if s["name"] in set(只跑表)]
        if len(候选) != 1:
            raise RuntimeError(f"指定的工作表不存在或不唯一：{'、'.join(map(str, 只跑表))}")
    elif len(工作表们) == 1:
        候选 = 工作表们
    else:
        raise RuntimeError("没有命名区域且不止一张工作表，无法判断表在哪张工作表上")
    表 = 候选[0]
    列们 = [c["column"] for c in (表.get("cells") or {}).values() if 起 <= c["row"] <= 止]
    左 = get_column_letter(min(列们)) if 列们 else "A"
    右 = get_column_letter(max(列们)) if 列们 else "A"
    return None, 表["name"], f"{左}{起}:{右}{止}"


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

def _例外排除(结构记录: dict) -> tuple:
    """执行第三遍交来的例外，返回（跳过的格, 明确不搬的格, 问题清单, 是否有执行不了的例外）。

    - 「合并」：只按左上角那一格算，其余不铺。
    - 「不搬」：这一格确定不搬，不进合集，记问题（可追溯）。
    - 处理方式缺失、范围读不懂：记问题，并标记“例外未执行”，
      上层据此阻止整张表自动搬运——不许当成普通格硬搬。
    """
    跳过, 不搬, 问题, 未执行 = set(), set(), [], False
    for e in 结构记录.get("例外") or []:
        if not isinstance(e, dict):
            问题.append({"原因": "例外格式不可解析", "原文": str(e)[:200],
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
            未执行 = True
            continue
        范围 = str(e.get("范围") or "").strip()
        处理 = str(e.get("处理") or "").strip()
        原因 = str(e.get("原因") or "").strip()
        if not 范围:
            问题.append({"原因": "例外没写范围", "原文": str(e)[:200],
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
            未执行 = True
            continue
        try:
            左, 顶, 右, 底 = range_boundaries(范围.replace("$", ""))
        except (ValueError, TypeError):
            问题.append({"原因": "例外范围不可解析", "原文": 范围,
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
            未执行 = True
            continue
        格们 = {(c, r) for r in range(顶, 底 + 1) for c in range(左, 右 + 1)}
        if 处理 == "合并":
            跳过 |= {g for g in 格们 if g != (左, 顶)}
        elif 处理 == "不搬":
            不搬 |= 格们
            问题.append({"原因": "识别标明该格不搬，交人工", "范围": 范围, "说明": 原因,
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
        else:
            问题.append({"原因": "例外处理方式不认识，未执行", "范围": 范围, "处理": 处理,
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
            未执行 = True
    return 跳过, 不搬, 问题, 未执行


def _拆期间(值) -> tuple:
    """把期间写法拆成（期间, 时点）：整角色拆两段，半截写法只填一处，未知两处都空。"""
    文 = str(值 or "").strip()
    if 文 in ("", "未知"):
        return (None, None)
    if 文 in _期间部件:
        return _期间部件[文]
    if 文 in _期间角色集:
        return (文[:2], 文[2:])       # 七种整写法都是“本期／上期”＋“期初／期末／发生”
    raise RuntimeError(f"第三遍期间角色不合规：{文}")


def _合期间(行值, 列值) -> str:
    """行与列共同决定这一格的期间角色。

    两处各写一半（列写本期／上期、行写期初／期末）就拼起来；只写期初／期末的按本期算。
    两处给了互相矛盾的一半，说明识别自相矛盾，交人工，不猜。
    """
    行期, 行点 = _拆期间(行值)
    列期, 列点 = _拆期间(列值)
    if 行期 and 列期 and 行期 != 列期:
        raise RuntimeError(f"第三遍同一格的期间冲突：行写「{行值}」，列写「{列值}」")
    if 行点 and 列点 and 行点 != 列点:
        raise RuntimeError(f"第三遍同一格的期间冲突：行写「{行值}」，列写「{列值}」")
    期间, 时点 = 行期 or 列期, 行点 or 列点
    if not 期间 and not 时点:
        return "未知"
    if not 期间:
        return "本期" + 时点          # 表里只有一套期初／期末，按本期记
    if not 时点:
        return "未知"                 # 只有本期／上期，拼不出时点，不猜
    return 期间 + 时点


def _维度值(作用: dict, 列: dict, 行: dict) -> dict:
    return {"期间角色": _合期间(行.get("期间角色"), 列.get("期间角色")),
            "期间实际": 列.get("期间实际") or 行.get("期间实际") or "未知",
            "期间跨度": 作用.get("期间跨度") or "未知",
            "选取范围": 行.get("名字") or "未知",
            "币种": 作用.get("币种") or "未知",
            "主体": 作用.get("主体") or "未知",
            "报表范围": 作用.get("报表范围") or "未知",
            "单位": 作用.get("单位") or "未知"}


def _项方式(数据列: list, 数据行: list) -> str:
    """名单排在列上还是行上。两边都有已在校验阶段拦下，这里只判一次。"""
    if any(c.get("序号") is not None for c in 数据列):
        return "列"
    return "行"


def 铺格子(结构记录: dict, 问题: list = None) -> list:
    """把一张表的结构铺成逐格记录。

    - 每格带「项方式」「项位置」：项是“同一行”还是“同一列”，由识别给的序号方向定，
      不再用物理行号硬套——横向名单的名字和数字也要能一起换。
    - 名单项生成项标签：行上是行标，列上是列标题。
    - 例外「不搬」的格不进合集；执行不了的例外记问题并标记，交给上层阻断搬运。
    """
    工作表 = 结构记录.get("工作表") or "sheet1"
    作用 = 结构记录.get("作用范围") or {}
    数据列 = [c for c in (结构记录.get("列") or []) if c.get("角色") == "数据列"]
    行名列 = next((c for c in (结构记录.get("列") or []) if c.get("角色") == "行名列"), None)
    数据行 = [r for r in (结构记录.get("行") or []) if r.get("角色") in _数据行角色]
    跳过, 不搬, 例外问题, 未执行 = _例外排除(结构记录)
    if 问题 is not None:
        问题.extend(例外问题)
    结构记录["例外未执行"] = bool(未执行)
    项方式 = _项方式(数据列, 数据行)

    def 列序(字母):
        try:
            return column_index_from_string(str(字母 or "").upper())
        except ValueError:
            return 0

    记 = []
    数据行号 = {r["行"] for r in 数据行}
    表头行们 = sorted(r["行"] for r in (结构记录.get("行") or []) if r.get("角色") == "表头")
    首个数据行 = min(数据行号, default=None)
    可用表头 = [h for h in 表头行们 if 首个数据行 is None or h < 首个数据行]
    默认标签行 = 可用表头[0] if len(可用表头) == 1 else None   # 只有一层表头就没有歧义
    列标签行 = {}
    结构记录["项标签缺失"] = False
    if 项方式 == "列":
        # 列上的名单，类别名称在第几行由识别交代；不许拿“最后一层表头”当名称——
        # 表头两层时那一层往往是“单位：元”，写进去就成了旧名字配新数字。
        for 列 in 数据列:
            if 列.get("序号") is None:
                continue
            标签行 = 列.get("标签行") or 默认标签行
            if 标签行 and 标签行 not in 数据行号:
                列标签行[str(列.get("列") or "").upper()] = 标签行
        缺 = [c.get("列") for c in 数据列 if c.get("序号") is not None
              and str(c.get("列") or "").upper() not in 列标签行]
        if 缺:
            结构记录["项标签缺失"] = True          # 名字写不进去，搬运时整表不动
            if 问题 is not None:
                问题.append({"原因": "名单在列上，识别没交代列标题写在那一行（标签行），该表不搬",
                           "列": 缺,
                           "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})

    for 行 in 数据行:
        行号 = 行.get("行")
        行路径 = [x for x in str(行.get("名字") or "").split("/") if x]
        槽表 = 行.get("槽位") or {}
        是汇总 = 行.get("角色") in ("小计", "合计")
        # 行标：行方式下名字随行搬；列方式下的行不是项，行标是身份，不搬
        if 行名列 and 项方式 == "行" and not 是汇总:
            字母 = str(行名列.get("列") or "").upper()
            序 = 列序(字母)
            if 序 and (序, 行号) not in 跳过:
                记.append({
                    "格": f"{工作表}!{字母}{行号}",
                    "科目": 结构记录.get("科目") or "",
                    "表名": 结构记录.get("表名") or "",
                    "表组": 结构记录.get("表组") or "",
                    "行号": 行号,
                    "序号": 行.get("序号"),
                    "项方式": "行", "项位置": 行号,
                    "角色": "项标签",
                    # 名字标了不搬，这一项就不许只换数——失败状态要跟着项走
                    "不搬": (序, 行号) in 不搬,
                    "行路径": 行路径,
                    "列路径": [x for x in str(行名列.get("路径") or "").split("/") if x],
                    "维度值": _维度值(作用, {"期间角色": None}, 行),
                    "槽位": "", "槽位名": "",
                    "原值": 行.get("名字"), "原值状态": "值",
                })
        for 列 in 数据列:
            字母 = str(列.get("列") or "").upper()
            序 = 列序(字母)
            if not 序 or (序, 行号) in 跳过:
                continue                       # 例外点名「合并」覆盖的格不铺
            记.append({
                "格": f"{工作表}!{字母}{行号}",
                "科目": 结构记录.get("科目") or "",
                "表名": 结构记录.get("表名") or "",
                "表组": 结构记录.get("表组") or "",
                "行号": 行号,
                "序号": 行.get("序号") if 项方式 == "行" else 列.get("序号"),
                "项方式": 项方式,
                "项位置": 行号 if 项方式 == "行" else 字母,
                "分组": 行.get("分组"),
                "角色": "汇总格" if 是汇总 else "数据格",
                # 识别说这格不搬：仍是本项要搬的格，留个占位好让整项一起保留
                "不搬": (序, 行号) in 不搬,
                "行路径": 行路径,
                "列路径": [x for x in str(列.get("路径") or "").split("/") if x],
                "维度值": _维度值(作用, 列, 行),
                "槽位": 槽表.get(字母, ""), "槽位名": "",
                "原值": None, "原值状态": "待读",
            })
    # 列方式：列标题也是内容，要随项一起换；写哪一行听识别的「标签行」
    for 列 in 数据列:
        字母 = str(列.get("列") or "").upper()
        序 = 列序(字母)
        标签行 = 列标签行.get(字母)
        if not 序 or not 标签行 or (序, 标签行) in 跳过:
            continue
        记.append({
            "格": f"{工作表}!{字母}{标签行}",
            "科目": 结构记录.get("科目") or "",
            "表名": 结构记录.get("表名") or "",
            "表组": 结构记录.get("表组") or "",
            "行号": 标签行,
            "序号": 列.get("序号"),
            "项方式": "列", "项位置": 字母,
            "角色": "项标签",
            "不搬": (序, 标签行) in 不搬,        # 列标题标了不搬，同样整项不动
            "行路径": [],
            "列路径": [x for x in str(列.get("路径") or "").split("/") if x],
            "维度值": _维度值(作用, 列, {"名字": ""}),
            "槽位": "", "槽位名": "",
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

def _列范围(范围: str) -> set:
    try:
        左, _, 右, _ = range_boundaries(str(范围))
    except (ValueError, TypeError):
        return set()
    return {get_column_letter(i) for i in range(左, 右 + 1)}


def _完整性(划分: list, 本侧: list, 快照) -> dict:
    """B 完不完整，由程序按范围覆盖判，不靠第三遍“校验通过”顶替。

    五条：第一遍覆盖全部有值行；第二遍覆盖真科目段内的有值行；第三遍每张表都做完；
    每张表原范围内实际有值的行都在行清单里有交代（漏一行就不完整）；没有数据行的表必须显式声明空名单。
    任一条不成立，清尾就不许做——「校验通过」只证明交上来的部分自洽，证明不了没漏。
    """
    有值行 = {c["row"] for s in 快照["sheets"] for c in (s.get("cells") or {}).values()}
    段 = [x["行范围"] for x in 划分 if x.get("行范围")]
    if 段 and 有值行:
        第一 = (min(a for a, _ in 段) <= min(有值行) and max(b for _, b in 段) >= max(有值行)
              and all(any(a <= r <= b for a, b in 段) for r in 有值行))
    else:
        第一 = bool(段)
    表段 = [(x["行范围"][0], x["行范围"][1]) for x in 本侧 if x.get("行范围")]
    真科目 = {x["科目"] for x in 本侧}
    真科目段 = [x["行范围"] for x in 划分 if x.get("科目") in 真科目]
    if 真科目段:
        归段 = all(any(a <= r <= b for a, b in 表段)
                  for 段x in 真科目段 for r in 有值行 if 段x[0] <= r <= 段x[1])
    else:
        归段 = bool(本侧)

    # 原表范围内实际有值的行，必须都能在行清单里找到；漏了就说明识别少交了行。
    # 按“整行任一格有值”判：名字还在、最后那个金额格空着，这一行照样存在，不能当成没有。
    有值格 = {s["name"]: [(c["row"], c["column"]) for c in (s.get("cells") or {}).values()
                       if c.get("value") not in (None, "")] for s in 快照["sheets"]}
    漏行, 漏列 = {}, {}
    for x in 本侧:
        范围 = str(x.get("范围") or "")
        if not 范围:
            continue
        try:
            左, 顶, 右, 底 = range_boundaries(范围)
        except (ValueError, TypeError):
            continue
        格们 = [(r, 列号) for r, 列号 in 有值格.get(x.get("工作表"), [])
              if 顶 <= r <= 底 and 左 <= 列号 <= 右]
        交代 = {r.get("行") for r in x.get("行") or []}
        缺 = sorted({r for r, _ in 格们} - 交代)
        if 缺:
            漏行[f"{x.get('科目')}／{x.get('表名')}"] = 缺
            x["漏行"] = 缺

        # 例外只在它真正盖住的那几格上作数，而且要分清两种：
        # 「合并」只有左上角以外的从属格可以不铺；左上角存的是真内容，照样要交代。
        # 「不搬」只说明这一格不交程序，它没有进项清单就等于这一项没有占位，
        # 拿它证明整列“已交代”，就会把识别失败当成“B 没有这一项”去清 A。
        列清单 = {column_index_from_string(str(c.get("列") or "").strip().upper())
               for c in x.get("列") or [] if str(c.get("列") or "").strip()}
        例外框 = []
        for 例 in x.get("例外") or []:
            if str(例.get("处理") or "") != "合并":
                continue
            try:
                例外框.append(range_boundaries(str(例.get("范围") or "")))
            except (ValueError, TypeError):
                continue
        未交代 = {}
        for r, 列号 in 格们:
            if 列号 in 列清单:
                continue
            if any(l <= 列号 <= rr and t <= r <= b and (r, 列号) != (t, l)
                   for l, t, rr, b in 例外框):
                continue                      # 合并的从属格，不算漏；左上角不算
            未交代.setdefault(列号, []).append(r)
        缺列 = [get_column_letter(列号) for 列号 in sorted(未交代)]
        if 缺列:
            漏列[f"{x.get('科目')}／{x.get('表名')}"] = 缺列
            x["漏列"] = 缺列
            x["漏列格"] = {名: sorted(行们) for 名, 行们 in
                       ((get_column_letter(列号), 行们) for 列号, 行们 in 未交代.items())}

    逐行 = (bool(本侧) and all(x.get("完成") for x in 本侧)
          and not 漏行 and not 漏列)
    空名单 = True
    for x in 本侧:
        有数据 = any(r.get("角色") in _数据行角色 for r in x.get("行") or [])
        if not 有数据 and not x.get("空名单"):
            空名单 = False
    return {"第一遍覆盖": bool(第一), "第二遍归段": bool(归段),
            "第三遍逐行": bool(逐行), "空名单已声明": bool(空名单),
            "漏行": 漏行, "漏列": 漏列}



def 识别一侧(配置, 表文件路径, 侧名, 运行目录, 金标准主文件, 上下文=None,
         只用已有答复=False, 取消=None, 只跑表=None) -> dict:
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
            文字 = 全文文字(快照, 只跑表)
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
            区域, 工作表, 范围 = _定位(快照, 表["行范围"], 只跑表)
            本侧.append({"侧": 侧名, "区域": 区域, "工作表": 工作表, "范围": 范围,
                       "行范围": 表["行范围"], "科目": 表["科目"], "表名": 表["表名"],
                       "表组": 表["表组"], "续表序": 表["续表序"],
                       "列": [], "行": [], "作用范围": {}, "空名单": [], "原始结构": {},
                       "例外": [], "完成": False})
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
        指令, 材料 = 第三遍材料(表文字, 记录["科目"], 记录["表名"], 候选,
                           _表上下文(上下文, 记录.get("区域")))
        目录 = 留证 / f"{侧名}_表{序号:03d}"
        答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复)
        if 答复 is None:
            待答复.append(str(目录))
            continue
        记录.update(_解析精读(答复, {c["id"] for c in 候选 if c.get("id")},
                          行范围=记录.get("行范围"), 列范围=_列范围(记录["范围"]),
                          续表序=记录.get("续表序") or 1))
        记录["原始结构"] = {"合并": _合并们(快照, 记录["工作表"], 记录["范围"])}
        记录["完成"] = True
        _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)

    _校验组序号(本侧)

    完整性 = _完整性(划分, 本侧, 快照)
    for 记录 in 本侧:
        记录["完整性"] = 完整性
    _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)

    # 合集：铺格子 + 程序读原值
    名表 = {}
    for 记录 in 本侧:
        for c in 候选槽位(索引, 记录["科目"], 记录["表名"]):
            if c.get("id"):
                名表[c["id"]] = c.get("槽位名")
    格们 = []
    识别问题 = []
    for 记录 in 本侧:
        if not 记录.get("列"):
            continue
        for 格 in 铺格子(记录, 识别问题):
            格["槽位名"] = 名表.get(格["槽位"], "")
            格["期间必填"] = 期间必填(索引, 格["槽位"])
            格们.append(格)
    格们 = 补原值(格们, 快照)
    for 记录 in 本侧:
        for 键, 说法 in (("漏行", "行"), ("漏列", "列")):
            if not 记录.get(键):
                continue
            条目 = {"原因": f"第三遍漏交这些{说法}，该表不清尾，交人工",
                  "科目": 记录.get("科目"), "表名": 记录.get("表名"),
                  说法: 记录[键]}
            if 键 == "漏列" and 记录.get("漏列格"):
                条目["行"] = 记录["漏列格"]       # 列要带上具体哪几行，方便补交
            识别问题.append(条目)
    # 铺格子会判“例外执行不了”“列标签缺失”，这些标记要落盘，配对时才拦得住
    _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)
    合集路径 = 运行目录 / f"{侧名}合集.jsonl"
    _写jsonl(合集路径, 格们)
    _写jsonl(运行目录 / f"识别问题_{侧名}.jsonl", 识别问题)
    统计["格子"] = len(格们)
    统计["识别问题"] = len(识别问题)

    索引状态 = 读索引(运行目录) or {"已完成范围": [], "产物": {}}
    索引状态.setdefault("已完成范围", []).append(f"识别{侧名}")
    索引状态["已完成范围"] = list(dict.fromkeys(索引状态["已完成范围"]))
    索引状态.setdefault("产物", {})[f"{侧名}合集"] = str(合集路径)
    索引状态["产物"]["结构清单"] = str(结构路径)
    索引状态[f"{侧名}完整性"] = 完整性
    写索引(运行目录, 索引状态)
    return {"结构清单": 结构路径, "合集": 合集路径, "统计": 统计, "待答复": 待答复}



# _旧实现已删除：按行切块、拼全部章节上下文、关键词挑金标准片段三处，均于本轮改为三遍线性。


# _按行切块与整章上下文拼接已删除：表按命名区域／行范围整张读，不切断表头。


# _旧入口（识别侧／准备材料／_块材料／_块清单）已由 识别一侧 取代，随三遍流程一并删除。
