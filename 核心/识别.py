# coding: utf-8
"""识别编排：三遍线性——先分科目，再分表，最后逐表精读挂槽位编号。

串行依赖：第二遍读第一遍的科目划分，第三遍读前两遍的科目与表名，全部顺序执行、不并发。
第一、二遍不给任何金标准条目；第三遍只给该表候选槽位。
模型只交结构判断；原值与原值状态由程序读原文件取得，答复里的值字段一律丢弃。

跨任务复用接在三个调用点上：每遍真正发请求之前，先用本次材料核对旧记录的证据；
命中就沿用解释并登记使用，未命中／冲突／版本不符就照原流程重认。调用次数在发请求处记账，
读已有答复或打开索引不算调用。复用只带解释，本次坐标、数值、标题、配对和回读都用本次文件。
"""
import json
from pathlib import Path

from openpyxl.utils import get_column_letter, range_boundaries, column_index_from_string
from openpyxl.utils.cell import coordinate_to_tuple

from .接续 import 读索引, 写索引
from .表格 import 读快照, 取可见集
from .表格文字 import 表格转文字
from .看表规范 import 分科目规范, 分表规范, 精读规范
from .金标准 import 加载金标准, 建索引, 科目清单, 表清单, 候选槽位, 期间必填
from .模型调用 import 调用识别, 校验指令
from .共用记录簿 import 记录簿, 已启用
from . import 复用证据 as 证据模块
from . import 复用接入

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
           "互不相干的两张表即使同名，也必须用不同「表组」。\n"
           "这张表前面那句日期标题需要换年份时，在该表里加 "
           '"标题": {"原文": "本次上下文里的完整原句", "日期": [{"片段": "原句里的日期原文", '
           '"期间角色": "本期期末"}]}；不改就不写这个字段。原文必须与本表上下文逐字一致；'
           "A端为原生Excel且标题在单元格内时，标题另加 \"位置\": \"B1\"，位置是本表实际格坐标；"
           "工作表由程序按本次读取部分确定，不另猜表名。"
           "日期只标出片段和期间角色，不要写换成什么——新标题在两侧配对后按本次 B 的期间日期生成。")

_第三遍指令 = ("只读这一张表：交列清单、行清单、作用范围、空名单、例外，并给每个数据格挑候选槽位编号。"
           "不交任何数值、不改原文。严格输出一个 JSON 对象，不写解释、不加代码块标记。")

_期间角色集 = {"本期期初", "本期期末", "上期期初", "上期期末", "本期发生", "上期发生", "未知"}

_第三遍格式 = """严格输出一个 JSON 对象：
{
  "列": [{"列": "B", "路径": "项目", "角色": "行名列", "期间角色": null, "序号": null},
         {"列": "C", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末", "序号": null,
          "标签行": null}],
  "行": [{"行": 56, "名字": "项目", "角色": "表头", "期间角色": null, "分组": null},
         {"行": 57, "名字": "库存现金", "角色": "数据", "序号": null, "期间角色": null,
          "期间角色按列": null, "期间实际": {"值": "2025-12-31", "位置": "C2"}, "分组": null,
          "槽位": {"C": "C-N087-T001-S001"}}],
  "作用范围": {"币种": "人民币", "单位": "元", "期间角色": "本期", "期间实际": "2025-12-31", "期间跨度": "时点",
    "单位依据": [{"来源":"上下文","位置":["unit_text_candidates",0],"原文":"金额单位：人民币元","片段":"人民币元","币种":"人民币","单位":"元","列":["C"]}]},
  "空名单": [],
  "例外": [{"范围": "B2", "处理": "不搬", "原因": "合并说明文字，不是数据"}]
}
要求：列字母、行号都用原文件的位置；路径多级用斜杠连接，不许拍平；数据行、小计行、合计行的槽位映射必须逐列给全，挑不到就写 ""；整行都不搬时写 "不搬原因"。
编号只能从候选里挑，不许编。
「作用范围」必须给「单位依据」数组，每条写来源（表内/上下文）、位置、完整原文、单位片段、币种、单位和适用的数据列。
例如 {"来源":"表内","位置":"E1","原文":"单位：人民币万元","片段":"人民币万元","币种":"人民币","单位":"万元","列":["C","D"]}。
上下文位置写 ["unit_text_candidates",0]、["chapter_context",0] 或 ["preceding_blocks",0]，索引从0开始，原文逐字照抄。
须判断说明实际管哪些列，不得将注册资本或其他表的单位当成本表单位。每个金额列都要有依据；分列单位分别列明；依据不足或冲突时不要强填，留空并写入例外。
「期间角色」只有七种：本期期初、本期期末、上期期初、上期期末、本期发生、上期发生、未知。
必须写完整的七种之一，不许写「本期」「上期」「期初」「期末」这种半截——程序不做拼合。
期间体现在列上就写在该列的「期间角色」，体现在行上就写在该行的「期间角色」；
一格同时由行和列决定（列是本期金额／上期金额、行是期初余额／期末余额）时，在该行加「期间角色按列」逐列给完整角色，
例如 {"C":"本期期初","D":"上期期初"}；同一行里各列角色不同时必须逐列写全，不许只写一处让程序去拼。
「期间实际」写成 {"值":"2025-12-31","位置":"C2"}：值是这一格对应的资产负债表日真实日期，位置是本表里写着这个日期的格（原坐标）；位置看不出来就只写值、位置留空。行上列上哪一处看得到就写哪一处，都看不到就不写。
有角色文字时，同一行/列项增加「角色依据」如 {"位置":"C1","原文":"期末余额"}，多层表头须标实际角色所在格。
单日期无角色文字时，只有明确独立期间说明才能填写「期间依据」：来源、位置、原文的格式同单位依据，另写「起日片段」「止日片段」和完整「期间角色」。成立日期、担保期限不能作为本期依据；确认不了将期间角色写「未知」。
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
        '"作用范围":{"币种":"人民币","单位":"元","单位依据":['
        '{"来源":"表内","位置":"B1","原文":"单位：人民币元","片段":"人民币元",'
        '"币种":"人民币","单位":"元","列":["C"]}]},"空名单":[],"例外":[]}',
    ])
    return _第三遍指令, 材料


def _表上下文(上下文, 区域) -> str:
    """按命名区域提供本表原始上下文，保留字段名和数组下标，供识别标注依据位置。

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
        return json.dumps({k: 项[k] for k in
                           ("title", "preceding_blocks", "unit_text_candidates", "chapter_context", "context_note")
                           if k in 项}, ensure_ascii=False, indent=2)
    return ""


# —— 文件与答复 ——

def _读jsonl(路径) -> list:
    p = Path(路径)
    if not p.is_file():
        return []
    return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]



def _带重试写入(路径: Path, 文本: str) -> None:
    """同步盘／杀软会短暂占用刚写的文件，重试几次再抛。"""
    import time as _time
    for 次 in range(20):
        try:
            with Path(路径).open("w", encoding="utf-8") as f:
                f.write(文本)
            return
        except PermissionError:
            if 次 == 19:
                raise
            _time.sleep(0.1)


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


def _取答复(目录: Path, 指令: str, 材料: str, 配置, 只用已有答复: bool, 记账=None, 取消=None):
    """有答复就导入；没有就写出材料，允许调用则调用；仍没有返回 None（留待答复）。

    “记账”只在真的要发请求时调用一次（读已有答复、写材料、打开索引都不算调用）。
    """
    if 取消 is not None and 取消.is_set():
        raise InterruptedError("用户中断")
    目录.mkdir(parents=True, exist_ok=True)
    答复路径 = 目录 / _答复名
    校验指令(指令)                                  # 简短指令硬校验，超 2000 字符直接拒
    if 答复路径.is_file():
        return 答复路径.read_text(encoding="utf-8")
    (目录 / "指令.txt").write_text(指令, encoding="utf-8")
    (目录 / "材料.txt").write_text(材料, encoding="utf-8")
    if 只用已有答复 or not 配置:
        return None
    if 记账 is not None:
        记账()
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
        标题 = 表.get("标题")
        if 标题 is not None:
            if not (isinstance(标题, dict)
                    and isinstance(标题.get("原文"), str) and 标题["原文"].strip()):
                raise RuntimeError(f"第二遍标题结论格式不对（要给出本次标题原文）：{表}")
            日期 = 标题.get("日期") or []
            if not isinstance(日期, list):
                raise RuntimeError(f"第二遍标题结论格式不对（日期须为列表）：{表}")
            if not 日期 and not 标题.get("新文"):
                raise RuntimeError(f"第二遍标题只说改、没给日期片段和期间角色：{表}")
            干净 = []
            for 项 in 日期:
                if not (isinstance(项, dict) and str(项.get("片段") or "").strip()):
                    raise RuntimeError(f"第二遍标题日期缺片段：{表}")
                角色 = str(项.get("期间角色") or "")
                if 角色 not in _期间角色集 or 角色 == "未知":
                    raise RuntimeError(f"第二遍标题日期的期间角色不合规：{项}")
                干净.append({"片段": str(项["片段"]).strip(), "期间角色": 角色})
            位置 = 标题.get("位置")
            if 位置 is not None:
                try:
                    coordinate_to_tuple(str(位置).replace("$", "").upper())
                except (TypeError, ValueError):
                    raise RuntimeError(f"第二遍Excel标题位置不是有效单元格：{位置}") from None
                位置 = str(位置).replace("$", "").upper()
            标题 = {"原文": 标题["原文"].strip(), "日期": 干净, "位置": 位置,
                  "旧新文": 标题.get("新文"), "旧新文已忽略": bool(标题.get("新文"))}
        出.append({"行范围": list(范围), "科目": 科目, "表名": 表名,
                  "表组": 组, "续表序": 续, "标题": 标题, "歧义": 表.get("歧义")})
    _校验表组(出)
    return 出


def _校验表组(表们: list) -> None:
    """续表序错误仍拒绝；同名不同组只标记受影响组，由配对前整组保留。"""
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
            for x in 表们:
                if x["表组"] in 组们:
                    x["歧义"] = f"同科目同名表分属多个表组：{科目}／{表名}"


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
    """期间角色只认七种完整写法；半截写法不再由程序拼合，缺完整角色就判不合规。"""
    if 原值 in (None, ""):
        return None
    文 = str(原值)
    if 文 not in _期间角色集:
        raise RuntimeError(f"第三遍{说明}期间角色不合规：{原值}"
                           f"（只认七种完整写法：{'、'.join(sorted(_期间角色集))}）")
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


def _期间实际(原值, 说明: str, 行范围=None, 列范围集=None) -> tuple:
    """期间实际及其在原表的写入位置。值可以只给字面，位置必须落在本表范围内。

    位置是“首次识别已标明用途的日期格”，复用与标题生成只认这个标注；没标就留空，
    程序不按内容猜哪一格是日期。
    """
    if 原值 in (None, ""):
        return None, None
    值, 位置 = 原值, None
    if isinstance(原值, dict):
        值, 位置 = 原值.get("值"), 原值.get("位置")
    if 值 in (None, ""):
        raise RuntimeError(f"第三遍{说明}期间实际缺值")
    文 = str(值).strip()
    if 位置 in (None, ""):
        return 文, None
    坐标 = str(位置).strip().replace("$", "").upper()
    try:
        行, 列 = coordinate_to_tuple(坐标)
    except (TypeError, ValueError):
        raise RuntimeError(f"第三遍{说明}期间实际位置不是有效单元格：{位置}") from None
    if 行范围 is not None and not (行范围[0] <= 行 <= 行范围[1]):
        raise RuntimeError(f"第三遍{说明}期间实际位置越界：{位置}")
    if 列范围集 is not None and get_column_letter(列) not in 列范围集:
        raise RuntimeError(f"第三遍{说明}期间实际位置越界：{位置}")
    return 文, 坐标

def _校验格期间(列们, 行们, 期间必填编号):
    for 行 in 行们:
        if 行.get("角色") not in _数据行角色:
            continue
        for 列 in 列们:
            字母 = 列.get("列")
            if 列.get("角色") != "数据列" or (行.get("槽位") or {}).get(字母) not in 期间必填编号:
                continue
            期间 = ((行.get("期间角色按列") or {}).get(字母)
                  or 行.get("期间角色") or 列.get("期间角色"))
            if 期间 not in _期间角色集:
                raise RuntimeError(f"第 {行.get('行')} 行 {字母} 列期间角色缺失或不合规")


def _解析精读(答复: str, 编号集: set, 行范围=None, 列范围=None, 续表序: int = 1,
         期间必填编号=None) -> dict:
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
        期间实际, 期间位置 = _期间实际(c.get("期间实际"), f"列 {字母} 的", 行范围, 列范围集)
        列出.append({"列": 字母, "路径": str(c.get("路径") or ""), "角色": 角色,
                   "角色依据": c.get("角色依据"), "期间依据": c.get("期间依据"),
                   "期间角色": 期间角色, "期间实际": 期间实际, "期间实际位置": 期间位置,
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
        按列 = {}
        for 字母, 值 in (r.get("期间角色按列") or {}).items():
            字母 = str(字母).strip().upper()
            if 列范围集 is not None and 字母 not in 列范围集:
                raise RuntimeError(f"第三遍第 {行号} 行逐格期间越界：列 {字母} 不在本表列范围内")
            按列[字母] = _期间值(值, f"第 {行号} 行 {字母} 列的")
        行期间, 行期间位置 = _期间实际(r.get("期间实际"), f"第 {行号} 行的", 行范围, 列范围集)
        行出.append({"行": 行号, "名字": str(r.get("名字") or ""),
                   "角色依据": r.get("角色依据"), "期间依据": r.get("期间依据"),
                   "角色": str(r.get("角色") or ""),
                   "序号": _序号值(r.get("序号"), f"第 {行号} 行的"),
                   "期间角色": _期间值(r.get("期间角色"), f"第 {行号} 行的"),
                   "期间角色按列": 按列,
                   "期间实际": 行期间, "期间实际位置": 行期间位置,
                   "分组": (str(r.get("分组")).strip() or None) if r.get("分组") not in (None, "") else None,
                   "槽位": 干净, "不搬原因": str(r.get("不搬原因") or "")})
    数据行 = [r for r in 行出 if r["角色"] in _数据行角色]
    for r in 数据行:
        if not r["槽位"] and not r["不搬原因"]:
            raise RuntimeError(f"第三遍第 {r['行']} 行没有交代：既没给槽位映射，也没写不搬原因")
    空名单 = 对象.get("空名单") or []
    if not isinstance(空名单, list) or any(
            not isinstance(项, dict) or 项.get("槽位") not in 编号集 for 项 in 空名单):
        raise RuntimeError("第三遍空名单声明的槽位编号无效")
    if not 数据行 and not 空名单:
        raise RuntimeError("第三遍只交了表头，也没在「空名单」里声明本期无行")
    _校验格期间(列出, 行出, 期间必填编号 or set())
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

        def 期间签名(r):
            集 = {str(v) for v in (r.get("期间角色按列") or {}).values() if v}
            if r.get("期间角色"):
                集.add(str(r["期间角色"]))
            return frozenset(集) or frozenset({"未知"})

        for 编号, 行们 in 按编号.items():
            签名 = [期间签名(r) for r in 行们]
            if (len(行们) >= 2 and len(set(签名)) < len(行们)
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


def 全文文字(快照, 只跑表=None, 用可见范围=False) -> str:
    """整个文件：有命名区域就逐区域出，没有就把该工作表整张出。

    只跑表 给定时只读这些工作表；工作表多于一张又没指定时报错，不默认取第一张。
    用可见范围=True 时（原生Excel两端）只呈现本次可见集合里的位置，隐藏与打印区域外不出现。
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
        if 用可见范围:
            可见 = 取可见集(快照, 表["name"])
            if 可见 is None:
                raise RuntimeError(f"工作表 {表['name']} 没有可见范围记录")
            if 可见.空():
                raise RuntimeError(f"工作表 {表['name']} 没有可见内容：{可见.原因 or '可见范围为空白'}")
            范围们 = [str(范围) for _, 范围 in 区域们] or [
                f"{get_column_letter(min(可见.列们()))}{min(可见.行们())}:"
                f"{get_column_letter(max(可见.列们()))}{max(可见.行们())}"]
            for 范围 in 范围们:
                段.append(表格转文字(快照, {"sheet": 表["name"], "range": 范围}, 用可见范围=True))
            continue
        if 区域们:
            段 += [表格转文字(快照, {"sheet": 表["name"], "range": str(范围)}) for _, 范围 in 区域们]
        else:
            范围 = _整表范围(表)
            if 范围:
                段.append(表格转文字(快照, {"sheet": 表["name"], "range": 范围}))
    return "\n\n".join(段)


def _定位(快照, 行范围: list, 只跑表=None, 用可见范围=False) -> tuple:
    """返回（命名区域名或 None, 工作表, 精确范围）。最小覆盖该行段的区域优先。

    指定了工作表就不认命名区域——那类文件里的命名多半是书签，不是表边界。
    用可见范围=True 时列范围只取该行段里可见的列，不把隐藏列算进表边界。
    """
    起, 止 = 行范围
    最优 = None
    for 名, 位置们 in ([] if 只跑表 is not None or 用可见范围 else (快照.get("names") or {}).items()):
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
    if 用可见范围:
        可见 = 取可见集(快照, 表["name"])
        if 可见 is None or 可见.空():
            raise RuntimeError(f"工作表 {表['name']} 没有可见内容，不能定位表范围")
        列们 = set()
        for r in range(起, 止 + 1):
            for 左段, 右段 in 可见.行段.get(r, ()):
                列们.update(range(左段, 右段 + 1))
        if not 列们:
            raise RuntimeError(f"{表['name']} 第 {起}–{止} 行没有可见单元格，不能定位表范围")
        左 = get_column_letter(min(列们))
        右 = get_column_letter(max(列们))
        return None, 表["name"], f"{左}{起}:{右}{止}"
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


def _维度值(作用: dict, 列: dict, 行: dict, 字母=None) -> dict:
    """期间角色按「逐格 > 整行 > 整列」取更具体的那一个，不拼、不补默认值。"""
    逐格 = 行.get("期间角色按列") or {}
    口径 = (作用.get("单位按列") or {}).get(字母, 作用)
    return {"期间角色": (逐格.get(字母) or 行.get("期间角色")
                     or 列.get("期间角色") or "未知"),
            "期间实际": 列.get("期间实际") or 行.get("期间实际") or "未知",
            "期间跨度": 作用.get("期间跨度") or "未知",
            "选取范围": 行.get("名字") or "未知",
            "币种": 口径.get("币种") or "未知",
            "主体": 作用.get("主体") or "未知",
            "报表范围": 作用.get("报表范围") or "未知",
            "单位": 口径.get("单位") or "未知"}


def _项方式(数据列: list, 数据行: list) -> str:
    """名单排在列上还是行上。两边都有已在校验阶段拦下，这里只判一次。"""
    if any(c.get("序号") is not None for c in 数据列):
        return "列"
    return "行"


def 铺格子(结构记录: dict, 问题: list = None, 可见集=None) -> list:
    """把一张表的结构铺成逐格记录。

    - 每格带「项方式」「项位置」：项是“同一行”还是“同一列”，由识别给的序号方向定，
      不再用物理行号硬套——横向名单的名字和数字也要能一起换。
    - 名单项生成项标签：行上是行标，列上是列标题。
    - 例外「不搬」的格不进合集；执行不了的例外记问题并标记，交给上层阻断搬运。
    - 给了可见集（原生Excel）时，越出集合的坐标不铺，并标记“越界解释”整表不搬。
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
    结构记录["越界解释"] = False
    项方式 = _项方式(数据列, 数据行)

    def 列序(字母):
        try:
            return column_index_from_string(str(字母 or "").upper())
        except ValueError:
            return 0

    def 可见(列号, 行号):
        if 可见集 is None:
            return True
        return 可见集.含(行号, 列号)

    越界 = []

    def 记越界(字母, 行号):
        越界.append(f"{字母}{行号}")

    记 = []
    数据行号 = {r["行"] for r in 数据行}
    列标签行 = {}
    结构记录["项标签缺失"] = False
    if 项方式 == "列":
        # 列上的名单，类别名称在第几行由识别交代；不许拿“最后一层表头”当名称——
        # 表头两层时那一层往往是“单位：元”，写进去就成了旧名字配新数字。
        for 列 in 数据列:
            if 列.get("序号") is None:
                continue
            标签行 = 列.get("标签行")
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
            if 序 and not 可见(序, 行号):
                记越界(字母, 行号)
            elif 序 and (序, 行号) not in 跳过:
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
            if not 可见(序, 行号):
                记越界(字母, 行号)              # 越出本次可见集合：拒绝这一格的解释
                continue
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
                "维度值": _维度值(作用, 列, 行, 字母),
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
        if 可见集 is not None and not 可见(序, 标签行):
            记越界(字母, 标签行)
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
    if 越界:
        结构记录["越界解释"] = True
        if 问题 is not None:
            问题.append({"原因": "识别给出的格不在本次可见范围（隐藏或打印区域外），该表不搬",
                       "范围": sorted(set(越界))[:20],
                       "科目": 结构记录.get("科目"), "表名": 结构记录.get("表名")})
    return 记


def 补原值(记录们: list, 快照, 用可见范围=False) -> list:
    索引 = {f"{表['name']}!{坐标}": 格
          for 表 in 快照["sheets"] for 坐标, 格 in (表.get("cells") or {}).items()}
    可见表 = {表["name"]: 取可见集(快照, 表["name"]) for 表 in 快照["sheets"]} if 用可见范围 else {}
    出 = []
    for 记录 in 记录们:
        格 = 索引.get(str(记录.get("格")))
        if 格 is None:
            出.append({**记录, "原值": None, "原值状态": "读取失败"})
            continue
        可见 = 可见表.get(str(记录.get("格")).rsplit("!", 1)[0])
        if 用可见范围 and (可见 is None or not 可见.含(格["row"], 格["column"])):
            出.append({**记录, "原值": None, "原值状态": "不可见"})
            continue
        if 格.get("formula"):
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


def _完整性(划分: list, 本侧: list, 快照, 用可见范围=False) -> dict:
    """B 完不完整，由程序按范围覆盖判，不靠第三遍“校验通过”顶替。

    五条：第一遍覆盖全部有值行；第二遍覆盖真科目段内的有值行；第三遍每张表都做完；
    每张表原范围内实际有值的行都在行清单里有交代（漏一行就不完整）；没有数据行的表必须显式声明空名单。
    任一条不成立，清尾就不许做——「校验通过」只证明交上来的部分自洽，证明不了没漏。
    用可见范围=True 时，隐藏与打印区域外的格不算有值格，也不参与漏行漏列。
    """
    本侧 = [x for x in 本侧 if not x.get("跳过")]   # 粗比跳过的表不在验收范围：它们本来就没做
    # 有值行只算这次参与识别的工作表：工作簿里多放一张无关表，不能拿它的行号来判完整
    用表 = {x.get("工作表") for x in 本侧 if x.get("工作表")}
    表们 = [s for s in 快照["sheets"] if not 用表 or s.get("name") in 用表] or 快照["sheets"]
    可见们 = {s["name"]: 取可见集(快照, s["name"]) for s in 表们} if 用可见范围 else {}

    def 见(名, 行, 列):
        集 = 可见们.get(名)
        return 集 is None or 集.含(行, 列)

    有值行 = {c["row"] for s in 表们 for c in (s.get("cells") or {}).values()
           if 见(s.get("name"), c["row"], c["column"])}
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
                       if c.get("value") not in (None, "") and 见(s.get("name"), c["row"], c["column"])]
           for s in 快照["sheets"]}
    漏行, 漏列 = {}, {}
    for x in 本侧:
        x.pop("漏行", None)                  # 每轮重算：旧值必须清掉，只增不清会一直背着假漏行
        x.pop("漏列", None)
        x.pop("漏列格", None)
        if not x.get("完成"):
            continue                         # 还没答复的表不参与，它不是“漏交”
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
         只用已有答复=False, 取消=None, 只跑表=None, 允许科目=None, 簿=None,
         材料类型=None, 根任务=None, 读取部分=None) -> dict:
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    快照 = 读快照(表文件路径, 上下文)
    材料类型 = 材料类型 or ("excel" if 侧名 == "B" else "word")
    if 材料类型 not in ("word", "excel"):
        raise ValueError(f"未知材料类型：{材料类型}")
    用可见范围 = 材料类型 == "excel"
    索引 = 建索引(加载金标准(金标准主文件)) if 金标准主文件 else {"科目": {}, "编号": set()}
    留证 = 运行目录 / "调用留证"
    统计 = {"科目划分": 0, "表数": 0, "格子": 0}
    文字 = None
    金标准版本 = 证据模块.内容摘要(金标准主文件) if 金标准主文件 else ""
    规则版本 = 证据模块.规则摘要()
    阅 = 复用接入.取复用(簿) if (簿 is not None and 已启用()) else {
        "可用": False, "原因": "跨任务复用已关闭", "索引": {}, "秒": 0.0}
    任务标识 = str(根任务 or 运行目录)
    工作表名 = (快照["sheets"][0]["name"] if len(快照["sheets"]) == 1
             else (只跑表[0] if 只跑表 else None))
    标题们 = 复用接入.区域标题们(上下文)
    复用记录 = {"侧": 侧名, "任务": 任务标识,
             "读取部分": (读取部分 or {}).get("标识"), "工作表": 工作表名, "材料类型": 材料类型,
             "库": {"可用": 阅["可用"], "原因": 阅["原因"], "代次": (阅["索引"] or {}).get("代次")},
             "调用": {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0},
             "复用": [], "回退": [], "跳过": [], "首认": [], "发布": [],
             "耗时": {"索引": round(阅["秒"], 4), "查找": 0.0, "发布": 0.0}}

    def 记账(遍, 组=None, 重试=False):
        复用记录["调用"][遍] = 复用记录["调用"].get(遍, 0) + 1
        if 重试:
            复用记录["调用"]["重试"] += 1

    def 记回退(步骤, 组, 原因):
        if 原因 and ("没有同侧" in 原因 or "没有同结构记录" in 原因):
            复用记录["首认"].append({"步骤": 步骤, "表组": 组, "说明": 原因})
        else:
            复用记录["回退"].append({"步骤": 步骤, "表组": 组, "原因": 原因 or "未命中"})

    def 登记本次(版本, 组):
        return 复用接入.登记使用(簿, 版本, 任务标识, 侧名, 组,
                              (阅["索引"] or {}).get("代次"),
                              读取部分=(读取部分 or {}).get("标识"), 工作表=工作表名,
                              原文件类型=材料类型)

    def 来源本次(组):
        return {"任务": 任务标识, "侧": 侧名, "表组": 组,
                "读取部分": (读取部分 or {}).get("标识"), "工作表": 工作表名,
                "原文件类型": 材料类型}

    def 共用文字():
        nonlocal 文字
        if 文字 is None:
            文字 = 全文文字(快照, 只跑表, 用可见范围)
        return 文字

    # 第一遍：分科目
    划分路径 = 运行目录 / f"科目划分_{侧名}.jsonl"
    划分 = [r for r in _读jsonl(划分路径) if r.get("侧") == 侧名]
    划分来源 = "已存在" if 划分 else ""
    if not 划分:
        import time as _time
        起 = _time.perf_counter()
        复用 = 复用接入.试复用科目(阅, 侧名, 快照, 用可见范围, 工作表名, 标题们,
                              金标准版本, 规则版本)
        复用记录["耗时"]["查找"] += _time.perf_counter() - 起
        命中 = False
        if 复用["结果"] == "唯一":
            登记, 因 = 登记本次(复用["版本"], "整侧")
            if 登记:
                划分 = [{"侧": 侧名, **x} for x in 复用["数据"]]
                _写jsonl(划分路径, 划分)
                复用记录["复用"].append({"步骤": "第一遍", "表组": "整侧", "版本": 复用["版本"]})
                划分来源 = "复用"
                命中 = True
            else:
                记回退("第一遍", "整侧", 因)
        else:
            记回退("第一遍", "整侧", 复用["原因"])
        if not 命中:
            指令, 材料 = 第一遍材料(共用文字(), 科目清单(索引))
            目录 = 留证 / f"{侧名}_第一遍"
            答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复,
                       记账=lambda: 记账("第一遍"), 取消=取消)
            if 答复 is None:
                return {"结构清单": None, "合集": None, "统计": 统计, "待答复": [str(目录)]}
            划分 = [{"侧": 侧名, **x} for x in _解析划分(答复, set(科目清单(索引)))]
            _写jsonl(划分路径, 划分)
            划分来源 = "首认"
    统计["科目划分"] = len(划分)
    if 取消 is not None and 取消.is_set():
        raise InterruptedError("用户中断")

    # 第二遍：分表
    结构路径 = 运行目录 / "结构清单.jsonl"
    全部 = _读jsonl(结构路径)
    本侧 = [r for r in 全部 if r.get("侧") == 侧名]
    # 完成状态不代表编号仍有效；拒绝旧缓存，保留原结构和答复供重新识别。
    for 记录 in 本侧:
        if not 记录.get("完成"):
            continue
        旧号 = {编号 for 行 in 记录.get("行") or []
              for 编号 in (行.get("槽位") or {}).values() if 编号}
        声明 = 记录.get("空名单") or []
        if not isinstance(声明, list) or any(
                not isinstance(项, dict) or 项.get("槽位") not in 索引["编号"] for 项 in 声明):
            raise RuntimeError("旧任务含失效空名单编号；请新建任务重新识别，原答复保留")
        旧号.update(项["槽位"] for 项 in 声明)
        失效 = sorted(旧号 - 索引["编号"])
        if 失效:
            raise RuntimeError("旧任务含失效槽位编号：" + "、".join(失效)
                               + "；请新建任务重新识别，原任务与答复保留")
        try:
            _校验格期间(记录.get("列") or [], 记录.get("行") or [],
                      {编号 for 编号 in 旧号 if 期间必填(索引, 编号)})
        except RuntimeError as e:
            raise RuntimeError(f"旧任务期间结论不完整：{e}；请新建任务重新识别，原答复保留") from e
    分表来源 = "已存在" if 本侧 else ""
    if not 本侧:
        import time as _time
        起 = _time.perf_counter()
        复用 = 复用接入.试复用分表(阅, 侧名, 快照, 划分, 用可见范围, 工作表名, 标题们,
                              金标准版本, 规则版本)
        复用记录["耗时"]["查找"] += _time.perf_counter() - 起
        命中 = False
        if 复用["结果"] == "唯一":
            登记, 因 = 登记本次(复用["版本"], "整侧")
            if 登记:
                本侧, 组名 = [], {}
                for 表 in 复用["数据"]:
                    区域, 工作表, 范围 = _定位(快照, 表["行范围"], 只跑表, 用可见范围)
                    旧组 = str(表.get("表组") or "")
                    组名.setdefault(旧组, f"复用{len(组名) + 1}")
                    本侧.append({"侧": 侧名, "区域": 区域, "工作表": 工作表, "范围": 范围,
                               "行范围": 表["行范围"], "科目": 表["科目"], "表名": 表["表名"],
                               "表组": 组名[旧组], "续表序": 表["续表序"], "标题": None,
                               "标题规则": 表.get("标题规则") or {},
                               "列": [], "行": [], "作用范围": {}, "空名单": [], "原始结构": {},
                               "例外": [], "完成": False})
                _写jsonl(结构路径, [r for r in 全部 if r.get("侧") != 侧名] + 本侧)
                复用记录["复用"].append({"步骤": "第二遍", "表组": "整侧", "版本": 复用["版本"]})
                分表来源 = "复用"
                命中 = True
            else:
                记回退("第二遍", "整侧", 因)
        else:
            记回退("第二遍", "整侧", 复用["原因"])
        if not 命中:
            表名册 = 待表科目(划分, 索引)
            指令, 材料 = 第二遍材料(共用文字(), 划分, 表名册)
            目录 = 留证 / f"{侧名}_第二遍"
            答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复,
                       记账=lambda: 记账("第二遍"), 取消=取消)
            if 答复 is None:
                return {"结构清单": None, "合集": None, "统计": 统计, "待答复": [str(目录)]}
            本侧 = []
            for 表 in _解析表(答复, 表名册, 划分):
                区域, 工作表, 范围 = _定位(快照, 表["行范围"], 只跑表, 用可见范围)
                本侧.append({"侧": 侧名, "区域": 区域, "工作表": 工作表, "范围": 范围,
                           "行范围": 表["行范围"], "科目": 表["科目"], "表名": 表["表名"],
                           "表组": 表["表组"], "续表序": 表["续表序"], "标题": 表.get("标题"),
                           "列": [], "行": [], "作用范围": {}, "空名单": [], "原始结构": {},
                           "例外": [], "完成": False})
            _写jsonl(结构路径, [r for r in 全部 if r.get("侧") != 侧名] + 本侧)
            分表来源 = "首认"
    # 分表已明确范围后补齐内部空白；仍由同一可见集扣除打印外及隐藏位置。
    if 用可见范围 and 本侧:
        补齐快照 = 读快照(表文件路径, 上下文, 业务范围=[
            {"sheet": r["工作表"], "range": r["范围"]} for r in 本侧])
        if 补齐快照["sha256"] != 快照["sha256"]:
            raise RuntimeError("原文件在识别期间发生变化，请新建任务重新识别")
        快照 = 补齐快照
    # 第二遍记录就绪后统一补齐标题一次：本次已有结论保留；复用来的规则按本次原文重建，
    # 不能确认就保持无标题，由标题生成阶段按既有保护撤回。第三遍不得再补或盖标题。
    for 表记录 in 本侧:
        if isinstance(表记录.get("标题"), dict):
            continue
        规则 = 表记录.get("标题规则") or {}
        if not 规则.get("有标题"):
            continue
        表记录["标题"] = 复用接入.按规则重建标题(表记录, 规则, 快照, 用可见范围, 上下文)
    统计["表数"] = len(本侧)

    # 粗比门槛：科目不在 A 侧出现过的表不做第三遍（B 独有科目，识别了也配不上）。
    # 只跳过精读，表组、行范围原样保留在结构清单里，可追溯、可人工复核。
    if 允许科目 is not None:
        改 = False
        for 记录 in 本侧:
            if not 记录.get("跳过") and 记录.get("科目") not in 允许科目:
                记录["跳过"] = "粗比：科目不在A"
                改 = True
        if 改:
            _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)

    # 第三遍：按业务表组精读（整组复用或整组重认，不逐格拼接新旧解释）
    待答复 = []
    序号表 = {id(x): i for i, x in enumerate(本侧, 1)}
    组们 = {}
    for 记录 in 本侧:
        组们.setdefault(记录.get("表组") or f"{记录.get('科目')}／{记录.get('表名')}", []).append(记录)
    for 组键, 记录们 in 组们.items():
        未完成 = [x for x in 记录们 if not x.get("完成")]
        if not 未完成:
            continue
        跳过因 = next((x.get("跳过") for x in 记录们 if x.get("跳过")), None)
        if 跳过因:
            复用记录["跳过"].append({"步骤": "第三遍", "表组": 组键, "原因": 跳过因})
            continue
        if 取消 is not None and 取消.is_set():
            raise InterruptedError("用户中断")
        import time as _time
        起 = _time.perf_counter()
        复用 = 复用接入.试复用精读(阅, 侧名, 快照, 记录们, 索引, 用可见范围,
                              金标准版本, 规则版本, 上下文)
        复用记录["耗时"]["查找"] += _time.perf_counter() - 起
        命中 = False
        if 复用["结果"] == "唯一":
            登记, 因 = 登记本次(复用["版本"], 组键)
            if 登记:
                复用接入.应用精读(记录们, 复用["数据"], 复用["版本"], 快照, 上下文)
                复用记录["复用"].append({"步骤": "第三遍", "表组": 组键, "版本": 复用["版本"]})
                _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)
                命中 = True
            else:
                记回退("第三遍", 组键, 因)
        else:
            记回退("第三遍", 组键, 复用["原因"])
        if 命中:
            continue
        for 记录 in 未完成:
            if 取消 is not None and 取消.is_set():
                raise InterruptedError("用户中断")
            表文字 = 表格转文字(快照, {"sheet": 记录["工作表"], "range": 记录["范围"]}, 用可见范围)
            候选 = 候选槽位(索引, 记录["科目"], 记录["表名"])
            指令, 材料 = 第三遍材料(表文字, 记录["科目"], 记录["表名"], 候选,
                               _表上下文(上下文, 记录.get("区域")))
            目录 = 留证 / f"{侧名}_表{序号表[id(记录)]:03d}"
            答复 = _取答复(目录, 指令, 材料, 配置, 只用已有答复,
                       记账=lambda: 记账("第三遍", 组键), 取消=取消)
            if 答复 is None:
                待答复.append(str(目录))
                continue
            记录.update(_解析精读(答复, {c["id"] for c in 候选 if c.get("id")},
                              行范围=记录.get("行范围"), 列范围=_列范围(记录["范围"]),
                              续表序=记录.get("续表序") or 1,
                              期间必填编号={c["id"] for c in 候选 if (c.get("维度") or {}).get("period")}))
            记录["原始结构"] = {"合并": _合并们(快照, 记录["工作表"], 记录["范围"])}
            记录["完成"] = True
            _写jsonl(结构路径, [r for r in _读jsonl(结构路径) if r.get("侧") != 侧名] + 本侧)

    _校验组序号(本侧)

    完整性 = _完整性(划分, 本侧, 快照, 用可见范围)
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
        if not 记录.get("复用版本"):
            期间因 = 证据模块.确认本次期间(记录, 快照, 上下文, 用可见范围)
            if 期间因:
                记录["期间问题"] = 期间因
                识别问题.append({"原因": 期间因, "科目": 记录.get("科目"), "表名": 记录.get("表名")})
            单位因 = 证据模块.确认本次单位(记录, 快照, 上下文, 用可见范围)
            if 单位因:
                记录["单位问题"] = 单位因
                识别问题.append({"原因": 单位因, "科目": 记录.get("科目"), "表名": 记录.get("表名")})
        可见 = 取可见集(快照, 记录["工作表"]) if 用可见范围 else None
        for 格 in 铺格子(记录, 识别问题, 可见集=可见):
            格["槽位名"] = 名表.get(格["槽位"], "")
            格["期间必填"] = bool(格["槽位"]) and 期间必填(索引, 格["槽位"])
            格们.append(格)
    已读格 = {f"{表['name']}!{格}" for 表 in 快照["sheets"] for 格 in 表.get("cells") or {}}
    if any(格["格"] not in 已读格 for 格 in 格们):
        范围们 = [{"sheet": 表["工作表"], "range": 表["范围"]}
                 for 表 in 本侧 if 表.get("列") and 表.get("范围")]
        补齐快照 = 读快照(表文件路径, 上下文, 业务范围=范围们)
        if 补齐快照["sha256"] != 快照["sha256"]:
            raise RuntimeError("原文件在识别期间发生变化，请新建任务重新识别")
        快照 = 补齐快照
    格们 = 补原值(格们, 快照, 用可见范围=用可见范围)
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
    # —— 分步发布：上游布局各自覆盖完整才发布；第三遍只交候选，配对后再定 ——
    标题映射 = {(x.get("区域") or ""): (x.get("标题") or "") for x in 标题们}
    import time as _time
    发布起 = _time.perf_counter()
    候选发布 = []
    科目版本 = [x["版本"] for x in 复用记录["复用"] if x.get("版本") and x["步骤"] == "第一遍"]
    分表版本 = [x["版本"] for x in 复用记录["复用"] if x.get("版本") and x["步骤"] == "第二遍"]
    if 簿 is not None and 阅["可用"]:
        if 划分来源 == "首认" and 完整性.get("第一遍覆盖"):
            证据 = 证据模块.建证据_分科目(侧名, 划分, 快照, 用可见范围, 工作表名,
                                      本侧, 标题映射)
            干净划分 = [{k: v for k, v in x.items() if k != "侧"} for x in 划分]
            成功, 值 = 复用接入.发布记录(簿, 侧名, "第一遍", 证据, {"划分": 干净划分},
                                    金标准版本, 规则版本, 来源本次("整侧"), 任务标识,
                                    依赖=[])
            if 成功:
                科目版本.append(值)
            复用记录["发布"].append({"步骤": "第一遍", "状态": "已发布" if 成功 else "未发布",
                                "版本": 值 if 成功 else None, "原因": "" if 成功 else 值})
        if 分表来源 == "首认" and 完整性.get("第二遍归段"):
            证据 = 证据模块.建证据_分表(侧名, 本侧, 快照, 用可见范围, 工作表名, 标题映射)
            组序 = {}
            for x in 本侧:
                组序.setdefault(x.get("表组"), len(组序) + 1)
            表清单 = [{"科目": x.get("科目"), "表名": x.get("表名"),
                      "表组": f"G{组序[x.get('表组')]}", "续表序": x.get("续表序"),
                      "行数": x["行范围"][1] - x["行范围"][0] + 1}
                     for x in 本侧]
            成功, 值 = 复用接入.发布记录(簿, 侧名, "第二遍", 证据, {"表清单": 表清单},
                                    金标准版本, 规则版本, 来源本次("整侧"), 任务标识,
                                    依赖=科目版本)
            if 成功:
                分表版本.append(值)
            复用记录["发布"].append({"步骤": "第二遍", "状态": "已发布" if 成功 else "未发布",
                                "版本": 值 if 成功 else None, "原因": "" if 成功 else 值})
        for 组键, 记录们 in 组们.items():
            if any(x.get("复用版本") for x in 记录们):
                continue
            可, 因 = 证据模块.可发布(记录们, 完整性)
            if not 可 or not 完整性.get("第一遍覆盖") or not 完整性.get("第二遍归段"):
                复用记录["发布"].append({"步骤": "第三遍", "表组": 组键, "状态": "未发布",
                                    "原因": 因 or "整侧覆盖或完整性未通过"})
                continue
            证据 = 证据模块.建证据_精读(侧名, 记录们, 快照, 用可见范围, 上下文)
            视图, 分表, _ = 复用接入.组形状(快照, 记录们, 用可见范围)
            候选发布.append({"步骤": "第三遍", "表组": 组键, "证据": 证据,
                        "解释": 证据["解释"], "形状键": 证据模块.形状摘要(视图.形状(分表)),
                        "来源": 来源本次(组键),
                        "依赖": list(dict.fromkeys(科目版本 + 分表版本))})
    复用记录["耗时"]["发布"] = round(_time.perf_counter() - 发布起, 4)
    _写jsonl(运行目录 / f"候选发布_{侧名}.jsonl", 候选发布)
    (运行目录 / f"复用记录_{侧名}.json").write_text(
        json.dumps(复用记录, ensure_ascii=False, indent=2), encoding="utf-8")

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
