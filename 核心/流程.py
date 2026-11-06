# coding: utf-8
"""总流程编排：提取→识别A→识别B→比对→搬运→写回→回读，接续索引逐步登记。

三条保护写在流程里：
- 每个产物一保存就登记路径和哈希；文件被人改过、或没有可信记录，一律另存不覆盖。
- 任务开始登记输入指纹（三个输入的路径与哈希、数据表工作表）；续跑对不上就拒绝，不混用旧结果。
- 某张表回写失败时，把那张表的变更恢复原值再交付其余表，绝不因为一张表全盘作废。

跨任务复用也在这里收口：任务开始刷新本机识别记录簿；上游布局在自身覆盖检查通过后发布，
第三遍候选等配对检查后才发布；标题在两侧配对后按本次 B 的期间日期生成，生成不了整组撤回，
保证 Excel 与 Word 对同一组使用同一结果。
"""
import hashlib
import json
import os
import uuid
from datetime import datetime
from pathlib import Path

from .比对 import 建立对应
from .写回 import 读标题更新, 写回Word
from .回读 import 回读验证
from .接续 import 读索引, 写索引
from .搬运 import 生成更新表, 保存更新表
from .结构维护 import 读结构操作
from .标题 import 生成本次标题
from .识别 import 识别一侧
from .文件输入 import 规范化输入, 准备一侧
from .Excel交付 import 交付Excel, 收集保留公式
from .表格 import 读快照
from .共用记录簿 import 取记录簿, 已启用
from . import 复用接入 as 接入
from . import 复用证据 as 证据模块


def _哈希(路径) -> str:
    p = Path(路径)
    if not p.is_file():
        return ""
    h = hashlib.sha256()
    with p.open("rb") as f:
        for 块 in iter(lambda: f.read(1024 * 1024), b""):
            h.update(块)
    return h.hexdigest()


def _避让(路径: Path, 旧哈希):
    """产物被人改过、或者没有可信记录，都不覆盖：另存带时间戳的新文件，原文件留着。"""
    if not 路径.exists():
        return 路径, False
    if 旧哈希 and _哈希(路径) == 旧哈希:
        return 路径, False
    新 = 路径.with_name(f"{路径.stem}_人工改过_{datetime.now():%Y%m%d_%H%M%S}{路径.suffix}")
    return 新, True


def _登记(索引: dict, 名: str, 路径) -> None:
    索引.setdefault("产物", {})[名] = str(路径)
    索引.setdefault("产物哈希", {})[名] = _哈希(路径)


def _保存交付事实(目录, 索引, A, B, 对应, 来源, 变更, 实值, 问题, 差异,
              结构, 标题, 输出, 公式=()):
    """只记录既有执行结果；展示记录失败不改变已经交付的文件和业务结论。"""
    try:
        说明路径 = 目录 / "任务说明.json"
        说明 = json.loads(说明路径.read_text(encoding="utf-8-sig")) if 说明路径.is_file() else {}
        公共 = {"task_id": 目录.name, "执行次数": 说明.get("执行次数", 0),
                "侧": "A", "输出": str(输出), "输出哈希": _哈希(输出)}
        对应表 = {x.get("a格"): x for x in 对应}
        B表 = {x.get("格"): x for x in B}
        公式表 = {x.get("a格"): x for x in 公式}
        差异格 = {x.get("格") for x in 差异}
        明细 = []
        for 甲 in A:
            格 = 甲.get("格")
            条 = 来源.get(格) or 对应表.get(格) or {}
            乙 = B表.get(条.get("b格")) or {}
            动作 = "原值保留"
            if 格 in 变更:
                动作 = ("已清空" if 变更[格] in (None, "") else
                      "数值未变" if 甲.get("原值状态") == "值" and 甲.get("原值") == 变更[格]
                      else "已写入")
            if 格 in 公式表:
                动作 = "公式保留"
            if 格 in 差异格:
                动作 = "回读有差异"
            原因 = [i for i, x in enumerate(问题) if x.get("a格") == 格
                  or (甲.get("表组") and (x.get("表组") == 甲["表组"]
                      or 甲["表组"] in (x.get("组") or [])))]
            明细.append({**公共, "类型": "格子", "a格": 格, "b格": 条.get("b格"),
                       **{k: 甲.get(k) for k in ("表组", "读取部分", "科目", "表名", "角色",
                                               "行路径", "列路径", "维度值", "原值", "原值状态")},
                       "来源值": 乙.get("原值"), "来源值状态": 乙.get("原值状态"),
                       "最终动作": 动作, "实际值": 实值.get(格), "实际值已取得": 格 in 实值,
                       "原因引用": 原因, "公式核对": 公式表.get(格)})
        for 表 in 结构:
            原标题 = 表.get("标题") or 表.get("标题规则")
            if not 原标题:
                continue
            动作们 = [x for x in 标题 if
                    (x.get("range_name") == 表.get("区域") if 表.get("区域")
                     else x.get("格") == f"{表.get('工作表')}!{原标题.get('位置')}"
                     if isinstance(原标题, dict) else False)]
            明细.append({**公共, "类型": "标题", "表组": 表.get("表组"),
                       "读取部分": 表.get("读取部分"), "区域": 表.get("区域"),
                       "工作表": 表.get("工作表"), "范围": 表.get("范围"),
                       "原结论": 原标题, "最终动作": "已写回" if 动作们 else "原样保留",
                       "动作": 动作们})
        路径 = 目录 / "交付明细.jsonl"
        临时 = 路径.with_name(f"交付明细_{uuid.uuid4().hex}.tmp")
        临时.write_text("".join(json.dumps(x, ensure_ascii=False, default=str) + "\n"
                               for x in 明细), encoding="utf-8")
        os.replace(临时, 路径)
        _登记(索引, "交付明细", 路径)
        索引["交付明细执行次数"] = 公共["执行次数"]
        索引.pop("展示错误", None)
    except Exception as e:
        索引["展示错误"] = f"结果明细不完整：{e}"
        索引.setdefault("产物", {}).pop("交付明细", None)
        索引.setdefault("产物哈希", {}).pop("交付明细", None)


def _读jsonl(路径):
    p = Path(路径)
    return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.is_file() else []



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


def _写jsonl(路径, 记录):
    with Path(路径).open("w", encoding="utf-8") as f:
        for r in 记录:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def _指定工作表(输入: dict):
    值 = 输入.get("b工作表")
    if not 值:
        return None
    return [值] if isinstance(值, str) else list(值)


def _输入指纹(输入: dict) -> dict:
    """旧输入保留原指纹结构；新输入同时固定两端类型和工作表集合。"""
    if "A" not in 输入 and "B" not in 输入:
        return {名: {"路径": str(输入.get(键) or ""), "哈希": _哈希(输入.get(键))}
                for 名, 键 in (("word", "word"), ("b表", "b表"), ("金标准", "金标准"))} | \
               {"b工作表": _指定工作表(输入), "规则版本": 证据模块.规则摘要()}
    出 = {"金标准": {"路径": str(输入.get("金标准") or ""), "哈希": _哈希(输入.get("金标准"))},
          "规则版本": 证据模块.规则摘要()}
    for 侧 in ("A", "B"):
        描述 = 输入.get(侧) or {}
        出[侧] = {"路径": str(描述.get("路径") or ""), "哈希": _哈希(描述.get("路径")),
                "类型": 描述.get("类型"), "工作表": sorted(描述.get("工作表") or [])}
    return 出


def _指纹差异(旧: dict, 新: dict) -> list:
    差 = []
    for 名 in ("word", "b表", "金标准", "A", "B"):
        a, b = (旧 or {}).get(名) or {}, (新 or {}).get(名) or {}
        if not a and not b:
            continue
        if a.get("路径") != b.get("路径"):
            差.append(f"{名} 路径不同")
        elif a.get("哈希") != b.get("哈希"):
            差.append(f"{名} 文件内容已变")
        elif a.get("类型") != b.get("类型"):
            差.append(f"{名} 文件类型不同")
        elif a.get("工作表") != b.get("工作表"):
            差.append(f"{名} 工作表不同")
    if (旧 or {}).get("b工作表") != (新 or {}).get("b工作表"):
        差.append("数据表工作表不同")
    if not (旧 or {}).get("规则版本") or 旧["规则版本"] != 新.get("规则版本"):
        差.append("识别规则版本已变或旧任务缺少可核验版本")
    return 差


def _错误里的区域(消息: str, 区域到组: dict):
    for 区域 in 区域到组:
        if 区域 and 区域 in 消息:
            return 区域
    return None


def _汇总一侧(侧, 读取部分们, 根目录) -> dict:
    根目录 = Path(根目录)
    多部分 = len(读取部分们) > 1
    结构们, 合集, 候选, 问题, 部分记录, 划分们 = [], [], [], [], [], []

    def 新组(记录, 标识):
        出 = dict(记录)
        出["读取部分"] = 标识
        if 多部分:
            出["表组"] = f"{标识}:{记录.get('表组') or ''}"
        return 出

    for 部分 in 读取部分们:
        目录 = Path(部分["目录"])
        标识 = 部分["标识"]
        结构们 += [新组(x, 标识) for x in _读jsonl(目录 / "结构清单.jsonl") if x.get("侧") == 侧]
        划分们 += [{"读取部分": 标识, **x} for x in _读jsonl(目录 / f"科目划分_{侧}.jsonl")
                if x.get("侧") == 侧]
        合集 += [新组(x, 标识) for x in _读jsonl(目录 / f"{侧}合集.jsonl")]
        问题 += [{"读取部分": 标识, **x} for x in _读jsonl(目录 / f"识别问题_{侧}.jsonl")]
        for x in _读jsonl(目录 / f"候选发布_{侧}.jsonl"):
            项 = 新组(x, 标识)
            来源 = dict(项.get("来源") or {})
            来源.update({"任务": str(根目录), "侧": 侧, "读取部分": 标识,
                       "工作表": 部分.get("工作表"), "原文件类型": 部分.get("类型")})
            项["来源"] = 来源
            候选.append(项)
        路径 = 目录 / f"复用记录_{侧}.json"
        if 路径.is_file():
            部分记录.append(json.loads(路径.read_text(encoding="utf-8")))

    同名 = {}
    for 记录 in 结构们:
        键 = (str(记录.get("科目") or ""), str(记录.get("表名") or ""))
        同名.setdefault(键, set()).add(str(记录.get("表组") or ""))
    歧义键 = {键 for 键, 组们 in 同名.items() if len(组们) > 1}
    for 记录 in 结构们:
        键 = (str(记录.get("科目") or ""), str(记录.get("表名") or ""))
        if 键 in 歧义键:
            记录["歧义"] = f"同科目同名表分属多个表组：{键[0]}／{键[1]}"

    结构路径 = 根目录 / "结构清单.jsonl"
    其他 = [x for x in _读jsonl(结构路径) if x.get("侧") != 侧]
    _写jsonl(结构路径, 其他 + 结构们)
    合集路径 = 根目录 / f"{侧}合集.jsonl"
    _写jsonl(合集路径, 合集)
    _写jsonl(根目录 / f"候选发布_{侧}.jsonl", 候选)
    _写jsonl(根目录 / f"识别问题_{侧}.jsonl", 问题)
    _写jsonl(根目录 / f"科目划分_{侧}.jsonl", 划分们)

    汇总 = {"侧": 侧, "任务": str(根目录), "部分": 部分记录,
            "库": (部分记录[-1].get("库") if 部分记录 else {}),
            "调用": {"第一遍": 0, "第二遍": 0, "第三遍": 0, "重试": 0}}
    for 名 in ("复用", "回退", "跳过", "首认", "发布"):
        汇总[名] = []
    for 部分, 记录 in zip(读取部分们, 部分记录):
        for 遍, 数 in (记录.get("调用") or {}).items():
            汇总["调用"][遍] = 汇总["调用"].get(遍, 0) + 数
        for 名 in ("复用", "回退", "跳过", "首认", "发布"):
            汇总[名] += [{"读取部分": 部分["标识"], **x} for x in (记录.get(名) or [])]
    (根目录 / f"复用记录_{侧}.json").write_text(
        json.dumps(汇总, ensure_ascii=False, indent=2), encoding="utf-8")
    统计 = {}
    for 记录 in 部分记录:
        for 名, 数 in (记录.get("统计") or {}).items():
            统计[名] = 统计.get(名, 0) + 数
    return {"结构清单": 结构路径, "合集": 合集路径, "统计": 统计, "待答复": []}


def _识别整侧(侧, 读取部分们, 根目录, 配置, 金标准, 簿, 允许科目=None, 取消=None) -> dict:
    根目录 = Path(根目录)
    待答复 = []
    for 部分 in 读取部分们:
        if 取消 is not None and 取消.is_set():
            raise InterruptedError("用户中断")
        索引 = 读索引(根目录) or {"已完成范围": [], "产物": {}}
        已完成 = 索引.setdefault("已完成范围", [])
        标记 = f"识别{侧}:{部分['标识']}"
        部分目录 = Path(部分["目录"])
        已完成产物 = (部分目录 / f"{侧}合集.jsonl").is_file() and (部分目录 / "结构清单.jsonl").is_file()
        if 部分目录 == 根目录 or 标记 not in 已完成 or not 已完成产物:
            结果 = 识别一侧(配置, 部分["材料"], 侧, 部分目录, 金标准,
                       上下文=部分.get("上下文") or [], 取消=取消,
                       只跑表=[部分["工作表"]] if 部分.get("类型") == "excel" else None,
                       允许科目=允许科目, 簿=簿, 材料类型=部分.get("类型"),
                       根任务=str(根目录), 读取部分=部分)
            if 结果["待答复"]:
                待答复 += 结果["待答复"]
                break
            索引 = 读索引(根目录) or 索引
            索引.setdefault("已完成范围", []).append(标记)
            索引["已完成范围"] = list(dict.fromkeys(索引["已完成范围"]))
            写索引(根目录, 索引)
    if 取消 is not None and 取消.is_set():
        raise InterruptedError("用户中断")
    if 待答复:
        return {"结构清单": None, "合集": None, "统计": {}, "待答复": 待答复}
    结果 = _汇总一侧(侧, 读取部分们, 根目录)
    索引 = 读索引(根目录) or {"已完成范围": [], "产物": {}}
    索引["产物"].update({f"{侧}合集": str(结果["合集"]), "结构清单": str(结果["结构清单"])})
    索引[f"{侧}完整性"] = [x.get("完整性") for x in _读jsonl(结果["结构清单"])
                       if x.get("侧") == 侧]
    写索引(根目录, 索引)
    return 结果


_配对错误 = ("同侧重复", "同名业务表归属不明", "口径冲突", "例外未执行", "列标签缺失", "越界解释", "单位依据不足或冲突", "期间角色未知", "槽位编号缺失")


def _发布第三遍(簿, 运行目录, 清单, A结构, B结构, 金标准版本, 规则版本) -> list:
    """第三遍在本组检查与适用配对检查之后发布；配对暴露的识别错误阻止入库。"""
    if 簿 is None:
        return []
    # a格、b格已经给出来源侧；不能用工作表名推断侧，两份文件可以同名。
    from openpyxl.utils.cell import coordinate_to_tuple
    from openpyxl.utils import range_boundaries
    错误 = set()
    for 条 in 清单 or []:
        if 条.get("状态") != "未对应" or str(条.get("原因") or "") not in _配对错误:
            continue
        for 键, 侧, 结构 in (("a格", "A", A结构), ("b格", "B", B结构)):
            格 = 条.get(键)
            if not 格:
                continue
            表, 坐标 = str(格).rsplit("!", 1)
            行, 列 = coordinate_to_tuple(坐标)
            for 记录 in 结构:
                if 记录.get("工作表") != 表:
                    continue
                左, 顶, 右, 底 = range_boundaries(记录["范围"])
                if 左 <= 列 <= 右 and 顶 <= 行 <= 底:
                    错误.add((侧, 记录.get("表组")))
    结果 = []
    for 侧 in ("A", "B"):
        候选 = _读jsonl(运行目录 / f"候选发布_{侧}.jsonl")
        记录路径 = 运行目录 / f"复用记录_{侧}.json"
        if not 候选:
            continue
        记录 = json.loads(记录路径.read_text(encoding="utf-8")) if 记录路径.is_file() else {}
        发布们 = 记录.setdefault("发布", [])
        for 项 in 候选:
            if (侧, 项.get("表组")) in 错误:
                发布们.append({"步骤": "第三遍", "表组": 项.get("表组"), "状态": "未发布",
                            "原因": "配对暴露识别错误，本组不入库"})
                结果.append({"侧": 侧, "表组": 项.get("表组"), "状态": "未发布",
                          "原因": "配对暴露识别错误"})
                continue
            成功, 值 = 接入.发布记录(簿, 侧, "第三遍", 项["证据"], 项["解释"],
                                金标准版本, 规则版本, 项.get("来源"), str(运行目录),
                                形状键=项.get("形状键"), 依赖=项.get("依赖"))
            发布们.append({"步骤": "第三遍", "表组": 项.get("表组"),
                        "状态": "已发布" if 成功 else "未发布",
                        "版本": 值 if 成功 else None, "原因": "" if 成功 else 值})
            结果.append({"侧": 侧, "表组": 项.get("表组"),
                      "状态": "已发布" if 成功 else "未发布", "版本": 值 if 成功 else None,
                      "原因": "" if 成功 else 值})
        记录路径.write_text(json.dumps(记录, ensure_ascii=False, indent=2), encoding="utf-8")
    return 结果


def 运行任务(输入: dict, 运行目录, 配置: dict, 取消=None) -> dict:
    原输入 = 输入
    输入 = 规范化输入(输入)
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    索引 = 读索引(运行目录) or {"步骤": "extract", "已完成范围": [], "产物": {}}
    产物 = 索引.setdefault("产物", {})

    指纹 = _输入指纹(原输入)
    旧指纹 = 索引.get("输入指纹")
    if 旧指纹:
        差 = _指纹差异(旧指纹, 指纹)
        if 差:
            raise RuntimeError("续跑输入与原任务不一致（" + "；".join(差) + "）："
                               "请新建任务，不要复用旧结果")
    elif "表A" in 产物:
        raise RuntimeError("这个任务没有可信的输入指纹（可能是升级前建的），"
                           "不能判断旧结果是否来自当前输入：请新建任务")
    # 续跑校验通过后才清上次状态，拒绝续跑时不改旧索引。
    索引["步骤"] = "extract"
    for 键 in ("错误", "识别错误", "Excel错误", "待答复"):
        索引.pop(键, None)
    索引["输入指纹"] = 指纹
    写索引(运行目录, 索引)

    簿, 复used = None, {"可用": False, "原因": "跨任务复用已关闭"}
    if 已启用(输入):
        try:
            簿 = 取记录簿(输入.get("识别记录目录"))
            快照 = 簿.刷新()
            复used = {"可用": bool(快照.get("可用")), "代次": 快照.get("代次"),
                     "原因": 快照.get("原因") or ""}
        except Exception as e:
            簿, 复used = None, {"可用": False, "原因": f"记录簿不可用：{e}"}
    索引["复用"] = 复used
    写索引(运行目录, 索引)

    材料 = 索引.get("材料")
    if not 材料 and "表A" in 产物:
        上下文 = json.loads(Path(产物["上下文"]).read_text(encoding="utf-8-sig"))
        材料 = {"A": [{"标识": "P001", "类型": "word", "材料": 产物["表A"],
                    "工作表": None, "上下文": 上下文.get("tables", []),
                    "上下文路径": 产物["上下文"], "带连接Word": 产物["带连接Word"],
                    "目录": str(运行目录)}],
                "B": [{"标识": "P001", "类型": "excel", "材料": 输入["B"]["路径"],
                    "工作表": 输入["B"]["工作表"][0], "上下文": [],
                    "上下文路径": None, "带连接Word": None, "目录": str(运行目录)}]}
    if not 材料:
        材料 = {}
    for 侧 in ("A", "B"):
        if 材料.get(侧):
            continue                        # 该端材料已准备完成，续跑不重复提取
        部分们 = 准备一侧(侧, 输入[侧], 运行目录 / 侧)
        材料[侧] = [{k: str(v) if isinstance(v, Path) else v for k, v in 部分.items()}
                  for 部分 in 部分们]
        索引 = 读索引(运行目录) or 索引
        索引["材料"] = 材料
        if 侧 == "A" and 输入[侧]["类型"] == "word":
            部分 = 材料[侧][0]
            索引["产物"].update({"表A": 部分["材料"], "带连接Word": 部分["带连接Word"],
                              "上下文": 部分["上下文路径"]})
            部分索引 = 读索引(部分["目录"]) or {}
            索引["桥接提示"] = 部分索引.get("桥接提示") or []
        索引["步骤"] = "identify_a"
        写索引(运行目录, 索引)
    产物 = 索引.setdefault("产物", {})
    上下文 = 材料["A"][0].get("上下文") or []
    金标准 = 输入.get("金标准")

    for 侧 in ("A", "B"):
        允许科目 = None
        if 侧 == "B":
            甲侧 = [x for x in _读jsonl(运行目录 / "结构清单.jsonl")
                    if x.get("侧") == "A" and x.get("完成")]
            if 甲侧:
                允许科目 = {x["科目"] for x in 甲侧}
        索引 = 读索引(运行目录) or 索引
        索引["步骤"] = "identify_a" if 侧 == "A" else "identify_b"
        写索引(运行目录, 索引)
        try:
            结果 = _识别整侧(侧, 材料[侧], 运行目录, 配置, 金标准, 簿,
                          允许科目=允许科目, 取消=取消)
        except InterruptedError:
            索引 = 读索引(运行目录) or 索引
            索引["步骤"] = "用户中断"
            索引.pop("待答复", None)
            写索引(运行目录, 索引)
            return 索引
        except Exception as e:
            索引 = 读索引(运行目录) or 索引
            索引.update({"步骤": "识别失败", "识别错误": str(e)})
            写索引(运行目录, 索引)
            e.任务步骤 = 索引["步骤"]  # 只传本次异常类别，服务不借用旧任务的失败状态。
            raise
        if 结果["待答复"]:
            索引 = 读索引(运行目录) or 索引
            索引["步骤"] = "identify_a" if 侧 == "A" else "identify_b"
            索引["待答复"] = 结果["待答复"]
            写索引(运行目录, 索引)
            return 索引

    索引 = 读索引(运行目录) or 索引
    索引.pop("待答复", None)      # 两侧识别都过了，旧的“待答复”目录不再提示
    索引.pop("识别错误", None)
    a记录 = _读jsonl(运行目录 / "A合集.jsonl")
    b记录 = _读jsonl(运行目录 / "B合集.jsonl")
    结构 = _读jsonl(运行目录 / "结构清单.jsonl")
    A结构 = [x for x in 结构 if x.get("侧") == "A"]
    B结构 = [x for x in 结构 if x.get("侧") == "B"]

    # 粗比单子：科目层级先对大方向；A 独有的科目列出来交人工，不许静默漏
    科A = {x["科目"] for x in A结构}
    科B = {x["科目"] for x in B结构}
    名B = {(x.get("科目"), x.get("表名")) for x in B结构}
    跳过 = [x for x in B结构 if x.get("跳过")]
    粗比 = {"A独有科目": sorted(科A - 科B), "B独有科目": sorted(科B - 科A),
            "B跳过表数": len(跳过), "B精读表数": sum(1 for x in B结构 if not x.get("跳过")),
            "A有B无的表": [f"{x.get('科目')}／{x.get('表名')}（{x.get('区域')}）"
                        for x in A结构 if (x.get("科目"), x.get("表名")) not in 名B]}
    (运行目录 / "粗比_B.json").write_text(json.dumps(粗比, ensure_ascii=False, indent=2),
                                       encoding="utf-8")
    粗比问题 = []
    if 粗比["A有B无的表"]:
        粗比问题.append({"原因": "粗比：A 有这张表，B 找不到对应（可能并进了别的表，也可能 B 没有这一节），交人工",
                         "表": 粗比["A有B无的表"]})

    清单 = 建立对应(a记录, b记录, A结构, B结构)
    _写jsonl(运行目录 / "对应清单.jsonl", 清单)
    _发布第三遍(簿, 运行目录, 清单, A结构, B结构,
            证据模块.内容摘要(输入.get("金标准")), 证据模块.规则摘要())
    索引["步骤"] = "carry"
    写索引(运行目录, 索引)

    搬运来源 = {}
    变更, 问题 = 生成更新表({}, 清单, b记录, a记录, 留证=搬运来源)
    超额问题 = [{"原因": x.get("原因"), "表组": x.get("表组"), "槽位": x.get("槽位")}
            for x in 清单 if str(x.get("原因") or "").startswith("B多出项")
            or str(x.get("原因") or "").startswith("B空项")]
    问题 = (_读jsonl(运行目录 / "识别问题_A.jsonl") + _读jsonl(运行目录 / "识别问题_B.jsonl")
          + 问题 + 粗比问题 + 超额问题)
    a按格 = {r.get("格"): r for r in a记录}
    标题更新, 标题问题, 标题撤回组, 标题证据 = 生成本次标题(
        A结构, B结构, 上下文, 变更, b记录)
    问题 += 标题问题
    if 标题撤回组:
        变更 = {格: 值 for 格, 值 in 变更.items()
                if a按格.get(格, {}).get("表组") not in 标题撤回组}
        问题.append({"原因": "标题无法按本次 B 日期生成，相关组的标题与数值一并保持原样",
                    "组": sorted(标题撤回组)})
    (运行目录 / "标题证据.json").write_text(
        json.dumps(标题证据, ensure_ascii=False, indent=2), encoding="utf-8")
    _写jsonl(运行目录 / "问题清单.jsonl", 问题)

    if 取消 is not None and 取消.is_set():
        索引["步骤"] = "用户中断"
        写索引(运行目录, 索引)
        return 索引

    if 输入["A"]["类型"] == "excel":
        A快照 = 读快照(材料["A"][0]["材料"], 业务范围=[
            {"sheet": x.get("工作表"), "range": x.get("范围")} for x in A结构
            if x.get("工作表") and x.get("范围")])
        公式清单 = 收集保留公式(A快照, A结构, 清单, b记录)
        # 标题撤回的组本次没有搬运，保留公式不能拿旧数值宣称核对通过。
        for 项 in 公式清单:
            if 项.get("表组") in 标题撤回组 and not 项.get("不能核对原因"):
                项["不能核对原因"] = "所在组因标题未能确认而保留，未按本次搬运结果核对"
        Excel标题 = {x["格"]: {"原文": x["old_text"], "新文": x["text"], "表组": x["表组"]}
                   for x in 标题更新 if x.get("格")}
        Excel路径, 改 = _避让(运行目录 / "Excel新副本.xlsx",
                              (索引.get("产物哈希") or {}).get("Excel新副本"))
        索引["步骤"] = "writeback"
        写索引(运行目录, 索引)
        try:
            结果 = 交付Excel(材料["A"][0]["材料"], Excel路径, 变更, A快照, A结构,
                           公式清单, Excel标题, 取消=取消)
        except Exception as e:
            索引 = 读索引(运行目录) or 索引
            中断 = isinstance(e, InterruptedError)
            索引["步骤"] = "用户中断" if 中断 else "Excel交付失败"
            if not 中断:
                索引["Excel错误"] = str(e)
            if Excel路径.exists():
                _登记(索引, "Excel候选副本", Excel路径)
            写索引(运行目录, 索引)
            if 中断:
                return 索引
            e.任务步骤 = 索引["步骤"]
            raise
        问题 += 结果["问题"]
        差异 = 结果["回读差异"]
        if 差异:
            问题 += [{"原因": "Excel最终回读不一致", "差异": x} for x in 差异]
        _写jsonl(运行目录 / "回读记录.jsonl", 差异)
        _写jsonl(运行目录 / "公式核对记录.jsonl", 结果["公式核对"])
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)
        索引 = 读索引(运行目录) or 索引
        _登记(索引, "Excel新副本", Excel路径)
        索引["产物"]["A端新副本"] = str(Excel路径)
        索引["产物"]["Excel新副本"] = str(Excel路径)
        索引["回读差异"] = len(差异)
        索引["公式待核数"] = 结果["公式待核数"]
        索引["撤回表"] = 结果["撤回组"]
        if 改:
            问题.append({"原因": "原Excel新副本被改过，已另存新文件，原文件保留",
                        "路径": str(Excel路径)})
            _写jsonl(运行目录 / "问题清单.jsonl", 问题)
        索引["步骤"] = ("回读不一致" if 差异
                     else "部分保留" if (结果["撤回组"] or 标题撤回组)
                     else "公式待核" if 结果["公式待核数"] else "done")
        最终标题 = [x for x in 标题更新 if x.get("格") in 结果.get("实际标题", {})]
        _保存交付事实(运行目录, 索引, a记录, b记录, 清单, 搬运来源,
                  结果.get("实际变更", {}), 结果.get("实际值", {}), 问题, 差异,
                  A结构, 最终标题, Excel路径, 结果["公式核对"])
        写索引(运行目录, 索引)
        return 索引

    已录哈希 = 索引.get("产物哈希") or {}
    更新表路径, 改1 = _避让(运行目录 / "更新后附注表格.xlsx", 已录哈希.get("更新后附注表格"))
    保存更新表(产物["表A"], 变更, 更新表路径)
    _登记(索引, "更新后附注表格", 更新表路径)      # 一保存就登记，中途失败也不留“无记录产物”
    if 改1:
        问题.append({"原因": "原更新表被改过，已另存新文件，原文件保留",
                    "路径": str(更新表路径)})
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)
    索引["步骤"] = "writeback"
    写索引(运行目录, 索引)

    区域到组 = {(x.get("区域") or ""): (x.get("表组") or "")
             for x in A结构 if x.get("区域")}
    Word路径, 改2 = _避让(运行目录 / "Word新副本.docx", 已录哈希.get("Word新副本"))
    撤回 = set()
    while True:
        try:
            回写 = 写回Word(产物["带连接Word"], 更新表路径, Word路径,
                       operations=[], caption_updates=标题更新)
            if 回写.get("backup_warning"):
                # 备份没做成不影响正文有效性：写回结果以桥接的逐格校验为准，这条只作记录。
                问题.append({"原因": "旧结果未能备份到 BackUp，正文已写入并校验；本条仅作记录",
                            "错误": str(回写["backup_warning"])[:300], "输出": str(Word路径)})
                _写jsonl(运行目录 / "问题清单.jsonl", 问题)
            break
        except Exception as e:
            区域 = _错误里的区域(str(e), 区域到组)
            if not 区域 or 区域 in 撤回:
                raise
            组 = 区域到组[区域]
            同组区域 = {名 for 名, 所属组 in 区域到组.items() if 所属组 == 组}
            撤回.update(同组区域)                 # 续表按整组恢复，标题与数字同进同退
            变更 = {格: 值 for 格, 值 in 变更.items()
                    if a按格.get(格, {}).get("表组") != 组}
            标题更新 = [u for u in 标题更新 if u.get("range_name") not in 同组区域]
            保存更新表(产物["表A"], 变更, 更新表路径)
            _登记(索引, "更新后附注表格", 更新表路径)
            组记录 = next((x for x in A结构 if x.get("表组") == 组), {})
            问题.append({"原因": "该表回写失败，已保持原值，其余表照常交付",
                        "区域": 区域, "科目": 组记录.get("科目"), "表名": 组记录.get("表名"),
                        "表组": 组, "错误": str(e)[:300]})
            _写jsonl(运行目录 / "问题清单.jsonl", 问题)
            写索引(运行目录, 索引)
    _登记(索引, "Word新副本", Word路径)
    if 改2:
        问题.append({"原因": "原 Word 新副本被改过，已另存新文件，原文件保留",
                    "路径": str(Word路径)})
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)
    索引["步骤"] = "readback"
    写索引(运行目录, 索引)

    实际值 = {}
    差异 = 回读验证(Word路径, 变更, 运行目录 / "回读临时", 实际值=实际值, 留证格=a按格)
    _写jsonl(运行目录 / "回读记录.jsonl", 差异)      # 差异文件保留，不删证据
    if 差异:
        问题 += [{"原因": "回读不一致：写进 Word 后回读对不上，明细见 回读记录.jsonl",
                  "差异": d} for d in 差异]
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)   # 流程末尾把所有问题收齐在一份清单里
    索引["产物哈希"] = {"更新后附注表格": _哈希(更新表路径), "Word新副本": _哈希(Word路径)}
    索引["回读差异"] = len(差异)
    索引["撤回表"] = sorted(撤回)
    # 标题生成或实际写回有撤回，都只能算部分保留。
    索引["步骤"] = ("回读不一致" if 差异
                 else "部分保留" if (撤回 or 标题撤回组) else "done")
    索引["产物"].update({"更新后附注表格": str(更新表路径),
                   "Word新副本": str(Word路径)})
    _保存交付事实(运行目录, 索引, a记录, b记录, 清单, 搬运来源, 变更,
              实际值, 问题, 差异, A结构, 标题更新, Word路径)
    写索引(运行目录, 索引)
    return 索引


def 人工回写(输入: dict) -> dict:
    """人工直接回写：按真实路径重读用户已保存的 Excel。

    结构操作由调用方给定（输入["结构操作"]，json 文件路径或数组）；不给就只改内容，
    不插行、不删行。程序不再比较两份文件去推断结构变化。
    输出路径前端叫「输出」、内部叫 output，两种都认，免得字段名对不上白跑一趟。
    """
    输出 = 输入.get("output") or 输入.get("输出")
    if not 输出:
        raise ValueError("缺少输出路径（输出／output）")
    ops = 读结构操作(输入["结构操作"]) if 输入.get("结构操作") else []
    return 写回Word(输入["带连接word"], 输入["excel"], 输出, operations=ops)
