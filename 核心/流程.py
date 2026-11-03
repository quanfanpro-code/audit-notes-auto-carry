# coding: utf-8
"""总流程编排：提取→识别A→识别B→比对→搬运→写回→回读，接续索引逐步登记。

三条保护写在流程里：
- 每个产物一保存就登记路径和哈希；文件被人改过、或没有可信记录，一律另存不覆盖。
- 任务开始登记输入指纹（三个输入的路径与哈希、数据表工作表）；续跑对不上就拒绝，不混用旧结果。
- 某张表回写失败时，把那张表的变更恢复原值再交付其余表，绝不因为一张表全盘作废。
"""
import hashlib
import json
from datetime import datetime
from pathlib import Path

from .比对 import 建立对应
from .写回 import 读标题更新, 写回Word
from .回读 import 回读验证
from .提取 import 提取任务
from .接续 import 读索引, 写索引
from .搬运 import 生成更新表, 保存更新表
from .结构维护 import 读结构操作
from .识别 import 识别一侧


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
    """续跑必须认得出“还是不是同一个任务”：路径换了、内容换了、工作表换了，都不许复用。"""
    return {名: {"路径": str(输入.get(键) or ""), "哈希": _哈希(输入.get(键))}
            for 名, 键 in (("word", "word"), ("b表", "b表"), ("金标准", "金标准"))} | \
           {"b工作表": _指定工作表(输入)}


def _指纹差异(旧: dict, 新: dict) -> list:
    差 = []
    for 名 in ("word", "b表", "金标准"):
        a, b = (旧 or {}).get(名) or {}, (新 or {}).get(名) or {}
        if a.get("路径") != b.get("路径"):
            差.append(f"{名} 路径不同")
        elif a.get("哈希") != b.get("哈希"):
            差.append(f"{名} 文件内容已变")
    if (旧 or {}).get("b工作表") != (新 or {}).get("b工作表"):
        差.append("数据表工作表不同")
    return 差


def _错误里的区域(消息: str, 区域到组: dict):
    for 区域 in 区域到组:
        if 区域 and 区域 in 消息:
            return 区域
    return None


def 运行任务(输入: dict, 运行目录, 配置: dict, 取消=None) -> dict:
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    索引 = 读索引(运行目录) or {"步骤": "extract", "已完成范围": [], "产物": {}}
    产物 = 索引.setdefault("产物", {})

    指纹 = _输入指纹(输入)
    旧指纹 = 索引.get("输入指纹")
    if 旧指纹:
        差 = _指纹差异(旧指纹, 指纹)
        if 差:
            raise RuntimeError("续跑输入与原任务不一致（" + "；".join(差) + "）："
                               "请新建任务，不要复用旧结果")
    elif "表A" in 产物:
        # 升级前建的任务没有输入指纹：认不出旧产物是不是这次输入的，不认领
        raise RuntimeError("这个任务没有可信的输入指纹（可能是升级前建的），"
                           "不能判断旧结果是否来自当前输入：请新建任务")
    索引["输入指纹"] = 指纹
    写索引(运行目录, 索引)

    if "表A" not in 产物:
        for 键, 值 in 提取任务(输入["word"], 运行目录).items():
            产物[键] = 值
        索引 = 读索引(运行目录) or 索引      # 提取里写了桥接提示，重新读，别被旧快照覆盖
        索引["产物"].update(产物)
        索引["步骤"] = "identify_a"
        写索引(运行目录, 索引)
    上下文 = json.loads(Path(产物["上下文"]).read_text(encoding="utf-8-sig"))
    金标准 = 输入.get("金标准")

    for 侧, 表文件, 表上下文, 只跑表 in (("A", 产物["表A"], 上下文.get("tables", []), None),
                                  ("B", 输入["b表"], None, _指定工作表(输入))):
        允许科目 = None
        if 侧 == "B":
            结构路径 = 运行目录 / "结构清单.jsonl"
            if 结构路径.exists():
                甲侧 = [x for x in _读jsonl(结构路径)
                        if x.get("侧") == "A" and x.get("完成")]
                if 甲侧:
                    允许科目 = {x["科目"] for x in 甲侧}   # 粗比：B 只精读 A 也有的科目
        结果 = 识别一侧(配置, 表文件, 侧, 运行目录, 金标准, 表上下文,
                   取消=取消, 只跑表=只跑表, 允许科目=允许科目)
        if 结果["待答复"]:
            索引 = 读索引(运行目录) or 索引
            索引["步骤"] = "identify_a" if 侧 == "A" else "identify_b"
            索引["待答复"] = 结果["待答复"]
            写索引(运行目录, 索引)
            return 索引

    索引 = 读索引(运行目录) or 索引
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
    索引["步骤"] = "carry"
    写索引(运行目录, 索引)

    变更, 问题 = 生成更新表({}, 清单, b记录, a记录)
    超额问题 = [{"原因": x.get("原因"), "表组": x.get("表组"), "槽位": x.get("槽位")}
            for x in 清单 if str(x.get("原因") or "").startswith("B多出项")
            or str(x.get("原因") or "").startswith("B空项")]
    问题 = (_读jsonl(运行目录 / "识别问题_A.jsonl") + _读jsonl(运行目录 / "识别问题_B.jsonl")
          + 问题 + 粗比问题 + 超额问题)
    a按格 = {r.get("格"): r for r in a记录}
    标题更新 = 读标题更新(A结构)          # 改哪段、改成什么由识别给，没给就不动标题
    _写jsonl(运行目录 / "问题清单.jsonl", 问题)

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

    区域到组 = {(x.get("区域") or ""): (x.get("科目") or "", x.get("表名") or "")
             for x in A结构 if x.get("区域")}
    Word路径, 改2 = _避让(运行目录 / "Word新副本.docx", 已录哈希.get("Word新副本"))
    撤回 = set()
    while True:
        try:
            写回Word(产物["带连接Word"], 更新表路径, Word路径,
                  operations=[], caption_updates=标题更新)
            break
        except Exception as e:
            区域 = _错误里的区域(str(e), 区域到组)
            if not 区域 or 区域 in 撤回:
                raise
            撤回.add(区域)                       # 这张表写不进去：恢复原值，其余照常交付
            组 = 区域到组[区域]
            变更 = {格: 值 for 格, 值 in 变更.items()
                    if (a按格.get(格, {}).get("科目"), a按格.get(格, {}).get("表名")) != 组}
            标题更新 = [u for u in 标题更新 if u.get("range_name") != 区域]
            保存更新表(产物["表A"], 变更, 更新表路径)
            _登记(索引, "更新后附注表格", 更新表路径)
            问题.append({"原因": "该表回写失败，已保持原值，其余表照常交付",
                        "区域": 区域, "科目": 组[0], "表名": 组[1], "错误": str(e)[:300]})
            _写jsonl(运行目录 / "问题清单.jsonl", 问题)
            写索引(运行目录, 索引)
    _登记(索引, "Word新副本", Word路径)
    if 改2:
        问题.append({"原因": "原 Word 新副本被改过，已另存新文件，原文件保留",
                    "路径": str(Word路径)})
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)
    索引["步骤"] = "readback"
    写索引(运行目录, 索引)

    差异 = 回读验证(Word路径, 变更, 运行目录 / "回读临时")
    _写jsonl(运行目录 / "回读记录.jsonl", 差异)      # 差异文件保留，不删证据
    if 差异:
        问题 += [{"原因": "回读不一致：写进 Word 后回读对不上，明细见 回读记录.jsonl",
                  "差异": d} for d in 差异]
        _写jsonl(运行目录 / "问题清单.jsonl", 问题)   # 流程末尾把所有问题收齐在一份清单里
    索引["产物哈希"] = {"更新后附注表格": _哈希(更新表路径), "Word新副本": _哈希(Word路径)}
    索引["回读差异"] = len(差异)
    索引["撤回表"] = sorted(撤回)
    索引["步骤"] = "回读不一致" if 差异 else "done"   # 有差异不许标完成
    索引["产物"].update({"更新后附注表格": str(更新表路径),
                   "Word新副本": str(Word路径)})
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
