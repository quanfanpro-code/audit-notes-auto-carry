# coding: utf-8
"""总流程编排：提取→识别A→识别B→比对→搬运→写回→回读，接续索引逐步登记。"""
import hashlib
import json
from datetime import datetime
from pathlib import Path

from .比对 import 建立对应
from .写回 import 构造标题更新, 写回Word
from .回读 import 回读验证
from .提取 import 提取任务
from .接续 import 读索引, 写索引
from .搬运 import 生成更新表, 保存更新表
from .结构维护 import 生成结构操作
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
    """产物被人改过就不覆盖：另存带时间戳的新文件，原文件留着。"""
    if not 旧哈希:
        return 路径, False
    现在 = _哈希(路径)
    if not 现在 or 现在 == 旧哈希:
        return 路径, False
    新 = 路径.with_name(f"{路径.stem}_人工改过_{datetime.now():%Y%m%d_%H%M%S}{路径.suffix}")
    return 新, True


def _读jsonl(路径):
    p = Path(路径)
    return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.is_file() else []


def _写jsonl(路径, 记录):
    with Path(路径).open("w", encoding="utf-8") as f:
        for r in 记录:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def _指定工作表(输入: dict):
    值 = 输入.get("b工作表")
    if not 值:
        return None
    return [值] if isinstance(值, str) else list(值)


def 运行任务(输入: dict, 运行目录, 配置: dict, 取消=None) -> dict:
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    索引 = 读索引(运行目录) or {"步骤": "extract", "已完成范围": [], "产物": {}}
    产物 = 索引.setdefault("产物", {})
    if "表A" not in 产物:
        产物.update(提取任务(输入["word"], 运行目录))
        索引["步骤"] = "identify_a"
        写索引(运行目录, 索引)
    上下文 = json.loads(Path(产物["上下文"]).read_text(encoding="utf-8-sig"))
    金标准 = 输入.get("金标准")

    for 侧, 表文件, 表上下文, 只跑表 in (("A", 产物["表A"], 上下文.get("tables", []), None),
                                  ("B", 输入["b表"], None, _指定工作表(输入))):
        结果 = 识别一侧(配置, 表文件, 侧, 运行目录, 金标准, 表上下文,
                   取消=取消, 只跑表=只跑表)
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
    清单 = 建立对应(a记录, b记录,
                 [x for x in 结构 if x.get("侧") == "A"],
                 [x for x in 结构 if x.get("侧") == "B"])
    _写jsonl(运行目录 / "对应清单.jsonl", 清单)
    索引["步骤"] = "carry"
    写索引(运行目录, 索引)

    变更, 问题 = 生成更新表({}, 清单, b记录, a记录)
    a按格 = {r.get("格"): r for r in a记录}
    标题更新, 标题失败 = 构造标题更新(清单, a记录, 上下文)
    if 标题失败:
        # 标题改不了的表，这张表一个数都不换：不留下“旧标题 + 新数字”
        失败表 = {(x["科目"], x["表名"]) for x in 标题失败}
        变更 = {格: 值 for 格, 值 in 变更.items()
                if (a按格.get(格, {}).get("科目"), a按格.get(格, {}).get("表名")) not in 失败表}
        问题 = 问题 + [{"原因": "标题改不了，该表不换数（" + x["原因"] + "）",
                      "科目": x["科目"], "表名": x["表名"]} for x in 标题失败]
    _写jsonl(运行目录 / "问题清单.jsonl", 问题)
    已录哈希 = 索引.get("产物哈希") or {}
    更新表路径, _改1 = _避让(运行目录 / "更新后附注表格.xlsx", 已录哈希.get("更新后附注表格"))
    保存更新表(产物["表A"], 变更, 更新表路径)
    索引["步骤"] = "writeback"
    写索引(运行目录, 索引)

    Word路径, _改2 = _避让(运行目录 / "Word新副本.docx", 已录哈希.get("Word新副本"))
    写回Word(产物["带连接Word"], 更新表路径, Word路径,
          operations=[], caption_updates=标题更新)
    索引["步骤"] = "readback"
    写索引(运行目录, 索引)

    差异 = 回读验证(Word路径, 变更, 运行目录 / "回读临时")
    _写jsonl(运行目录 / "回读记录.jsonl", 差异)      # 差异文件保留，不删证据
    索引["产物哈希"] = {"更新后附注表格": _哈希(更新表路径), "Word新副本": _哈希(Word路径)}
    索引["回读差异"] = len(差异)
    索引["步骤"] = "回读不一致" if 差异 else "done"   # 有差异不许标完成
    索引["产物"].update({"更新后附注表格": str(更新表路径),
                   "Word新副本": str(Word路径)})
    写索引(运行目录, 索引)
    return 索引


def 人工回写(输入: dict) -> dict:
    """人工直接回写：按真实路径重读用户已保存的 Excel；给基准表A时先推导结构操作。

    输出路径前端叫「输出」、内部叫 output，两种都认，免得字段名对不上白跑一趟。
    """
    输出 = 输入.get("output") or 输入.get("输出")
    if not 输出:
        raise ValueError("缺少输出路径（输出／output）")
    ops = []
    if 输入.get("基准表A") and 输入.get("excel"):
        ops = 生成结构操作(输入["基准表A"], 输入["excel"])
    return 写回Word(输入["带连接word"], 输入["excel"], 输出, operations=ops)

