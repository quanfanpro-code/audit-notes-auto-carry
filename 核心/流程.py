# coding: utf-8
"""总流程编排：提取→识别A→识别B→比对→搬运→写回→回读，接续索引逐步登记。"""
import json
from pathlib import Path

from .比对 import 建立对应
from .写回 import 构造标题更新, 写回Word
from .回读 import 回读验证
from .提取 import 提取任务
from .接续 import 读索引, 写索引
from .搬运 import 生成更新表, 保存更新表
from .结构维护 import 生成结构操作
from .识别 import 识别一侧


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

    变更, 问题 = 生成更新表({}, 清单, b记录)
    _写jsonl(运行目录 / "问题清单.jsonl", 问题)
    更新表路径 = 运行目录 / "更新后附注表格.xlsx"
    保存更新表(产物["表A"], 变更, 更新表路径)
    索引["步骤"] = "writeback"
    写索引(运行目录, 索引)

    标题更新 = 构造标题更新(清单, a记录, 上下文)
    写回Word(产物["带连接Word"], 更新表路径, 运行目录 / "Word新副本.docx",
          operations=[], caption_updates=标题更新)
    索引["步骤"] = "readback"
    写索引(运行目录, 索引)

    差异 = 回读验证(运行目录 / "Word新副本.docx", 变更, 运行目录 / "回读临时")
    _写jsonl(运行目录 / "回读记录.jsonl", 差异)
    索引["步骤"] = "done"
    索引["产物"].update({"更新后附注表格": str(更新表路径),
                   "Word新副本": str(运行目录 / "Word新副本.docx")})
    写索引(运行目录, 索引)
    return 索引


def 人工回写(输入: dict) -> dict:
    """人工直接回写：按真实路径重读用户已保存的 Excel；给基准表A时先推导结构操作。"""
    ops = []
    if 输入.get("基准表A") and 输入.get("excel"):
        ops = 生成结构操作(输入["基准表A"], 输入["excel"])
    return 写回Word(输入["带连接word"], 输入["excel"], 输入["output"], operations=ops)

