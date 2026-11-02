# coding: utf-8
"""提取编排：Word → 表A＋带连接Word＋上下文，产物登记接续索引。"""
from pathlib import Path

from .桥接 import 提取 as 调用桥接提取
from .接续 import 读索引, 写索引


def 提取任务(word路径, 运行目录) -> dict:
    运行目录 = Path(运行目录)
    运行目录.mkdir(parents=True, exist_ok=True)
    表A = 运行目录 / "表A.xlsx"
    带连接 = 运行目录 / "带连接Word.docx"
    上下文 = 运行目录 / "上下文.json"
    for 路径 in (表A, 带连接, 上下文):
        if 路径.exists():
            raise FileExistsError(f"提取产物已存在，拒绝覆盖：{路径}")
    结果 = 调用桥接提取(str(word路径), str(表A), str(带连接), str(上下文)) or {}
    产物 = {"表A": str(表A), "带连接Word": str(带连接), "上下文": str(上下文)}
    索引 = 读索引(运行目录) or {"已完成范围": [], "产物": {}}
    索引["产物"].update(产物)
    索引["桥接提示"] = [str(x) for x in (结果.get("warnings") or [])]   # 桥接的提示要留痕
    索引["步骤"] = "identify_a"          # 提取完成，下一步待办
    写索引(运行目录, 索引)
    return 产物
