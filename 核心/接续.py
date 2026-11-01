# coding: utf-8
"""准确接续索引：记录进度快照，断点续接的依据。"""
import json
import os
from pathlib import Path


def 写索引(运行目录, 状态: dict) -> None:
    路径 = Path(运行目录) / "接续索引.json"
    临时 = 路径.with_suffix(".json.tmp")
    临时.write_text(json.dumps(状态, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(临时, 路径)


def 读索引(运行目录):
    路径 = Path(运行目录) / "接续索引.json"
    if not 路径.is_file():
        return None
    return json.loads(路径.read_text(encoding="utf-8"))


def 待办步骤(索引: dict):
    步骤 = 索引.get("步骤")
    return None if 步骤 == "done" else 步骤
