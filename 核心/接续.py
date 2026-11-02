# coding: utf-8
"""准确接续索引：记录进度快照，断点续接的依据。"""
import json
import os
import time
from pathlib import Path


def 写索引(运行目录, 状态: dict) -> None:
    路径 = Path(运行目录) / "接续索引.json"
    临时 = 路径.with_suffix(".json.tmp")
    临时.write_text(json.dumps(状态, ensure_ascii=False, indent=2), encoding="utf-8")
    # shortcut: Windows 上杀软／索引器会短暂占用刚写的文件，改名偶尔报 WinError 5；重试几次，仍失败就照实抛
    for 次 in range(5):
        try:
            os.replace(临时, 路径)
            return
        except PermissionError:
            if 次 == 4:
                raise
            time.sleep(0.05)


def 读索引(运行目录):
    路径 = Path(运行目录) / "接续索引.json"
    if not 路径.is_file():
        return None
    return json.loads(路径.read_text(encoding="utf-8"))


def 待办步骤(索引: dict):
    步骤 = 索引.get("步骤")
    return None if 步骤 == "done" else 步骤
