# coding: utf-8
"""Word 桥接封装：子进程调用 WordBridge，UTF-8 JSON 请求/响应。
主体抄自老项目 附注更新/流程.py 的 bridge()（带错误处理），接口按新计划命名。"""
import json
import os
import subprocess
import uuid
from pathlib import Path

桥接程序 = Path(__file__).resolve().parent.parent / "Word桥接" / "发布" / "WordBridge.exe"
桥接超时秒 = 1800      # 大附注打开+写回要几分钟，给足；卡住就明确报错，不再无限等


def 调用桥接(请求: dict, 工作目录=None, exe=None, 前缀=None):
    exe = Path(exe) if exe else 桥接程序
    if not exe.is_file():
        raise FileNotFoundError(f"Word桥接程序不存在：{exe}")
    工作目录 = (Path(工作目录) if 工作目录 else
              Path(__file__).resolve().parent.parent / "运行" / "桥接留证" / uuid.uuid4().hex)
    工作目录.mkdir(parents=True, exist_ok=True)
    请求文件 = 工作目录 / ("Word请求_" + uuid.uuid4().hex[:8] + ".json")
    请求文件.write_text(json.dumps(请求, ensure_ascii=False), encoding="utf-8-sig")
    命令 = ([str(前缀)] if 前缀 else []) + [str(exe), str(请求文件)]
    # shortcut: 系统 %TEMP% 对该 exe 的临时文件访问被拒（已定位复现）, 环境修复后可去掉 TEMP 重定向
    临时 = 工作目录 / "桥接临时"
    临时.mkdir(exist_ok=True)
    环境 = {**os.environ, "TEMP": str(临时), "TMP": str(临时)}
    try:
        完成 = subprocess.run(命令, capture_output=True, encoding="utf-8", errors="strict",
                         env=环境, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                         timeout=桥接超时秒)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"Word桥接超过 {桥接超时秒} 秒没有返回，已强制结束：{命令[-1]}") from e
    try:
        结果 = json.loads(完成.stdout.lstrip("﻿").strip())
    except json.JSONDecodeError as e:
        raise RuntimeError("Word桥接没有返回有效结果：" + (完成.stderr[:1000] or "")) from e
    if 完成.returncode or not 结果.get("ok"):
        raise ValueError(结果.get("error") or 完成.stderr or "Word操作失败")
    return 结果


def 提取(word, excel, 带连接word, 上下文, exe=None, 前缀=None, 工作目录=None):
    return 调用桥接({"action": "extract", "word": str(word), "excel": str(excel),
                 "linked_word": str(带连接word), "context": str(上下文)},
                exe=exe, 前缀=前缀, 工作目录=工作目录)


def 读取(word, exe=None, 前缀=None, 工作目录=None):
    return 调用桥接({"action": "read", "word": str(word)},
                exe=exe, 前缀=前缀, 工作目录=工作目录)


def 更新(word, excel, output, operations=None, caption_updates=None, exe=None, 前缀=None, 工作目录=None):
    return 调用桥接({"action": "update", "word": str(word), "excel": str(excel), "output": str(output),
                 "operations": operations or [], "caption_updates": caption_updates or []},
                exe=exe, 前缀=前缀, 工作目录=工作目录)
