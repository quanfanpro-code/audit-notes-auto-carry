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
NET下载地址 = "https://dotnet.microsoft.com/zh-cn/download/dotnet/10.0"


def _检查宿主错误(完成):
    # .NET宿主标准退出码；不依赖错误消息语言，也不把普通业务错误当成缺环境。
    if 完成.returncode & 0xFFFFFFFF in {0x80008083, 0x80008096}:
        raise RuntimeError("Word桥接缺少可用的 .NET 10 运行时（Windows x64）。\n"
                           "请安装微软 .NET Runtime 10 的 Windows x64 版本，安装后重新启动本程序。\n"
                           "已安装其他版本或仅32位版本，仍可能无法运行此桥接。\n" + NET下载地址)


def 检查桥接运行环境(exe=None):
    程序 = Path(exe) if exe else 桥接程序
    if not 程序.is_file():
        raise RuntimeError("Word桥接尚未编译或发布文件不完整，请先运行项目内的 编译Word桥接.py。")
    try:
        完成 = subprocess.run([str(程序)], capture_output=True, encoding="utf-8", errors="replace",
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=15)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError("Word桥接启动检查超时，请检查发布文件和系统环境。") from e
    except OSError as e:
        raise RuntimeError("Word桥接无法启动，请检查发布文件和Windows x64兼容性：" + str(e)) from e
    _检查宿主错误(完成)
    try:
        答复 = json.loads(完成.stdout.lstrip("\ufeff").strip())
    except (ValueError, TypeError):
        答复 = {}
    if (完成.returncode != 1 or not isinstance(答复, dict) or 答复.get("ok") is not False
            or 答复.get("error") != "请提供一个 UTF-8 JSON 请求文件路径。"):
        raise RuntimeError("Word桥接发布文件无法正常启动，请重新编译：" + (完成.stderr or 完成.stdout)[:1000])


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
    _检查宿主错误(完成)
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
