# coding: utf-8
"""在当前新版目录编译精简的Windows x64桥接，不依赖旧版目录。"""
import ctypes
import hashlib
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

根 = Path(__file__).resolve().parent
sys.path.insert(0, str(根))
from 核心.桥接 import 检查桥接运行环境, NET下载地址


def 查找SDK():
    候选 = [shutil.which("dotnet")]
    for 目录 in (os.environ.get("DOTNET_ROOT"), os.environ.get("ProgramW6432"), os.environ.get("ProgramFiles")):
        if 目录:
            候选.append(str(Path(目录) / ("dotnet.exe" if 目录 == os.environ.get("DOTNET_ROOT") else "dotnet/dotnet.exe")))
    for 候 in dict.fromkeys(候选):
        if not 候 or not Path(候).is_file():
            continue
        try:
            结果 = subprocess.run([候, "--list-sdks"], capture_output=True, encoding="utf-8", errors="replace",
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=15)
            if 结果.returncode == 0 and any(行.split(".")[0].isdigit() and int(行.split(".")[0]) >= 10
                                         for 行 in 结果.stdout.splitlines()):
                return str(Path(候).resolve())
        except (OSError, subprocess.TimeoutExpired):
            continue
    raise RuntimeError("编译需要 .NET 10 或更新版本的 SDK。仅安装运行时不能编译。\n"
                       "请安装Windows x64 SDK后重新运行 编译Word桥接.py。\n" + NET下载地址)


def 编译():
    sdk = 查找SDK()
    项目 = 根 / "Word桥接" / "Word桥接.csproj"
    if not 项目.is_file():
        raise RuntimeError("缺少Word桥接源码，请重新下载完整的新版仓库。")
    本次 = 根 / "运行" / ("桥接编译_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
    本次.mkdir(parents=True)
    暂存 = 本次 / "发布"
    命令 = [sdk, "publish", str(项目), "-c", "Release", "-r", "win-x64", "--self-contained", "false",
          "-o", str(暂存), "-p:BaseIntermediateOutputPath=" + str(本次 / "obj") + os.sep,
          "-p:BaseOutputPath=" + str(本次 / "bin") + os.sep]
    结果 = subprocess.run(命令, cwd=根, capture_output=True, encoding="utf-8", errors="replace",
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=900)
    日志 = 本次 / "编译日志.txt"
    日志.write_text(结果.stdout + "\n" + 结果.stderr, encoding="utf-8-sig")
    if 结果.returncode:
        raise RuntimeError("Word桥接编译失败，原发布文件未替换。详细原因见：" + str(日志))
    for 名 in ("WordBridge.exe", "WordBridge.dll", "WordBridge.deps.json", "WordBridge.runtimeconfig.json"):
        if not (暂存 / 名).is_file():
            raise RuntimeError("编译输出缺少 " + 名 + "，原发布文件未替换。")
    检查桥接运行环境(暂存 / "WordBridge.exe")
    发布 = 根 / "Word桥接" / "发布"
    已有 = list(发布.rglob("*")) if 发布.exists() else []
    if any(p.is_file() for p in 已有):
        备份 = Path.home() / "BackUp" / ("audit-notes-auto-carry_桥接发布_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        for 原 in 已有:
            if 原.is_file():
                副本 = 备份 / "Word桥接" / "发布" / 原.relative_to(发布)
                副本.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(原, 副本)
                if hashlib.sha256(原.read_bytes()).digest() != hashlib.sha256(副本.read_bytes()).digest():
                    raise RuntimeError("现有发布文件备份校验失败，已停止替换。")
    发布.mkdir(parents=True, exist_ok=True)
    shutil.copytree(暂存, 发布, dirs_exist_ok=True)
    检查桥接运行环境(发布 / "WordBridge.exe")
    return 发布


if __name__ == "__main__":
    try:
        发布 = 编译()
        消息 = "Word桥接编译成功：\n" + str(发布) + "\n以后使用 启动服务.pyw。换电脑运行需安装 .NET 10 Windows x64运行时。"
        状态 = 0
    except Exception as e:
        消息 = "Word桥接编译未完成：\n" + str(e)
        状态 = 1
    if sys.stdout is not None:
        print(消息)
    ctypes.windll.user32.MessageBoxW(None, 消息, "Word桥接编译", 0x40 if 状态 == 0 else 0x10)
    raise SystemExit(状态)
