# coding: utf-8
"""双击打开桌面工作台；失败通过Windows窗口说明。"""
import ctypes
import traceback
import sys
import runpy

# 冻结后的Excel子进程仍执行原模块，不再误开一个桌面窗口。
if getattr(sys, "frozen", False) and sys.argv[1:3] == ["-m", "核心.Excel交付"]:
    sys.argv = sys.argv[2:]
    runpy.run_module("核心.Excel交付", run_name="__main__")
    raise SystemExit(0)
from 桌面窗口 import 启动桌面
from 服务.启动 import 记启动

if __name__ == "__main__":
    try:
        启动桌面()
    except Exception as e:
        记启动("桌面启动失败：" + traceback.format_exc())
        ctypes.windll.user32.MessageBoxW(None, str(e), "审计附注搬运启动失败", 0x10)
