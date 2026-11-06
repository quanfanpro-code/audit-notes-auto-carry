# coding: utf-8
"""双击打开桌面工作台；失败通过Windows窗口说明。"""
import ctypes
import traceback
from 桌面窗口 import 启动桌面
from 服务.启动 import 记启动

if __name__ == "__main__":
    try:
        启动桌面()
    except Exception as e:
        记启动("桌面启动失败：" + traceback.format_exc())
        ctypes.windll.user32.MessageBoxW(None, str(e), "审计附注搬运启动失败", 0x10)
