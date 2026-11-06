# coding: utf-8
"""双击启动旧版浏览器页面（仅本机127.0.0.1），与桌面版共用单实例和任务入口。

无控制台启动时 stdout／stderr 是 None，所以这里给 uvicorn 一份写文件的日志配置，
不碰标准输出；等服务真的起来了才打开浏览器。
"""
import sys
import traceback
import webbrowser
from pathlib import Path

根 = Path(__file__).resolve().parent
sys.path.insert(0, str(根))

from 服务.启动 import 记启动


def 说(文本: str) -> None:
    """有控制台就顺手打一行；没有控制台也不许因此报错。"""
    try:
        if sys.stdout is not None:
            print(文本)
    except Exception:
        pass


if __name__ == "__main__":
    import ctypes
    from 桌面窗口 import 实例锁, 本机服务
    锁 = 实例锁()
    if not 锁.取得():
        ctypes.windll.user32.MessageBoxW(None, "程序已在运行，请使用已有窗口。", "审计附注搬运", 0x40)
        raise SystemExit(0)
    服务 = 本机服务(需要桌面=False)
    try:
        服务.启动()
        webbrowser.open(服务.地址 + "/")
        服务.线程.join()
    except Exception:
        记启动("启动失败：" + traceback.format_exc())
        raise
    finally:
        服务.停止()
        锁.关闭()
