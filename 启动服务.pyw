# coding: utf-8
"""双击启动本地 WebUI（仅本机 127.0.0.1）。端口被占就往后顺延，并记启动日志。

无控制台启动时 stdout／stderr 是 None，所以这里给 uvicorn 一份写文件的日志配置，
不碰标准输出；等服务真的起来了才打开浏览器。
"""
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

import uvicorn

根 = Path(__file__).resolve().parent
sys.path.insert(0, str(根))

from 服务.启动 import 选端口, 记启动, 日志配置, 准备桥接


def 说(文本: str) -> None:
    """有控制台就顺手打一行；没有控制台也不许因此报错。"""
    try:
        if sys.stdout is not None:
            print(文本)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        if not 准备桥接():
            raise SystemExit(1)
        端口 = 选端口()
        地址 = f"http://127.0.0.1:{端口}/"
        记启动("启动 " + 地址)
        说("服务已启动：" + 地址)
        配置 = uvicorn.Config("服务.web:应用", host="127.0.0.1", port=端口,
                            log_level="info", log_config=日志配置())
        服务器 = uvicorn.Server(配置)
        线程 = threading.Thread(target=服务器.run, daemon=True)
        线程.start()
        for _ in range(300):                    # 最多等 30 秒
            if 服务器.started:
                break
            time.sleep(0.1)
        if not 服务器.started:
            记启动("启动失败：服务没有起来，详见 运行\服务日志.txt")
            raise SystemExit(1)
        webbrowser.open(地址)                   # 服务确认起来后才开页面
        线程.join()
    except SystemExit:
        raise
    except Exception:
        记启动("启动失败：" + traceback.format_exc())
        raise
