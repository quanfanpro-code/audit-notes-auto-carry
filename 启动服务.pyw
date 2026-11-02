# coding: utf-8
"""双击启动本地 WebUI（仅本机 127.0.0.1）。端口被占就往后顺延，并记启动日志。"""
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

import uvicorn

根 = Path(__file__).resolve().parent
sys.path.insert(0, str(根))

from 服务.启动 import 选端口, 记启动

if __name__ == "__main__":
    try:
        端口 = 选端口()
        地址 = f"http://127.0.0.1:{端口}/"
        记启动("启动 " + 地址)
        print("服务已启动：" + 地址)              # 有控制台时能看到实际端口

        def 开浏览器():
            time.sleep(2)
            webbrowser.open(地址)

        threading.Thread(target=开浏览器, daemon=True).start()
        uvicorn.run("服务.web:应用", host="127.0.0.1", port=端口, log_level="warning")
    except Exception:
        记启动("启动失败：\n" + traceback.format_exc())
        raise
