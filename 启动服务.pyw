# coding: utf-8
"""启动本地 WebUI 并打开浏览器（仅本机 127.0.0.1）。"""
import sys
import threading
import time
import webbrowser
from pathlib import Path

import uvicorn

sys.path.insert(0, str(Path(__file__).resolve().parent))

if __name__ == "__main__":
    端口 = 8765

    def 开浏览器():
        time.sleep(2)
        webbrowser.open(f"http://127.0.0.1:{端口}/")

    threading.Thread(target=开浏览器, daemon=True).start()
    uvicorn.run("服务.web:应用", host="127.0.0.1", port=端口)
