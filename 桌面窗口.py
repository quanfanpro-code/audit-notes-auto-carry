# coding: utf-8
"""Windows桌面入口；窗口只管理本机服务和关闭，不接管财务处理。"""
import ctypes
from ctypes import wintypes
import secrets
import sys
import socket
import threading
import time
from pathlib import Path
from 运行位置 import 软件目录

import uvicorn
from 服务.启动 import 日志配置, 记启动

根 = 软件目录()


class 实例锁:
    def __init__(self, 名="审计附注自动搬运"):
        self.名 = "Local\\" + 名
        self.句柄 = None
        self.库 = ctypes.WinDLL("kernel32", use_last_error=True)
        self.库.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        self.库.CreateMutexW.restype = wintypes.HANDLE
        self.库.CloseHandle.argtypes = [wintypes.HANDLE]

    def 取得(self):
        self.句柄 = self.库.CreateMutexW(None, False, self.名)
        if not self.句柄:
            raise ctypes.WinError(ctypes.get_last_error())
        已有 = ctypes.get_last_error() == 183
        if 已有:
            self.关闭()
        return not 已有

    def 关闭(self):
        if self.句柄:
            self.库.CloseHandle(self.句柄)
            self.句柄 = None


class 本机服务:
    def __init__(self, 页面=None, 需要桌面=True):
        self.页面 = Path(页面 or Path(__file__).resolve().parent / "desktop_ui" / "index.html")
        self.需要桌面 = 需要桌面
        self.服务器 = self.线程 = self.连接 = None
        self.令牌 = secrets.token_urlsafe(32)
        self.地址 = ""
        self.端口 = None

    def 启动(self):
        if self.服务器:
            return
        if self.需要桌面 and not self.页面.is_file():
            raise FileNotFoundError("桌面页面缺失，请取得含desktop_ui的完整程序")
        from 服务.web import 应用
        应用.state.桌面页面 = self.页面
        应用.state.会话令牌 = self.令牌
        应用.state.退出中 = False
        self.连接 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.连接.bind(("127.0.0.1", 0))
        self.端口 = self.连接.getsockname()[1]
        self.地址 = f"http://127.0.0.1:{self.端口}"
        self.服务器 = uvicorn.Server(uvicorn.Config(应用, log_config=日志配置(), log_level="info"))
        self.线程 = threading.Thread(target=self.服务器.run, kwargs={"sockets": [self.连接]}, daemon=True)
        self.线程.start()
        截止 = time.monotonic() + 30
        while not self.服务器.started and self.线程.is_alive() and time.monotonic() < 截止:
            time.sleep(0.05)
        if not self.服务器.started:
            self.停止()
            raise RuntimeError("本机服务未能启动，请查看运行目录中的服务日志")
        记启动("桌面服务就绪 " + self.地址)

    def 停止(self):
        if self.服务器:
            self.服务器.should_exit = True
        if self.线程 and self.线程.is_alive():
            self.线程.join(timeout=10)
        if self.连接:
            self.连接.close()
        from 服务.web import 应用
        if getattr(应用.state, "会话令牌", None) == self.令牌:
            应用.state.会话令牌 = None
        self.服务器 = self.线程 = self.连接 = None


class 窗口桥:
    def __init__(self):
        self._窗口 = None
        self.未保存 = False
        self.确认中 = False
        self.允许关闭 = False

    def 标记未保存(self, 有未保存):
        self.未保存 = bool(有未保存)

    def _选文件(self, 类型):
        from webview import FileDialog
        格式 = {"报表": "Word或Excel (*.docx;*.xlsx)", "Word": "Word (*.docx)", "Excel": "Excel (*.xlsx)",
                "金标准": "金标准 (*.jsonl;*.json)", "结构操作": "结构操作 (*.json)", "另存为Word": "Word (*.docx)"}
        选择 = self._窗口.create_file_dialog(
            FileDialog.FOLDER if 类型 == "文件夹" else FileDialog.SAVE if 类型 == "另存为Word" else FileDialog.OPEN,
            directory=str(根), file_types=(格式.get(类型, "所有文件 (*.*)"),),
            save_filename="新副本.docx" if 类型 == "另存为Word" else "")
        return 选择[0] if 选择 else ""

    def 关闭事件(self):
        from 服务 import web
        from 服务.任务服务 import 正在执行, 执行锁
        if self.允许关闭:
            return True
        with 执行锁:
            if not self.未保存 and not 正在执行(web):
                web.应用.state.退出中 = True
                return True
        if not self.确认中:
            self.确认中 = True
            threading.Thread(target=self._确认退出, daemon=True).start()
        return False

    def _确认退出(self):
        from 服务 import web
        from 服务.任务服务 import 正在执行, 中断任务, 读说明, 执行锁
        try:
            if self.未保存 and not self._窗口.create_confirmation_dialog("尚未保存", "设置或答复草稿尚未保存。放弃草稿并退出？"):
                return
            with 执行锁:
                web.应用.state.退出中 = True
            if 正在执行(web):
                目录 = web.应用.state.状态["运行目录"]
                人工 = 读说明(目录).get("类型") == "人工回写"
                提示 = ("人工回写正在进行。等待写回安全结束后退出？" if 人工 else
                      "任务正在运行。中断后退出？当前Office操作可能需要等待完成。")
                if not self._窗口.create_confirmation_dialog("任务尚未结束", 提示):
                    return
                if not 人工:
                    中断任务(web, 目录)
                self._窗口.set_title("审计附注搬运 · 等待当前操作安全结束")
                while 正在执行(web):
                    time.sleep(0.2)
            self.允许关闭 = True
            self._窗口.destroy()
        finally:
            self.确认中 = False
            if not self.允许关闭:
                web.应用.state.退出中 = False


def 启动桌面():
    锁 = 实例锁()
    if not 锁.取得():
        ctypes.windll.user32.MessageBoxW(None, "程序已在运行，请回到已经打开的窗口。", "审计附注搬运", 0x40)
        return
    服务 = 本机服务()
    try:
        if getattr(sys, "frozen", False):
            import os
            os.environ["PYTHONNET_RUNTIME"] = "coreclr"
            from pythonnet import load
            load("coreclr", runtime_config=str(根 / "运行库" / "桌面.runtimeconfig.json"),
                 dotnet_root=str(根 / "运行库" / "dotnet"))
        import webview
        if getattr(sys, "frozen", False):
            浏览器 = 根 / "运行库" / "WebView2"
            if not (浏览器 / "msedgewebview2.exe").is_file():
                raise FileNotFoundError("软件文件夹缺少WebView2运行库，请使用完整的软件文件夹")
            webview.settings["WEBVIEW2_RUNTIME_PATH"] = str(浏览器)
        服务.启动()
        桥 = 窗口桥()
        屏幕 = webview.screens[0]
        工作区 = 屏幕.frame
        宽 = min(1440, 工作区.Width - 40)
        高 = min(900, 工作区.Height - 60)
        窗口 = webview.create_window("审计附注搬运", 服务.地址 + "/桌面", js_api=桥,
                                 width=宽, height=高, min_size=(min(900, 宽), min(560, 高)),
                                 x=工作区.X + (工作区.Width - 宽) // 2,
                                 y=工作区.Y + (工作区.Height - 高) // 2,
                                 background_color="#F4F6FA", text_select=True)
        桥._窗口 = 窗口
        from 服务.web import 应用
        应用.state.文件选择器 = 桥._选文件
        窗口.events.closing += 桥.关闭事件
        webview.start(gui="edgechromium", storage_path=str(根 / "运行" / "浏览器缓存"), private_mode=False)
    finally:
        from 服务.web import 应用
        应用.state.文件选择器 = None
        服务.停止()
        锁.关闭()
