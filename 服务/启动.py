# coding: utf-8
"""启动入口的公共逻辑：挑一个没被占用的端口，把启动情况写进日志。

双击启动时没有控制台，出错也看不见，所以启动成功和失败都落一行日志，
位置在「运行\\启动日志.txt」。
"""
import socket
from datetime import datetime
from pathlib import Path

日志路径 = Path(__file__).resolve().parent.parent / "运行" / "启动日志.txt"


def 端口可用(端口: int) -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", 端口))
        return True
    except OSError:
        return False
    finally:
        sock.close()


def 选端口(起始: int = 8765, 次数: int = 20) -> int:
    """从起始端口往后找第一个能用的；全被占用就明确报错，不静默退出。"""
    for 端口 in range(起始, 起始 + 次数):
        if 端口可用(端口):
            return 端口
    raise RuntimeError(f"{起始} 到 {起始 + 次数 - 1} 的端口都被占用，无法启动")


def 记启动(文本: str) -> None:
    日志路径.parent.mkdir(parents=True, exist_ok=True)
    with 日志路径.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {文本}\n")


def 服务日志路径() -> Path:
    return 日志路径.parent / "服务日志.txt"


def 准备桥接() -> bool:
    from 核心.桥接 import 检查桥接运行环境, NET下载地址
    try:
        检查桥接运行环境()
        return True
    except RuntimeError as e:
        import ctypes
        import webbrowser
        原因 = str(e)
        记启动("启动失败：" + 原因)
        if NET下载地址 in 原因:
            选择 = ctypes.windll.user32.MessageBoxW(None, 原因 + "\n\n是否打开微软官方下载页面？",
                                                "需要安装 .NET 10", 0x14)
            if 选择 == 6:
                webbrowser.open(NET下载地址)
        else:
            ctypes.windll.user32.MessageBoxW(None, 原因, "Word桥接尚未准备好", 0x10)
        return False


def 日志配置() -> dict:
    """给 uvicorn 一份写文件的日志配置。

    双击启动没有控制台，sys.stdout／sys.stderr 是 None；uvicorn 默认格式器会去问
    输出流 isatty，直接抛 ValueError: Unable to configure formatter 'default'。
    这里一律写文件，不碰标准输出。
    """
    路径 = 服务日志路径()
    路径.parent.mkdir(parents=True, exist_ok=True)
    return {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {"简单": {"format": "%(asctime)s %(levelname)s %(message)s"}},
        "handlers": {"文件": {"class": "logging.FileHandler", "filename": str(路径),
                            "formatter": "简单", "encoding": "utf-8"}},
        "loggers": {名: {"handlers": ["文件"], "level": "INFO", "propagate": False}
                    for 名 in ("uvicorn", "uvicorn.error", "uvicorn.access")},
    }
