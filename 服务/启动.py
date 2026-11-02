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
