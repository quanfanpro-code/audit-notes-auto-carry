# coding: utf-8
"""验证真实桥接可运行，以及缺运行环境时普通用户能得到安装指引。"""
import importlib.util
import json
import runpy
import subprocess
import sys
from pathlib import Path

import pytest

根 = Path(__file__).resolve().parent.parent


def 编译入口():
    路径 = 根 / "编译Word桥接.py"
    assert 路径.is_file(), "新版缺少中文Python编译入口"
    spec = importlib.util.spec_from_file_location("桥接编译测试", 路径)
    模块 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(模块)
    return 模块


def test_新版桥接可真实启动():
    from 核心.桥接 import 桥接程序, 检查桥接运行环境
    assert 桥接程序.is_file(), "新版尚未从自己的源码编译桥接"
    检查桥接运行环境()
    配置 = json.loads(桥接程序.with_suffix(".runtimeconfig.json").read_text(encoding="utf-8"))
    assert 配置["runtimeOptions"]["framework"]["name"] == "Microsoft.NETCore.App"
    assert 配置["runtimeOptions"]["framework"]["version"].startswith("10.")


def test_没有SDK时给出编译安装指引(monkeypatch):
    模块 = 编译入口()
    monkeypatch.setattr(模块.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 0, "9.0.320 [SDK]", ""))
    with pytest.raises(RuntimeError, match=r"\.NET 10.*SDK"):
        模块.查找SDK()


@pytest.mark.parametrize("退出码", [0x80008096, -2147450730, 0x80008083])
def test_缺运行时提示正确版本和下载地址(tmp_path, monkeypatch, 退出码):
    import 核心.桥接 as 桥接
    exe = tmp_path / "WordBridge.exe"
    exe.write_bytes("占位，仅替代外部进程".encode("utf-8"))
    monkeypatch.setattr(桥接.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 退出码, "", "host error"))
    with pytest.raises(RuntimeError) as e:
        桥接.检查桥接运行环境(exe)
    assert ".NET 10" in str(e.value)
    assert "Windows x64" in str(e.value)
    assert "https://dotnet.microsoft.com/" in str(e.value)
    assert "重新启动" in str(e.value)


def test_实际调用缺运行时也有中文安装提示(tmp_path, monkeypatch):
    import 核心.桥接 as 桥接
    exe = tmp_path / "WordBridge.exe"
    exe.write_bytes("占位".encode("utf-8"))
    monkeypatch.setattr(桥接.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 0x80008096, "", "host error"))
    with pytest.raises(RuntimeError, match=r"\.NET 10"):
        桥接.调用桥接({"action": "read", "word": "不会读取.docx"}, exe=exe, 工作目录=tmp_path / "请求")


def test_缺发布文件不能谎称只缺NET(tmp_path):
    from 核心.桥接 import 检查桥接运行环境
    with pytest.raises(RuntimeError, match="编译Word桥接.py"):
        检查桥接运行环境(tmp_path / "不存在.exe")


def test_损坏发布包不能误报缺运行时(tmp_path, monkeypatch):
    import 核心.桥接 as 桥接
    exe = tmp_path / "WordBridge.exe"
    exe.write_bytes("占位".encode("utf-8"))
    monkeypatch.setattr(桥接.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 1, "", "缺少应用程序集"))
    with pytest.raises(RuntimeError, match="缺少应用程序集") as e:
        桥接.检查桥接运行环境(exe)
    assert "安装 .NET" not in str(e.value)


def test_桥接返回错误结构时也能明确提示(tmp_path, monkeypatch):
    import 核心.桥接 as 桥接
    exe = tmp_path / "WordBridge.exe"
    exe.write_bytes(b"placeholder")
    monkeypatch.setattr(桥接.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 1, "[]", "损坏的发布包"))
    with pytest.raises(RuntimeError, match="重新编译"):
        桥接.检查桥接运行环境(exe)


def test_启动探针超时明确停止(tmp_path, monkeypatch):
    import 核心.桥接 as 桥接
    exe = tmp_path / "WordBridge.exe"
    exe.write_bytes(b"placeholder")
    def 卡住(*a, **k):
        raise subprocess.TimeoutExpired(a[0], 15)
    monkeypatch.setattr(桥接.subprocess, "run", 卡住)
    with pytest.raises(RuntimeError, match="启动检查超时"):
        桥接.检查桥接运行环境(exe)


def test_桥接检查失败时启动入口停止(monkeypatch):
    import 服务.启动 as 启动
    monkeypatch.setattr(启动, "准备桥接", lambda: False, raising=False)
    def 不应选端口(*a, **k):
        pytest.fail("环境失败后仍然启动了服务")
    monkeypatch.setattr(启动, "选端口", 不应选端口)
    with pytest.raises(SystemExit) as e:
        runpy.run_path(str(根 / "启动服务.pyw"), run_name="__main__")
    assert e.value.code == 1


def test_真实隔离NET安装仍能给出中文提示(tmp_path, monkeypatch):
    from 核心.桥接 import 检查桥接运行环境
    空环境 = tmp_path / "无NET"
    空环境.mkdir()
    for 名 in ("DOTNET_ROOT_X64", "DOTNET_ROOT"):
        monkeypatch.setenv(名, str(空环境))
    monkeypatch.setenv("DOTNET_DISABLE_GUI_ERRORS", "1")
    with pytest.raises(RuntimeError, match=r"\.NET 10"):
        检查桥接运行环境()


def test_NET不在PATH但桥接可运行时不误报(monkeypatch):
    from 核心.桥接 import 检查桥接运行环境
    monkeypatch.setenv("PATH", "")
    检查桥接运行环境()


@pytest.mark.parametrize("选择,应打开", [(6, True), (7, False)])
def test_安装提示经用户选择才打开下载页(tmp_path, monkeypatch, 选择, 应打开):
    import ctypes
    import webbrowser
    import 服务.启动 as 启动
    import 核心.桥接 as 桥接
    def 缺环境():
        raise RuntimeError("请安装 .NET 10 Windows x64，安装后重新启动\n" + 桥接.NET下载地址)
    monkeypatch.setattr(桥接, "检查桥接运行环境", 缺环境)
    monkeypatch.setattr(启动, "日志路径", tmp_path / "启动日志.txt")
    提示, 打开 = [], []
    def 窗口(_, 文本, 标题, 样式):
        提示.append(文本)
        return 选择
    monkeypatch.setattr(ctypes.windll.user32, "MessageBoxW", 窗口)
    monkeypatch.setattr(webbrowser, "open", lambda 地址: 打开.append(地址))
    assert 启动.准备桥接() is False
    assert len(提示) == 1 and ".NET 10" in 提示[0]
    assert 打开 == ([桥接.NET下载地址] if 应打开 else [])
    assert "启动失败" in 启动.日志路径.read_text(encoding="utf-8")


def test_已有环境不弹安装提示(monkeypatch):
    import ctypes
    import 服务.启动 as 启动
    import 核心.桥接 as 桥接
    monkeypatch.setattr(桥接, "检查桥接运行环境", lambda: None)
    def 不应弹窗(*args):
        pytest.fail("已有环境仍提示安装")
    monkeypatch.setattr(ctypes.windll.user32, "MessageBoxW", 不应弹窗)
    assert 启动.准备桥接() is True


def test_编译失败不替换原发布文件(tmp_path, monkeypatch):
    模块 = 编译入口()
    monkeypatch.setattr(模块, "根", tmp_path)
    monkeypatch.setattr(模块, "查找SDK", lambda: sys.executable)
    项目 = tmp_path / "Word桥接/Word桥接.csproj"
    项目.parent.mkdir(parents=True)
    项目.write_text("<Project/>", encoding="utf-8")
    发布 = 项目.parent / "发布"
    发布.mkdir()
    原 = 发布 / "WordBridge.exe"
    原.write_bytes(b"existing-working-release")
    monkeypatch.setattr(模块.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 1, "故意模拟编译错误", ""))
    with pytest.raises(RuntimeError, match="原发布文件未替换"):
        模块.编译()
    assert 原.read_bytes() == b"existing-working-release"
