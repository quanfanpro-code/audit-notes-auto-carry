# coding: utf-8
"""工作进程结果边界：具体错误要传出，旧结果和异常退出不能冒充成功。"""
import json
from pathlib import Path

import pytest

from 核心 import Excel交付 as 模块


@pytest.mark.parametrize("情形,退出码", [
    ("本次错误", 7), ("旧成功", 0), ("无标识", 0), ("损坏", 0),
    ("缺文件", 0), ("异常退出却写成功", 7), ("本次成功", 0),
])
def test_只采纳本次有效结果(tmp_path, monkeypatch, 情形, 退出码):
    # 仅控制进程边界，不用此用例证明真实 Office 成功；正常链另有真实文件验收。
    def 启动(命令, **参数):
        请求 = Path(命令[-1])
        内容 = {"请求": 请求.name, "成功": True, "结果": {"值": 42}}
        if 情形 == "本次错误":
            内容.update({"成功": False, "错误": "资金!B2不能唯一归组"})
        elif 情形 == "旧成功":
            内容["请求"] = "Excel请求_上一次.json"
        elif 情形 == "无标识":
            内容.pop("请求")
        if 情形 != "缺文件":
            请求.with_name("Excel结果.json").write_text(
                "{损坏" if 情形 == "损坏" else json.dumps(内容), encoding="utf-8")
        参数["stdout"].write("本次进程日志")
        参数["stdout"].flush()

        class 已退出:
            returncode = 退出码

            def poll(self):
                return self.returncode

        return 已退出()

    monkeypatch.setattr(模块.subprocess, "Popen", 启动)
    if 情形 == "本次成功":
        assert 模块.交付Excel("不启动Office", tmp_path / "结果.xlsx", {}, {}, [], [], {}) == {"值": 42}
    else:
        with pytest.raises(RuntimeError) as 错误:
            模块.交付Excel("不启动Office", tmp_path / "结果.xlsx", {}, {}, [], [], {})
        if 情形 == "本次错误":
            assert "资金!B2不能唯一归组" in str(错误.value)
        else:
            assert "本次进程日志" in str(错误.value)
            assert str(退出码) in str(错误.value)


@pytest.mark.parametrize("状态进程号", [123, 999])
def test_停止不使用上一次Excel实例记录(tmp_path, monkeypatch, 状态进程号):
    import ctypes

    状态路径 = tmp_path / "Excel实例.json"
    状态路径.write_text(json.dumps({"python": 状态进程号, "excel_hwnd": 456}), encoding="utf-8")
    结束 = []

    class 本次进程:
        pid = 123

        def terminate(self):
            结束.append(self.pid)

        def wait(self, timeout):
            return 0

    def 查进程(hwnd, 指针):
        指针._obj.value = 789

    monkeypatch.setattr(ctypes.windll.user32, "GetWindowThreadProcessId", 查进程)
    monkeypatch.setattr(模块.subprocess, "run", lambda 命令, **参数: 结束.append(命令[2]))
    模块._终止(本次进程(), 状态路径)
    assert 结束 == ([123, "789"] if 状态进程号 == 123 else [123])
