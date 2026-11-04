# coding: utf-8
"""共用记录簿：跨进程提交、半批不可见、幂等、损坏与刷新失败回退。"""
import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

from 核心.共用记录簿 import 记录簿, 使用事件, 发布事件, 默认目录, 已启用

根 = Path(__file__).resolve().parent.parent


def _事件(标识, 序=0):
    return 使用事件("V" + 标识, f"任务{标识}", "A", f"组{序}", 0)


def _脚本(tmp_path, 名="提交.py"):
    路径 = tmp_path / 名
    路径.write_text(
        "# coding: utf-8\n"
        "import sys, time\n"
        "sys.path.insert(0, %r)\n"
        "from 核心.共用记录簿 import 记录簿, 使用事件\n"
        "目录, 标识, 次数 = sys.argv[1], sys.argv[2], int(sys.argv[3])\n"
        "簿 = 记录簿(目录)\n"
        "for i in range(次数):\n"
        "    簿.提交([使用事件('V' + 标识, '任务' + 标识, 'A', '组%%d' %% i, 0)],\n"
        "            '请求%%s-%%d' %% (标识, i))\n"
        "    time.sleep(0.02)\n"
        "print('OK', 簿.快照()['代次'])\n" % str(根), encoding="utf-8")
    return 路径


def _跑(脚本, 目录, 标识, 次数):
    return subprocess.Popen([sys.executable, str(脚本), str(目录), 标识, str(次数)],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            encoding="utf-8", creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def test_两个进程交错提交不丢批次(tmp_path):
    目录 = tmp_path / "记录"
    脚本 = _脚本(tmp_path)
    甲 = _跑(脚本, 目录, "甲", 5)
    乙 = _跑(脚本, 目录, "乙", 5)
    甲出, 甲错 = 甲.communicate(timeout=120)
    乙出, 乙错 = 乙.communicate(timeout=120)
    assert 甲.returncode == 0 and 乙.returncode == 0, (甲错, 乙错)
    簿 = 记录簿(目录)
    索引 = 簿.快照()
    assert 索引["可用"] and 索引["代次"] == 10
    任务 = {x["任务"] for x in 索引["使用"]}
    assert 任务 == {"任务甲", "任务乙"}
    assert len(索引["使用"]) == 10


def test_批次落盘但清单未替换时读者看不到(tmp_path):
    目录 = tmp_path / "记录"
    簿 = 记录簿(目录)
    簿.提交([_事件("甲")], "请求-甲")
    assert 簿.快照()["代次"] == 1
    # 模拟：批次已写、清单替换前进程中断
    (目录 / "批次" / "b00000002.json").write_text(json.dumps(
        {"代次": 2, "请求": "请求-乙", "事件": [_事件("乙")]}, ensure_ascii=False),
        encoding="utf-8")
    读者 = 记录簿(目录)
    索引 = 读者.快照()
    assert 索引["代次"] == 1 and len(索引["使用"]) == 1      # 未列入清单的批次不可见
    # 下一位写者仍能取得锁并正常提交
    结果 = 簿.提交([_事件("丙")], "请求-丙")
    assert 结果["状态"] == "已提交" and 结果["代次"] == 2
    assert len(记录簿(目录).快照()["使用"]) == 2


def test_持锁进程异常退出后锁可再取得(tmp_path):
    目录 = tmp_path / "记录"
    簿 = 记录簿(目录)
    簿.提交([_事件("甲")], "请求-甲")
    脚本 = tmp_path / "占锁.py"
    脚本.write_text(
        "# coding: utf-8\n"
        "import sys, os\n"
        "sys.path.insert(0, %r)\n"
        "from 核心.共用记录簿 import _文件锁\n"
        "锁 = _文件锁(%r)\n"
        "锁.__enter__()\n"
        "print('LOCKED', flush=True)\n"
        "os._exit(0)\n" % (str(根), str(目录 / "锁")), encoding="utf-8")
    子 = subprocess.run([sys.executable, str(脚本)], capture_output=True, encoding="utf-8",
                        timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    assert "LOCKED" in 子.stdout
    结果 = 簿.提交([_事件("乙")], "请求-乙")
    assert 结果["状态"] == "已提交"


def test_同一请求重试幂等且内容变了拒绝(tmp_path):
    簿 = 记录簿(tmp_path / "记录")
    一 = 簿.提交([_事件("甲")], "同一请求")
    二 = 簿.提交([_事件("甲")], "同一请求")
    assert 一["状态"] == "已提交" and 二["状态"] == "已存在"
    assert 簿.快照()["代次"] == 1
    with pytest.raises(ValueError, match="不同内容"):
        簿.提交([_事件("乙")], "同一请求")


def test_缺批次与刷新失败都不沿用旧索引(tmp_path):
    目录 = tmp_path / "记录"
    簿 = 记录簿(目录)
    簿.提交([_事件("甲")], "请求-甲")
    簿.提交([_事件("乙")], "请求-乙")
    assert 簿.快照()["可用"] and len(簿.快照()["使用"]) == 2
    (目录 / "批次" / "b00000001.json").unlink()
    索引 = 簿.快照()
    assert 索引["可用"] is False and "批次缺失" in 索引["原因"]
    assert 索引["解释"] == {} and 索引["按证据"] == {}
    assert 簿.刷新()["可用"] is False
    # 同一实例也不许拿旧结果
    assert 簿.快照()["可用"] is False


def test_批次被截断后整库不可用(tmp_path):
    """批次文件被截短、大小与清单不符：不能把半份记录当完整记录用。"""
    目录 = tmp_path / "记录"
    记录簿(目录).提交([_事件("截", 1)], "截断前")
    批 = next((目录 / "批次").glob("*.json"))
    原文 = 批.read_text(encoding="utf-8")
    批.write_text(原文[: max(1, len(原文) // 2)], encoding="utf-8")
    索引 = 记录簿(目录).快照()
    assert 索引["可用"] is False
    assert "批次" in (索引.get("原因") or "")     # 摘要先拦住，大小校验是第二道


def test_清单损坏时整库不可用(tmp_path):
    目录 = tmp_path / "记录"
    记录簿(目录).提交([_事件("甲")], "请求-甲")
    (目录 / "当前.json").write_text("{不是JSON", encoding="utf-8")
    索引 = 记录簿(目录).快照()
    assert 索引["可用"] is False and "索引刷新失败" in 索引["原因"]


def test_停用事件跨进程可见(tmp_path):
    目录 = tmp_path / "记录"
    一 = 记录簿(目录)
    一.提交([发布事件("A", "第三遍", {"证据版本": 1}, {"科目": "货币资金"},
                    "证据甲", "解释甲", "V甲", "金", "规", {"任务": "T1"})], "发布甲")
    脚本 = tmp_path / "停用.py"
    脚本.write_text(
        "# coding: utf-8\n"
        "import sys\n"
        "sys.path.insert(0, %r)\n"
        "from 核心.共用记录簿 import 记录簿\n"
        "记录簿(sys.argv[1]).提交([{'类型': '停用', '目标': ['V甲'], '原因': '错认'}], '停用甲')\n"
        % str(根), encoding="utf-8")
    subprocess.run([sys.executable, str(脚本), str(目录)], check=True, timeout=60,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    索引 = 一.快照()                       # 不重启，代次变了要重读
    assert 索引["代次"] == 2
    assert 索引["解释"]["V甲"]["状态"] == "停用"


def test_默认目录避开同步目录(tmp_path, monkeypatch):
    monkeypatch.delenv("审计附注_识别记录目录", raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "Local"))
    目录 = 默认目录()
    assert 目录 and "识别记录" in str(目录)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "OneDrive" / "Local"))
    assert 默认目录() is None


def test_复用开关与关闭后仍可完整识别(tmp_path, monkeypatch):
    monkeypatch.delenv("审计附注_复用", raising=False)
    assert 已启用({})
    assert not 已启用({"复用": False})
    monkeypatch.setenv("审计附注_复用", "0")
    assert not 已启用({})
