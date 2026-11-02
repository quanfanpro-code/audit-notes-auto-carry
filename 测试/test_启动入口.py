# coding: utf-8
"""双击入口测试：真的去编译那份 .pyw，不测辅助函数。

上一轮的问题是入口第 54 行字符串断成两行，compileall 没检出来，146 条测试全绿，
双击却根本起不来。所以这里只认“真文件能通过语法编译”这一件事。
"""
import sys
import pathlib

根 = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(根))


def test_双击入口能通过语法编译():
    入口 = 根 / "启动服务.pyw"
    源码 = 入口.read_text(encoding="utf-8")
    compile(源码, str(入口), "exec")          # 语法错就在这里抛 SyntaxError


def test_入口失败时才记日志():
    源码 = (根 / "启动服务.pyw").read_text(encoding="utf-8")
    assert "启动失败：" in 源码
    assert 'traceback.format_exc()' in 源码     # 失败原因要落到日志，不许只弹一下
