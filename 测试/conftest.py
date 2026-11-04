# coding: utf-8
"""测试环境的记录簿隔离：绝不写本机真实识别记录目录。

默认关闭跨任务复用，老用例行为与接入前一致；复用用例自行打开开关并指定临时目录。
"""
import os

import pytest


@pytest.fixture(scope="session", autouse=True)
def _隔离识别记录(tmp_path_factory):
    目录 = tmp_path_factory.mktemp("识别记录")
    os.environ["审计附注_识别记录目录"] = str(目录)
    os.environ["审计附注_复用"] = "0"
    yield 目录
    os.environ.pop("审计附注_识别记录目录", None)
    os.environ.pop("审计附注_复用", None)
