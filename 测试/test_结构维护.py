# coding: utf-8
"""结构操作只执行给定清单：程序不比较两份文件、不推断插行删行。"""
import json

import pytest

from 核心.结构维护 import 读结构操作


def test_照单执行(tmp_path):
    p = tmp_path / "结构操作.json"
    p.write_text(json.dumps([{"range_name": "TA_001", "axis": "row", "action": "insert",
                              "index": 3, "count": 1, "template_index": 2}],
                            ensure_ascii=False), encoding="utf-8")
    assert 读结构操作(p) == [{"range_name": "TA_001", "axis": "row", "action": "insert",
                              "index": 3, "count": 1, "template_index": 2}]


def test_也可以直接给数组():
    出 = 读结构操作([{"range_name": "TA_002", "axis": "row",
                  "action": "delete", "index": 1, "count": 2}])
    assert 出[0]["action"] == "delete" and 出[0]["count"] == 2


def test_不是数组就报错():
    with pytest.raises(RuntimeError):
        读结构操作({"range_name": "TA_001"})
