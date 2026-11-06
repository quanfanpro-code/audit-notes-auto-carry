# coding: utf-8
"""复用既有真实文件反例，并增加最终展示事实的独立核验。"""
import json
from pathlib import Path
import pytest
from openpyxl import load_workbook
from 测试 import test_复核五项修复 as 五项
from 测试 import test_真实桥接集成 as 桥接
from 测试 import test_Excel公式 as 公式


@pytest.mark.parametrize("用例", [
    五项.test_叹号表名_Excel接收金额名单合计正确,
    五项.test_来源零值空白和清尾仍按原规则,
    五项.test_文字名单保持原文与文本类型,
    五项.test_金额相同仍换年_复用,
    五项.test_续表标题Excel三份均更新两处标题,
    五项.test_续表标题失败整组保留独立组正常,
    桥接.test_三组标题与数值_重新打开核对,
    公式.test_公式合计差异保留待核副本,
])
def test_既有真实文件与交付事实一致(tmp_path, monkeypatch, 用例):
    用例(tmp_path, monkeypatch)
    已核 = 0
    for p in tmp_path.rglob("交付明细.jsonl"):
        记录 = [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines()]
        if not 记录:
            continue
        输出 = Path(记录[0]["输出"])
        实表 = 输出 if 输出.suffix == ".xlsx" else p.parent / "回读临时" / "回读表.xlsx"
        wb = load_workbook(实表, data_only=True)
        try:
            for 条 in 记录:
                if 条["类型"] == "标题":
                    for 动作 in 条["动作"]:
                        if 动作.get("格"):
                            表, 格 = 动作["格"].rsplit("!", 1)
                            assert wb[表][格].value == 动作["text"]
                    continue
                assert 条["实际值已取得"], 条
                表, 格 = 条["a格"].rsplit("!", 1)
                assert 条["实际值"] == wb[表][格].value
                if 条["最终动作"] == "已清空":
                    assert wb[表][格].value in (None, "")
                if 条["最终动作"] == "已写入" and 条["角色"] == "项标签":
                    assert 条["b格"] and 条["实际值"] == 条["来源值"]
                已核 += 1
        finally:
            wb.close()
    assert 已核 > 0
