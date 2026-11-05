# coding: utf-8
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook

from 测试.test_二次修复 import _B单位, _带说明Word, _有依据模型
from 测试.test_复核反例 import _单组金标准, _跑一份, 读回Word
from 测试.安全目录 import 安全记录目录


@pytest.mark.parametrize("标注,换序", [("正常", True), ("漏标", True), ("原文失实", True), ("坐标失实", True), ("漏标", False), ("正常", False)])
def test_同格角色日期对调不得错搬(tmp_path, monkeypatch, 标注, 换序):
    monkeypatch.setenv("审计附注_复用", "1")
    说明 = "金额单位：人民币元"
    a = _带说明Word(tmp_path, "原件", 说明)
    金 = _单组金标准(tmp_path)
    库 = 安全记录目录("独立角色补查")

    def 模型(年, 换序):
        基 = _有依据模型("元", "元", 说明)

        def 调用(配置, 指令, 材料, 目录):
            答 = json.loads(基(配置, 指令, 材料, 目录))
            if "列" in 答 and Path(目录).name.startswith("B_"):
                for i, c in enumerate(答["列"][1:3]):
                    是期初 = (i == 0) != 换序
                    角色 = "期初" if 是期初 else "期末"
                    c["路径"] = f"{角色}{年}-12-31"
                    c["期间角色"] = "本期期初" if 是期初 else "本期期末"
                    c["期间实际"] = {"值": f"{年}-12-31", "位置": c["列"] + "1"}
                    if 标注 != "漏标":
                        c["角色依据"] = {"位置": "Z99" if 标注 == "坐标失实" else c["列"] + "1", "原文": "不存在的原文" if 标注 == "原文失实" else c["路径"]}
            return json.dumps(答, ensure_ascii=False)

        return 调用

    for 次, 年 in enumerate((2025, 2026)):
        本次换序 = 换序 and 次 == 1
        初, 末 = (901, 102) if 本次换序 else (102, 901)
        b = _B单位(tmp_path, f"{次}.xlsx", "元", 初, 末)
        wb = load_workbook(b)
        wb.active["C1"] = f"{'期末' if 本次换序 else '期初'}{年}-12-31"
        wb.active["D1"] = f"{'期初' if 本次换序 else '期末'}{年}-12-31"
        wb.save(b)
        wb.close()
        monkeypatch.setattr("核心.识别.调用识别", 模型(年, 本次换序))
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)

    assert 读回Word(出["Word"]).表格()[1] == ["甲公司", "102", "901"]
    assert 读回Word(出["Word"]).表格()[2] == ["乙公司", "103", "902"]
    for 侧 in ("A", "B"):
        日志 = json.loads((出["任务"] / f"复用记录_{侧}.json").read_text(encoding="utf-8"))
        assert [日志["调用"][步] for 步 in ("第一遍", "第二遍", "第三遍")] == [0, 0, 1 if 侧 == "B" and 换序 else 0]
