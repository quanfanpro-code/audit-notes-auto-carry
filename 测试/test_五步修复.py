# coding: utf-8
"""五步修复的边界：不把失效号、未知位置或读取失败解释成可搬的空白。"""
import pytest
import json
from openpyxl import Workbook

from 核心.表格 import 读快照, file_hash, WorkbookError
from 核心.识别 import 补原值
from 测试.test_第五阶段独立验收 import 金原, _两侧识别


def test_无效编号不能被解释成不要求期间():
    from 核心.金标准 import 建索引, 期间必填
    索引 = 建索引([{"id": "文字", "note": {"name": "现金"},
                "table": {"name": "余额"}, "slot": {"name": "名称"}}])
    assert 期间必填(索引, "文字") is False
    with pytest.raises(ValueError, match="无效槽位编号"):
        期间必填(索引, "不存在")


def test_只补识别确认范围内的普通空格(tmp_path):
    wb = Workbook()
    s = wb.active
    s.title = "s"
    s["A1"] = "标题"
    s.merge_cells("A1:B1")
    s["A2"], s["C2"] = "名称", "=1+1"
    s["Z100"] = 9
    p = tmp_path / "源.xlsx"
    wb.save(p)
    wb.close()
    前 = file_hash(p)
    上下文 = [{"sheet": "s", "structural_omissions": [
        {"sheet": "s", "range": "B2", "reason": "Word原表没有这一格"}]}]
    快照 = 读快照(p, 上下文, 业务范围=[{"sheet": "s", "range": "A1:C3"}])
    出 = 补原值([{"格": "s!" + c} for c in ("A3", "B1", "B2", "C2", "D3")], 快照)
    assert [r["原值状态"] for r in 出] == ["真实空白", "读取失败", "读取失败", "公式无保存结果", "读取失败"]
    assert file_hash(p) == 前


def test_补空白拒绝超大范围(tmp_path):
    wb = Workbook()
    wb.active.title = "s"
    wb.active["A1"] = 1
    p = tmp_path / "源.xlsx"
    wb.save(p)
    wb.close()
    with pytest.raises(WorkbookError, match="50万格"):
        读快照(p, 业务范围=[{"sheet": "s", "range": "A1:XFD1048576"}])


@pytest.mark.parametrize("中项", [(None, None), (0, None), (None, 0), ("0", "0")])
def test_空白或文字零不能冒充两期数值全零(tmp_path, monkeypatch, 金原, 中项):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原, [(101, 111), 中项, (303, 313)])
    assert (值["B3"], 值["C3"]) == 中项
    assert 值["A3"] == "新2"
    assert (值["B4"], 值["C4"]) == (303, 313)


def test_公式无缓存保留原项并说明原因(tmp_path, monkeypatch, 金原):
    *_, 问题, 值 = _两侧识别(tmp_path, monkeypatch, 金原,
                          [(101, 111), ("=1+1", "=2+2"), (303, 313)])
    assert (值["B3"], 值["C3"]) == (200, 210)
    assert any("公式无保存结果" in x.get("原因", "") for x in 问题)
    assert 值["B4"] == 303


def test_一个期间为零不能证明两个期间全零():
    from 核心.比对 import 建立对应
    from 核心.搬运 import 生成更新表
    from 测试.test_比对 import _记, _构
    a = [_记("a!B2", "S1", 序号=1, 原值=100)]
    b = [_记("b!B2", "S1", 序号=1, 原值=0)]
    变更, _ = 生成更新表({}, 建立对应(a, b, _构(), _构()), b, a)
    assert 变更 == {"a!B2": 0}


@pytest.mark.parametrize("期间", [None, "未知", "本期期末"])
def test_期间缺字段与明确未知分别处理(期间):
    from 核心.识别 import _解析精读
    列 = {"列": "A", "角色": "数据列"}
    if 期间 is not None:
        列["期间角色"] = 期间
    答 = json.dumps({"列": [列], "行": [{"行": 1, "角色": "数据", "槽位": {"A": "S1"}}]})
    if 期间 is None:
        with pytest.raises(RuntimeError, match="期间角色缺失"):
            _解析精读(答, {"S1"}, 期间必填编号={"S1"})
    else:
        assert _解析精读(答, {"S1"}, 期间必填编号={"S1"})["列"][0]["期间角色"] == 期间


def test_空名单声明不能引用失效号():
    from 核心.识别 import _解析精读
    答 = json.dumps({"列": [{"列": "A", "角色": "数据列", "期间角色": "本期期末"}],
                    "行": [{"行": 1, "角色": "表头"}],
                    "空名单": [{"槽位": "失效号", "说明": "本期无行"}]})
    with pytest.raises(RuntimeError, match="空名单.*编号"):
        _解析精读(答, {"有效号"})


def test_旧空名单声明同样要查有效编号(tmp_path, monkeypatch, 金原):
    from 测试.test_修复后独立复验 import _旧号任务
    from 核心.识别 import 识别一侧
    _旧号任务(tmp_path, monkeypatch, 金原, True, 仅建旧结构=True)
    p = tmp_path / "任务" / "结构清单.jsonl"
    结构 = [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines()]
    for 表 in 结构:
        if 表["侧"] == "A":
            表["行"] = [{"行": 1, "角色": "表头", "槽位": {}}]
            表["空名单"] = [{"槽位": "C-N054-T001-S002", "说明": "本期无行"}]
    p.write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 结构), encoding="utf-8")
    前 = p.read_bytes()
    with pytest.raises(RuntimeError, match="失效.*新建任务"):
        识别一侧({"测试": True}, tmp_path / "A.xlsx", "A", tmp_path / "任务", tmp_path / "金标准.jsonl")
    assert p.read_bytes() == 前
