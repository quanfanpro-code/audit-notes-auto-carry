# coding: utf-8
"""第五阶段第四轮复核：合并表头例外不能替整列作证。

反例：B 的 B1:D1 是真实合并表头，识别漏交 D 列，却把这条例外写进了答复。
上一版把“例外涉及的列”整列当成已交代，于是 D2/D3 被放过，B 判完整，A 第三项被清空。
"""
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from openpyxl import Workbook

from 核心.搬运 import 生成更新表
from 核心.比对 import 建立对应
from 核心.接续 import 读索引
from 核心.识别 import 识别一侧, 铺格子, _解析精读


编号 = "C-N038-T001-S001"


def _金标准(tmp_path):
    路径 = tmp_path / "金标准.jsonl"
    路径.write_text(json.dumps({
        "note": {"name": "固定资产", "display_order": 1},
        "table": {"name": "固定资产情况", "display_order": 1},
        "slot": {"name": "账面原值 / 期末余额"}, "id": 编号,
        "row_path": ["账面原值"], "column_path": ["<资产类别>"],
        "dimensions": [{"name": "period", "required": True}]}, ensure_ascii=False) + "\n",
                  encoding="utf-8")
    return 路径


def _合并表头样本(tmp_path):
    """B 第一层 B1:D1 合并写“资产类别”，第二行名称，第三行金额。"""
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    ws["A1"] = "项目"; ws["B1"] = "资产类别"
    ws.merge_cells("B1:D1")
    ws["A2"] = ""
    ws["B2"], ws["C2"], ws["D2"] = "专用设备", "办公设备", "运输设备"
    ws["A3"] = "期末余额"
    ws["B3"], ws["C3"], ws["D3"] = 110, 220, 330
    p = tmp_path / "B.xlsx"; wb.save(p)
    return p


def _全列(带D=True):
    列们 = [{"列": "A", "路径": "项目", "角色": "行名列", "期间角色": None},
            {"列": "B", "路径": "专用设备", "角色": "数据列", "期间角色": "本期期末",
             "序号": 1, "标签行": 2},
            {"列": "C", "路径": "办公设备", "角色": "数据列", "期间角色": "本期期末",
             "序号": 2, "标签行": 2}]
    if 带D:
        列们.append({"列": "D", "路径": "运输设备", "角色": "数据列", "期间角色": "本期期末",
                   "序号": 3, "标签行": 2})
    return 列们


def _行们(带D=True):
    槽 = {"B": 编号, "C": 编号}
    if 带D:
        槽["D"] = 编号
    return [{"行": 1, "名字": "资产类别", "角色": "表头"},
            {"行": 2, "名字": "名称", "角色": "表头"},
            {"行": 3, "名字": "期末余额", "角色": "数据", "槽位": 槽}]


def _跑识别(tmp_path, monkeypatch, 列们, 行们, 例外):
    def 假调用(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,3],"科目":"固定资产"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,3],"科目":"固定资产","表名":"固定资产情况",'
                    '"表组":"G1","续表序":1}]}')
        return json.dumps({"列": 列们, "行": 行们, "作用范围": {}, "空名单": [],
                           "例外": 例外}, ensure_ascii=False)

    monkeypatch.setattr("核心.识别.调用识别", 假调用)
    运行 = tmp_path / "任务"
    识别一侧({"url": "u", "api_key": "k", "model": "m", "timeout": 1},
          str(_合并表头样本(tmp_path)), "B", 运行, _金标准(tmp_path))
    索引 = 读索引(运行)
    乙 = [json.loads(x) for x in (运行 / "B合集.jsonl").read_text(encoding="utf-8").splitlines()
          if x.strip()]
    乙结构 = [json.loads(x) for x in (运行 / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()
           if x.strip() and json.loads(x).get("侧") == "B"]
    问题 = [json.loads(x) for x in (运行 / "识别问题_B.jsonl").read_text(encoding="utf-8").splitlines()
          if x.strip()]
    return 索引["B完整性"], 乙, 乙结构, 问题


def _甲侧():
    """A：电子设备 / 办公家具 / 运输工具 三项，名称在列上。"""
    列们 = [{"列": "A", "路径": "项目", "角色": "行名列", "期间角色": None},
            {"列": "B", "路径": "电子设备", "角色": "数据列", "期间角色": "本期期末",
             "序号": 1, "标签行": 1},
            {"列": "C", "路径": "办公家具", "角色": "数据列", "期间角色": "本期期末",
             "序号": 2, "标签行": 1},
            {"列": "D", "路径": "运输工具", "角色": "数据列", "期间角色": "本期期末",
             "序号": 3, "标签行": 1}]
    行们 = [{"行": 1, "名字": "类别", "角色": "表头"},
            {"行": 2, "名字": "期末余额", "角色": "数据",
             "槽位": {"B": 编号, "C": 编号, "D": 编号}}]
    结构 = {"工作表": "s1", "科目": "固定资产", "表名": "固定资产情况", "表组": "G1",
          **_解析精读(json.dumps({"列": 列们, "行": 行们, "作用范围": {}, "空名单": [],
                                "例外": []}, ensure_ascii=False), {编号},
                     行范围=[1, 2], 列范围={"A", "B", "C", "D"})}
    格们 = 铺格子(结构)
    for 条 in 格们:
        条["期间必填"] = 条["角色"] != "项标签"
    return 结构, 格们


合并例外 = [{"范围": "B1:D1", "处理": "合并", "原因": "真实合并表头"}]


def test_第1项_合并表头例外不能替漏掉的列作证(tmp_path, monkeypatch):
    完, _, _, 问题 = _跑识别(tmp_path, monkeypatch, _全列(带D=False), _行们(带D=False), 合并例外)
    assert 完["第三遍逐行"] is False                 # B1:D1 只是表头，不能证明 D2、D3 已交代
    assert 完["漏列"]["固定资产／固定资产情况"] == ["D"]
    assert any("漏交这些列" in (x.get("原因") or "") for x in 问题)


def test_第1项_合并表头漏列时A第三项保留(tmp_path, monkeypatch):
    完, 乙, 乙结构, _ = _跑识别(tmp_path, monkeypatch, _全列(带D=False), _行们(带D=False), 合并例外)
    结构, 甲 = _甲侧()
    清 = 建立对应(甲, 乙, [结构], 乙结构)
    查 = {x["a格"]: (x["状态"], x["原因"]) for x in 清 if x["a格"]}
    assert 查["s1!D2"] == ("未对应", "仅A侧有")     # 第三项保留，不许当“B 无此项”
    变更, _ = 生成更新表({}, 清, 乙, 甲)
    assert "s1!D1" not in 变更 and "s1!D2" not in 变更


def test_第1项_列交齐时合并表头照常更新(tmp_path, monkeypatch):
    完, 乙, 乙结构, _ = _跑识别(tmp_path, monkeypatch, _全列(带D=True), _行们(带D=True), 合并例外)
    assert 完["第三遍逐行"] is True and 完["漏列"] == {}
    结构, 甲 = _甲侧()
    清 = 建立对应(甲, 乙, [结构], 乙结构)
    查 = {x["a格"]: x["状态"] for x in 清 if x["a格"]}
    assert 查["s1!D2"] == "对应"                    # 三列交齐才正常更新
