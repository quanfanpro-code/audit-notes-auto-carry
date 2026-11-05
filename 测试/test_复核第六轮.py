# coding: utf-8
"""第五阶段第六轮复核：合并格左上角存着名称，不能当成从属格免检。

反例：B 第三项名称放在 D1:D2 上下合并格，左上角 D1 就是“运输设备”。
识别漏交 D 列、只写了“D1:D2 合并”，上一版把整个合并范围都当作可从属格跳过，
于是 D1 也算“已交代”，B 判完整，A 第三项被当成“B 无此项”清空。
"""
from 测试.单位样本 import 单位作用
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from openpyxl import Workbook
from openpyxl.styles import Font

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


def _合并名称样本(tmp_path):
    """B：B1:C1 合并写“资产类别”；第三项名称在 D1:D2 上下合并，D3 空白但有格式。"""
    wb = Workbook(); ws = wb.active; ws.title = "s1"
    ws["A1"] = "项目（单位：人民币元）"; ws["B1"] = "资产类别"
    ws.merge_cells("B1:C1")
    ws["A2"] = "类别"
    ws["B2"], ws["C2"] = "专用设备", "办公设备"
    ws["D1"] = "运输设备"
    ws.merge_cells("D1:D2")                     # 名称在左上角 D1，D2 是从属格
    ws["A3"] = "期末余额"
    ws["B3"], ws["C3"] = 110, 220
    ws["D3"].font = Font(bold=True)             # 金额真实空白，只是有格式
    p = tmp_path / "B.xlsx"; wb.save(p)
    return p


def _列们(带D=True):
    列们 = [{"列": "A", "路径": "项目", "角色": "行名列", "期间角色": None},
            {"列": "B", "路径": "专用设备", "角色": "数据列", "期间角色": "本期期末",
             "序号": 1, "标签行": 2},
            {"列": "C", "路径": "办公设备", "角色": "数据列", "期间角色": "本期期末",
             "序号": 2, "标签行": 2}]
    if 带D:
        列们.append({"列": "D", "路径": "运输设备", "角色": "数据列", "期间角色": "本期期末",
                   "序号": 3, "标签行": 1})     # 名称写在第 1 行（合并左上角）
    return 列们


def _行们(带D=True):
    槽 = {"B": 编号, "C": 编号}
    if 带D:
        槽["D"] = 编号
    return [{"行": 1, "名字": "资产类别", "角色": "表头"},
            {"行": 2, "名字": "类别", "角色": "表头"},
            {"行": 3, "名字": "期末余额", "角色": "数据", "槽位": 槽}]


例外 = [{"范围": "B1:C1", "处理": "合并", "原因": "第一层表头横跨两列"},
      {"范围": "D1:D2", "处理": "合并", "原因": "名称上下合并"}]


def _跑识别(tmp_path, monkeypatch, 带D):
    def 假调用(配置, 指令, 材料, 调用目录):
        if 指令.startswith("把整份文件"):
            return '{"划分":[{"行范围":[1,3],"科目":"固定资产"}]}'
        if 指令.startswith("在第一遍"):
            return ('{"表":[{"行范围":[1,3],"科目":"固定资产","表名":"固定资产情况",'
                    '"表组":"G1","续表序":1}]}')
        return json.dumps({"列": _列们(带D), "行": _行们(带D), "作用范围": 单位作用("A1", "项目（单位：人民币元）", [c["列"] for c in _列们(带D) if c["角色"] == "数据列"]),
                           "空名单": [], "例外": 例外}, ensure_ascii=False)

    monkeypatch.setattr("核心.识别.调用识别", 假调用)
    运行 = tmp_path / "任务"
    识别一侧({"url": "u", "api_key": "k", "model": "m", "timeout": 1},
          str(_合并名称样本(tmp_path)), "B", 运行, _金标准(tmp_path))
    乙 = [json.loads(x) for x in (运行 / "B合集.jsonl").read_text(encoding="utf-8").splitlines()
          if x.strip()]
    乙结构 = [json.loads(x) for x in (运行 / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()
           if x.strip() and json.loads(x).get("侧") == "B"]
    问题 = [json.loads(x) for x in (运行 / "识别问题_B.jsonl").read_text(encoding="utf-8").splitlines()
          if x.strip()]
    return 读索引(运行)["B完整性"], 乙, 乙结构, 问题


def _甲侧():
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
          **_解析精读(json.dumps({"列": 列们, "行": 行们, "作用范围": 单位作用("A1", "项目（单位：人民币元）", [c["列"] for c in 列们 if c["角色"] == "数据列"]), "空名单": [],
                                "例外": []}, ensure_ascii=False), {编号},
                     行范围=[1, 2], 列范围={"A", "B", "C", "D"})}
    格们 = 铺格子(结构)
    for 条 in 格们:
        条["期间必填"] = 条["角色"] != "项标签"
        条["原值"] = {"s1!B1": "电子设备", "s1!C1": "办公家具", "s1!D1": "运输工具",
                    "s1!B2": 100, "s1!C2": 200, "s1!D2": 300}.get(条["格"])
        条["原值状态"] = "值" if 条["原值"] is not None else "真实空白"
    return 结构, 格们


def test_第6轮_合并左上角有值不能免检(tmp_path, monkeypatch):
    完, _, _, 问题 = _跑识别(tmp_path, monkeypatch, 带D=False)
    assert 完["第三遍逐行"] is False                 # D1 是名称，不是可从属格
    assert 完["漏列"]["固定资产／固定资产情况"] == ["D"]
    条 = [x for x in 问题 if "漏交这些列" in (x.get("原因") or "")][0]
    assert 条["行"]["D"] == [1]                      # 说清漏的是第 1 行那个左上角


def test_第6轮_合并左上角漏识别时A第三项保留(tmp_path, monkeypatch):
    _, 乙, 乙结构, _ = _跑识别(tmp_path, monkeypatch, 带D=False)
    结构, 甲 = _甲侧()
    清 = 建立对应(甲, 乙, [结构], 乙结构)
    查 = {x["a格"]: (x["状态"], x["原因"]) for x in 清 if x["a格"]}
    assert 查["s1!D2"] == ("未对应", "仅A侧有")
    变更, _ = 生成更新表({}, 清, 乙, 甲)
    assert "s1!D1" not in 变更 and "s1!D2" not in 变更


def test_第6轮_列交齐时合并名称照常更新(tmp_path, monkeypatch):
    完, 乙, 乙结构, _ = _跑识别(tmp_path, monkeypatch, 带D=True)
    assert 完["第三遍逐行"] is True and 完["漏列"] == {}
    结构, 甲 = _甲侧()
    清 = 建立对应(甲, 乙, [结构], 乙结构)
    变更, _ = 生成更新表({}, 清, 乙, 甲)
    assert 变更["s1!D1"] == "运输设备"               # 名称从合并左上角搬过来
    assert 变更["s1!D2"] == ""                       # B 金额真实空白，按政策清空
