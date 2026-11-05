# coding: utf-8
"""真实桥接集成用例（-m 集成）：用 测试\样本 真实 docx＋真桥接，断言实际结果。
覆盖：①提取命名区域/上下文 ②普通值回写 ③增行 ④删行 ⑧命名区域缺失报错 ⑨原件哈希不变。
⑤列调整/⑥标题更新/⑦跨合并拒绝 需真实业务文件，列入天宇油脂验收演练（见计划 T15）。"""
import hashlib
from 测试.单位样本 import 单位作用
import json
import shutil
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from 核心.桥接 import 提取, 更新, 桥接程序

样本 = Path(__file__).resolve().parent / "样本" / "01_普通表格.docx"
pytestmark = [pytest.mark.集成, pytest.mark.skipif(
    not 桥接程序.is_file(), reason="缺少WordBridge发布包，真实Word集成验证未执行")]


def md5(p):
    return hashlib.md5(Path(p).read_bytes()).hexdigest()




def 提取到(tmp_path):
    d = tmp_path / "任务"
    d.mkdir()
    提取(str(样本), str(d / "表A.xlsx"), str(d / "带连接.docx"), str(d / "上下文.json"))
    return d




def 回读(d, 名):
    提取(str(d / 名), str(d / "回读.xlsx"), str(d / "回读连.docx"), str(d / "回读上下文.json"))
    return load_workbook(d / "回读.xlsx")




def test_提取命名区域与上下文(tmp_path):
    d = 提取到(tmp_path)
    wb = load_workbook(d / "表A.xlsx")
    assert "TA_001" in wb.defined_names
    ctx = json.loads((d / "上下文.json").read_text(encoding="utf-8-sig"))
    assert ctx["tables"][0]["range_name"] == "TA_001"
    assert "chapter_context" in ctx["tables"][0]




def test_普通值回写并回读核对(tmp_path):
    前 = md5(样本)
    d = 提取到(tmp_path)
    表A = load_workbook(d / "表A.xlsx")
    表A["sheet1"]["C2"] = "新值"
    表A.save(d / "改后.xlsx")
    更新(str(d / "带连接.docx"), str(d / "改后.xlsx"), str(d / "新副本.docx"))
    结果表 = 回读(d, "新副本.docx")
    assert 结果表["sheet1"]["C2"].value == "新值"       # 实际写入
    assert 结果表["sheet1"]["B2"].value != "新值"       # 旁格不动（C2 之外）
    assert md5(样本) == 前                              # 原件哈希不变




def test_命名区域缺失报错(tmp_path):
    d = 提取到(tmp_path)
    wb = Workbook(); wb.active.title = "sheet1"; wb.active["B2"] = "x"
    wb.save(d / "无命名.xlsx")
    with pytest.raises(ValueError) as e:
        更新(str(d / "带连接.docx"), str(d / "无命名.xlsx"), str(d / "新副本.docx"))
    assert "命名区域" in str(e.value)




# —— 2026-11-02 追加：⑩ 多区域下删中间行、⑪ 跨区域明确报错 ——

# —— 2026-11-04 追加：跨任务复用交付前必验（真实桥接，三组标题与数值） ——
import re  # noqa: E402
from zipfile import ZipFile  # noqa: E402
from lxml import etree as E  # noqa: E402
from 核心.流程 import 运行任务  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _三组样本(tmp_path):
    """三张表、三句日期标题：第一组正常换年，第二组日期冲突，第三组不动标题。"""
    with ZipFile(样本) as z:
        树 = E.fromstring(z.read("word/document.xml"))
    body = 树.find(W + "body")
    for 子 in list(body):
        body.remove(子)

    def 段落(文本):
        p = E.Element(W + "p")
        r = E.SubElement(p, W + "r")
        E.SubElement(r, W + "t").text = 文本
        return p

    def 表(行们):
        tbl = E.Element(W + "tbl")
        pr = E.SubElement(tbl, W + "tblPr")
        w = E.SubElement(pr, W + "tblW")
        w.set(W + "w", "0")
        w.set(W + "type", "auto")
        grid = E.SubElement(tbl, W + "tblGrid")
        for _ in 行们[0]:
            E.SubElement(grid, W + "gridCol").set(W + "w", "2400")
        for 行 in 行们:
            tr = E.SubElement(tbl, W + "tr")
            for 值 in 行:
                tc = E.SubElement(tr, W + "tc")
                tcPr = E.SubElement(tc, W + "tcPr")
                E.SubElement(tcPr, W + "tcW").set(W + "w", "2400")
                p = E.SubElement(tc, W + "p")
                r = E.SubElement(p, W + "r")
                E.SubElement(r, W + "t").text = str(值)
        return tbl

    for 名, 期初, 期末 in (("第一组", 100, 110), ("第二组", 200, 210), ("第三组", 300, 310)):
        body.append(段落(f"截至2024年12月31日（{名}）"))
        body.append(表([["项目（单位：人民币元）", "期初余额", "期末余额"],
                       ["甲公司", 期初, 期末],
                       ["乙公司", 期初 + 1, 期末 + 1]]))
    word = tmp_path / "三组样本.docx"
    with ZipFile(样本) as zin, ZipFile(word, "w") as zout:
        for part in zin.infolist():
            内容 = (E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True)
                  if part.filename == "word/document.xml" else zin.read(part.filename))
            zout.writestr(part, 内容)
    return word


def _三组金标准(tmp_path):
    行们 = [{"id": "N001", "note": {"name": "货币资金"}, "table": {"name": "余额构成"},
            "slot": {"name": "余额"}, "row_path": ["甲公司"], "column_path": ["期末余额"],
            "dimensions": [{"name": "period", "required": True}]},
           {"id": "N002", "note": {"name": "应收账款"}, "table": {"name": "按账龄披露"},
            "slot": {"name": "余额（左）"}, "row_path": ["甲公司"], "column_path": ["期末余额"],
            "dimensions": [{"name": "period", "required": True}]},
           {"id": "N002B", "note": {"name": "应收账款"}, "table": {"name": "按账龄披露"},
            "slot": {"name": "余额（右）"}, "row_path": ["甲公司"], "column_path": ["期末余额"],
            "dimensions": [{"name": "period", "required": True}]},
           {"id": "N003", "note": {"name": "存货"}, "table": {"name": "存货明细"},
            "slot": {"name": "余额"}, "row_path": ["甲公司"], "column_path": ["期末余额"],
            "dimensions": [{"name": "period", "required": True}]}]
    p = tmp_path / "金标准三组.jsonl"
    p.write_text("\n".join(json.dumps(x, ensure_ascii=False) for x in 行们), encoding="utf-8")
    return p


def _三组模型():
    """固定答复：A 只交标题定位（原文＋日期片段＋期间角色），B 给本次日期。

    A 的表格由桥接提取在 B:D 列、行间夹着标题段落；B 是自己造的三列连续区块。
    """
    布局 = {"A": {"顶": [2, 8, 14], "列": ("B", "C", "D"), "段": [[1, 6], [7, 12], [13, 16]]},
          "B": {"顶": [1, 4, 7], "列": ("A", "B", "C"), "段": [[1, 3], [4, 6], [7, 9]]}}

    def 调用(配置, 指令, 材料, 目录):
        侧 = Path(目录).name[:1]
        位 = 布局[侧]
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": 段, "科目": 科}
                                    for 段, 科 in zip(位["段"], ("货币资金", "应收账款", "存货"))]},
                           ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            表们 = [{"行范围": [顶, 顶 + 2], "科目": 科, "表名": 表,
                    "表组": f"G{i}", "续表序": 1}
                   for i, (顶, 科, 表) in enumerate(
                       zip(位["顶"], ("货币资金", "应收账款", "存货"),
                           ("余额构成", "按账龄披露", "存货明细")), 1)]
            if 侧 == "A":
                for 表 in 表们[:2]:
                    表["标题"] = {"原文": f"截至2024年12月31日（{'第一组' if 表['表组'] == 'G1' else '第二组'}）",
                               "日期": [{"片段": "2024年12月31日", "期间角色": "本期期末"}]}
            return json.dumps({"表": 表们}, ensure_ascii=False)
        科目 = re.search(r"科目：(.+)", 材料).group(1).strip()
        表名 = re.search(r"表名：(.+)", 材料).group(1).strip()
        序 = {"货币资金": 1, "应收账款": 2, "存货": 3}[科目]
        编号 = f"N00{序}"
        顶 = 位["顶"][序 - 1]
        行列, 期初列, 期末列 = 位["列"]
        行们 = [{"行": 顶, "角色": "表头"},
               {"行": 顶 + 1, "角色": "数据", "序号": 1, "槽位": {期初列: 编号, 期末列: 编号}},
               {"行": 顶 + 2, "角色": "数据", "序号": 2, "槽位": {期初列: 编号, 期末列: 编号}}]
        列们 = [{"列": 行列, "路径": "项目", "角色": "行名列"},
               {"列": 期初列, "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
                "期间实际": {"值": "2025-01-01", "位置": f"{期初列}{顶}"}},
               {"列": 期末列, "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
                "期间实际": {"值": "2025-12-31", "位置": f"{期末列}{顶}"}}]
        if 序 == 2:
            # 两列都是本期期末、日期却不一致：标题无法唯一确定，整组应保持原样；
            # 两列映射到不同编号，配对不受影响，金额本来会变。
            列们[1] = {"列": 期初列, "路径": "期末余额（左）", "角色": "数据列",
                      "期间角色": "本期期末",
                      "期间实际": {"值": "2025-12-31", "位置": f"{期初列}{顶}"}}
            列们[2]["期间实际"] = {"值": "2025-06-30", "位置": f"{期末列}{顶}"}
            for 行 in 行们[1:]:
                行["槽位"] = {期初列: "N002", 期末列: "N002B"}
        return json.dumps({"列": 列们, "行": 行们,
                          "作用范围": 单位作用(f"{行列}{顶}", "项目（单位：人民币元）", [期初列, 期末列]),
                          "空名单": [], "例外": []}, ensure_ascii=False)
    return 调用


def _造B(tmp_path):
    """B：与 A 同布局、数值都不同，保证“本来应变的金额”。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "word附注"
    for i, (期初, 期末) in enumerate(((900, 910), (800, 810), (700, 710))):
        顶 = 3 * i + 1
        ws.cell(顶, 1, "项目（单位：人民币元）"); ws.cell(顶, 2, "期初余额"); ws.cell(顶, 3, "期末余额")
        ws.cell(顶 + 1, 1, "甲公司"); ws.cell(顶 + 1, 2, 期初); ws.cell(顶 + 1, 3, 期末)
        ws.cell(顶 + 2, 1, "乙公司"); ws.cell(顶 + 2, 2, 期初 + 1); ws.cell(顶 + 2, 3, 期末 + 1)
    p = tmp_path / "三组B.xlsx"
    wb.save(p)
    wb.close()
    return p


def test_三组标题与数值_重新打开核对(tmp_path, monkeypatch):
    word = _三组样本(tmp_path)
    B = _造B(tmp_path)
    金 = _三组金标准(tmp_path)
    原哈希 = word.read_bytes()
    monkeypatch.setattr("核心.识别.调用识别", _三组模型())
    任务 = tmp_path / "任务"
    状态 = 运行任务({"word": str(word), "b表": str(B), "金标准": str(金), "b工作表": "word附注"},
                任务, {"测试": True})
    # 第二组标题日期冲突整组保留，状态如实为“部分保留”（五项修复的状态语义收紧）。
    assert 状态["步骤"] == "部分保留", 状态.get("步骤")
    assert word.read_bytes() == 原哈希                     # 原 Word 未被改动

    输出 = Path(状态["产物"]["Word新副本"])
    assert 输出.is_file()
    d = tmp_path / "重开"
    d.mkdir()
    提取(str(输出), str(d / "表.xlsx"), str(d / "连接.docx"), str(d / "上下文.json"))

    with ZipFile(输出) as z:
        body = E.fromstring(z.read("word/document.xml")).find(W + "body")
    段落 = ["".join(p.itertext()) for p in body.findall(W + "p")]
    表们 = [[["".join(tc.itertext()) for tc in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]
          for tbl in body.findall(W + "tbl")]
    assert [x for x in 段落 if x.strip()] == [
        "截至2025年12月31日（第一组）",        # 用本次 B 日期更新
        "截至2024年12月31日（第二组）",        # 同角色日期冲突：标题保持原样
        "截至2024年12月31日（第三组）",        # 没有标题结论：保持原样
    ]
    assert 表们[0][1][2] == "910" and 表们[0][2][2] == "911"    # 第一组按 B 写入
    assert 表们[1][1][2] == "210" and 表们[1][2][2] == "211"    # 第二组整组保持原值
    assert 表们[2][1][2] == "710" and 表们[2][2][2] == "711"    # 第三组照常写入

    问题 = (任务 / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "应收账款" in 问题 and "标题" in 问题
    Excel = load_workbook(Path(状态["产物"]["更新后附注表格"]))
    try:
        # 提取后的表在 B:D 列，三组分别在第 2、8、14 行起
        assert Excel["sheet1"]["D3"].value == 910              # 第一组按 B 写入
        assert Excel["sheet1"]["D9"].value == 210              # 第二组整组保持原值
        assert Excel["sheet1"]["D15"].value == 710             # 第三组照常写入
    finally:
        Excel.close()
    assert (任务 / "回读记录.jsonl").read_text(encoding="utf-8").strip() == ""

