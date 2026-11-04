# coding: utf-8
"""仅用本库人工样本，验证迁入桥接的标题、结构保护与真实写回。"""
from copy import deepcopy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
import pytest

from 核心.桥接 import 提取, 更新, 读取

样本 = Path(__file__).resolve().parent / "样本/01_普通表格.docx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _树(path):
    with ZipFile(path) as z:
        return E.fromstring(z.read("word/document.xml"))


def _造样本(tmp_path, *, 双表=False, 纵向合并=False, 行首省略=False):
    树 = _树(样本)
    body = 树.find(W + "body")
    表 = body.find(W + "tbl")
    表.append(deepcopy(表.findall(W + "tr")[-1]))
    标题 = E.Element(W + "p")
    run = E.SubElement(标题, W + "r")
    prop = E.SubElement(run, W + "rPr")
    E.SubElement(prop, W + "b")
    E.SubElement(run, W + "t").text = "2024年期末"
    body.insert(body.index(表), 标题)
    if 双表:
        i = body.index(表) + 1
        body.insert(i, deepcopy(标题))
        body.insert(i + 1, deepcopy(表))
    if 纵向合并:
        for i, 行 in enumerate(表.findall(W + "tr")[:2]):
            格 = 行.findall(W + "tc")[0]
            pr = 格.find(W + "tcPr")
            if pr is None:
                pr = E.Element(W + "tcPr")
                格.insert(0, pr)
            E.SubElement(pr, W + "vMerge").set(W + "val", "restart" if i == 0 else "continue")
            if i:
                for t in 格.iter(W + "t"):
                    t.text = ""
    if 行首省略:
        行 = 表.findall(W + "tr")[1]
        行.remove(行.findall(W + "tc")[0])
        pr = E.Element(W + "trPr")
        E.SubElement(pr, W + "gridBefore").set(W + "val", "1")
        行.insert(0, pr)
    word = tmp_path / "人工样本.docx"
    with ZipFile(样本) as zin, ZipFile(word, "w") as zout:
        for part in zin.infolist():
            内容 = E.tostring(树, encoding="UTF-8", xml_declaration=True, standalone=True) if part.filename == "word/document.xml" else zin.read(part.filename)
            zout.writestr(part, 内容)
    return word


def _提取(word, folder):
    folder.mkdir()
    excel, linked, context = folder / "表.xlsx", folder / "关联.docx", folder / "上下文.json"
    提取(word, excel, linked, context, 工作目录=folder / "请求")
    return excel, linked, json.loads(context.read_text(encoding="utf-8-sig"))


def _指纹(*paths):
    return [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths]


def test_真实标题仅更新所属表且保留粗体(tmp_path):
    word = _造样本(tmp_path, 双表=True)
    excel, linked, context = _提取(word, tmp_path / "提取")
    原 = _指纹(word, excel, linked)
    output = tmp_path / "新副本.docx"
    更新(linked, excel, output, caption_updates=[{"range_name": "TA_001", "old_text": "2024年期末", "text": "2025年期末"}], 工作目录=tmp_path / "更新")
    body = _树(output).find(W + "body")
    标题 = [p for p in body.findall(W + "p") if "".join(p.itertext()).startswith(("2024年期末", "2025年期末"))]
    assert ["".join(p.itertext()) for p in 标题] == ["2025年期末", "2024年期末"]
    assert 标题[0].find(".//" + W + "b") is not None
    assert 原 == _指纹(word, excel, linked)


@pytest.mark.parametrize("动作", ["insert", "delete"])
def test_真实增删行与原件保护(tmp_path, 动作):
    word = _造样本(tmp_path)
    excel, linked, ctx = _提取(word, tmp_path / "提取")
    原 = _指纹(word, excel, linked)
    wb = load_workbook(excel)
    sheet = wb[ctx["tables"][0]["sheet"]]
    name = wb.defined_names["TA_001"]
    名, 范围 = list(name.destinations)[0]
    左, 上, 右, 下 = range_boundaries(范围)
    if 动作 == "insert":
        sheet.insert_rows(上 + 1)
        sheet.cell(上 + 1, 左, "新增项目")
        sheet.cell(上 + 1, 左 + 1, 123.45)
        下 += 1
    else:
        sheet.delete_rows(上 + 1)
        下 -= 1
    from openpyxl.utils import get_column_letter
    name.attr_text = f"'{名}'!${get_column_letter(左)}${上}:${get_column_letter(右)}${下}"
    changed = tmp_path / "改后.xlsx"
    wb.save(changed)
    op = {"range_name": "TA_001", "axis": "row", "action": 动作, "index": 1, "count": 1}
    if 动作 == "insert":
        op["template_index"] = 1
    output = tmp_path / "新副本.docx"
    更新(linked, changed, output, operations=[op], 工作目录=tmp_path / "更新")
    actual = _树(output).find(".//" + W + "tbl")
    assert len(actual.findall(W + "tr")) == (4 if 动作 == "insert" else 2)
    if 动作 == "insert":
        assert "新增项目" in "".join(actual.itertext())
        assert "123.45" in "".join(actual.itertext())
    assert 原 == _指纹(word, excel, linked)


def test_真实标题不匹配拒绝发布(tmp_path):
    word = _造样本(tmp_path)
    excel, linked, _ = _提取(word, tmp_path / "提取")
    output = tmp_path / "不应生成.docx"
    with pytest.raises(ValueError, match="标题"):
        更新(linked, excel, output, caption_updates=[{"range_name": "TA_001", "old_text": "不存在的标题", "text": "2025"}], 工作目录=tmp_path / "请求")
    assert not output.exists()


def test_真实纵向合并不被增行切开(tmp_path):
    word = _造样本(tmp_path, 纵向合并=True)
    excel, linked, _ = _提取(word, tmp_path / "提取")
    output = tmp_path / "不应生成.docx"
    with pytest.raises(ValueError, match="合并"):
        更新(linked, excel, output, operations=[{"range_name": "TA_001", "axis": "row", "action": "insert", "index": 1, "count": 1, "template_index": 2}], 工作目录=tmp_path / "请求")
    assert not output.exists()


@pytest.mark.parametrize("违规填值", [False, True])
def test_真实省略格不被补造或写入(tmp_path, 违规填值):
    word = _造样本(tmp_path, 行首省略=True)
    excel, linked, ctx = _提取(word, tmp_path / "提取")
    原 = _指纹(word, excel, linked)
    assert ctx["tables"][0]["structural_omissions"]
    output = tmp_path / "输出.docx"
    if 违规填值:
        wb = load_workbook(excel)
        wb["sheet1"]["B3"] = "原Word不存在的格"
        changed = tmp_path / "违规.xlsx"
        wb.save(changed)
        with pytest.raises(ValueError, match="省略"):
            更新(linked, changed, output, 工作目录=tmp_path / "请求")
        assert not output.exists()
    else:
        更新(linked, excel, output, 工作目录=tmp_path / "请求")
        行 = _树(output).find(".//" + W + "tbl").findall(W + "tr")[1]
        assert len(行.findall(W + "tc")) == 2
        assert 行.find(W + "trPr/" + W + "gridBefore").get(W + "val") == "1"
    assert 原 == _指纹(word, excel, linked)
