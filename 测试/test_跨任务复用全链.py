# coding: utf-8
"""跨任务复用总体验收：同一语料，完整识别与启用复用逐项对照。

对照的是“本次结果”，不是模型答复：科目、表组、槽位与期间、本次坐标、配对、清尾、
产物与回读。另测缩小打印区域、隐藏列影响可见值、同结构冲突、金标准与规则版本变化。
"""
import json
import shutil
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from 核心.流程 import 运行任务
from 核心.共用记录簿 import 记录簿, 发布事件
from 核心.复用证据 import 证据键, 解释键, 版本号
from 测试.test_三遍复用 import _金标准, _造, _模型, 科目表
from 测试.安全目录 import 安全记录目录

语料 = [("一", 科目表, 2025), ("二", [科目表[0], 科目表[2]], 2026), ("三", 科目表, 2027)]


def _跑(tmp_path, 名, 科目A, 年, 记录目录=None, 金=None, monkeypatch=None, 打印区域=None,
       科目B=None, 值偏移=0, 隐藏行=None, 辅助=None):
    tmp_path.mkdir(parents=True, exist_ok=True)
    科目B = 科目B or 科目A
    A, 上下文, _ = _造(tmp_path, f"{名}A", 科目A, 年)
    _, _, B = _造(tmp_path, f"{名}B", 科目B, 年)
    if 打印区域 or 值偏移 or 隐藏行 or 辅助 is not None:
        底 = load_workbook(B)
        ws = 底["word附注"]
        if 打印区域:
            ws.print_area = 打印区域
        if 值偏移:
            for 行 in ws.iter_rows(min_row=1, max_row=9, min_col=1, max_col=3):
                for 格 in 行:
                    if isinstance(格.value, (int, float)) and not isinstance(格.value, bool):
                        格.value = 格.value + 值偏移
        if 隐藏行:
            ws.row_dimensions[隐藏行].hidden = True
        if 辅助 is not None:
            ws["Z40"] = 辅助["值"]          # 区域外的辅助公式所在格
            ws.row_dimensions[40].hidden = True
        底.save(B)
        底.close()

    def 假提取(word, excel, 带连接, 路径):
        shutil.copyfile(A, excel)
        Path(带连接).write_text("连接体", encoding="utf-8")
        Path(路径).write_text(json.dumps(上下文, ensure_ascii=False), encoding="utf-8")
        return {"ok": True, "warnings": []}

    def 假更新(word, excel, output, operations, caption_updates):
        Path(output).write_text(json.dumps({"excel": str(excel)}, ensure_ascii=False), encoding="utf-8")
        return {"ok": True}

    def 假回读(word, excel, linked, 路径):
        shutil.copyfile(json.loads(Path(word).read_text(encoding="utf-8"))["excel"], excel)

    monkeypatch.setattr("核心.提取.调用桥接提取", 假提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 假更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 假回读)
    模型, _ = _模型(年, {"A": 科目A, "B": 科目B})
    monkeypatch.setattr("核心.识别.调用识别", 模型)
    任务 = tmp_path / f"任务_{名}"
    入 = {"word": "原件.docx", "b表": str(B), "金标准": str(金), "b工作表": "word附注"}
    if 记录目录:
        入["识别记录目录"] = str(记录目录)
    状态 = 运行任务(入, 任务, {"测试": True})
    assert 状态["步骤"] == "done", 状态.get("步骤")
    return 任务


def _规范(任务):
    def 读(名):
        p = 任务 / 名
        return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] \
            if p.is_file() else []
    结构 = sorted(
        [(x["侧"], x["科目"], x["表名"], x.get("续表序"), tuple(x["行范围"]),
          tuple(sorted((c["列"], c.get("期间角色")) for c in x.get("列") or [])),
          tuple(sorted((r["行"], tuple(sorted((r.get("槽位") or {}).items())),
                        r.get("期间角色")) for r in x.get("行") or [])),
          x.get("完整性", {}).get("第三遍逐行"))
         for x in 读("结构清单.jsonl")], key=str)
    对应 = sorted([(x["a格"], x["b格"], x["状态"], x.get("原因"), x.get("槽位"), x.get("序号"))
                 for x in 读("对应清单.jsonl")], key=str)
    产物 = load_workbook(任务 / "更新后附注表格.xlsx")
    try:
        值 = {f"{s.title}!{c.coordinate}": c.value for s in 产物 for row in s for c in row}
    finally:
        产物.close()
    标题 = None
    if (任务 / "标题证据.json").is_file():
        标题 = json.loads((任务 / "标题证据.json").read_text(encoding="utf-8")) or None
        标题 = 标题[0] if isinstance(标题, list) and 标题 else 标题
    return {"结构": 结构, "对应": 对应, "值": 值, "标题": 标题,
            "回读": 读("回读记录.jsonl"), "问题": 读("问题清单.jsonl")}


def _基线(tmp_path, monkeypatch, 金, 语料, 打印=None):
    monkeypatch.setenv("审计附注_复用", "0")
    出 = []
    for 名, 科目, 年 in 语料:
        # 基线用独立任务目录；B 与 A 的文件名按“名”区分，打印区域只作用于该份
        出.append(_规范(_跑(tmp_path / f"基线{名}", 名, 科目, 年, None, 金, monkeypatch,
                          打印区域=(打印 or {}).get(名))))
    return 出


def _复用跑(tmp_path, monkeypatch, 金, 语料, 打印=None, 记录目录=None, 语料override=None):
    monkeypatch.delenv("审计附注_复用", raising=False)
    记录目录 = 记录目录 or (安全记录目录())
    出 = []
    for 名, 科目, 年 in (语料override or 语料):
        出.append(_规范(_跑(tmp_path / f"复用{名}", 名, 科目, 年, 记录目录, 金, monkeypatch,
                          打印区域=(打印 or {}).get(名))))
    return 出


def test_完整识别与启用复用结果逐项一致(tmp_path, monkeypatch):
    金 = _金标准(tmp_path)
    基线 = _基线(tmp_path, monkeypatch, 金, 语料)
    复用 = _复用跑(tmp_path, monkeypatch, 金, 语料)
    for 甲, 乙 in zip(基线, 复用):
        assert 甲["结构"] == 乙["结构"], "科目、表组、槽位或期间不一致"
        assert 甲["对应"] == 乙["对应"], "配对或清尾不一致"
        assert 甲["值"] == 乙["值"], "写入值不一致"
        assert 甲["回读"] == 乙["回读"] == []
        assert (甲["标题"] or {}).get("新文") == (乙["标题"] or {}).get("新文")
    # 第二、三份确实命中过复用
    记录 = json.loads((tmp_path / "复用二" / "任务_二" / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["复用"] and 记录["调用"]["第三遍"] == 0


def test_缩小打印区域与隐藏列后结果仍与基线一致(tmp_path, monkeypatch):
    金 = _金标准(tmp_path)
    打印 = {"三": "A1:D6"}                       # 第三份只打印前两组，第三组在打印区域外
    语料三 = [("三", 科目表, 2027)]
    # 基线：A 仍有三组，B 只有可见的两组
    基线 = []
    for 名, 科目, 年 in 语料三:
        基线.append(_规范(_跑(tmp_path / f"基线{名}", 名, 科目, 年, None, 金, monkeypatch,
                            打印区域=打印.get(名), 科目B=科目[:2])))
    复用 = []
    for 名, 科目, 年 in 语料三:
        复用.append(_规范(_跑(tmp_path / f"复用{名}", 名, 科目, 年, tmp_path / "记录三", 金,
                            monkeypatch, 打印区域=打印.get(名), 科目B=科目[:2])))
    assert 基线[0] == 复用[0]
    # 打印区域外的第三组不进入本次识别
    assert not any(x[0] == "B" and x[2] == "存货明细" for x in 复用[0]["结构"])


def test_金标准版本变化后旧记录退出自动复用(tmp_path, monkeypatch):
    金 = _金标准(tmp_path)
    基线 = _基线(tmp_path, monkeypatch, 金, 语料)
    记录目录 = 安全记录目录()
    _复用跑(tmp_path, monkeypatch, 金, 语料[:1], 记录目录=记录目录)
    # 金标准内容变了：追加一条有效编号
    金.write_text(金.read_text(encoding="utf-8") + "\n" + json.dumps(
        {"id": "N999", "note": {"name": "货币资金"}, "table": {"name": "余额构成"},
         "slot": {"name": "补充"}, "dimensions": [{"name": "period", "required": True}]},
        ensure_ascii=False), encoding="utf-8")
    任务 = _跑(tmp_path / "改金标准", "二", 语料[1][1], 2026, 记录目录, 金, monkeypatch)
    记录 = json.loads((任务 / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第一遍"] >= 1                     # 版本不符，回退重认
    assert any("金标准版本" in (x.get("原因") or "") for x in 记录["回退"])
    规范 = _规范(任务)
    assert 规范["结构"] == 基线[1]["结构"]
    assert 规范["对应"] == 基线[1]["对应"]


def test_识别规则版本变化后旧记录退出自动复用(tmp_path, monkeypatch):
    金 = _金标准(tmp_path)
    基线 = _基线(tmp_path, monkeypatch, 金, 语料)
    记录目录 = 安全记录目录()
    _复用跑(tmp_path, monkeypatch, 金, 语料[:1], 记录目录=记录目录)
    from 核心 import 复用证据
    原 = 复用证据.规则摘要
    monkeypatch.setattr(复用证据, "规则摘要", lambda: 原() + "-改")
    任务 = _跑(tmp_path / "改规则", "二", 语料[1][1], 2026, 记录目录, 金, monkeypatch)
    记录 = json.loads((任务 / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第一遍"] >= 1
    assert any("规则版本" in (x.get("原因") or "") for x in 记录["回退"])
    assert _规范(任务)["结构"] == 基线[1]["结构"]


def test_同结构冲突时与基线结果一致(tmp_path, monkeypatch):
    金 = _金标准(tmp_path)
    基线 = _基线(tmp_path, monkeypatch, 金, 语料)
    记录目录 = 安全记录目录()
    _复用跑(tmp_path, monkeypatch, 金, 语料[:1], 记录目录=记录目录)
    簿 = 记录簿(记录目录)
    索引 = 簿.快照()
    项 = next(x for x in 索引["解释"].values() if x["侧"] == "A" and x["步骤"] == "第三遍")
    冲突 = json.loads(json.dumps(项["解释"], ensure_ascii=False))
    冲突["行"][1]["槽位"] = {"B": "N002", "C": "N002"}
    键, 解 = 证据键(项["证据"]), 解释键(冲突)
    簿.提交([发布事件("A", "第三遍", 项["证据"], 冲突, 键, 解,
                  版本号(键, 解, 项["金标准版本"], 项["规则版本"]),
                  项["金标准版本"], 项["规则版本"],
                  {"任务": "另一任务", "侧": "A", "表组": "G1"}, 形状键=None)], "注入冲突")
    任务 = _跑(tmp_path / "撞", "二", 语料[1][1], 2026, 记录目录, 金, monkeypatch)
    记录 = json.loads((任务 / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第三遍"] >= 1
    assert any("冲突" in (x.get("原因") or "") for x in 记录["回退"])
    assert _规范(任务) == 基线[1]


def test_隐藏辅助公式变化可见值更新仍复用(tmp_path, monkeypatch):
    """区域外／隐藏的辅助格变了、可见值也变了：结构没变就复用，读值用本次文件。"""
    金 = _金标准(tmp_path)
    记录目录 = 安全记录目录()
    语料一 = [("一", 科目表, 2025)]
    _复用跑(tmp_path, monkeypatch, 金, 语料一, 记录目录=记录目录)
    # 第二份：同布局，可见数值整体 +50，隐藏辅助格改成别的值
    任务 = _跑(tmp_path / "复用二", "二", 科目表, 2026, 记录目录, 金, monkeypatch,
              值偏移=50, 辅助={"值": "辅助已改"})
    记录 = json.loads((任务 / "复用记录_A.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第三遍"] == 0 and len(记录["复用"]) >= 3
    规范 = _规范(任务)
    值 = {v for v in 规范["值"].values() if isinstance(v, (int, float))}
    assert {152, 252} <= 值                      # 102+50、202+50：本次可见值
    基线 = _规范(_跑(tmp_path / "基线二", "二", 科目表, 2026, None, 金, monkeypatch,
                   值偏移=50, 辅助={"值": "辅助已改"}))
    assert 规范["值"] == 基线["值"]                # 与同一份文件的完整识别一致


def test_重新隐藏数据行后不沿用旧坐标(tmp_path, monkeypatch):
    """把数据行藏起来：可见结构变了，回退重认，结果与同一份文件的基线一致。"""
    金 = _金标准(tmp_path)
    记录目录 = 安全记录目录()
    _复用跑(tmp_path, monkeypatch, 金, [("一", 科目表, 2025)], 记录目录=记录目录)
    任务 = _跑(tmp_path / "复用二", "二", 科目表[:1], 2026, 记录目录, 金, monkeypatch,
              隐藏行=2, 科目B=科目表[:1])
    记录 = json.loads((任务 / "复用记录_B.json").read_text(encoding="utf-8"))
    assert 记录["调用"]["第一遍"] >= 1 or 记录["调用"]["第二遍"] >= 1
    基线 = _基线(tmp_path, monkeypatch, 金, [("二", 科目表[:1], 2026)])
    assert _规范(任务)["结构"] == 基线[0]["结构"]
