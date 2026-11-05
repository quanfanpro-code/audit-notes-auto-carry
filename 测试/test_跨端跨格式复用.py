# coding: utf-8
import json
from pathlib import Path

from 测试.组合样本 import 造Word, 造金标准, 文件哈希, 换端Word模型, 读Word格, 跑场景
from 测试.安全目录 import 安全记录目录


def _输入(A, B, 金, 记录):
    return {"A": {"路径": str(A), "工作表": []}, "B": {"路径": str(B), "工作表": []},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}


def _调用为零(结果):
    assert 结果["调用"]
    assert all((x["第一遍"], x["第二遍"], x["第三遍"]) == (0, 0, 0)
               for x in 结果["调用"]), 结果["调用"]


def test_换端连续三任务及最终名单清尾(tmp_path, monkeypatch):
    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    两 = 造Word(tmp_path / "两.docx", 100, 110, 名单=("甲公司", "乙公司"))
    三 = 造Word(tmp_path / "三.docx", 901, 102, 名单=("客户甲", "客户乙", "客户丙"))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    前 = {p: 文件哈希(p) for p in (两, 三)}

    一 = 跑场景(_输入(两, 三, 金, 记录), tmp_path / "任务一", 换端Word模型(), monkeypatch)
    assert 一["状态"]["步骤"] == "done", 一["状态"]
    assert (读Word格(一["输出"], 0, 1, 0), 读Word格(一["输出"], 0, 1, 1),
            读Word格(一["输出"], 0, 1, 2)) == ("客户甲", "901", "102")
    assert all(x["第三遍"] == 1 for x in 一["调用"])

    二 = 跑场景(_输入(三, 两, 金, 记录), tmp_path / "任务二", 换端Word模型(), monkeypatch)
    _调用为零(二)
    assert (读Word格(二["输出"], 0, 1, 0), 读Word格(二["输出"], 0, 1, 1),
            读Word格(二["输出"], 0, 1, 2)) == ("甲公司", "100", "110")
    assert (读Word格(二["输出"], 0, 2, 0), 读Word格(二["输出"], 0, 2, 1),
            读Word格(二["输出"], 0, 2, 2)) == ("乙公司", "200", "210")
    assert [读Word格(二["输出"], 0, 3, i) for i in range(3)] == ["", "", ""]

    三份 = 跑场景(_输入(三, 两, 金, 记录), tmp_path / "任务三", 换端Word模型(), monkeypatch)
    _调用为零(三份)
    assert 读Word格(三份["输出"], 0, 1, 0) == "甲公司"
    assert 读Word格(三份["输出"], 0, 1, 1) == "100"
    assert {p: 文件哈希(p) for p in (两, 三)} == 前


def test_换端使用与停用追溯到工作表(tmp_path, monkeypatch):
    from 核心.共用记录簿 import 记录簿
    from 核心.复用管理 import 停用

    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    两 = 造Word(tmp_path / "两.docx", 100, 110, 名单=("甲公司", "乙公司"))
    三 = 造Word(tmp_path / "三.docx", 901, 102, 名单=("客户甲", "客户乙", "客户丙"))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    跑场景(_输入(两, 三, 金, 记录), tmp_path / "任务一", 换端Word模型(), monkeypatch)
    任务二 = tmp_path / "任务二"
    跑场景(_输入(三, 两, 金, 记录), 任务二, 换端Word模型(), monkeypatch)

    簿 = 记录簿(记录)
    使用 = 簿.快照()["使用"]
    本次 = [x for x in 使用 if x.get("任务") == str(任务二)]
    assert 本次 and all(x.get("读取部分") == "P001" and x.get("工作表") == "sheet1" for x in 本次)
    版本 = next(x["解释版本"] for x in 本次 if x.get("侧") == "B")
    结果 = 停用(簿, "跨端复核停用", [版本], "复核测试", "本轮记录")
    条目 = next(x for x in 结果["直接"]["任务表组"] if x["任务"] == str(任务二))
    assert 条目["侧"] == "B" and 条目["读取部分"] == "P001" and 条目["工作表"] == "sheet1"


def test_旧证据版本4退出自动复用():
    from 核心.复用接入 import 版本可用
    from 核心.复用证据 import 规则摘要

    项 = {"状态": "有效", "金标准版本": "金", "规则版本": 规则摘要(),
          "证据": {"证据版本": 4}}
    可用, 原因 = 版本可用(项, "金", 规则摘要())
    assert not 可用 and "证据版本" in 原因


def _跨输入(A, B, 金, 记录):
    return {"A": {"路径": str(A), "工作表": []},
            "B": {"路径": str(B), "工作表": ["余额"] if Path(B).suffix == ".xlsx" else []},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}


def _跨Excel(路径, 期初, 期末, 名单, 表内单位=True):
    from 测试.组合样本 import 造Excel

    表头 = "项目（单位：人民币元）" if 表内单位 else "项目"
    return 造Excel(路径, {"余额": [[表头, "2025-01-01", "2025-12-31"],
        [名单[0], 期初, 期末], [名单[1], 期初 + 100, 期末 + 100]]}, 起列=2)


def test_跨格式双向复用且最终名单金额为本次(tmp_path, monkeypatch):
    from 测试.组合样本 import 跨格式模型

    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    A = 造Word(tmp_path / "A.docx")
    金 = 造金标准(tmp_path / "金标准.jsonl")
    B1 = _跨Excel(tmp_path / "B1.xlsx", 901, 102, ("Excel甲", "Excel乙"))
    一 = 跑场景(_跨输入(A, B1, 金, 记录), tmp_path / "任务一", 跨格式模型(), monkeypatch)
    assert 一["状态"]["步骤"] == "done"
    assert (读Word格(一["输出"], 0, 1, 0), 读Word格(一["输出"], 0, 1, 1),
            读Word格(一["输出"], 0, 1, 2)) == ("Excel甲", "901", "102")

    B2 = 造Word(tmp_path / "B2.docx", 501, 502, 名单=("Word甲", "Word乙"),
               表内单位=True, 表前单位=False)
    二 = 跑场景(_跨输入(A, B2, 金, 记录), tmp_path / "任务二", 跨格式模型(), monkeypatch)
    B调用 = next(x for x in 二["调用"] if x["侧"] == "B")
    assert (B调用["第一遍"], B调用["第二遍"], B调用["第三遍"]) == (1, 1, 0), B调用
    assert (读Word格(二["输出"], 0, 1, 0), 读Word格(二["输出"], 0, 1, 1),
            读Word格(二["输出"], 0, 1, 2)) == ("Word甲", "501", "502")

    B3 = _跨Excel(tmp_path / "B3.xlsx", 601, 602, ("新Excel甲", "新Excel乙"))
    三 = 跑场景(_跨输入(A, B3, 金, 记录), tmp_path / "任务三", 跨格式模型(), monkeypatch)
    B调用 = next(x for x in 三["调用"] if x["侧"] == "B")
    assert (B调用["第一遍"], B调用["第二遍"], B调用["第三遍"]) == (0, 0, 0), B调用
    assert (读Word格(三["输出"], 0, 1, 0), 读Word格(三["输出"], 0, 1, 1),
            读Word格(三["输出"], 0, 1, 2)) == ("新Excel甲", "601", "602")


def test_跨格式证据不足重认且最终不搬(tmp_path, monkeypatch):
    from 测试.组合样本 import 跨格式模型

    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    A = 造Word(tmp_path / "A.docx")
    B1 = 造Word(tmp_path / "B1.docx", 901, 102)
    金 = 造金标准(tmp_path / "金标准.jsonl")
    一 = 跑场景(_跨输入(A, B1, 金, 记录), tmp_path / "任务一", 跨格式模型(), monkeypatch)
    assert 一["状态"]["步骤"] == "done"

    B2 = _跨Excel(tmp_path / "B2.xlsx", 999, 888, ("错甲", "错乙"), 表内单位=False)
    二 = 跑场景(_跨输入(A, B2, 金, 记录), tmp_path / "任务二", 跨格式模型(), monkeypatch)
    B调用 = next(x for x in 二["调用"] if x["侧"] == "B")
    assert B调用["第三遍"] == 1, B调用
    assert (读Word格(二["输出"], 0, 1, 0), 读Word格(二["输出"], 0, 1, 1),
            读Word格(二["输出"], 0, 1, 2)) == ("甲公司", "100", "110")
    问题 = (二["任务"] / "问题清单.jsonl").read_text(encoding="utf-8")
    assert "单位" in 问题


def _Excel布局(路径, 项数, 期初, 期末, 表名, 名称):
    from 测试.组合样本 import 造Excel

    行们 = [["项目（单位：人民币元）", "2025-01-01", "2025-12-31"]]
    for i in range(项数):
        行们.append([名称[i], 期初 + i * 100, 期末 + i * 100])
    return 造Excel(路径, {表名: 行们}, 起列=2)


def _换端Excel模型():
    def 调用(配置, 指令, 材料, 目录):
        实际 = 材料.split("【二、", 1)[0]
        三项 = "丙公司" in 实际 or "客户丙" in 实际
        底 = 4 if 三项 else 3
        if 指令.startswith("把整份文件"):
            return json.dumps({"划分": [{"行范围": [1, 底], "科目": "货币资金"}]}, ensure_ascii=False)
        if 指令.startswith("在第一遍"):
            return json.dumps({"表": [{"行范围": [1, 底], "科目": "货币资金", "表名": "余额构成",
                "表组": "G1", "续表序": 1}]}, ensure_ascii=False)
        列 = [{"列": "B", "路径": "项目", "角色": "行名列"},
              {"列": "C", "路径": "期初余额", "角色": "数据列", "期间角色": "本期期初",
               "期间实际": {"值": "2025-01-01", "位置": "C1"}},
              {"列": "D", "路径": "期末余额", "角色": "数据列", "期间角色": "本期期末",
               "期间实际": {"值": "2025-12-31", "位置": "D1"}}]
        行 = [{"行": 1, "名字": "项目", "角色": "表头"}] + [
            {"行": i, "名字": f"名单{i - 1}", "角色": "数据", "序号": i - 1,
             "槽位": {"C": "N001", "D": "N001"}} for i in range(2, 底 + 1)]
        作用 = {"币种": "人民币", "单位": "元", "期间跨度": "时点", "单位依据": [
            {"来源": "表内", "位置": "B1", "原文": "项目（单位：人民币元）", "片段": "元",
             "币种": "人民币", "单位": "元", "列": ["C", "D"]}]}
        return json.dumps({"列": 列, "行": 行, "作用范围": 作用, "空名单": [], "例外": []},
                          ensure_ascii=False)
    return 调用


def test_Excel双向换端及工作表改名仍复用(tmp_path, monkeypatch):
    from 测试.组合样本 import 读Excel格

    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    两 = _Excel布局(tmp_path / "两.xlsx", 2, 100, 110, "资金", ("甲公司", "乙公司"))
    三 = _Excel布局(tmp_path / "三.xlsx", 3, 901, 102, "余额", ("客户甲", "客户乙", "客户丙"))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    一输入 = {"A": {"路径": str(两), "工作表": ["资金"]}, "B": {"路径": str(三), "工作表": ["余额"]},
             "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    一 = 跑场景(一输入, tmp_path / "任务一", _换端Excel模型(), monkeypatch)
    assert 一["状态"]["步骤"] == "done", 一["状态"]
    assert 读Excel格(一["输出"], "资金", "C2") == 901

    wb = __import__("openpyxl").load_workbook(三)
    wb["余额"].title = "来源改名"
    wb.save(三)
    wb.close()
    二输入 = {"A": {"路径": str(三), "工作表": ["来源改名"]}, "B": {"路径": str(两), "工作表": ["资金"]},
             "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    二 = 跑场景(二输入, tmp_path / "任务二", _换端Excel模型(), monkeypatch)
    _调用为零(二)
    assert 读Excel格(二["输出"], "来源改名", "B2") == "甲公司"
    assert 读Excel格(二["输出"], "来源改名", "C2") == 100
    assert 读Excel格(二["输出"], "来源改名", "D2") == 110
    assert 读Excel格(二["输出"], "来源改名", "C4") is None


def test_A端跨格式复用后最终Excel用本次值(tmp_path, monkeypatch):
    from 测试.组合样本 import 跨格式模型, 读Excel格

    monkeypatch.setenv("审计附注_复用", "1")
    记录 = 安全记录目录()
    monkeypatch.setenv("审计附注_识别记录目录", str(记录))
    金 = 造金标准(tmp_path / "金标准.jsonl")
    A1 = 造Word(tmp_path / "A1.docx", 100, 110, 表内单位=True, 表前单位=False)
    B1 = _跨Excel(tmp_path / "B1.xlsx", 901, 102, ("Excel甲", "Excel乙"))
    一 = 跑场景(_跨输入(A1, B1, 金, 记录), tmp_path / "任务一", 跨格式模型(), monkeypatch)
    assert 一["状态"]["步骤"] == "done"

    A2 = _跨Excel(tmp_path / "A2.xlsx", 100, 110, ("甲公司", "乙公司"))
    B2 = 造Word(tmp_path / "B2.docx", 501, 502, 名单=("Word甲", "Word乙"),
               表内单位=True, 表前单位=False)
    输入 = {"A": {"路径": str(A2), "工作表": ["余额"]}, "B": {"路径": str(B2), "工作表": []},
            "金标准": str(金), "识别记录目录": str(记录), "复用": True}
    二 = 跑场景(输入, tmp_path / "任务二", 跨格式模型(), monkeypatch)
    assert all(x["第三遍"] == 0 for x in 二["调用"]), 二["调用"]
    assert 读Excel格(二["输出"], "余额", "B2") == "Word甲"
    assert 读Excel格(二["输出"], "余额", "C2") == 501
    assert 读Excel格(二["输出"], "余额", "D2") == 502
