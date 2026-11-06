# coding: utf-8
import hashlib
import json
from pathlib import Path
from 测试.GUI样本 import 准备, 开始, 等待


def 哈希们(p):
    return {str(x.relative_to(p)): hashlib.sha256(x.read_bytes()).hexdigest()
            for x in p.rglob("*") if x.is_file()}


def test_真实结果分页名单公式与文件变动(tmp_path, monkeypatch):
    客户, 输入 = 准备(tmp_path, monkeypatch)
    初 = 开始(客户, 输入, "结果查询")
    等待(客户, 初["task_id"])
    基址 = f"/api/桌面/任务/{初['task_id']}"
    p = Path(初["运行目录"])
    前 = 哈希们(p)
    r = 客户.get(基址 + "/结果", params={"view": "明细"})
    assert r.status_code == 200
    数据 = r.json()
    名单 = next(x for x in 数据["条目"] if x["A位置"]["格"] == "B2")
    assert 名单["B本次值"]["原样文本"] == "=客户甲"
    assert 名单["A实际值"]["原样文本"] == "=客户甲"
    assert 名单["最终动作"] == "已写入"
    assert 数据["计数"]["已写入"] == 6
    式 = 客户.get(基址 + "/结果", params={"view": "公式"}).json()["条目"]
    assert 式[0]["结论"] == "本次金额一致"
    assert 客户.get(基址 + "/结果", params={"view": "明细", "结果版本": "旧版本"}).status_code == 409
    文件 = 客户.get(基址 + "/文件").json()["条目"]
    成品 = next(x for x in 文件 if x["名称"] == "Excel新副本")
    assert 成品["文件状况"] == "未改动" and 成品["正式"]
    assert 哈希们(p) == 前
    Path(成品["路径"]).write_bytes("本例模拟用户改动".encode("utf-8"))
    新 = 客户.get(基址 + "/文件").json()["条目"]
    assert next(x for x in 新 if x["file_id"] == 成品["file_id"])["文件状况"] == "文件已被改动"


def test_业务表筛选与历史文件状态来自已有记录(tmp_path, monkeypatch):
    客户, 输入 = 准备(tmp_path, monkeypatch)
    r = 开始(客户, 输入, "历史核查")
    等待(客户, r["task_id"])
    基址 = f"/api/桌面/任务/{r['task_id']}"
    表 = 客户.get(基址 + "/结果", params={"view": "表", "侧": "A", "仅需查看": True}).json()
    assert 表["条目"] == []
    历史 = 客户.get("/api/桌面/任务").json()["条目"][0]
    assert 历史["需查看表数"] == 0
    assert 历史["文件状况"] == "未改动"


def test_交付留证被截短不能显示为完整结果(tmp_path, monkeypatch):
    客户, 输入 = 准备(tmp_path, monkeypatch)
    初 = 开始(客户, 输入, "明细完整性")
    等待(客户, 初["task_id"])
    p = Path(初["运行目录"])
    明细 = p / "交付明细.jsonl"
    文 = 明细.read_text(encoding="utf-8")
    明细.write_text(文.splitlines()[0] + "\n", encoding="utf-8")
    数据 = 客户.get(f"/api/桌面/任务/{初['task_id']}/结果", params={"view": "明细"}).json()
    assert 数据["明细完整性"] == "交付明细已变化，无法确认最终动作"
    assert 数据["计数"].get("已写入", 0) == 0
    assert all(not x["A实际值"]["是否取得"] for x in 数据["条目"])


def test_旧任务未知与类型不猜也不改写(tmp_path, monkeypatch):
    客户, _ = 准备(tmp_path, monkeypatch)
    from 服务 import web
    p = Path(web.运行根) / "任务_旧资料"
    p.mkdir(parents=True)
    (p / "接续索引.json").write_text(json.dumps({"步骤": "done", "撤回表": ["区域1"], "产物": {}}))
    行们 = [{"格": f"余额!本期!C{i+1}", "表组": "G1", "读取部分": "P001", "原值": v,
           "原值状态": "值", "角色": "数据格"} for i, v in enumerate([0, "00123", -12.345678901, None])]
    (p / "A合集.jsonl").write_text("\n".join(json.dumps(x) for x in 行们), encoding="utf-8")
    (p / "对应清单.jsonl").write_text(json.dumps({"a格": 行们[0]["格"], "b格": "来源!C1", "状态": "对应"}))
    前 = 哈希们(p)
    数据 = 客户.get(f"/api/桌面/任务/{p.name}/结果", params={"view": "明细"}).json()
    assert 数据["明细完整性"] == "旧任务未记录最终动作"
    assert "旧记录状态不一致" in 数据["提示"]
    assert 数据["条目"][0]["A位置"] == {"工作表": "余额!本期", "格": "C1"}
    assert 数据["条目"][1]["A原值"]["原样文本"] == "00123"
    assert 数据["条目"][0]["A原值"]["原样文本"] == "0"
    assert all(x["最终动作"] == "未记录" and not x["A实际值"]["是否取得"] for x in 数据["条目"])
    assert 数据["计数"].get("已写入", 0) == 0
    assert 哈希们(p) == 前


def test_按业务表查看不漏掉留证已关联的组级问题(tmp_path, monkeypatch):
    c, _ = 准备(tmp_path, monkeypatch)
    from 服务 import web
    p = Path(web.运行根) / "任务_组级问题"
    p.mkdir(parents=True)
    (p / "任务说明.json").write_text('{"执行次数":1}', encoding="utf-8")
    (p / "接续索引.json").write_text(json.dumps({"步骤": "部分保留", "产物": {"交付明细": str(p / "交付明细.jsonl")}}), encoding="utf-8")
    结构 = [{"侧": s, "表组": "G1", "读取部分": "P1", "科目": "货币资金", "表名": "余额"} for s in ("A", "B")]
    问题 = [{"表组": "G1", "原因": "标题失败"}, {"组": ["G1"], "原因": "整组保留"}]
    明细 = [{"类型": "格子", "执行次数": 1, "侧": "A", "读取部分": "P1", "表组": "G1", "a格": "资金!C2", "最终动作": "原值保留", "原因引用": [0, 1]}]
    for 名, 内容 in (("结构清单.jsonl", 结构), ("问题清单.jsonl", 问题), ("交付明细.jsonl", 明细)):
        (p / 名).write_text("\n".join(json.dumps(x) for x in 内容), encoding="utf-8")
    基址 = f"/api/桌面/任务/{p.name}/结果"
    表 = c.get(基址, params={"view": "表", "侧": "A"}).json()["条目"][0]
    组内 = c.get(基址, params={"view": "问题", "组": 表["组键"]}).json()
    assert {x["原因"] for x in 组内["条目"]} == {"标题失败", "整组保留"}
    assert 表["问题数"] == 2
    assert c.get(基址, params={"view": "问题"}).json()["问题条数"] == 2
