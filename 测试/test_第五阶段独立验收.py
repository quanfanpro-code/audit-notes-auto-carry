# coding: utf-8
"""仅复核：使用现有接口验证完整两期、标题、金标准和日常操作，不修改功能。"""
import datetime
import json
import shutil
import subprocess
import sys
import shutil
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.workbook.defined_name import DefinedName

from 核心.比对 import 建立对应
from 核心.搬运 import 生成更新表, 保存更新表
from 核心.识别 import 识别一侧
from 核心.金标准 import 加载金标准, 建索引, 候选槽位, 编号存在

根 = Path(__file__).resolve().parent.parent


@pytest.fixture(scope="session")
def 金原():
    return 加载金标准(next((根 / "金标准").rglob("*以此为准.jsonl")))


def _写证据(目录, **数据):
    (目录 / "实际结果.json").write_text(json.dumps(数据, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def _造场景(tmp_path, monkeypatch, 金原, B值, 方向="行", 其它表=False,
          标题=False, 给期间=True, 空格有样式=False):
    编号 = "C-N049-T001-S001" if 方向 == "行" else "C-N038-T001-S001"
    原 = next(r for r in 金原 if r["id"] == 编号)
    词 = {k: 原[k] for k in ("id", "note", "table", "slot", "row_path", "column_path", "dimensions", "value_type")}
    金 = tmp_path / "金标准.jsonl"
    金.write_text(json.dumps(词, ensure_ascii=False), encoding="utf-8")
    科目, 表名 = 原["note"]["name"], 原["table"]["name"]
    A值 = [(100, 110), (200, 210), (300, 310)]
    文件 = {}
    for 侧, 数组, 工作表 in (("A", A值, "sheet1"), ("B", B值, "word附注")):
        wb = Workbook()
        s = wb.active
        s.title = 工作表
        if 方向 == "行":
            s.append(["项目", "期初余额", "期末余额"])
            for i, 两期 in enumerate(数组, 2):
                s.cell(i, 1, ("旧" if 侧 == "A" else "新") + str(i - 1))
                for c, v in zip((2, 3), 两期):
                    s.cell(i, c, v)
                    if v is None and 空格有样式:
                        s.cell(i, c).font = Font(name="宋体")
            范围 = f"A1:C{len(数组)+1}"
        else:
            s.cell(2, 1, "期初余额")
            s.cell(3, 1, "期末余额")
            for c, 两期 in enumerate(数组, 2):
                s.cell(1, c, ("旧" if 侧 == "A" else "新") + str(c - 1))
                for r, v in zip((2, 3), 两期):
                    s.cell(r, c, v)
            范围 = f"A1:{chr(65+len(数组))}3"
        if 侧 == "A":
            wb.defined_names.add(DefinedName("TA_001", attr_text=f"'sheet1'!{范围}"))
        if 其它表 and 侧 == "B":
            wb.create_sheet("不参与识别的表")["A100"] = 999
        文件[侧] = tmp_path / f"{侧}.xlsx"
        wb.save(文件[侧])
        wb.close()

    def 模型(配置, 指令, 材料, 调用目录):
        侧 = Path(调用目录).name[0]
        数量 = len(A值 if 侧 == "A" else B值)
        行尾 = 数量 + 1 if 方向 == "行" else 3
        if 指令.startswith("把整份文件"):
            答 = {"划分": [{"行范围": [1, 行尾], "科目": 科目}]}
        elif 指令.startswith("在第一遍"):
            t = {"行范围": [1, 行尾], "科目": 科目, "表名": 表名, "表组": "G", "续表序": 1}
            if 标题 and 侧 == "A":
                t["标题"] = {"原文": "截至2024年12月31日", "新文": "截至2025年12月31日"}
            答 = {"表": [t]}
        elif 方向 == "行":
            列 = [{"列": "A", "角色": "行名列", "路径": "项目"}]
            for 字母, 期间 in (("B", "本期期初"), ("C", "本期期末")):
                c = {"列": 字母, "角色": "数据列", "路径": 期间}
                if 给期间:
                    c["期间角色"] = 期间
                列.append(c)
            答 = {"列": 列, "行": [{"行": 1, "角色": "表头"}] + [
                {"行": i+2, "角色": "数据", "序号": i+1, "槽位": {"B": 编号, "C": 编号}}
                for i in range(数量)], "作用范围": {"币种": "人民币", "单位": "元", "期间跨度": "时点"}}
        else:
            字母们 = [chr(66+i) for i in range(数量)]
            列 = [{"列": "A", "角色": "行名列", "路径": "期间"}] + [
                {"列": c, "角色": "数据列", "序号": i+1, "标签行": 1, "路径": "资产类别"}
                for i, c in enumerate(字母们)]
            答 = {"列": 列, "行": [{"行": 1, "角色": "表头"}] + [
                {"行": i, "角色": "数据", "期间角色": 期间, "槽位": {c: 编号 for c in 字母们}}
                for i, 期间 in ((2, "本期期初"), (3, "本期期末"))],
                "作用范围": {"币种": "人民币", "单位": "元", "期间跨度": "时点"}}
        return json.dumps(答, ensure_ascii=False)

    monkeypatch.setattr("核心.识别.调用识别", 模型)
    return 文件, 金, 模型


def _两侧识别(tmp_path, monkeypatch, 金原, B值, **设置):
    文件, 金, _ = _造场景(tmp_path, monkeypatch, 金原, B值, **设置)
    任务 = tmp_path / "任务"
    识别一侧({"测试": True}, 文件["A"], "A", 任务, 金)
    识别一侧({"测试": True}, 文件["B"], "B", 任务, 金, 只跑表=["word附注"])
    def 读(名):
        return [json.loads(x) for x in (任务 / 名).read_text(encoding="utf-8").splitlines()]
    a, b, 结构 = 读("A合集.jsonl"), 读("B合集.jsonl"), 读("结构清单.jsonl")
    A结 = [x for x in 结构 if x["侧"] == "A"]
    B结 = [x for x in 结构 if x["侧"] == "B"]
    清单 = 建立对应(a, b, A结, B结)
    变更, 问题 = 生成更新表({}, 清单, b, a)
    输出 = tmp_path / "实际更新表.xlsx"
    保存更新表(文件["A"], 变更, 输出)
    wb = load_workbook(输出)
    try:
        值 = {cell.coordinate: cell.value for row in wb.active for cell in row}
    finally:
        wb.close()
    _写证据(tmp_path, B输入=B值, A结构=A结, B结构=B结, A合集=a, B合集=b, 对应=清单, 变更=变更, 问题=问题, 实际Excel=值)
    return a, b, A结, B结, 变更, 问题, 值


@pytest.mark.parametrize("方向", ["行", "列"])
def test_两期全零之后的非零项目不能丢失(tmp_path, monkeypatch, 金原, 方向):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原,
                      [(101, 111), (0, 0), (303, 313)], 方向=方向)
    # 不预设零项是否压缩；任何排法都不能把仍在B中的303和313全部弄丢。
    assert 303 in 值.values() and 313 in 值.values(), 值


@pytest.mark.parametrize("中项", [(0, 211), (201, 0), (201, 211)])
def test_只有一期为零照常搬运_对照(tmp_path, monkeypatch, 金原, 中项):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原, [(101, 111), 中项, (303, 313)])
    assert (值["B3"], 值["C3"]) == 中项
    assert (值["B4"], 值["C4"]) == (303, 313)


def test_全表两期均零不搬_对照(tmp_path, monkeypatch, 金原):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原, [(0, 0), (0, 0), (0, 0)])
    assert [值[c] for c in ("B2", "C2", "B3", "C3", "B4", "C4")] == [100, 110, 200, 210, 300, 310]


def test_公式没缓存不能使后续非零项目被清空(tmp_path, monkeypatch, 金原):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原, [(101, 111), ("=1+1", "=2+2"), (303, 313)])
    assert 303 in 值.values() and 313 in 值.values(), 值


@pytest.mark.parametrize("空格有样式", [False, True])
def test_真实空白不应因单元格样式不同而失效(tmp_path, monkeypatch, 金原, 空格有样式):
    *_, 值 = _两侧识别(tmp_path, monkeypatch, 金原,
                      [(101, 111), (None, 211), (303, 313)], 空格有样式=空格有样式)
    assert 值["B3"] is None and 值["C3"] == 211 and 值["A3"] == "新2", 值


@pytest.mark.parametrize("其它表", [False, True])
def test_指定工作表完整性不受别表影响(tmp_path, monkeypatch, 金原, 其它表):
    _, _, _, B结, *余 = _两侧识别(tmp_path, monkeypatch, 金原,
                               [(101, 111), (201, 211)], 其它表=其它表)
    assert B结[0]["完整性"]["第一遍覆盖"] is True, B结[0]["完整性"]
    assert 余[-1]["B4"] is None and 余[-1]["C4"] is None


def test_已交付标题必须进入最终写回清单(tmp_path, monkeypatch, 金原):
    from 核心.写回 import 读标题更新
    _, _, A结, *_ = _两侧识别(tmp_path, monkeypatch, 金原,
                            [(101, 111), (201, 211), (303, 313)], 标题=True)
    更新 = 读标题更新(A结)
    _写证据(tmp_path, 识别给的标题={"原文": "截至2024年12月31日", "新文": "截至2025年12月31日"},
          保存的结构=A结, 最终标题更新=更新)
    assert 更新 == [{"range_name": "TA_001", "old_text": "截至2024年12月31日", "text": "截至2025年12月31日"}]


def test_必需期间漏填不能登记为识别完成(tmp_path, monkeypatch, 金原):
    # 走真实识别入口，不向不存在的拟议接口传参数。
    文件, 金, _ = _造场景(tmp_path, monkeypatch, 金原, [(101, 111)], 给期间=False)
    try:
        识别一侧({"测试": True}, 文件["A"], "A", tmp_path / "任务", 金)
    except RuntimeError:
        return
    结构 = [json.loads(x) for x in (tmp_path / "任务" / "结构清单.jsonl").read_text(encoding="utf-8").splitlines()]
    _写证据(tmp_path, 结构=结构)
    assert not any(x.get("完成") for x in 结构), "缺少期间仍登记完成"


@pytest.mark.parametrize("期望,实际,应一致", [("00123", "123", False), ("1.00", "1", False),
                                               (96.91, 97, True), (3.09, 3.1, True), (0, None, False)])
def test_回读区分文本金额和空白(tmp_path, monkeypatch, 期望, 实际, 应一致):
    from 核心.回读 import 回读验证
    def 回读提取(word, excel, linked, context):
        wb = Workbook()
        wb.active.title = "sheet1"
        wb.active["A1"] = 实际
        wb.save(excel)
        wb.close()
    monkeypatch.setattr("核心.回读.调用桥接提取", 回读提取)
    差异 = 回读验证("模拟输出.docx", {"sheet1!A1": 期望}, tmp_path)
    _写证据(tmp_path, 期望=期望, 实际=实际, 差异=差异)
    assert (差异 == []) is 应一致


@pytest.mark.parametrize("状态", ["superseded", "deprecated"])
def test_已停用金标准不得继续发给识别(金原, 状态):
    旧 = next(r for r in 金原 if r.get("status") == 状态)
    索引 = 建索引(金原)
    候选 = 候选槽位(索引, 旧["note"]["name"], 旧["table"]["name"])
    assert 旧["id"] not in {c["id"] for c in 候选}, (旧["id"], 状态, 旧.get("superseded_by"))


def test_停用号不能继续通过合法性检查(金原):
    旧 = next(r for r in 金原 if r.get("status") == "superseded")
    assert not 编号存在(建索引(金原), 旧["id"]), 旧["id"]


def test_正常标题格式与非法格式在入口已区分_对照():
    from 核心.识别 import _解析表
    表 = {"行范围": [1, 3], "科目": "现金", "表名": "余额", "表组": "G", "续表序": 1}
    with pytest.raises(RuntimeError, match="标题结论格式不对"):
        _解析表(json.dumps({"表": [{**表, "标题": {"新文": "新日期"}}]}),
             {"现金": ["余额"]}, [{"行范围": [1, 3], "科目": "现金"}])
    assert _解析表(json.dumps({"表": [表]}), {"现金": ["余额"]},
                [{"行范围": [1, 3], "科目": "现金"}])[0]["标题"] is None


def test_B多于A按现行提交先搬容纳部分_对照(tmp_path, monkeypatch, 金原):
    _, _, _, _, 变更, _, 值 = _两侧识别(tmp_path, monkeypatch, 金原,
          [(101, 111), (201, 211), (303, 313), (404, 414)])
    assert 值["B2"] == 101 and 值["B4"] == 303
    assert 404 not in 变更.values() and 414 not in 变更.values()


def test_桥接等待应有上限(tmp_path, monkeypatch):
    import 核心.桥接 as 桥接
    收到 = {}
    脚本 = tmp_path / "模拟桥接.py"
    脚本.write_text("print('模拟')", encoding="utf-8")
    def 调用(*args, **kwargs):
        收到["timeout"] = kwargs.get("timeout")
        return subprocess.CompletedProcess(args[0], 0, '{"ok":true}', "")
    monkeypatch.setattr(桥接.subprocess, "run", 调用)
    桥接.调用桥接({}, 工作目录=tmp_path, exe=脚本, 前缀=sys.executable)
    _写证据(tmp_path, 设置等待秒数=收到.get("timeout"))
    assert isinstance(收到.get("timeout"), (int, float)) and 收到["timeout"] > 0


def test_桥接默认留证应与发布文件分开(tmp_path):
    from 核心.桥接 import 调用桥接
    发布 = tmp_path / "发布"
    发布.mkdir()
    脚本 = 发布 / "模拟桥接.py"
    脚本.write_text("print('{\"ok\":true}')", encoding="utf-8")
    assert 调用桥接({}, exe=脚本, 前缀=sys.executable)["ok"]
    留证 = list(发布.glob("Word请求_*.json"))
    _写证据(tmp_path, 发布目录请求文件=[str(x) for x in 留证])
    assert not 留证, 留证


def test_人工回写结构清单接口确实可执行_对照(tmp_path, monkeypatch):
    from fastapi.testclient import TestClient
    from 服务.web import 应用
    收到 = {}
    def 更新(word, excel, output, operations, caption_updates):
        收到["操作"] = operations
        return {"ok": True}
    monkeypatch.setattr("核心.写回.调用桥接更新", 更新)
    操作 = [{"range_name": "TA_001", "axis": "row", "action": "insert", "index": 1, "count": 1, "template_index": 0}]
    with TestClient(应用) as c:
        r = c.post("/api/人工回写", json={"带连接word": "模拟.docx", "excel": "模拟.xlsx",
                                    "输出": "新副本.docx", "结构操作": 操作}).json()
    assert r["ok"] and 收到["操作"] == 操作


def test_从待答复续跑完成后不应继续提示旧待答复(tmp_path, monkeypatch):
    from 测试.test_全链 import _装stub, _造金标准
    from 核心.流程 import 运行任务
    _装stub(monkeypatch)
    b = tmp_path / "B.xlsx"
    wb = Workbook()
    wb.active.title = "word附注"
    wb.active["A1"] = 9
    wb.save(b)
    wb.close()
    入 = {"word": "模拟原件.docx", "b表": str(b), "金标准": str(_造金标准(tmp_path))}
    任务 = tmp_path / "任务"
    初 = 运行任务(入, 任务, {})
    assert 初["待答复"]
    后 = 运行任务(入, 任务, {"测试": True})
    _写证据(tmp_path, 开始步骤=初["步骤"], 完成步骤=后["步骤"], 完成后待答复=后.get("待答复"))
    assert 后["步骤"] == "done"
    assert not 后.get("待答复"), 后["待答复"]


def test_全链标题不是只测辅助函数(tmp_path, monkeypatch, 金原):
    from 核心.流程 import 运行任务
    文件, 金, _ = _造场景(tmp_path, monkeypatch, 金原,
                     [(101, 111), (201, 211), (303, 313)], 标题=True)
    原Word = tmp_path / "模拟原件.docx"
    原Word.write_text("仅模拟桥接边界，不是真实Word", encoding="utf-8")
    标题们 = []
    def 提取(word, excel, linked, context):
        shutil.copyfile(文件["A"], excel)
        Path(linked).write_text("模拟连接", encoding="utf-8")
        Path(context).write_text(json.dumps({"tables": [{"range_name": "TA_001", "title": "截至2024年12月31日"}]}),
                                encoding="utf-8")
        return {"ok": True}
    def 更新(word, excel, output, operations, caption_updates):
        标题们.extend(caption_updates)
        Path(output).write_text(json.dumps({"excel": str(excel)}), encoding="utf-8")
        return {"ok": True}
    def 回读提取(word, excel, linked, context):
        src = json.loads(Path(word).read_text(encoding="utf-8"))["excel"]
        shutil.copyfile(src, excel)
    monkeypatch.setattr("核心.提取.调用桥接提取", 提取)
    monkeypatch.setattr("核心.写回.调用桥接更新", 更新)
    monkeypatch.setattr("核心.回读.调用桥接提取", 回读提取)
    入 = {"word": str(原Word), "b表": str(文件["B"]), "金标准": str(金), "b工作表": "word附注"}
    前 = {k: Path(入[k]).read_bytes() for k in ("word", "b表", "金标准")}
    状态 = 运行任务(入, tmp_path / "完整任务", {"测试": True})
    _写证据(tmp_path, 步骤=状态["步骤"], 桥接实际收到标题=标题们)
    assert all(Path(入[k]).read_bytes() == v for k, v in 前.items())
    assert 标题们 and 标题们[0]["text"] == "截至2025年12月31日", 状态["步骤"]


def test_同秒新任务也必须使用不同目录(tmp_path, monkeypatch):
    from fastapi.testclient import TestClient
    from 服务 import web
    class 固定时间:
        @staticmethod
        def now():
            return datetime.datetime(2026, 11, 4, 10, 0, 0)
    class 立即线程:
        def __init__(self, target, **kwargs):
            self.target = target
        def start(self):
            self.target()
        def is_alive(self):
            return False
    def 跑(输入, 目录, 配置, **kwargs):
        Path(目录).mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr(web, "datetime", SimpleNamespace(datetime=固定时间))
    monkeypatch.setattr(web.threading, "Thread", 立即线程)
    monkeypatch.setattr(web, "运行任务", 跑)
    monkeypatch.setattr(web, "运行根", tmp_path / "运行")
    monkeypatch.setattr(web.应用.state, "状态", {"线程": None, "运行目录": None, "日志": []})
    monkeypatch.setattr(web.应用.state, "配置路径", tmp_path / "不存在.json")
    # 直接调用路由，线程替身只固定时间顺序，不启动服务和窗口。
    入 = {"word": "模拟.docx", "b表": "模拟.xlsx", "金标准": "模拟.jsonl"}
    一 = web.开始任务(入)["运行目录"]
    二 = web.开始任务(入)["运行目录"]
    _写证据(tmp_path, 第一个新任务=一, 第二个新任务=二)
    assert 一 != 二


def test_网页显示续跑失败的实际原因(tmp_path):
    # 执行仓库原始前端脚本；只替换DOM和网络边界，不打开浏览器或原生窗口。
    脚本 = r'''
const fs = require("node:fs"), vm = require("node:vm");
const 元素 = {}, 请求 = [], 定时 = [];
function 取(id) { return 元素[id] ??= {value:"",textContent:"", listeners:{},
  addEventListener(n,f) { this.listeners[n]=f; }}; }
const 沙盒 = {document:{getElementById:取,querySelectorAll:()=>[]},
 clearInterval:()=>{},setInterval:(f,n)=>{定时.push(n);return 1;},
 fetch:async (url,opt)=>{请求.push({url,body:opt?.body?JSON.parse(opt.body):null});
  return {status:url==="/api/任务"?400:200,json:async()=>url==="/api/任务"?{detail:"没有可续跑的任务"}:{}};}};
vm.createContext(沙盒);
vm.runInContext(fs.readFileSync("服务/页面/脚本.js","utf8"),沙盒);
(async()=>{for(const id of ["word","b表","金标准"])取(id).value="模拟文件";
 await 取("续跑").listeners.click();
 console.log(JSON.stringify({进度:取("进度").textContent,开始轮询次数:定时.length,请求}));})().catch(e=>{console.error(e);process.exitCode=1;});
'''
    node = shutil.which("node")
    if not node:
        pytest.skip("本机未安装Node.js，前端脚本检查未执行")
    完成 = subprocess.run([node, "-e", 脚本], cwd=根,
                     capture_output=True, encoding="utf-8", timeout=10,
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    assert 完成.returncode == 0, 完成.stderr
    实际 = json.loads(完成.stdout)
    _写证据(tmp_path, **实际)
    assert "没有可续跑的任务" in 实际["进度"], 实际
