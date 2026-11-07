# coding: utf-8
"""待答复材料的薄接口；保存格式校验继续使用原解析器。"""
import hashlib
import json
import os
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from 运行位置 import 备份目录

from 核心.接续 import 读索引
from 核心.识别 import _解析对象
from .任务服务 import 任务目录, 正在执行, 执行锁, 原子JSON, 拒绝


def 备份已有(路径):
    p = Path(路径)
    if not p.is_file():
        return
    出 = 备份目录() / ("audit-notes-auto-carry_界面保存_" + datetime.now().strftime("%Y%m%d_%H%M%S_%f")) / p.name
    出.parent.mkdir(parents=True, exist_ok=False)
    前 = hashlib.sha256(p.read_bytes()).digest()
    shutil.copy2(p, 出)
    if 前 != hashlib.sha256(出.read_bytes()).digest() or 前 != hashlib.sha256(p.read_bytes()).digest():
        raise OSError("原文件正在变化或备份校验失败，未覆盖")


def 材料列表(目录):
    出 = []
    for 路径 in (读索引(目录) or {}).get("待答复") or []:
        p = Path(路径).resolve()
        if not p.is_relative_to(Path(目录).resolve()) or not (p / "指令.txt").is_file() or not (p / "材料.txt").is_file():
            continue
        身份 = hashlib.sha256(str(p).encode() + (p / "指令.txt").read_bytes() + (p / "材料.txt").read_bytes()).hexdigest()[:24]
        出.append({"id": 身份, "名称": p.name, "路径": str(p), "已有答复": (p / "答复.txt").is_file()})
    return 出


def _材料(目录, 标识):
    条 = next((x for x in 材料列表(目录) if x["id"] == 标识), None)
    if not 条:
        拒绝(404, "这份待答复材料已变化或不属于此任务，请刷新")
    return 条


def 挂接人工接口(应用, 宿主):
    @应用.get("/api/桌面/任务/{task_id}/答复")
    def 列表(task_id: str):
        return {"条目": 材料列表(任务目录(宿主, task_id))}

    @应用.get("/api/桌面/任务/{task_id}/答复/{material_id}")
    def 读答复(task_id: str, material_id: str):
        条 = _材料(任务目录(宿主, task_id), material_id)
        p = Path(条["路径"])
        文 = (p / "答复.txt").read_bytes() if (p / "答复.txt").is_file() else b""
        return {**条, "指令": (p / "指令.txt").read_text(encoding="utf-8-sig"),
                "材料": (p / "材料.txt").read_text(encoding="utf-8-sig"),
                "已有答复": 文.decode("utf-8-sig"), "答复哈希": hashlib.sha256(文).hexdigest()}

    @应用.post("/api/桌面/任务/{task_id}/答复/{material_id}")
    def 保存答复(task_id: str, material_id: str, 请求: dict):
        with 执行锁:
            if 正在执行(宿主):
                拒绝(409, "任务正在处理，不能同时替换答复")
            目录 = 任务目录(宿主, task_id)
            条 = _材料(目录, material_id)
            p = Path(条["路径"]) / "答复.txt"
            新文 = 请求.get("答复")
            if not isinstance(新文, str) or not 新文.strip():
                拒绝(400, "请粘贴模型答复")
            try:
                _解析对象(新文, "人工")
            except RuntimeError as e:
                拒绝(400, str(e), "答复格式错误")
            新值 = 新文.encode("utf-8")
            标识 = str(请求.get("请求标识") or "")
            if not 标识 or len(标识) > 128:
                拒绝(400, "缺少有效请求标识")
            回执路径 = 目录 / "答复提交.json"
            回执 = json.loads(回执路径.read_text(encoding="utf-8")) if 回执路径.is_file() else {}
            内容键 = hashlib.sha256(material_id.encode() + 新值).hexdigest()
            if 标识 in 回执:
                if 回执[标识] != 内容键:
                    拒绝(409, "同一请求标识不能提交不同答复")
                return {"ok": True, "说明": "该次答复已保存"}
            原 = p.read_bytes() if p.is_file() else b""
            if 请求.get("原答复哈希") != hashlib.sha256(原).hexdigest():
                拒绝(409, "已有答复已变化，请重新读取后决定")
            if 原 and 原 != 新值 and not 请求.get("替换确认"):
                拒绝(409, "已有答复将被替换，请先确认")
            if 原 != 新值:
                备份已有(p)
                临时 = p.with_name("答复_" + uuid.uuid4().hex + ".tmp")
                临时.write_bytes(新值)
                os.replace(临时, p)
            回执[标识] = 内容键
            备份已有(回执路径)
            原子JSON(回执路径, 回执)
            return {"ok": True, "说明": "答复已保存；继续任务后由原识别流程核验内容"}
