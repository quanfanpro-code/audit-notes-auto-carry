# coding: utf-8
"""非秘密的界面默认值；不改变已启动任务的原始输入。"""
import json
import os
import sys
from pathlib import Path
from 运行位置 import 软件目录, 检查数据位置
from 核心.共用记录簿 import 默认目录, 已启用

根 = 软件目录()
设置路径 = 根 / "界面设置.json"
基准路径 = 根 / "金标准" / "精简新副本_20261101_010814" / "财务报表附注语义金标准_当前基准_以此为准.jsonl"


def 读取默认():
    默认 = {"金标准": str(基准路径), "复用": True, "识别记录目录": str(默认目录() or ""), "字号": 14}
    if 设置路径.is_file():
        已存 = json.loads(设置路径.read_text(encoding="utf-8-sig"))
        默认.update({k: v for k, v in 已存.items() if k in 默认})
    if getattr(sys, "frozen", False):
        for 名 in ("金标准", "识别记录目录"):
            if 默认.get(名) and not Path(默认[名]).is_absolute():
                默认[名] = str(根 / 默认[名])
        检查数据位置(默认["识别记录目录"])
    return 默认


def 概况补充(宿主):
    默认 = 读取默认()
    锁定 = os.environ.get("审计附注_复用") is not None
    默认["复用"] = 已启用(默认)
    try:
        模型 = 宿主.读设置()
    except (OSError, ValueError):
        模型 = {}
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            Excel = True
    except OSError:
        Excel = False
    from 核心.桥接 import 桥接程序
    return {"默认设置": 默认, "模型": 模型, "复用锁定": 锁定,
            "复用说明": ("环境设置已将复用固定为" + ("开启" if 默认["复用"] else "关闭")) if 锁定 else "按本次选择执行",
            "组件": {"Word桥接文件存在": 桥接程序.is_file(), "Excel已登记": Excel}}


def 挂接设置接口(应用, 宿主):
    from .任务服务 import 原子JSON, 拒绝
    from .人工操作 import 备份已有
    @应用.get("/api/桌面/设置")
    def 读():
        return 概况补充(宿主)

    @应用.post("/api/桌面/设置")
    def 存(请求: dict):
        默认 = 请求.get("默认设置") or {}
        旧 = 读取默认()
        新 = {**旧, **{k: v for k, v in 默认.items() if k in 旧}}
        if 新.get("字号") not in (14, 15, 16, 18):
            拒绝(400, "字号请选择14、15、16或18")
        if not Path(新.get("金标准") or "").is_file():
            拒绝(400, "金标准文件不存在，请重新选择")
        目录 = 新.get("识别记录目录")
        if 目录 and not Path(目录).is_absolute():
            拒绝(400, "请用选择窗口选择识别记录目录")
        try:
            if 目录:
                检查数据位置(目录)
        except ValueError as e:
            拒绝(400, str(e))
        if getattr(sys, "frozen", False):
            for 名 in ("金标准", "识别记录目录"):
                if 新.get(名) and Path(新[名]).resolve().is_relative_to(根):
                    新[名] = str(Path(新[名]).resolve().relative_to(根))
        宿主.存设置(请求.get("模型") or {})
        备份已有(设置路径)
        原子JSON(设置路径, 新)
        return {"ok": True}
