# coding: utf-8
"""本地 WebUI（FastAPI）：任务、结果、设置、人工回写四区；仅监听 127.0.0.1。"""
import datetime
import json
import os
import subprocess
import threading
from pathlib import Path

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, PlainTextResponse

from 核心.流程 import 运行任务, 人工回写
from 核心.接续 import 读索引, 写索引
from 核心.文件输入 import 列出工作表, 规范化输入
from 核心.设置 import 保存模型配置, 读取模型配置, 默认配置路径

页面目录 = Path(__file__).resolve().parent / "页面"
运行根 = Path(__file__).resolve().parent.parent / "运行"

应用 = FastAPI(title="审计附注自动搬运系统")
应用.state.配置路径 = 默认配置路径()
应用.state.状态 = {"线程": None, "运行目录": None, "日志": []}


def _选文件(类型="文件"):
    """后端弹 Windows 原生选择窗口（浏览器文件框拿不到真实路径）。"""
    if 类型 == "文件夹":
        脚本 = ("Add-Type -AssemblyName System.Windows.Forms; "
              "$d=New-Object System.Windows.Forms.FolderBrowserDialog; "
              "if($d.ShowDialog() -eq 'OK'){Write-Output $d.SelectedPath}")
    else:
        脚本 = ("Add-Type -AssemblyName System.Windows.Forms; "
              "$d=New-Object System.Windows.Forms.OpenFileDialog; "
              "if($d.ShowDialog() -eq 'OK'){Write-Output $d.FileName}")
    完成 = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", 脚本],
                     capture_output=True, text=True, encoding="utf-8",
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return 完成.stdout.strip()


@应用.get("/")
def 首页():
    return HTMLResponse((页面目录 / "index.html").read_text(encoding="utf-8"))


@应用.get("/static")
def 静态(name: str = ""):
    if name not in ("样式.css", "脚本.js"):
        return PlainTextResponse("未找到", status_code=404)
    类型 = "text/css; charset=utf-8" if name.endswith(".css") else "application/javascript; charset=utf-8"
    return PlainTextResponse((页面目录 / name).read_text(encoding="utf-8"), media_type=类型)


@应用.post("/api/选文件")
def 选文件(请求: dict):
    return {"路径": _选文件(请求.get("类型", "文件"))}


@应用.post("/api/工作表")
def 工作表(请求: dict):
    try:
        return {"工作表": 列出工作表(请求.get("路径") or "")}
    except (ValueError, OSError) as e:
        raise HTTPException(status_code=400, detail=str(e)) from e


@应用.post("/api/任务")
def 开始任务(请求: dict):
    if "A" in 请求 or "B" in 请求:
        缺 = [k for k in ("A", "B", "金标准") if not 请求.get(k)]
        if 缺:
            raise HTTPException(status_code=400, detail="缺少字段：" + "、".join(缺))
        try:
            规范化输入(请求)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e)) from e
    else:
        缺 = [k for k in ("word", "b表", "金标准") if not 请求.get(k)]
        if 缺:
            raise HTTPException(status_code=400, detail="缺少字段：" + "、".join(缺))
    线程 = 应用.state.状态.get("线程")
    if 线程 and 线程.is_alive():
        当前 = 读索引(应用.state.状态.get("运行目录")) or {}
        终态 = {"done", "公式待核", "部分保留", "回读不一致", "识别失败",
               "Excel交付失败", "任务失败", "用户中断"}
        if 当前.get("步骤") not in 终态:
            return {"错误": "已有任务在运行"}
    if 请求.get("续跑"):
        上次 = 应用.state.状态.get("运行目录")
        if 上次 and Path(上次).is_dir():
            运行目录 = Path(上次)
        else:
            候选 = sorted(p for p in 运行根.glob("任务_*") if p.is_dir())
            if not 候选:
                raise HTTPException(status_code=400, detail="没有可续跑的任务")
            运行目录 = 候选[-1]            # 续跑：沿用原来的目录，不新建
    else:
        基 = 运行根 / ("任务_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        运行目录, 序 = 基, 1
        while 运行目录.exists():          # 同一秒连开两个任务，目录名要避让，不能覆盖
            序 += 1
            运行目录 = 基.with_name(f"{基.name}_{序}")
    try:
        配置 = 读取模型配置(应用.state.配置路径)
    except Exception:
        配置 = {}                                  # 没配过或配置坏了，都按“没配”处理

    def 跑():
        try:
            运行任务(请求, 运行目录, 配置, 取消=应用.state.状态.get("取消"))
        except Exception as e:
            应用.state.状态["日志"].append("失败：" + str(e))
            运行目录.mkdir(parents=True, exist_ok=True)
            索引 = 读索引(运行目录) or {"已完成范围": [], "产物": {}}
            索引.update({"步骤": "任务失败", "错误": str(e)})
            写索引(运行目录, 索引)

    运行目录.mkdir(parents=True, exist_ok=True)
    取消 = threading.Event()
    线程 = threading.Thread(target=跑, daemon=True)
    应用.state.状态.update({"线程": 线程, "运行目录": 运行目录, "日志": [], "取消": 取消})
    线程.start()
    return {"运行目录": str(运行目录)}


def _复用统计(运行目录):
    出 = []
    for 侧 in ("A", "B"):
        路径 = Path(运行目录) / f"复用记录_{侧}.json"
        if not 路径.is_file():
            continue
        记录 = json.loads(路径.read_text(encoding="utf-8"))
        for 部分 in 记录.get("部分") or []:
            出.append({"侧": 侧, "读取部分": 部分.get("读取部分"),
                      "工作表": 部分.get("工作表"), "材料类型": 部分.get("材料类型"),
                      **(部分.get("调用") or {})})
    return 出


@应用.get("/api/进度")
def 进度():
    运行目录 = 应用.state.状态.get("运行目录")
    索引 = (读索引(运行目录) if 运行目录 else None) or {}
    return {"步骤": 索引.get("步骤"),
            "已完成范围": len(索引.get("已完成范围", [])),
            "待答复": 索引.get("待答复", []),
            "复用统计": _复用统计(运行目录) if 运行目录 else [],
            "运行目录": str(运行目录) if 运行目录 else None,
            "日志": 应用.state.状态.get("日志", [])[-20:]}


@应用.post("/api/中断")
def 中断():
    事件 = 应用.state.状态.get("取消")
    if 事件:
        事件.set()
    应用.state.状态["日志"].append("已请求中断（当前这张表做完后不再继续）")
    return {"ok": True}


@应用.get("/api/结果")
def 结果():
    运行目录 = 应用.state.状态.get("运行目录")
    if not 运行目录:
        return {"错误": "还没有任务"}

    def 读(名):
        p = Path(运行目录) / 名
        return [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.is_file() else []

    对应 = 读("对应清单.jsonl")
    原因 = {}
    for x in 对应:
        if x.get("状态") == "对应":
            continue
        名 = x.get("原因") or "未分类"
        原因[名] = 原因.get(名, 0) + 1
    索引 = 读索引(运行目录) or {}
    产物 = 索引.get("产物") or {}
    return {"对应": 对应, "问题": 读("问题清单.jsonl"), "回读": 读("回读记录.jsonl"),
            "公式核对": 读("公式核对记录.jsonl"), "公式待核数": 索引.get("公式待核数", 0),
            "复用统计": _复用统计(运行目录), "桥接提示": 索引.get("桥接提示") or [],
            "文件类型": ("excel" if 产物.get("Excel新副本") else
                       "word" if 产物.get("Word新副本") else None),
            "摘要": {"对应条数": sum(1 for x in 对应 if x.get("状态") == "对应"),
                    "清尾条数": sum(1 for x in 对应 if x.get("状态") == "清尾"),
                    "未对应": 原因,
                    "回读差异": 索引.get("回读差异", 0),
                    "公式待核数": 索引.get("公式待核数", 0),
                    "步骤": 索引.get("步骤")},
            "产物": {"运行目录": str(运行目录),
                    "更新后附注表格": 产物.get("更新后附注表格"),
                    "A端新副本": 产物.get("A端新副本"),
                    "Excel新副本": 产物.get("Excel新副本"),
                    "Word新副本": 产物.get("Word新副本")}}


@应用.post("/api/打开目录")
def 打开目录(请求: dict):
    目标 = Path(str(请求.get("路径") or 应用.state.状态.get("运行目录") or ""))
    if not 目标.exists():
        return {"ok": False, "错误": "目录不存在"}
    路径 = 目标 if 目标.is_dir() else 目标.parent
    try:
        os.startfile(str(路径))            # 只面向 Windows
        return {"ok": True, "路径": str(路径)}
    except Exception as e:
        return {"ok": False, "错误": str(e)}


@应用.get("/api/设置")
def 读设置():
    路径 = 应用.state.配置路径
    配置 = 读取模型配置(路径) if Path(路径).is_file() else {}
    配置.pop("api_key", None)                 # 不回传密钥
    return 配置


@应用.post("/api/设置")
def 存设置(请求: dict):
    路径 = 应用.state.配置路径
    旧 = 读取模型配置(路径) if Path(路径).is_file() else {}
    合并 = {**旧, **{k: v for k, v in 请求.items() if v not in ("", None)}}
    合并.setdefault("api_key", "")
    保存模型配置(路径, 合并)
    return {"ok": True}


@应用.post("/api/设置/测试")
def 设置测试():
    配置 = 读取模型配置(应用.state.配置路径)
    if not 配置.get("url"):
        return {"ok": False, "说明": "还没填 API 地址"}
    try:
        响应 = httpx.get(配置["url"].rstrip("/") + "/models", timeout=15,
                   headers={"Authorization": "Bearer " + 配置.get("api_key", "")})
        码 = 响应.status_code
        if 200 <= 码 < 300:
            return {"ok": True, "状态码": 码, "说明": "连通"}
        if 码 in (401, 403):
            return {"ok": False, "状态码": 码, "说明": "密钥无效或没有权限"}
        if 码 == 404:
            return {"ok": False, "状态码": 码, "说明": "地址不对，找不到这个接口"}
        return {"ok": False, "状态码": 码, "说明": "服务端返回错误"}
    except Exception as e:
        return {"ok": False, "错误": str(e), "说明": "连不上"}


@应用.post("/api/人工回写")
def 人工回写入口(请求: dict):
    try:
        return {"ok": True, "结果": 人工回写(请求)}
    except Exception as e:
        return {"ok": False, "错误": str(e)}
