# coding: utf-8
"""本地 WebUI（FastAPI）：任务、结果、设置、人工回写四区；仅监听 127.0.0.1。"""
import datetime
import json
import subprocess
import threading
from pathlib import Path

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, PlainTextResponse

from 核心.流程 import 运行任务, 人工回写
from 核心.接续 import 读索引
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


@应用.post("/api/任务")
def 开始任务(请求: dict):
    缺 = [k for k in ("word", "b表", "金标准") if not 请求.get(k)]
    if 缺:
        raise HTTPException(status_code=400, detail="缺少字段：" + "、".join(缺))
    if 应用.state.状态.get("线程") and 应用.state.状态["线程"].is_alive():
        return {"错误": "已有任务在运行"}
    运行目录 = 运行根 / ("任务_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    配置 = 读取模型配置(应用.state.配置路径) if Path(应用.state.配置路径).is_file() else {}

    def 跑():
        try:
            运行任务(请求, 运行目录, 配置, 取消=应用.state.状态.get("取消"))
        except Exception as e:                       # 失败如实进日志，不吞
            应用.state.状态["日志"].append("失败：" + str(e))

    取消 = threading.Event()
    线程 = threading.Thread(target=跑, daemon=True)
    应用.state.状态.update({"线程": 线程, "运行目录": 运行目录, "日志": [], "取消": 取消})
    线程.start()
    return {"运行目录": str(运行目录)}


@应用.get("/api/进度")
def 进度():
    运行目录 = 应用.state.状态.get("运行目录")
    索引 = (读索引(运行目录) if 运行目录 else None) or {}
    return {"步骤": 索引.get("步骤"),
            "已完成范围": len(索引.get("已完成范围", [])),
            "待答复": 索引.get("待答复", []),          # 卡在哪份材料，要谁去答
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

    return {"对应": 读("对应清单.jsonl"), "问题": 读("问题清单.jsonl"), "回读": 读("回读记录.jsonl")}


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
    try:
        响应 = httpx.get(配置["url"].rstrip("/") + "/models", timeout=15,
                   headers={"Authorization": "Bearer " + 配置.get("api_key", "")})
        return {"ok": 响应.status_code < 500, "状态码": 响应.status_code}
    except Exception as e:
        return {"ok": False, "错误": str(e)}


@应用.post("/api/人工回写")
def 人工回写入口(请求: dict):
    try:
        return {"ok": True, "结果": 人工回写(请求)}
    except Exception as e:
        return {"ok": False, "错误": str(e)}
