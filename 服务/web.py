# coding: utf-8
"""本地 WebUI（FastAPI）：任务、结果、设置、人工回写四区；仅监听 127.0.0.1。"""
import datetime
import json
import os
import subprocess
import threading
from pathlib import Path

import httpx
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, PlainTextResponse, JSONResponse

from 核心.流程 import 运行任务, 人工回写
from 核心.接续 import 读索引, 写索引
from 核心.文件输入 import 列出工作表, 规范化输入
from 核心.设置 import 保存模型配置, 读取模型配置, 默认配置路径

页面目录 = Path(__file__).resolve().parent / "页面"
运行根 = Path(__file__).resolve().parent.parent / "运行"

应用 = FastAPI(title="审计附注自动搬运系统")
应用.state.配置路径 = 默认配置路径()
应用.state.状态 = {"线程": None, "运行目录": None, "日志": []}


@应用.middleware("http")
async def 会话保护(request: Request, call_next):
    令牌 = getattr(应用.state, "会话令牌", None)
    if 令牌 and request.method not in ("GET", "HEAD", "OPTIONS"):
        if request.headers.get("X-Audit-Token") != 令牌:
            return JSONResponse({"detail": {"类别": "会话已失效", "说明": "请重新打开本机程序页面"}}, status_code=403)
    return await call_next(request)


def _带会话(文):
    令牌 = json.dumps(getattr(应用.state, "会话令牌", "") or "")
    return 文.replace("<head>", "<head><script>window.__AUDIT_TOKEN__=" + 令牌 + ";</script>", 1)


@应用.get("/桌面")
def 桌面页():
    p = Path(getattr(应用.state, "桌面页面", 页面目录.parent.parent / "desktop_ui" / "index.html"))
    if not p.is_file():
        return HTMLResponse("桌面页面缺失，请取得完整程序目录。", status_code=503)
    return HTMLResponse(_带会话(p.read_text(encoding="utf-8")))


def _选文件(类型="文件"):
    原生选择 = getattr(应用.state, "文件选择器", None)
    if 原生选择:
        return 原生选择(类型)
    """后端弹 Windows 原生选择窗口（浏览器文件框拿不到真实路径）。"""
    if 类型 == "文件夹":
        脚本 = ("Add-Type -AssemblyName System.Windows.Forms; "
              "$d=New-Object System.Windows.Forms.FolderBrowserDialog; "
              "if($d.ShowDialog() -eq 'OK'){Write-Output $d.SelectedPath}")
    else:
        类 = "SaveFileDialog" if 类型 == "另存为Word" else "OpenFileDialog"
        筛选 = {"报表": "Word或Excel|*.docx;*.xlsx", "Word": "Word|*.docx", "Excel": "Excel|*.xlsx",
                "金标准": "金标准|*.jsonl;*.json", "结构操作": "结构操作|*.json", "另存为Word": "Word|*.docx"}.get(类型, "所有文件|*.*")
        脚本 = ("Add-Type -AssemblyName System.Windows.Forms; "
              f"$d=New-Object System.Windows.Forms.{类}; $d.Filter='{筛选}'; "
              "if($d.ShowDialog() -eq 'OK'){Write-Output $d.FileName}")
    脚本 = "[Console]::OutputEncoding=[Text.Encoding]::UTF8; " + 脚本
    程序 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    完成 = subprocess.run([str(程序), "-NoProfile", "-STA", "-Command", 脚本],
                     capture_output=True, text=True, encoding="utf-8",
                     creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if 完成.returncode:
        raise RuntimeError("Windows选择窗口未能打开：" + 完成.stderr.strip())
    return 完成.stdout.strip()


@应用.get("/")
def 首页():
    return HTMLResponse(_带会话((页面目录 / "index.html").read_text(encoding="utf-8")))


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
    import sys
    from .任务服务 import 接受任务, 目录列表
    宿主 = sys.modules[__name__]
    续跑目录 = None
    if 请求.get("续跑"):
        续跑目录 = 应用.state.状态.get("运行目录")
        if not 续跑目录 or not Path(续跑目录).is_dir():
            候选 = 目录列表(宿主)
            if not 候选:
                raise HTTPException(400, "没有可续跑的任务")
            续跑目录 = 候选[0]
    try:
        return 接受任务(宿主, 请求, 续跑目录=续跑目录)
    except HTTPException as e:
        if e.status_code == 409 and "已有任务" in str(e.detail):
            return {"错误": "已有任务在运行"}
        raise


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
    import sys
    from .任务服务 import 中断任务
    目录 = 应用.state.状态.get("运行目录")
    if 目录:
        中断任务(sys.modules[__name__], 目录)
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
    配置["密钥已保存"] = bool(配置.get("api_key"))
    配置.pop("api_key", None)                 # 不回传密钥
    return 配置


@应用.post("/api/设置")
def 存设置(请求: dict):
    路径 = 应用.state.配置路径
    旧 = 读取模型配置(路径) if Path(路径).is_file() else {}
    合并 = {**旧, **{k: v for k, v in 请求.items() if v not in ("", None)}}
    合并.setdefault("api_key", "")
    合并.pop("密钥已保存", None)
    from .人工操作 import 备份已有
    备份已有(路径)
    保存模型配置(路径, 合并)
    return {"ok": True}


@应用.post("/api/设置/测试")
def 设置测试():
    if not Path(应用.state.配置路径).is_file():
        return {"ok": False, "说明": "还没有保存模型连接"}
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
    import sys
    from .任务服务 import 接受任务
    try:
        返回 = 接受任务(sys.modules[__name__], 请求, 类型="人工回写", 后台=False)
        if 返回.get("错误"):
            return {"ok": False, "错误": 返回["错误"]}
        return {"ok": True, "结果": 返回.get("结果")}
    except Exception as e:
        return {"ok": False, "错误": str(e)}


import sys
from .任务服务 import 挂接任务接口
挂接任务接口(应用, sys.modules[__name__])
from .结果展示 import 挂接结果接口
挂接结果接口(应用, sys.modules[__name__])
from .人工操作 import 挂接人工接口
挂接人工接口(应用, sys.modules[__name__])
from .记录展示 import 挂接记录接口
挂接记录接口(应用, sys.modules[__name__])
from .界面设置 import 挂接设置接口
挂接设置接口(应用, sys.modules[__name__])
