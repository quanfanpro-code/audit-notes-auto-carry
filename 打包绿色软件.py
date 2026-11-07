# coding: utf-8
"""制作Windows目录式绿色包；输出必须为新目录，不覆盖已有软件或用户数据。"""
import argparse
import hashlib
import importlib.metadata
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

根 = Path(__file__).resolve().parent

def 执行(命令, 日志, 环境=None):
    with Path(日志).open("w", encoding="utf-8") as f:
        完成 = subprocess.run(命令, cwd=根, env=环境, stdout=f, stderr=subprocess.STDOUT,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if 完成.returncode:
        raise RuntimeError(f"制作未完成（退出码{完成.returncode}），见：{日志}")

def 打包(输出, 工作, 浏览器, dotnet):
    输出, 工作, 浏览器, dotnet = map(lambda p: Path(p).resolve(), (输出, 工作, 浏览器, dotnet))
    if 输出.exists() or 工作.exists():
        raise FileExistsError("输出或构建目录已存在，请指定新目录，避免覆盖")
    if not (浏览器 / "msedgewebview2.exe").is_file():
        raise FileNotFoundError("固定WebView2运行库不完整")
    版本 = "10.0.14"
    运行目录 = [dotnet / "host" / "fxr" / 版本] + [dotnet / "shared" / 名 / 版本
                 for 名 in ("Microsoft.NETCore.App", "Microsoft.WindowsDesktop.App")]
    if not all(p.is_dir() for p in 运行目录):
        raise FileNotFoundError("缺少本轮已验证的.NET 10.0.14运行时，请核对依赖")
    工作.mkdir(parents=True)
    环境 = {**os.environ, "PYINSTALLER_CONFIG_DIR": str(工作 / "PyInstaller缓存"),
            "DOTNET_CLI_HOME": str(工作 / "dotnet缓存"), "DOTNET_CLI_TELEMETRY_OPTOUT": "1",
            "DOTNET_SKIP_FIRST_TIME_EXPERIENCE": "1"}
    桥接 = 工作 / "Word发布"
    执行([str(dotnet / "dotnet.exe"), "publish", str(根 / "Word桥接" / "Word桥接.csproj"),
          "-c", "Release", "-r", "win-x64", "--self-contained", "false", "-o", str(桥接),
          "-p:BaseIntermediateOutputPath=" + str(工作 / "obj") + os.sep,
          "-p:BaseOutputPath=" + str(工作 / "bin") + os.sep], 工作 / "桥接构建.log", 环境)
    命令 = [sys.executable, "-m", "PyInstaller", "--onedir", "--windowed",
          "--name", 输出.name, "--contents-directory", "程序",
          "--distpath", str(输出.parent), "--workpath", str(工作 / "分析"),
          "--specpath", str(工作), "--runtime-hook", str(根 / "绿色运行钩子.py"),
          "--add-data", str(根 / "desktop_ui") + ":desktop_ui",
          "--add-data", str(根 / "服务" / "页面") + ":服务/页面",
          "--add-data", str(桥接) + ":Word桥接/发布",
          "--hidden-import", "核心.Excel交付", "--hidden-import", "pythonnet",
          "--hidden-import", "webview.platforms.edgechromium", "--hidden-import", "webview.platforms.winforms",
          "--collect-submodules", "uvicorn", "--noupx"]
    for 包 in ("PyQt5", "PyQt6", "PySide2", "PySide6", "pytest", "IPython", "matplotlib", "pandas", "numpy"):
        命令 += ["--exclude-module", 包]
    命令.append(str(根 / "启动桌面.pyw"))
    执行(命令, 工作 / "Python构建.log", 环境)
    运行库 = 输出 / "运行库"
    运行库.mkdir()
    shutil.copytree(浏览器, 运行库 / "WebView2")
    dotnet出 = 运行库 / "dotnet"
    for 目录 in 运行目录:
        shutil.copytree(目录, dotnet出 / 目录.relative_to(dotnet))
    for 名 in ("dotnet.exe", "LICENSE.txt", "ThirdPartyNotices.txt"):
        shutil.copy2(dotnet / 名, dotnet出 / 名)
    (运行库 / "桌面.runtimeconfig.json").write_text(json.dumps({"runtimeOptions": {
        "tfm": "net10.0", "frameworks": [{"name": n, "version": 版本} for n in
        ("Microsoft.NETCore.App", "Microsoft.WindowsDesktop.App")] }}, indent=2), encoding="utf-8")
    shutil.copytree(根 / "金标准", 输出 / "金标准")
    for 名 in ("运行", "识别记录", "备份"):
        (输出 / 名).mkdir(exist_ok=True)
    许可 = 输出 / "许可"
    许可.mkdir()
    shutil.copy2(根 / "LICENSE", 许可 / "程序-LICENSE")
    shutil.copy2(Path(sys.executable).parent / "LICENSE.txt", 许可 / "Python-LICENSE.txt")
    shutil.copy2(根 / "requirements.txt", 许可 / "构建依赖.txt")
    包清单 = {}
    for 包 in ("pyinstaller", "pywebview", "pythonnet", "clr_loader", "fastapi", "uvicorn", "httpx",
               "openpyxl", "pywin32", "pydantic", "starlette", "anyio", "certifi"):
        dist = importlib.metadata.distribution(包)
        包清单[包] = dist.version
        for 文件 in dist.files or []:
            if any(x in str(文件).lower() for x in ("license", "copying", "notice")):
                原 = Path(dist.locate_file(文件))
                if 原.is_file():
                    目标 = 许可 / 包 / str(文件).replace("..", "_")
                    目标.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(原, 目标)
    for 包 in ("vue", "element-plus"):
        for 名 in ("LICENSE", "LICENSE.md"):
            原 = 根 / "frontend" / "node_modules" / 包 / 名
            if 原.is_file():
                shutil.copy2(原, 许可 / (包 + "-" + 名))
    提交 = subprocess.check_output(["C:/Program Files/Git/bin/git.exe", "rev-parse", "HEAD"], cwd=根).decode().strip()
    清单 = {"源码基线": 提交, "构建时包含本轮未提交的运行适配": True, "Python": sys.version,
            ".NET": 版本, "依赖": 包清单, "WebView2来源目录": 浏览器.name,
            "平台": "Windows x64", "外部依赖": "本机Word和Excel；识别使用已有模型连接或人工答复"}
    (输出 / "版本清单.json").write_text(json.dumps(清单, ensure_ascii=False, indent=2), encoding="utf-8")
    说明 = f"""# 审计附注搬运绿色版

双击本目录的“{输出.name}.exe”启动。请保留完整文件夹，不能只复制exe。

仅要求本机已经安装Word和Excel；程序所用Python、.NET和浏览器运行库均在本目录内。
模型连接在“设置”填写；没有模型连接时仍可使用既有人工答复流程，不自带离线模型。
模型密钥沿用Windows加密保护，换Windows账户或电脑时需要重新填写。

- 程序、运行库：程序自身依赖。
- 金标准：随包的财务语义标准。
- 识别记录：自动保存已识别的解释。
- 运行：任务材料、结果、日志、临时文件和浏览器缓存。
- 备份：保存配置、答复和覆盖已有成果前的备份。

A/B输入原件可以放在别处，程序不改原件。绿色版的记录目录和人工另存位置请选择软件文件夹内。
软件可在其他目录启动；历史留证仍保存原路径，旧任务路径失效时需重新选择原文件新建任务。
不要在程序运行时移动软件文件夹。更新前保留整个旧文件夹，不覆盖正在使用的数据。
"""
    (输出 / "使用说明.md").write_text(说明, encoding="utf-8-sig")
    文件摘要 = {str(p.relative_to(输出)): hashlib.sha256(p.read_bytes()).hexdigest()
                for p in 输出.rglob("*") if p.is_file()}
    (工作 / "构建文件SHA256.json").write_text(json.dumps(文件摘要, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"输出": str(输出), "文件数": len(文件摘要), "总字节": sum(p.stat().st_size for p in 输出.rglob("*") if p.is_file())}

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--输出", required=True)
    p.add_argument("--工作目录", required=True)
    p.add_argument("--WebView2", required=True)
    p.add_argument("--dotnet", default="C:/Program Files/dotnet")
    a = p.parse_args()
    print(json.dumps(打包(a.输出, a.工作目录, a.WebView2, a.dotnet), ensure_ascii=False))
