# coding: utf-8
"""先于第三方运行钩子设置包内缓存，避免首次启动写进系统临时目录。"""
import os
import sys
import tempfile
from pathlib import Path

根 = Path(sys.executable).resolve().parent
临时 = 根 / "运行" / "临时"
临时.mkdir(parents=True, exist_ok=True)
os.environ["TEMP"] = os.environ["TMP"] = str(临时)
tempfile.tempdir = str(临时)
sys.dont_write_bytecode = True
os.environ["DOTNET_ROOT"] = os.environ["DOTNET_ROOT_X64"] = str(根 / "运行库" / "dotnet")
os.environ["DOTNET_MULTILEVEL_LOOKUP"] = "0"
os.environ["审计附注_备份目录"] = str(根 / "备份")
os.environ["审计附注_识别记录目录"] = str(根 / "识别记录")
