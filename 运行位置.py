# coding: utf-8
"""只统一运行位置；冻结版所有自有数据留在软件目录。"""
import sys
from pathlib import Path

def 软件目录():
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

def 备份目录():
    return 软件目录() / "备份" if getattr(sys, "frozen", False) else Path(r"C:\Users\27651\BackUp")

def 检查数据位置(路径):
    p = Path(路径).resolve()
    if getattr(sys, "frozen", False) and not p.is_relative_to(软件目录()):
        raise ValueError("绿色版产生的数据必须保存在软件文件夹内，请重新选择位置")
    return p
