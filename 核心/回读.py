# coding: utf-8
"""回读验证：对写回后的 Word 新副本重新提取成 Excel，逐格核对变更清单。
（桥接 read 只返回结构元数据、不含值，因此以重提取为核对依据。）"""
from pathlib import Path

from openpyxl import load_workbook

from .桥接 import 提取 as 调用桥接提取


def 回读验证(word输出, 变更清单: dict, 临时目录) -> list:
    临时目录 = Path(临时目录)
    临时目录.mkdir(parents=True, exist_ok=True)
    回读表 = 临时目录 / "回读表.xlsx"
    回读连 = 临时目录 / "回读连接.docx"
    回读文 = 临时目录 / "回读上下文.json"
    调用桥接提取(str(word输出), str(回读表), str(回读连), str(回读文))
    wb = load_workbook(回读表)
    差异 = []
    try:
        for 格, 期望 in 变更清单.items():
            表名, 坐标 = 格.split("!", 1)
            实际 = wb[表名][坐标].value if 表名 in wb.sheetnames else None
            期望文 = "" if 期望 in (None, "") else str(期望)
            实际文 = "" if 实际 in (None, "") else str(实际)
            if 期望文 != 实际文:
                差异.append({"格": 格, "期望": 期望, "实际": 实际})
    finally:
        wb.close()
    return 差异
