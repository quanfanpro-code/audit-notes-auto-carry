# coding: utf-8
"""回读验证：对写回后的 Word 新副本重新提取成 Excel，逐格核对变更清单。
（桥接 read 只返回结构元数据、不含值，因此以重提取为核对依据。）"""
from pathlib import Path

from openpyxl import load_workbook

from .桥接 import 提取 as 调用桥接提取


def _小数位(值) -> int:
    if isinstance(值, int) and not isinstance(值, bool):
        return 0
    文 = repr(float(值))
    if "e" in 文 or "E" in 文:
        return 15
    return len(文.split(".")[1]) if "." in 文 else 0


def _按显示精度一致(期望, 实际) -> bool:
    """Word 单元格按自己的格式显示，回读只能比到它显示出来的精度。

    只有两边都是数字才按精度比。文字格（编号、账号）逐字比：
    按数字比会把 "00123" 和 "123"、"1.00" 和 "1" 判成一致，文本被改坏就漏报了。
    """
    if isinstance(期望, bool) or isinstance(实际, bool):
        return False
    if isinstance(期望, (int, float)) and isinstance(实际, (int, float)):
        甲, 乙 = float(期望), float(实际)
        if 甲 == 乙:
            return True
        return abs(甲 - 乙) <= 0.5 * 10 ** (-_小数位(实际)) + 1e-9
    return False


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
            表名, 坐标 = 格.rsplit("!", 1)
            实际 = wb[表名][坐标].value if 表名 in wb.sheetnames else None
            期望文 = "" if 期望 in (None, "") else str(期望)
            实际文 = "" if 实际 in (None, "") else str(实际)
            if 期望文 != 实际文 and not _按显示精度一致(期望, 实际):
                差异.append({"格": 格, "期望": 期望, "实际": 实际})
    finally:
        wb.close()
    return 差异
