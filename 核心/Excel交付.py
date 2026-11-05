# coding: utf-8
import ast
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.formula.tokenizer import Tokenizer

from .表格 import 取可见集, 读快照
from .搬运 import _是错误值


def 是纯数字算式(公式: str) -> bool:
    if not isinstance(公式, str) or not 公式.startswith("="):
        return False
    try:
        标记 = Tokenizer(公式).items
        有数字 = False
        for t in 标记:
            if t.type == "OPERAND" and t.subtype == "NUMBER":
                有数字 = True
            elif t.type == "OPERATOR-INFIX" and t.value in ("+", "-", "*", "/"):
                continue
            elif t.type == "OPERATOR-PREFIX" and t.value in ("+", "-"):
                continue
            elif t.type == "PAREN" or t.type == "WHITE-SPACE":
                continue
            else:
                return False
        return 有数字 and ast.parse(公式[1:].strip(), mode="eval") is not None
    except (SyntaxError, ValueError, TypeError):
        return False


def _快照格(快照):
    return {表["name"]: 表.get("cells") or {} for 表 in 快照.get("sheets") or []}


def _拆(格):
    return str(格).rsplit("!", 1)


def _组表(A结构):
    出 = []
    for 记录 in A结构 or []:
        if not 记录.get("工作表") or not 记录.get("范围"):
            continue
        出.append((记录["工作表"], range_boundaries(str(记录["范围"])), 记录.get("表组") or ""))
    return 出


def _所属组(格, 组表):
    表, 坐标 = _拆(格)
    行, 列 = coordinate_to_tuple(坐标)
    命中 = [组 for 名, (左, 顶, 右, 底), 组 in 组表
            if 名 == 表 and 左 <= 列 <= 右 and 顶 <= 行 <= 底]
    if len(命中) != 1:
        raise ValueError(f"拟写格不能唯一归入业务表组：{格}")
    return 命中[0]


def _写入(cell, 值):
    if 值 is None or 值 == "":
        cell.ClearContents()
    elif isinstance(值, str):
        cell.Value2 = "'" + 值          # 文字一律按文本落格，避免 Excel 自动转数字或日期
    else:
        cell.Value2 = 值


def _错误文本(格, 组, 原因):
    return {"原因": 原因, "a格": 格, "表组": 组}


def _实际交付(请求):
    输出 = Path(请求["输出"])
    shutil.copy2(请求["原件"], 输出)
    变更 = dict(请求.get("变更") or {})
    标题 = dict(请求.get("标题变更") or {})
    快照格 = _快照格(请求["A快照"])
    组表 = _组表(请求.get("A结构"))
    # 标题与普通格共用同一份本次可见集合；取不到集合就不放行，不能当可见。
    可见们 = {表["name"]: 取可见集(请求["A快照"], 表["name"])
            for 表 in 请求["A快照"].get("sheets") or []}
    保留 = {x["a格"] for x in 请求.get("公式清单") or []}
    for 格 in list(变更):
        表, 坐标 = _拆(格)
        原 = (快照格.get(表) or {}).get(坐标) or {}
        if 原.get("formula") and (格 in 保留 or not 是纯数字算式(原["formula"])):
            变更.pop(格)

    目标组 = {}
    for 格 in list(变更) + list(标题):
        目标组[格] = (标题[格].get("表组") if 格 in 标题 and 标题[格].get("表组")
                     else _所属组(格, 组表))
    问题, 撤回 = [], set()
    app = wb = None
    状态路径 = Path(请求["状态"])
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        状态路径.write_text(json.dumps({"python": os.getpid(), "excel_hwnd": app.Hwnd}), encoding="utf-8")
        wb = app.Workbooks.Open(str(输出), UpdateLinks=0, ReadOnly=False,
                               IgnoreReadOnlyRecommended=True)
        if wb.ReadOnly:
            raise RuntimeError("Excel候选副本以只读方式打开，不能写入")

        for 格 in list(变更) + list(标题):
            组 = 目标组[格]
            if 组 in 撤回:
                continue
            表, 坐标 = _拆(格)
            集 = 可见们.get(表)
            if 集 is None or not 集.含(*coordinate_to_tuple(坐标)):
                撤回.add(组)
                问题.append(_错误文本(
                    格, 组, "缺少可见范围依据或格不在本次可见范围，已保留该组"))
                continue
            cell = wb.Worksheets(表).Range(坐标)
            数组合格 = getattr(cell, "HasArray", False)
            合并从属 = False
            if cell.MergeCells:
                合并从属 = cell.MergeArea.Cells(1, 1).Address != cell.Address
            溢出 = bool(getattr(cell, "HasSpill", False))
            保护 = bool(wb.Worksheets(表).ProtectContents and cell.Locked)
            标题因 = ""
            if 格 in 标题:
                原 = (快照格.get(表) or {}).get(坐标) or {}
                当前 = 原.get("cached_value") if 原.get("formula") else 原.get("value")
                if 原.get("formula"):
                    标题因 = "标题格是保留公式，不能替换显示文字，已保留该组"
                elif 当前 != 标题[格].get("原文"):
                    标题因 = "标题格本次原文与识别依据不符，已保留该组"
            if 数组合格 or 合并从属 or 溢出 or 保护 or 标题因:
                撤回.add(组)
                问题.append(_错误文本(格, 组, 标题因 or "目标格受保护、为合并从属格或数组/溢出范围，已保留该组"))

        原变更 = dict(变更)
        原标题 = dict(标题)
        for 格 in list(变更) + list(标题):
            组 = 目标组[格]
            if 组 in 撤回:
                continue
            表, 坐标 = _拆(格)
            try:
                值 = 标题[格]["新文"] if 格 in 标题 else 变更[格]
                _写入(wb.Worksheets(表).Range(坐标), 值)
            except Exception as e:
                try:
                    for 原格 in list(原变更) + list(原标题):
                        if 目标组[原格] != 组:
                            continue
                        表2, 坐标2 = _拆(原格)
                        原 = (快照格.get(表2) or {}).get(坐标2) or {}
                        格对象 = wb.Worksheets(表2).Range(坐标2)
                        if 原.get("formula"):
                            格对象.Formula = 原["formula"]
                        else:
                            _写入(格对象, 原.get("value"))
                except Exception as 回退:
                    raise RuntimeError(f"{格}写入失败且该组无法完整撤回：{回退}") from e
                撤回.add(组)
                问题.append(_错误文本(格, 组, f"写入失败，已恢复该组原内容：{e}"))

        app.CalculateFullRebuild()
        截止 = time.monotonic() + 280
        while app.CalculationState != 0:
            if time.monotonic() > 截止:
                raise TimeoutError("Excel计算未完成")
            time.sleep(0.1)
        wb.Save()
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        finally:
            if app is not None:
                app.Quit()

    if 撤回:
        变更 = {格: 值 for 格, 值 in 变更.items() if 目标组[格] not in 撤回}
        标题 = {格: 值 for 格, 值 in 标题.items() if 目标组[格] not in 撤回}
    公式 = load_workbook(输出, data_only=False)
    数值 = load_workbook(输出, data_only=True)
    差异 = []
    try:
        for 格, 期望 in 变更.items():
            表, 坐标 = _拆(格)
            实际 = 数值[表][坐标].value
            if 期望 in (None, ""):
                一致 = 实际 in (None, "")
            else:
                一致 = 实际 == 期望
            if not 一致:
                差异.append({"格": 格, "期望": 期望, "实际": 实际})
        for 格, 项 in 标题.items():
            表, 坐标 = _拆(格)
            实际 = 公式[表][坐标].value
            if 实际 != 项["新文"]:
                差异.append({"格": 格, "期望": 项["新文"], "实际": 实际})
        for 项 in 请求.get("公式清单") or []:
            表, 坐标 = _拆(项["a格"])
            实际 = 公式[表][坐标].value
            if 实际 != 项["原公式"]:
                差异.append({"格": 项["a格"], "期望": 项["原公式"], "实际": 实际})
    finally:
        公式.close()
        数值.close()

    import math
    最终格 = _快照格(读快照(输出))
    核对 = []
    for 项 in 请求.get("公式清单") or []:
        表, 坐标 = _拆(项["a格"])
        A结果 = ((最终格.get(表) or {}).get(坐标) or {}).get("cached_value")
        B金额 = 项.get("B原值")
        记录 = {"a格": 项["a格"], "表组": 项.get("表组"), "原公式": 项["原公式"],
                "A结果": A结果, "b格": 项.get("b格"), "B金额": B金额,
                "差额": None, "结论": "未能核对", "原因": ""}
        if 项.get("表组") in 撤回:
            记录["原因"] = "所在组写入未完成，未按本次搬运结果核对"
        elif 项.get("不能核对原因"):
            记录["原因"] = 项["不能核对原因"]
        elif 请求["A快照"].get("external_references"):
            记录["原因"] = "工作簿含外部依赖，未能证明本公式独立"
        elif 项.get("B原值状态") != "值" or not isinstance(B金额, (int, float)) \
                or isinstance(B金额, bool) or not math.isfinite(B金额):
            记录["原因"] = "B金额不可用：" + str(项.get("B原值状态") or "非有限数值")
        elif _是错误值(A结果):
            记录["结论"] = "计算错误"
            记录["原因"] = "Excel公式计算错误"
        elif A结果 == "":
            记录["原因"] = "公式返回空字符串，不能作为金额比较"
        elif not isinstance(A结果, (int, float)) or isinstance(A结果, bool) or not math.isfinite(A结果):
            记录["原因"] = "没有本次保存后的有限数值结果"
        else:
            记录["差额"] = A结果 - B金额
            if abs(记录["差额"]) <= 1e-9:
                记录["结论"] = "本次金额一致"
            else:
                记录["结论"] = "金额不一致"
                记录["原因"] = "本次金额不一致"
        核对.append(记录)
        if 记录["结论"] != "本次金额一致":
            问题.append({"原因": f"公式{记录['结论']}：{记录['原因']}",
                        "a格": 项["a格"], "表组": 项.get("表组")})
    待核 = sum(x["结论"] != "本次金额一致" for x in 核对)
    return {"输出": str(输出), "回读差异": 差异, "公式核对": 核对, "公式待核数": 待核,
            "问题": 问题, "撤回组": sorted(撤回)}


def 收集保留公式(A快照, A结构, 对应清单, B合集):
    对应 = {x.get("a格"): x for x in 对应清单 if x.get("状态") == "对应"}
    b按格 = {x.get("格"): x for x in B合集}
    格表 = _快照格(A快照)
    出, 见 = [], set()
    for 记录 in A结构:
        表 = 记录.get("工作表")
        if not 表 or not 记录.get("范围"):
            continue
        左, 顶, 右, 底 = range_boundaries(str(记录["范围"]))
        可见 = 取可见集(A快照, 表)
        for 坐标, 格 in (格表.get(表) or {}).items():
            if not (左 <= 格["column"] <= 右 and 顶 <= 格["row"] <= 底):
                continue
            if 可见 is not None and not 可见.含(格["row"], 格["column"]):
                continue
            公式 = 格.get("formula")
            键 = f"{表}!{坐标}"
            if not 公式 or 是纯数字算式(公式) or 键 in 见:
                continue
            见.add(键)
            条 = 对应.get(键) or {}
            乙 = b按格.get(条.get("b格")) or {}
            出.append({"a格": 键, "表组": 记录.get("表组"), "原公式": 公式,
                       "b格": 条.get("b格"), "B原值": 乙.get("原值"),
                       "B原值状态": 乙.get("原值状态"),
                       "不能核对原因": "" if 条 else "没有对应金额，未能核对"})
    return 出


def _终止(worker, 状态路径):
    try:
        状态 = json.loads(状态路径.read_text(encoding="utf-8")) if 状态路径.is_file() else {}
    except Exception:
        状态 = {}
    worker.terminate()
    try:
        worker.wait(timeout=10)
    except subprocess.TimeoutExpired:
        worker.kill()
    # Excel由COM宿主启动，父进程只能按本次登记的窗口进程号结束自己的实例。
    import ctypes
    hwnd = 状态.get("excel_hwnd")
    if hwnd:
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
        if pid.value:
            subprocess.run(["taskkill", "/PID", str(pid.value), "/F"],
                           capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def 交付Excel(原件, 输出, 变更: dict, A快照: dict, A结构: list, 公式清单: list,
             标题变更: dict, 取消=None) -> dict:
    输出 = Path(输出)
    输出.parent.mkdir(parents=True, exist_ok=True)
    请求路径 = 输出.parent / f"Excel请求_{int(time.time() * 1000)}.json"
    结果路径 = 请求路径.with_name("Excel结果.json")
    状态路径 = 请求路径.with_name("Excel实例.json")
    日志路径 = 请求路径.with_name("Excel工作进程.log")
    请求 = {"原件": str(原件), "输出": str(输出), "变更": 变更, "A快照": A快照,
            "A结构": A结构, "公式清单": 公式清单, "标题变更": 标题变更,
            "状态": str(状态路径)}
    请求路径.write_text(json.dumps(请求, ensure_ascii=False, default=str), encoding="utf-8")
    with 日志路径.open("w", encoding="utf-8") as 日志:
        worker = subprocess.Popen([sys.executable, "-m", "核心.Excel交付", "--worker", str(请求路径)],
                                  cwd=Path(__file__).resolve().parent.parent,
                                  stdout=日志, stderr=subprocess.STDOUT,
                                  creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    截止 = time.monotonic() + 300
    已终止 = False
    try:
        while worker.poll() is None:
            if 取消 is not None and 取消.is_set():
                _终止(worker, 状态路径)
                已终止 = True
                raise RuntimeError("用户中断，已停止本次Excel工作进程")
            if time.monotonic() > 截止:
                _终止(worker, 状态路径)
                已终止 = True
                raise TimeoutError("Excel计算或保存未完成")
            time.sleep(0.1)
        if worker.returncode or not 结果路径.is_file():
            raise RuntimeError("Excel工作进程失败：" + 日志路径.read_text(encoding="utf-8", errors="replace")[-1000:])
        结果 = json.loads(结果路径.read_text(encoding="utf-8"))
        if not 结果.get("成功"):
            raise RuntimeError(结果.get("错误") or "Excel交付失败")
        return 结果["结果"]
    except Exception:
        if worker.poll() is None and not 已终止:
            _终止(worker, 状态路径)
        raise


if __name__ == "__main__" and len(sys.argv) == 3 and sys.argv[1] == "--worker":
    请求路径 = Path(sys.argv[2])
    结果路径 = 请求路径.with_name("Excel结果.json")
    try:
        结果 = _实际交付(json.loads(请求路径.read_text(encoding="utf-8")))
        内容 = {"成功": True, "结果": 结果}
    except Exception as e:
        内容 = {"成功": False, "错误": f"{type(e).__name__}: {e}"}
    结果路径.write_text(json.dumps(内容, ensure_ascii=False, default=str), encoding="utf-8")
    sys.exit(0 if 内容["成功"] else 1)
