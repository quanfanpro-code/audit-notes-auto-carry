# coding: utf-8
"""共用识别记录簿：本地文件＋内存索引，批次不可变，跨进程独占锁内发布。

磁盘布局（目录只放本机、非同步位置）：
- ``批次/`` 每批一个完整 JSON，写入后不再修改；
- ``当前.json`` 只记录单调代次与已提交批次的编号和摘要，一次原子替换；
- ``锁`` 1 字节文件，所有写者按同一路径、同一字节区间加锁。

读者只按某份已提交清单取批次：未列入的批次不可见，读到一半崩溃不会留下半批状态。
索引刷新失败、批次缺失或摘要不符时，整库标记不可用，本次禁用复用，不拿旧索引凑。
"""
from __future__ import annotations

import hashlib
import ctypes
import json
import os
import threading
import time
from pathlib import Path

try:
    import msvcrt
except ImportError:                     # 只面向 Windows，保留导入期可读性
    msvcrt = None

_清单名 = "当前.json"
_批次目录 = "批次"
_锁名 = "锁"
_线程锁 = {}
_线程锁表锁 = threading.Lock()
_同步标记 = ("onedrive", "dropbox", "google drive", "googledrive", "icloud", "sync",
          "坚果云", "nutstore", "baidunetdisk", "百度网盘", "微云", "weiyun")


def 目录是否安全(路径) -> bool:
    """记录目录必须是绝对路径、不落在同步盘；无法确认就拒绝。"""
    if 路径 is None:
        return False
    原 = Path(路径)
    if not 原.is_absolute():
        return False                       # 相对路径随工作目录漂移，无法确认落在哪里
    try:
        实际 = 原.resolve()
        文 = str(实际).lower()
        if 原.drive.startswith("\\") or 实际.drive.startswith("\\"):
            return False
        # Windows 原生盘类型：只接受本地可写盘，映射网络盘不能绕过。
        类型 = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(实际.anchor))
        if 类型 not in (2, 3, 6):
            return False
    except (OSError, ValueError):
        return False
    return not any(x in 文 for x in _同步标记)


def 项目根() -> Path:
    """源码运行时的项目根：核心/ 的上一级。"""
    return Path(__file__).resolve().parents[1]


def 默认目录():
    """记录目录：环境变量优先，否则项目根下 识别记录；同步盘与相对路径仍拒。

    shortcut: 打包成单文件运行时 __file__ 在临时解包目录，需显式指定目录。
    """
    指定 = os.environ.get("审计附注_识别记录目录")
    候选 = Path(指定) if 指定 else 项目根() / "识别记录"
    return 候选 if 目录是否安全(候选) else None


def 已启用(输入: dict = None) -> bool:
    """跨任务复用的总开关：环境变量优先，其次任务输入；关闭后仍按新可见范围完整识别。"""
    值 = os.environ.get("审计附注_复用")
    if 值 is not None:
        return str(值).strip().lower() not in ("0", "off", "false", "否", "关闭")
    if isinstance(输入, dict) and 输入.get("复用") is not None:
        return 输入.get("复用") not in (False, 0, "0", "off", "关闭")
    return True


def _摘要(内容: bytes) -> str:
    return hashlib.sha256(内容).hexdigest()


def _取线程锁(键: str):
    with _线程锁表锁:
        return _线程锁.setdefault(键, threading.RLock())


class _文件锁:
    """同一规范化路径的跨进程字节区间锁；同进程线程也串行。"""

    def __init__(self, 路径: Path, 等待秒: float = 10.0):
        self.路径 = Path(路径)
        self.等待秒 = 等待秒
        self._句柄 = None
        self._线程锁 = _取线程锁(str(self.路径.resolve()).lower())

    def __enter__(self):
        self._线程锁.acquire()
        句柄 = open(self.路径, "a+b")
        截止 = time.monotonic() + self.等待秒
        while True:
            try:
                句柄.seek(0)
                msvcrt.locking(句柄.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= 截止:
                    句柄.close()
                    self._线程锁.release()
                    raise TimeoutError("识别记录簿被其他进程占用，等待超时；本次不使用旧记录")
                time.sleep(0.02)
        self._句柄 = 句柄
        return self

    def __exit__(self, *_):
        try:
            self._句柄.seek(0)
            msvcrt.locking(self._句柄.fileno(), msvcrt.LK_UNLCK, 1)
        finally:
            self._句柄.close()
            self._句柄 = None
            self._线程锁.release()


def _写文件并落盘(路径: Path, 内容: bytes) -> None:
    临时 = 路径.with_name(路径.name + ".tmp")
    with 临时.open("wb") as f:
        f.write(内容)
        f.flush()
        os.fsync(f.fileno())
    # 锁内只重试同一份内容；若替换已成功但调用报错，核对落盘结果。
    for 次 in range(6):
        try:
            os.replace(临时, 路径)
            return
        except PermissionError:
            try:
                if 路径.read_bytes() == 内容:
                    return
            except OSError:
                pass
            if 次 == 5:
                raise
            time.sleep(0.05 * (次 + 1))


def _读json(路径: Path):
    return json.loads(路径.read_text(encoding="utf-8"))


class 记录簿:
    """一个记录簿目录的读写入口；进程内缓存索引，代次变化时重读。"""

    def __init__(self, 目录=None):
        指定 = 目录 if 目录 else os.environ.get("审计附注_识别记录目录")
        候选 = Path(目录) if 目录 else 默认目录()
        self._缓存 = None                  # (代次, 索引)
        self._不可用 = ""                  # 缓存失效原因：非空时本次一律不复用
        self.最近原因 = ""
        if 候选 is not None and not 目录是否安全(候选):
            self.目录 = None
            self._不可用 = "记录目录不在允许的位置（同步盘、网络位置或无法确认位置）"
        elif 候选 is None and 指定:
            self.目录 = None
            self._不可用 = "记录目录不在允许的位置（同步盘、网络位置或无法确认位置）"
        else:
            self.目录 = 候选
            if 候选 is None:
                self._不可用 = "未找到可用的本机记录目录"

    # —— 读 ——
    def _清单(self):
        if self.目录 is None:
            raise FileNotFoundError("未找到可用的本机记录目录")
        路径 = self.目录 / _清单名
        if not 路径.is_file():
            return {"代次": 0, "批次": []}
        return _读json(路径)

    def 快照(self):
        """取一次已提交清单并构造索引；任何异常都标记本次不可用。"""
        if self.目录 is None:
            self.最近原因 = self._不可用 or "未找到可用的本机记录目录"
            return {"可用": False, "代次": None, "原因": self.最近原因,
                    "解释": {}, "按证据": {}, "按步": {}, "冲突": {}, "依赖": {},
                    "使用": [], "请求": {}, "批次": []}
        try:
            清单 = self._清单()
            代次 = int(清单.get("代次") or 0)
            if self._缓存 and self._缓存[0] == 代次 and not self._不可用:
                self._校验批次(清单)          # 同一代次也要确认批次还在、没被截断
                return self._缓存[1]
            索引 = self._建索引(清单)
            self._缓存 = (代次, 索引)
            self._不可用 = ""
            self.最近原因 = ""
            return 索引
        except Exception as e:                      # 缺批次、摘要不符、清单损坏都走这里
            self._不可用 = f"索引刷新失败：{e}"
            self.最近原因 = self._不可用
            self._缓存 = None
            return {"可用": False, "代次": None, "原因": self._不可用,
                    "解释": {}, "按证据": {}, "按步": {}, "冲突": {}, "依赖": {},
                    "使用": [], "请求": {}, "批次": []}

    def _校验批次(self, 清单):
        for 批 in 清单.get("批次") or []:
            文件 = self.目录 / _批次目录 / f"{批.get('编号')}.json"
            if not 文件.is_file():
                raise FileNotFoundError(f"批次缺失：{批.get('编号')}")
            if 批.get("大小") is not None and 文件.stat().st_size != 批.get("大小"):
                raise ValueError(f"批次大小不符：{批.get('编号')}")

    def _建索引(self, 清单):
        索引 = {"可用": True, "代次": int(清单.get("代次") or 0), "解释": {},
              "按证据": {}, "按步": {}, "冲突": {}, "依赖": {}, "使用": [],
              "请求": {}, "批次": []}
        for 批 in 清单.get("批次") or []:
            文件 = self.目录 / _批次目录 / f"{批.get('编号')}.json"
            if not 文件.is_file():
                raise FileNotFoundError(f"批次缺失：{批.get('编号')}")
            内容 = 文件.read_bytes()
            if _摘要(内容) != 批.get("摘要"):
                raise ValueError(f"批次摘要不符：{批.get('编号')}")
            if 批.get("大小") is not None and len(内容) != 批.get("大小"):
                raise ValueError(f"批次大小不符：{批.get('编号')}")
            数据 = json.loads(内容.decode("utf-8"))
            索引["批次"].append(批.get("编号"))
            for 事件 in 数据.get("事件") or []:
                self._吃事件(索引, 事件)
        # 冲突按解释内容判，不比较各自存盘的摘要：旧记录的解释键由旧算法算出，
        # 升级后同一条解释按新算法会得到不同摘要串，直接比字符串会把相同解释当成冲突。
        from .复用证据 import 解释键 as 计算解释键
        for 键, 版本们 in 索引["按证据"].items():
            有效 = [v for v in 版本们 if (索引["解释"].get(v) or {}).get("状态") == "有效"]
            解释们 = set()
            for v in 有效:
                项 = 索引["解释"][v]
                try:
                    解释们.add(计算解释键(项.get("解释") or {}))
                except Exception:
                    解释们.add(str(项.get("解释键") or v))   # 解释缺损时退回存盘摘要
            if len(解释们) > 1:
                索引["冲突"][键] = sorted(有效)
        return 索引

    @staticmethod
    def _吃事件(索引, 事件):
        类型 = 事件.get("类型")
        if 类型 == "发布":
            版本 = 事件.get("版本")
            if not 版本:
                return
            项 = 索引["解释"].setdefault(版本, {
                "版本": 版本, "证据": 事件.get("证据"), "解释": 事件.get("解释"),
                "解释键": 事件.get("解释键"), "证据键": 事件.get("证据键"),
                "侧": 事件.get("侧"), "步骤": 事件.get("步骤"),
                "金标准版本": 事件.get("金标准版本"), "规则版本": 事件.get("规则版本"),
                "证据版本": 事件.get("证据版本"), "来源": [], "状态": "有效", "停用原因": ""})
            for 来源 in 事件.get("来源") or []:
                if 来源 and 来源 not in 项["来源"]:
                    项["来源"].append(来源)
            if 事件.get("依赖"):
                索引["依赖"][版本] = list(dict.fromkeys((索引["依赖"].get(版本) or []) + 事件["依赖"]))
            索引["按证据"].setdefault(事件.get("证据键"), [])
            if 版本 not in 索引["按证据"][事件.get("证据键")]:
                索引["按证据"][事件.get("证据键")].append(版本)
            步骤 = 事件.get("步骤")
            if 步骤 in ("第一遍", "第二遍"):
                材料类型 = ((事件.get("证据") or {}).get("材料类型")
                         or ("excel" if 事件.get("侧") == "B" else "word"))
                槽 = f"{材料类型}|{步骤}|-"
            else:
                槽 = f"*|{步骤}|{事件.get('形状键') or '-'}"
            索引["按步"].setdefault(槽, [])
            if 版本 not in 索引["按步"][槽]:
                索引["按步"][槽].append(版本)
        elif 类型 == "使用":
            索引["使用"].append({键: 事件.get(键) for 键 in
                              ("解释版本", "任务", "侧", "表组", "读取部分", "工作表",
                               "原文件类型", "代次", "请求")})
        elif 类型 == "停用":
            for 版本 in 事件.get("目标") or []:
                if 版本 in 索引["解释"]:
                    索引["解释"][版本]["状态"] = "停用"
                    索引["解释"][版本]["停用原因"] = 事件.get("原因") or ""
            # 直接与间接下游同批次失效：停用一版不能只停它自己
            for 版本 in 事件.get("失效") or []:
                if 版本 in 索引["解释"] and 索引["解释"][版本]["状态"] == "有效":
                    索引["解释"][版本]["状态"] = "停用"
                    索引["解释"][版本]["停用原因"] = "因上游停用：" + (事件.get("原因") or "")
        elif 类型 == "裁决":
            保留 = [事件["保留"]] if 事件.get("保留") else []
            for 版本 in set(事件.get("停用") or []) | set(事件.get("失效") or []):
                if 版本 in 索引["解释"]:
                    索引["解释"][版本]["状态"] = "停用"
                    索引["解释"][版本]["停用原因"] = 事件.get("原因") or ""
            for 版本 in 保留:
                if 版本 in 索引["解释"]:
                    索引["解释"][版本]["裁决"] = 事件.get("含义") or "保留"
        if 事件.get("请求"):
            索引["请求"].setdefault(事件["请求"], {"类型": 类型, "事件": 事件})

    # —— 写 ——
    def 提交(self, 事件们: list, 请求标识: str, 预期代次=None, 校验=None) -> dict:
        """在跨进程锁内追加一批事件：先落盘核验批次，再一次替换当前清单。

        校验 在锁内、写批次之前调用（拿到的是刚读到的清单），用来拒绝依赖已失效状态
        的提交；它抛异常即整批不写。这样停用与登记交错时不会两边都成功。
        """
        if self.目录 is None:
            raise FileNotFoundError("未找到可用的本机记录目录")
        if not 请求标识:
            raise ValueError("提交必须带请求标识，便于同一请求重试时幂等")
        事件们 = list(事件们 or [])
        重试 = 0
        while True:
            try:
                return self._提交一次(事件们, 请求标识, 预期代次, 校验)
            except _代次不符:
                重试 += 1
                if 预期代次 is not None or 重试 > 20:
                    raise RuntimeError("记录簿代次已变化，请重读后重新提交") from None
                time.sleep(0.01)

    def _提交一次(self, 事件们, 请求标识, 预期代次, 校验=None):
        self.目录.mkdir(parents=True, exist_ok=True)
        (self.目录 / _批次目录).mkdir(exist_ok=True)
        with _文件锁(self.目录 / _锁名):
            清单 = self._清单()
            代次 = int(清单.get("代次") or 0)
            if 预期代次 is not None and int(预期代次) != 代次:
                raise _代次不符()
            if 校验 is not None:
                校验(清单)  # 管理操作在同一锁内生成结果，重试也在锁内核对原请求。
            for 批 in 清单.get("批次") or []:
                文件 = self.目录 / _批次目录 / f"{批.get('编号')}.json"
                if 文件.is_file():
                    数据 = _读json(文件)
                    if 数据.get("请求") == 请求标识:
                        if 数据.get("事件") == 事件们:
                            return {"代次": 代次, "状态": "已存在", "请求": 请求标识,
                                    "批次": 批.get("编号")}
                        raise ValueError("同一请求标识对应不同内容，拒绝提交")
            编号 = f"b{代次 + 1:08d}"
            数据 = {"代次": 代次 + 1, "请求": 请求标识, "事件": 事件们}
            内容 = json.dumps(数据, ensure_ascii=False, sort_keys=True,
                            indent=1).encode("utf-8")
            批次路径 = self.目录 / _批次目录 / f"{编号}.json"
            _写文件并落盘(批次路径, 内容)
            新清单 = {"代次": 代次 + 1,
                     "批次": (清单.get("批次") or []) + [
                         {"编号": 编号, "摘要": _摘要(内容), "大小": len(内容)}]}
            _写文件并落盘(self.目录 / _清单名,
                       json.dumps(新清单, ensure_ascii=False, indent=1).encode("utf-8"))
            索引 = self._建索引(新清单)
            self._缓存 = (新清单["代次"], 索引)
            self._不可用 = ""
            return {"代次": 新清单["代次"], "状态": "已提交", "请求": 请求标识, "批次": 编号}

    def 刷新(self):
        """发布或外部提交后同步刷新索引；失败即标记不可用，不沿用旧索引。"""
        if self.目录 is None:
            self.最近原因 = self._不可用 or "未找到可用的本机记录目录"
            return {"可用": False, "代次": None, "原因": self.最近原因,
                    "解释": {}, "按证据": {}, "按步": {}, "冲突": {}, "依赖": {},
                    "使用": [], "请求": {}, "批次": []}
        try:
            清单 = self._清单()
            索引 = self._建索引(清单)
            self._缓存 = (int(清单.get("代次") or 0), 索引)
            self._不可用 = ""
            self.最近原因 = ""
            return 索引
        except Exception as e:
            self._不可用 = f"索引刷新失败：{e}"
            self.最近原因 = self._不可用
            self._缓存 = None
            return {"可用": False, "代次": None, "原因": self._不可用,
                    "解释": {}, "按证据": {}, "按步": {}, "冲突": {}, "依赖": {},
                    "使用": [], "请求": {}, "批次": []}

    # —— 供内部管理与流程使用 ——
    def 可用(self) -> bool:
        return bool(self.快照().get("可用"))

    def 读索引(self, 清单) -> dict:
        """按给定清单构造索引（供锁内校验使用，不读盘、不改缓存）。"""
        return self._建索引(清单)

    def 解释们(self) -> dict:
        return (self.快照().get("解释") or {})

    def 按证据(self) -> dict:
        return (self.快照().get("按证据") or {})

    def 按步(self) -> dict:
        return (self.快照().get("按步") or {})

    def 冲突(self) -> dict:
        return (self.快照().get("冲突") or {})


class _代次不符(Exception):
    pass


_实例 = {}


def 取记录簿(目录=None):
    """同一进程内按目录共用一个记录簿实例；任务开始时刷新，跨进程提交靠代次发现。"""
    键 = str(Path(目录).resolve()) if 目录 else "默认"
    if 键 not in _实例:
        _实例[键] = 记录簿(目录)
    return _实例[键]


def 发布事件(侧, 步骤, 证据, 解释, 证据键值, 解释键值, 版本, 金标准版本, 规则版本,
         来源, 形状键=None, 依赖=None) -> dict:
    return {"类型": "发布", "侧": 侧, "步骤": 步骤, "证据": 证据, "解释": 解释,
            "证据键": 证据键值, "解释键": 解释键值, "版本": 版本, "形状键": 形状键,
            "金标准版本": 金标准版本, "规则版本": 规则版本,
            "证据版本": 证据.get("证据版本"), "来源": [来源] if 来源 else [],
            "依赖": list(dict.fromkeys(str(x) for x in (依赖 or []) if x))}


def 使用事件(版本, 任务, 侧, 表组, 代次, 请求="", 读取部分="", 工作表="", 原文件类型="") -> dict:
    return {"类型": "使用", "解释版本": 版本, "任务": str(任务), "侧": 侧,
            "表组": 表组, "读取部分": 读取部分, "工作表": 工作表, "原文件类型": 原文件类型,
            "代次": 代次,
            "请求": 请求 or f"使用:{任务}:{侧}:{读取部分 or '-'}:{表组}:{版本}"}
