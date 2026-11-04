# coding: utf-8
"""测试用的识别记录目录：必须落在系统临时目录，不能落在项目或同步盘。

产品代码会拒绝项目目录（`\\documents\\code\\` 等）和同步盘；pytest 的 basetemp
在仓库内，不能直接拿来当记录目录，否则复用会被正确地禁用。
"""
import tempfile
from pathlib import Path


def 安全记录目录(名: str = "记录") -> Path:
    """系统临时目录下的唯一目录；每个用例各用一个，互不干扰。"""
    return Path(tempfile.mkdtemp(prefix=f"审计附注_{名}_"))
