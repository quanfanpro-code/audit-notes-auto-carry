# coding: utf-8
"""测试用的识别记录目录：系统临时目录下各建一个，用例之间互不干扰。

产品默认放在项目根的 `识别记录`，那是要长期积累的资产；测试不能往那里写，
否则一次全量就会把测试语料混进真实记录。
"""
import tempfile
from pathlib import Path


def 安全记录目录(名: str = "记录") -> Path:
    """系统临时目录下的唯一目录；每个用例各用一个，互不干扰。"""
    return Path(tempfile.mkdtemp(prefix=f"审计附注_{名}_"))
