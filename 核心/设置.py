# coding: utf-8
"""模型配置的保存与读取；API 密钥经 Windows DPAPI 加密落盘，不存明文。"""
import json
from pathlib import Path

import win32crypt


def 默认配置路径() -> Path:
    return Path(__file__).resolve().parent.parent / "模型配置.json"


def 保存模型配置(路径, 配置: dict) -> None:
    数据 = {k: v for k, v in 配置.items() if k != "api_key"}
    密文 = win32crypt.CryptProtectData(配置["api_key"].encode("utf-8"), None, None, None, None, 0)
    数据["api_key_enc"] = 密文.hex()
    Path(路径).write_text(json.dumps(数据, ensure_ascii=False, indent=2), encoding="utf-8")


def 读取模型配置(路径) -> dict:
    数据 = json.loads(Path(路径).read_text(encoding="utf-8"))
    明文 = win32crypt.CryptUnprotectData(bytes.fromhex(数据.pop("api_key_enc")), None, None, None, 0)[1]
    数据["api_key"] = 明文.decode("utf-8")
    return 数据
