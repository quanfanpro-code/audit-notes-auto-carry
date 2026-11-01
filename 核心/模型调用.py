# coding: utf-8
"""模型调用：简短指令守卫、HTTP API 调用、调用留证。"""
import json
import time
from pathlib import Path

import httpx


def 校验指令(指令: str) -> None:
    if len(指令) > 2000:
        raise ValueError(f"简短指令{len(指令)}字符，超过2000上限；拒绝发送")


def 调用识别(配置: dict, 指令: str, 材料: str, 调用目录) -> str:
    校验指令(指令)
    调用目录 = Path(调用目录)
    调用目录.mkdir(parents=True, exist_ok=True)
    请求体 = {"model": 配置["model"],
             "messages": [{"role": "user", "content": 材料},
                          {"role": "user", "content": 指令}]}
    (调用目录 / "请求.json").write_text(json.dumps(请求体, ensure_ascii=False, indent=2), encoding="utf-8")
    开始 = time.monotonic()
    try:
        响应 = httpx.post(配置["url"].rstrip("/") + "/chat/completions", json=请求体,
                    timeout=float(配置.get("timeout", 600)),
                    headers={"Authorization": f"Bearer {配置['api_key']}"})
        响应.raise_for_status()
        文本 = 响应.json()["choices"][0]["message"]["content"]
    except Exception as e:
        (调用目录 / "耗时.json").write_text(json.dumps({"秒": time.monotonic() - 开始}, ensure_ascii=False), encoding="utf-8")
        raise RuntimeError(f"模型调用失败：{e}") from e
    (调用目录 / "响应.txt").write_text(文本, encoding="utf-8")
    (调用目录 / "耗时.json").write_text(json.dumps({"秒": time.monotonic() - 开始}, ensure_ascii=False), encoding="utf-8")
    return 文本
