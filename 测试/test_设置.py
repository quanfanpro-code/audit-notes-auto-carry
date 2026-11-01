# coding: utf-8
"""设置模块测试：模型配置往返、密钥加密落盘。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from 核心.设置 import 保存模型配置, 读取模型配置


def test_配置往返加密(tmp_path):
    p = tmp_path / "模型配置.json"
    保存模型配置(p, {"url": "https://api.example.com/v1", "api_key": "sk-秘密",
                 "model": "gpt-x", "timeout": 600})
    cfg = 读取模型配置(p)
    assert cfg["url"] == "https://api.example.com/v1"
    assert cfg["api_key"] == "sk-秘密"      # 解密后还原
    assert "sk-秘密" not in p.read_text(encoding="utf-8")  # 密钥非明文
