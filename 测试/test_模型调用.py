# coding: utf-8
"""模型调用测试：指令超长拒发、HTTP 调用与留证。"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import json, threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from 核心.模型调用 import 校验指令, 调用识别


def test_指令超长拒发():
    try:
        校验指令("字" * 2001)
        assert False, "应拒绝"
    except ValueError as e:
        assert "2000" in str(e)


def test_调用留证(tmp_path):
    class H(BaseHTTPRequestHandler):
        def do_POST(self):
            body = self.rfile.read(int(self.headers["Content-Length"]))
            req = json.loads(body)
            assert req["messages"][0]["content"].startswith("材料")  # 材料完整在前
            out = json.dumps({"choices": [{"message": {"content": "答复"}}]}).encode()
            self.send_response(200); self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(out))); self.end_headers()
            self.wfile.write(out)
        def log_message(self, *a): pass
    srv = HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    配置 = {"url": f"http://127.0.0.1:{srv.server_port}/v1", "api_key": "k",
            "model": "m", "timeout": 30}
    答复 = 调用识别(配置, "问法", "材料内容", tmp_path / "调用1")
    assert 答复 == "答复"
    assert (tmp_path / "调用1" / "请求.json").is_file()
    assert (tmp_path / "调用1" / "响应.txt").read_text(encoding="utf-8") == "答复"
    srv.shutdown()
