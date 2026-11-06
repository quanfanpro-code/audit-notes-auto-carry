# coding: utf-8
import json
from pathlib import Path
from 测试.GUI样本 import 准备, 开始, 等待


def test_答复绑定材料替换保护且能继续(tmp_path, monkeypatch):
    客户, 输入 = 准备(tmp_path, monkeypatch, 模型=False)
    初 = 开始(客户, 输入, "人工答复")
    等待(客户, 初["task_id"])
    基址 = f"/api/桌面/任务/{初['task_id']}"
    r = 客户.get(基址 + "/答复")
    assert r.status_code == 200
    材料 = r.json()["条目"][0]
    路径 = 基址 + "/答复/" + 材料["id"]
    读 = 客户.get(路径).json()
    assert 读["指令"] and 读["材料"] and 读["已有答复"] == ""
    请求 = {"答复": "不是JSON", "原答复哈希": 读["答复哈希"], "请求标识": "第一次"}
    assert 客户.post(路径, json=请求).status_code == 400
    from 测试.test_Excel公式 import _模型
    答复 = _模型()({}, 读["指令"], 读["材料"], Path(读["路径"]))
    请求["答复"] = 答复
    assert 客户.post(路径, json=请求).status_code == 200
    assert 客户.post(路径, json=请求).status_code == 200
    已存 = 客户.get(路径).json()
    改 = {**请求, "请求标识": "第二次", "答复": 答复 + "\n", "原答复哈希": 已存["答复哈希"]}
    assert 客户.post(路径, json=改).status_code == 409
    改["替换确认"] = True
    assert 客户.post(路径, json=改).status_code == 200
    assert 客户.post(路径, json={**改, "请求标识": "第三次", "答复": 答复 + "  "}).status_code == 409
    原材料 = Path(读["路径"]) / "材料.txt"
    前 = 原材料.read_bytes()
    assert 客户.post(基址 + "/中断").status_code == 200
    assert 原材料.read_bytes() == 前
    assert 客户.get(基址 + "/答复/不存在").status_code == 404
    继续 = 客户.post(基址 + "/继续", json={"请求标识": "继续"}).json()
    完 = 等待(客户, 继续["task_id"])
    assert 完["等待"] and "第二遍" in 完["待答复"][0]
    assert (Path(读["路径"]) / "答复.txt").read_text(encoding="utf-8") == 改["答复"]


def test_模型密钥不回显且界面默认值实际生效(tmp_path, monkeypatch):
    客户, 输入 = 准备(tmp_path, monkeypatch)
    from 服务 import web, 界面设置
    from 核心.设置 import 读取模型配置
    monkeypatch.setattr(web, '读取模型配置', 读取模型配置)
    monkeypatch.setattr(界面设置, '设置路径', tmp_path / '界面设置.json')
    入 = {'模型': {'url':'https://example.invalid/v1', 'model':'本例模型', 'api_key':'只用于加密测试', 'timeout':60},
          '默认设置': {'金标准':输入['金标准'], '复用':True, '识别记录目录':str(tmp_path / '识别记录'), '字号':16}}
    assert 客户.post('/api/桌面/设置', json=入).status_code == 200
    读 = 客户.get('/api/桌面/设置').json()
    assert 读['模型']['密钥已保存'] and 'api_key' not in 读['模型']
    assert '只用于加密测试' not in Path(web.应用.state.配置路径).read_text(encoding='utf-8')
    入['模型']['api_key'] = ''
    assert 客户.post('/api/桌面/设置', json=入).status_code == 200
    assert 读取模型配置(web.应用.state.配置路径)['api_key'] == '只用于加密测试'
    monkeypatch.setenv('审计附注_复用','0')
    概况 = 客户.get('/api/桌面/概况').json()
    assert 概况['复用锁定'] and 概况['默认设置']['复用'] is False
    assert 概况['默认设置']['字号'] == 16


def test_人工回写入口真实Word与原件保护(tmp_path, monkeypatch):
    import hashlib
    from openpyxl import load_workbook
    from 测试.组合样本 import 造Word, 读Word格
    from 核心.桥接 import 提取
    客户, _ = 准备(tmp_path, monkeypatch)
    原 = tmp_path / '原.docx'
    造Word(原, 单位='元')
    表, 连接, 文 = tmp_path/'提取.xlsx', tmp_path/'连接.docx', tmp_path/'上下文.json'
    提取(原, 表, 连接, 文, 工作目录=tmp_path/'提取留证')
    前 = hashlib.sha256(原.read_bytes()).hexdigest()
    wb = load_workbook(表); ws = wb.active; ws['B3']='人工名单'; ws['C3']=777; ws['D3']=888; wb.save(表); wb.close()
    输入 = {'带连接word':str(连接), 'excel':str(表), '输出':str(tmp_path/'人工新副本.docx')}
    r = 客户.post('/api/桌面/任务', json={'类型':'人工回写', '请求标识':'人工回写一次', '原始输入':输入})
    assert r.status_code == 200
    末 = 等待(客户, r.json()['task_id'])
    assert 末['步骤'] == 'done' and not 末['允许继续']
    assert 读Word格(输入['输出'],0,1,0) == '人工名单'
    assert 读Word格(输入['输出'],0,1,1) == '777'
    assert 读Word格(输入['输出'],0,1,2) == '888'
    assert hashlib.sha256(原.read_bytes()).hexdigest() == 前
