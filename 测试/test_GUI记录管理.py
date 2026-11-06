# coding: utf-8
from 测试.GUI样本 import 准备
from 测试.test_复用记录管理 import _链, _发布


def test_界面停用复用原影响算法且重复提交不增批次(tmp_path, monkeypatch):
    客户, _ = 准备(tmp_path, monkeypatch)
    簿 = _链(tmp_path)
    from 服务 import 记录展示
    monkeypatch.setattr(记录展示, "当前记录簿", lambda 宿主: 簿)
    列表 = 客户.get('/api/桌面/记录').json()
    assert 列表['总数'] == 3
    assert all(x['当前适用'] != '可自动复用' for x in 列表['条目'])
    预览 = 客户.post('/api/桌面/记录/影响预览', json={'版本': ['V甲']}).json()
    assert set(预览['待复核']) == {'任务甲', '任务乙', '任务丙'}
    请求 = {'请求标识': '界面停用', '版本': ['V甲'], '原因': '错认', '依据': '本次原表', '预览代次': 预览['代次']}
    一 = 客户.post('/api/桌面/记录/停用', json=请求)
    assert 一.status_code == 200
    批次 = 簿.快照()['代次']
    二 = 客户.post('/api/桌面/记录/停用', json=请求)
    assert 二.status_code == 200 and 簿.快照()['代次'] == 批次
    查 = 客户.get('/api/桌面/记录/请求/界面停用').json()
    assert set(查['待复核']) == {'任务甲', '任务乙', '任务丙'}
    assert all(x['状态'] == '停用' for x in 簿.快照()['解释'].values())


def test_冲突需明确选择且过时集合被原函数拒绝(tmp_path, monkeypatch):
    客户, _ = 准备(tmp_path, monkeypatch)
    簿 = _链(tmp_path)
    _发布(簿, 'V错', '证甲', '错解', '任务错', 解释={'科目':'存货'})
    from 服务 import 记录展示
    monkeypatch.setattr(记录展示, '当前记录簿', lambda 宿主: 簿)
    详情 = 客户.get('/api/桌面/记录/V甲').json()
    assert {x['版本'] for x in 详情['冲突解释']} == {'V甲','V错'}
    请求 = {'请求标识':'裁决旧', '证据组':'证甲', '预期冲突版本':['V甲'], '保留':'V甲', '原因':'复核', '依据':'原表'}
    assert 客户.post('/api/桌面/记录/裁决', json=请求).status_code == 409
    请求.update(请求标识='裁决新', 预期冲突版本=['V甲','V错'])
    assert 客户.post('/api/桌面/记录/裁决', json=请求).status_code == 200
    assert 簿.快照()['解释']['V甲']['状态'] == '有效'
