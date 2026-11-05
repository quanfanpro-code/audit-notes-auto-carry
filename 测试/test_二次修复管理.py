# coding: utf-8
"""管理入口和跨进程交错的回归，不以预制依赖代替真实任务发布。"""
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

from 核心.共用记录簿 import 记录簿, 目录是否安全
from 核心.复用管理 import 停用, 裁决
from 测试.安全目录 import 安全记录目录
from 测试.test_复用记录管理 import _发布


def test_裁决保留请求原样重试():
    簿 = 记录簿(安全记录目录())
    _发布(簿, 'X', '相同', '甲解释', '任务甲')
    _发布(簿, 'Y', '相同', '乙解释', '任务乙')
    a = 裁决(簿, '裁决1', '相同', ['X', 'Y'], 保留='X', 依据='复核')
    b = 裁决(簿, '裁决1', '相同', ['Y', 'X'], 保留='X', 依据='复核')
    assert a == b and 簿.快照()['代次'] == 3


def test_裁决部分停用状态与剩余一致():
    簿 = 记录簿(安全记录目录())
    _发布(簿, 'X', '相同', '甲解释', '任务甲')
    _发布(簿, 'Y', '相同', '乙解释', '任务乙')
    a = 裁决(簿, '裁决1', '相同', ['X', 'Y'], 停用们=['X'], 依据='复核')
    assert a['状态'] != '全部停用'
    assert 簿.快照()['解释']['Y']['状态'] == '有效'


def test_停用提交前另进程登记不会漏报(monkeypatch):
    簿 = 记录簿(安全记录目录())
    _发布(簿, 'X', '证', '解', '首次任务')
    提交 = 簿.提交
    def 交错(事件们, 请求标识, **设置):
        code = "from 核心.共用记录簿 import 记录簿; from 核心.复用接入 import 登记使用; import sys; assert 登记使用(记录簿(sys.argv[1]),'X','并发任务','A','G2',1)[0]"
        p = subprocess.run([sys.executable, '-c', code, str(簿.目录)], cwd=Path(__file__).resolve().parents[1],
                           capture_output=True, encoding='utf-8', timeout=30,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        assert p.returncode == 0, p.stderr
        return 提交(事件们, 请求标识, **设置)
    monkeypatch.setattr(簿, '提交', 交错)
    out = 停用(簿, '停用1', ['X'], '错认', '复核')
    assert '并发任务' in out['待复核']
    assert out['代次'] == 簿.快照()['代次'] == 3
    assert {'任务': '并发任务', '侧': 'A', '表组': 'G2'} in out['直接']['任务表组']


def test_依赖来自正常首次任务发布(tmp_path, monkeypatch):
    from 测试.test_二次修复 import _有依据模型, _带说明Word, _B单位, _单组金标准, _跑一份
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '元', 说明))
    库 = 安全记录目录()
    出 = _跑一份(tmp_path / '任务', _带说明Word(tmp_path, '原件', 说明),
                _B单位(tmp_path, '底表.xlsx', '元', 901, 911), _单组金标准(tmp_path), 库)
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    版本 = {x['步骤']: x['版本'] for x in 日志['发布'] if x['状态'] == '已发布'}
    簿 = 记录簿(库)
    out = 停用(簿, '停用科目', [版本['第一遍']], '科目误认', '复核')
    assert 版本['第二遍'] in out['受影响版本'] and 版本['第三遍'] in out['受影响版本']
    assert 簿.快照()['解释'][版本['第三遍']]['状态'] == '停用'


def test_目录拒绝网络路径但允许项目():
    assert not 目录是否安全(r'\\localhost\C$\临时审计缓存')
    根 = 安全记录目录('目录检验')
    (根 / '.git').mkdir()
    assert 目录是否安全(根 / '识别记录')


def test_目录按Windows驱动类型拒绝映射网络盘(monkeypatch):
    import ctypes
    根 = 安全记录目录('驱动检查')
    for 类型, 允许 in ((4, False), (0, False), (3, True)):
        monkeypatch.setattr(ctypes.windll.kernel32, 'GetDriveTypeW', lambda _, t=类型: t)
        assert 目录是否安全(根) is 允许


def test_目录联结按最终位置核对(tmp_path):
    目标 = tmp_path / 'OneDrive' / '记录'
    目标.mkdir(parents=True)
    联结 = tmp_path / '普通名字'
    命令 = Path(os.environ['SystemRoot']) / 'System32' / 'cmd.exe'
    assert 命令.is_file()
    p = subprocess.run([str(命令), '/c', 'mklink', '/J', str(联结), str(目标)],
                       capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    assert p.returncode == 0, p.stderr
    assert 联结.resolve() == 目标.resolve()
    assert not 目录是否安全(联结)


def test_续跑规则变更在消费产物前拒绝(tmp_path, monkeypatch):
    from 核心.流程 import 运行任务
    from 核心 import 看表规范, 识别
    from 测试.test_二次修复 import _带说明Word, _B单位, _单组金标准
    monkeypatch.setenv('审计附注_复用', '0')
    a = _带说明Word(tmp_path, '原件', '金额单位：人民币元')
    b = _B单位(tmp_path, '底表.xlsx', '元', 901, 911)
    入 = {'word': str(a), 'b表': str(b), '金标准': str(_单组金标准(tmp_path)), 'b工作表': 'word附注'}
    d = tmp_path / '任务'
    assert 运行任务(入, d, {})['步骤'] == 'identify_a'
    原索引 = (d / '接续索引.json').read_bytes()
    原 = 看表规范.精读规范()
    monkeypatch.setattr(看表规范, '精读规范', lambda: 原 + '\n新增识别规则')
    monkeypatch.setattr(识别, '精读规范', 看表规范.精读规范)
    with pytest.raises(RuntimeError, match='规则|版本'):
        运行任务(入, d, {})
    assert (d / '接续索引.json').read_bytes() == 原索引


@pytest.mark.parametrize('失败方式', ['一次', '替换后报错', '持续'])
def test_重试清单替换不得重复登记(monkeypatch, 失败方式):
    from 核心.复用接入 import 登记使用
    from 核心 import 共用记录簿 as 模块
    簿 = 记录簿(安全记录目录())
    _发布(簿, 'X', '证', '解', '首次')
    replace = 模块.os.replace
    次 = []
    def 替换(src, dst):
        if Path(dst).name == '当前.json':
            次.append(1)
            if 失败方式 == '替换后报错' and len(次) == 1:
                replace(src, dst)
                raise PermissionError('模拟替换完成但调用方收到异常')
            if 失败方式 == '持续' or len(次) == 1:
                raise PermissionError('模拟瞬时占用')
        return replace(src, dst)
    monkeypatch.setattr(模块.os, 'replace', 替换)
    ok, 因 = 登记使用(簿, 'X', '本次', 'A', 'G1', 1)
    out = 记录簿(簿.目录).快照()
    if 失败方式 == '持续':
        assert not ok and out['代次'] == 1 and out['使用'] == []
        assert 1 < len(次) <= 8
    else:
        assert ok, 因
        assert out['代次'] == 2 and len(out['使用']) == 1
