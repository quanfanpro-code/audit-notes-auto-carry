# coding: utf-8
"""ba96492 独立复核回归。最终金额必须读真实 Word；只固定模型答复。"""
import json
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree as E
from openpyxl import load_workbook

from 测试.test_复核反例 import (_单组Word, _单组金标准, _单组模型, _从起列造B,
                              _跑一份, 读回Word, W)
from 测试.安全目录 import 安全记录目录
from 测试.test_复用证据 import _表, _记录
from 核心.表格 import 读快照
from 核心.复用证据 import 建证据_精读, 核对精读, 表视图, _单位核对


def _带说明Word(tmp_path, 名, 说明):
    原 = _单组Word(tmp_path, 名 + '_底.docx')
    出 = tmp_path / (名 + '.docx')
    with ZipFile(原) as z, ZipFile(出, 'w') as dst:
        for p in z.infolist():
            data = z.read(p.filename)
            if p.filename == 'word/document.xml':
                tree = E.fromstring(data)
                tree.find('.//' + W + 'body/' + W + 'p/' + W + 'r/' + W + 't').text = 说明
                data = E.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            dst.writestr(p, data)
    return 出


def _有依据模型(A单位, B单位, 说明):
    原 = _单组模型('B')

    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        A = Path(目录).name.startswith('A_')
        if '划分' in 答 and A:
            答['划分'].insert(0, {'行范围': [1, 1], '科目': '非表格'})
        if '作用范围' in 答:
            单位 = A单位 if A else B单位
            依据 = {'来源': '上下文' if A else '表内',
                    '位置': ['unit_text_candidates', 0] if A else 'E1',
                    '原文': 说明 if A else '单位：人民币' + 单位,
                    '片段': '人民币' + 单位, '单位': 单位, '币种': '人民币',
                    '列': ['C', 'D']}
            答['作用范围'] = {'币种': '人民币', '单位': 单位, '单位依据': [依据]}
        return json.dumps(答, ensure_ascii=False)
    return 调用


def _B单位(tmp_path, 名, 单位, 初, 末):
    p = _从起列造B(tmp_path / 名, 'B', 初, 末)
    wb = load_workbook(p)
    wb.active['E1'] = '单位：人民币' + 单位
    wb.save(p); wb.close()
    return p


def test_单位同段注册资本不能掩盖本表单位变化(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    库 = 安全记录目录('二次修复单位')
    金 = _单组金标准(tmp_path)
    说明1 = '注册资本100万元，本表金额单位：人民币万元'
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('万元', '万元', 说明1))
    一 = _跑一份(tmp_path / '首次', _带说明Word(tmp_path, '首次', 说明1),
                _B单位(tmp_path, '首次.xlsx', '万元', 100, 110), 金, 库)
    assert 一['状态']['步骤'] == 'done'
    说明2 = '注册资本100万元，本表金额单位：人民币元'
    a = _带说明Word(tmp_path, '第二份', 说明2)
    b = _B单位(tmp_path, '第二份.xlsx', '万元', 901, 911)
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '万元', 说明2))
    二 = _跑一份(tmp_path / '开启复用', a, b, 金, 库)
    monkeypatch.setenv('审计附注_复用', '0')
    对照 = _跑一份(tmp_path / '关闭复用', a, b, 金, 库)
    assert 读回Word(对照['Word']).表格()[1] == ['甲公司', '100', '110']
    assert 读回Word(二['Word']).表格() == 读回Word(对照['Word']).表格()
    assert 读回Word(二['Word']).段落 == 读回Word(a).段落
    日志 = json.loads((二['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == 0, 日志


def test_单位没有证据不能放行():
    assert _单位核对([], None, None, None)


def test_单位分列必须分别核对(tmp_path):
    p = _表(tmp_path, '分列', [('甲', 1, 2)],
            首行=('项目', '2025-01-01', '2025-12-31', '单位：人民币元'))
    记录 = _记录(行数=2)
    记录['作用范围']['单位依据'] = [
        {'来源': '表内', '位置': 'D1', '原文': '单位：人民币元', '片段': '人民币元',
         '列': ['B'], '单位': '元', '币种': '人民币'}]
    证据 = 建证据_精读('A', [记录], 读快照(p))
    出 = 核对精读(证据, 表视图(读快照(p), 'sheet1', (1, 1, 4, 2)))
    assert 出['结果'] == '未命中', 'C列没有任何单位依据，不能继承B列的元'


@pytest.mark.parametrize('缺陷', ['没有依据', '同列冲突', '分列不一致', '标注失实'])
def test_单位重认仍未确认时整组保留真实Word(tmp_path, monkeypatch, 缺陷):
    monkeypatch.setenv('审计附注_复用', '0')
    说明 = '金额单位：人民币元'
    a = _带说明Word(tmp_path, '原件', 说明)
    b = _B单位(tmp_path, '底表.xlsx', '元', 901, 911)
    wb = load_workbook(b)
    wb.active['F1'] = '单位：人民币万元'
    wb.save(b); wb.close()
    模型 = _有依据模型('元', '元', 说明)

    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(模型(配置, 指令, 材料, 目录))
        if '作用范围' in 答 and Path(目录).name.startswith('B_'):
            答['列'].append({'列': 'F', '路径': '单位说明', '角色': '说明列'})
            if 缺陷 == '没有依据':
                答['作用范围']['单位依据'] = []
            else:
                if 缺陷 == '分列不一致':
                    答['作用范围']['单位依据'][0]['列'] = ['C']
                答['作用范围']['单位依据'].append(
                    {'来源': '表内', '位置': 'F1', '原文': '不存在的单位说明' if 缺陷 == '标注失实' else '单位：人民币万元', '片段': '人民币万元',
                     '列': ['D'], '单位': '万元', '币种': '人民币'})
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    结果 = _跑一份(tmp_path / '任务', a, b, _单组金标准(tmp_path))
    assert 读回Word(结果['Word']).表格() == 读回Word(a).表格()
    assert 读回Word(结果['Word']).段落 == 读回Word(a).段落


def test_期间无关日期不能作为独立期间(tmp_path):
    from 核心.复用证据 import _期间核对
    p = _表(tmp_path, '单日期', [('甲', 1, 2)], 首行=('项目', '2025-12-31', '', '单位：元'))
    项 = {'相对': [0, 1], '值': '2025-12-31', '期间角色': '本期期末', '角色词': ''}
    上下文 = {'tables': [{'range_name': 'TA_001', 'chapter_context':
                         ['公司成立于2010-01-01；担保到期于2025-12-31']} ]}
    assert _期间核对([项], 表视图(读快照(p), 'sheet1', (1, 1, 4, 2)), 上下文, 'TA_001')


def test_标题双日期骨架各自对应():
    from 核心.复用证据 import 按骨架取片段
    assert 按骨架取片段('上期{上期期末}，本期{本期期末}', '上期2023-12-31，本期2024-12-31') == [
        ('上期期末', '2023-12-31'), ('本期期末', '2024-12-31')]


def test_标题二遍命中三遍重认仍更新真实Word(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    库 = 安全记录目录('二次修复标题')
    金 = _单组金标准(tmp_path)
    说明 = '截至2024年12月31日，金额单位：人民币元'
    原 = _带说明Word(tmp_path, '原件', 说明)
    二A = tmp_path / '改表头.docx'
    with ZipFile(原) as zin, ZipFile(二A, 'w') as zout:
        for p in zin.infolist():
            data = zin.read(p.filename)
            if p.filename == 'word/document.xml':
                tree = E.fromstring(data)
                tree.findall('.//' + W + 'tbl/' + W + 'tr')[0].findall(W + 'tc')[1].find('.//' + W + 't').text = '期初余额'
                tree.findall('.//' + W + 'tbl/' + W + 'tr')[0].findall(W + 'tc')[2].find('.//' + W + 't').text = '期末2025-12-31'
                data = E.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            zout.writestr(p, data)

    def 模型(年):
        基 = _有依据模型('元', '元', 说明)
        def 调用(配置, 指令, 材料, 目录):
            答 = json.loads(基(配置, 指令, 材料, 目录))
            A = Path(目录).name.startswith('A_')
            if '表' in 答 and A:
                答['表'][0]['标题'] = {'原文': 说明, '日期': [
                    {'片段': '2024年12月31日', '期间角色': '本期期末'}]}
            if '列' in 答:
                for c in 答['列']:
                    if c.get('期间实际') and not A:
                        c['期间实际']['值'] = c['期间实际']['值'].replace('2025', str(年))
                if A and 年 == 2026:
                    答['列'][1]['期间实际'] = None
                    答['列'][2]['角色依据'] = {'位置': 'D2', '原文': '期末2025-12-31'}
            return json.dumps(答, ensure_ascii=False)
        return 调用
    monkeypatch.setattr('核心.识别.调用识别', 模型(2025))
    一 = _跑一份(tmp_path / '首次', 原, _B单位(tmp_path, '一.xlsx', '元', 100, 110), 金, 库)
    assert '截至2025年12月31日' in 读回Word(一['Word']).段落[0]
    二B = _B单位(tmp_path, '二.xlsx', '元', 901, 911)
    wb = load_workbook(二B)
    wb.active['C1'], wb.active['D1'] = '2026-01-01', '2026-12-31'
    wb.save(二B); wb.close()
    monkeypatch.setattr('核心.识别.调用识别', 模型(2026))
    二 = _跑一份(tmp_path / '第二次', 二A, 二B, 金, 库)
    日志 = json.loads((二['任务'] / '复用记录_A.json').read_text(encoding='utf-8'))
    assert [日志['调用'][步] for 步 in ('第一遍', '第二遍', '第三遍')] == [0, 0, 1], 日志
    assert 读回Word(二['Word']).表格()[1] == ['甲公司', '901', '911']
    assert 读回Word(二['Word']).段落[0] == '截至2026年12月31日，金额单位：人民币元'


def test_续表保留独立例外和作用范围():
    from 核心.复用证据 import _抽解释, 重定位
    from 核心.复用接入 import 应用精读
    主, 续 = _记录(行数=2), _记录(行数=2, 起=3, 续表序=2)
    续['例外'] = [{'范围': 'C4', '处理': '不搬', '原因': '续表独立例外'}]
    续['作用范围']['单位'] = '万元'
    续['行'][1]['槽位']['C'] = '独立编号'
    解释 = 重定位(_抽解释([主, 续], 1, 1), type('视图', (), {'顶': 1, '左': 1})())
    out = 应用精读([dict(主), dict(续)], 解释, 'V1', {'sheets': [{'name': 'sheet1', 'merges': []}]})
    assert out[1]['例外'] == [{'范围': 'C4:C4', '处理': '不搬', '原因': '续表独立例外'}]
    assert out[1]['作用范围']['单位'] == '万元'
    assert out[1]['行'][1]['槽位']['C'] == '独立编号'


def test_入库不能接受空编号值():
    from 核心.复用证据 import 可发布
    记 = _记录(行数=2)
    记['行'][1]['槽位']['C'] = ''
    assert not 可发布([记])[0]


def test_入库铺格才发现坏例外不能先生成候选(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'
    原模型 = _有依据模型('元', '元', 说明)
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原模型(配置, 指令, 材料, 目录))
        if '例外' in 答:
            答['例外'] = [{'范围': 'C3', '处理': '尚未决定'}]
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    出 = _跑一份(tmp_path / '任务', _带说明Word(tmp_path, '原件', 说明),
                _B单位(tmp_path, '底表.xlsx', '元', 901, 911), _单组金标准(tmp_path), 安全记录目录())
    assert not (出['任务'] / '候选发布_A.jsonl').read_text(encoding='utf-8').strip()
    assert not (出['任务'] / '候选发布_B.jsonl').read_text(encoding='utf-8').strip()


def test_可见命名表区保留内部空白和合并从属格(tmp_path):
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from 核心.表格 import 取可见集
    wb = Workbook(); s = wb.active; s.title = '底表'
    s['A1'], s['A3'], s['C1'], s['C3'] = '项目', '合计', '期末', 9
    s.merge_cells('A1:A2'); s['M88'] = '辅助'; s.column_dimensions['M'].hidden = True
    wb.defined_names.add(DefinedName('业务表', attr_text="'底表'!$A$1:$C$3"))
    p = tmp_path / '表区.xlsx'; wb.save(p); wb.close()
    v = 取可见集(读快照(p), '底表')
    assert v.含(1, 2) and v.含(2, 1)
    assert not v.含(88, 1)


def test_名单改名三遍复用且实际Word采用本次名称(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '元', 说明))
    库 = 安全记录目录(); 金 = _单组金标准(tmp_path)
    a = _带说明Word(tmp_path, '原件', 说明)
    _跑一份(tmp_path / '首次', a, _B单位(tmp_path, '一.xlsx', '元', 100, 110), 金, 库)
    b = _B单位(tmp_path, '二.xlsx', '元', 901, 911)
    wb = load_workbook(b); wb.active['B3'] = '丙公司'; wb.save(b); wb.close()
    出 = _跑一份(tmp_path / '改名', a, b, 金, 库)
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert [日志['调用'][步] for 步 in ('第一遍', '第二遍', '第三遍')] == [0, 0, 0], 日志
    assert 读回Word(出['Word']).表格()[2] == ['丙公司', '902', '912']


@pytest.mark.parametrize('换序', [False, True])
def test_期间双12月31日换年与对调真实Word(tmp_path, monkeypatch, 换序):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'; a = _带说明Word(tmp_path, '原件', 说明)
    金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    def 模型(年, 倒):
        原 = _有依据模型('元', '元', 说明)
        def 调用(配置, 指令, 材料, 目录):
            答 = json.loads(原(配置, 指令, 材料, 目录))
            if '列' in 答 and Path(目录).name.startswith('B_'):
                for i, c in enumerate(答['列'][1:3]):
                    初 = (i == 0) != 倒
                    c.update(期间角色='本期期初' if 初 else '本期期末',
                             期间实际={'值': f'{年-1 if 初 else 年}-12-31', '位置': c['列']+'1'})
            return json.dumps(答, ensure_ascii=False)
        return 调用
    for 次, 年 in enumerate((2025, 2026)):
        倒 = 换序 and 次 == 1
        b = _B单位(tmp_path, f'{次}.xlsx', '元', 901 if 倒 else 102, 102 if 倒 else 901)
        wb = load_workbook(b)
        wb.active['C1'], wb.active['D1'] = ((f'{年}-12-31', f'{年-1}-12-31') if 倒
                                            else (f'{年-1}-12-31', f'{年}-12-31'))
        wb.save(b); wb.close()
        monkeypatch.setattr('核心.识别.调用识别', 模型(年, 倒))
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    assert 读回Word(出['Word']).表格()[1] == ['甲公司', '102', '901']
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert [日志['调用'][步] for 步 in ('第一遍', '第二遍', '第三遍')] == [0, 0, int(换序)], 日志


def test_期间重认仍有未知整组保留真实Word(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '0')
    说明 = '公司成立于2010-01-01；担保到期于2025-12-31。金额单位：人民币元'
    a = _带说明Word(tmp_path, '原件', 说明)
    原 = _有依据模型('元', '元', 说明)
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        if '列' in 答 and Path(目录).name.startswith('B_'):
            答['行'][1]['期间角色按列'] = {'C': '未知'}
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    出 = _跑一份(tmp_path / '任务', a, _B单位(tmp_path, '二.xlsx', '元', 901, 911), _单组金标准(tmp_path))
    assert 读回Word(出['Word']).表格() == 读回Word(a).表格()
    assert 读回Word(出['Word']).段落 == 读回Word(a).段落


def test_续表独立例外与单位在真实Word保留(tmp_path, monkeypatch):
    import copy
    from 测试.单位样本 import 单位作用
    monkeypatch.setenv('审计附注_复用', '1')
    原件 = _单组Word(tmp_path)
    a = tmp_path / '主续表.docx'
    with ZipFile(原件) as zin, ZipFile(a, 'w') as zout:
        for p in zin.infolist():
            data = zin.read(p.filename)
            if p.filename == 'word/document.xml':
                tree = E.fromstring(data); body = tree.find(W+'body')
                表 = body.find(W+'tbl')
                表.find('.//'+W+'t').text = '项目（单位：人民币元）'
                续 = copy.deepcopy(表)
                续.find('.//'+W+'t').text = '项目（单位：人民币万元）'
                段 = copy.deepcopy(body.find(W+'p')); 段.find('.//'+W+'t').text = '续表（金额单位：人民币万元）'
                body.append(段); body.append(续)
                data = E.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            zout.writestr(p, data)
    金 = _单组金标准(tmp_path)
    gold = json.loads(金.read_text(encoding='utf-8')); gold['id'] = 'N002'; gold['slot']['name'] = '续表余额'
    金.write_text(金.read_text(encoding='utf-8')+'\n'+json.dumps(gold, ensure_ascii=False), encoding='utf-8')
    基 = _有依据模型('元', '元', '金额单位：人民币元')
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(基(配置, 指令, 材料, 目录)); A = Path(目录).name.startswith('A_')
        顶们 = [2, 8] if A else [1, 4]
        if '划分' in 答:
            答['划分'] = ([{'行范围': [1, 1], '科目': '非表格'}, {'行范围': [2, 6], '科目': '货币资金'},
                          {'行范围': [7, 7], '科目': '非表格'}, {'行范围': [8, 10], '科目': '货币资金'}] if A
                         else [{'行范围': [1, 6], '科目': '货币资金'}])
        elif '表' in 答:
            答['表'] = [{'行范围': [顶, 顶+2], '科目': '货币资金', '表名': '余额构成',
                         '表组': 'G1', '续表序': i+1} for i, 顶 in enumerate(顶们)]
        else:
            续 = Path(目录).name.endswith('002'); 顶 = 顶们[int(续)]; 原顶 = 2 if A else 1
            for c in 答['列']:
                if c.get('期间实际'): c['期间实际']['位置'] = c['列']+str(顶)
            for r in 答['行']:
                r['行'] += 顶-原顶
                if r.get('序号') is not None:
                    r['序号'] += 2*int(续)
                    r['槽位'] = {k: 'N002' if 续 else 'N001' for k in r['槽位']}
            单位 = '万元' if 续 else '元'
            答['作用范围'] = 单位作用(f'B{顶}' if A else f'E{顶}',
                f'项目（单位：人民币{单位}）' if A else '单位：人民币'+单位, ['C', 'D'], 单位)
            if 续 and A: 答['例外'] = [{'范围': f'D{顶+2}', '处理': '不搬', '原因': '续表独立不搬'}]
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    库 = 安全记录目录()
    for 次, 数 in enumerate((100, 901)):
        b = _B单位(tmp_path, f'{次}.xlsx', '元', 数, 数+10)
        wb = load_workbook(b); s = wb.active
        for 行 in range(1, 4):
            for 列 in range(2, 6): s.cell(行+3, 列, s.cell(行, 列).value)
        s['E4'] = '单位：人民币万元'
        wb.save(b); wb.close()
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    表 = 读回Word(出['Word'])
    assert 表.表格(0)[1] == ['甲公司', '901', '911']
    assert 表.表格(1)[1] == ['甲公司', '901', '911']
    assert 表.表格(1)[2] == ['乙公司', '200', '210']  # 既有规则：名单项有不搬格时整项保留
    for 侧 in ('A', 'B'):
        日志 = json.loads((出['任务'] / f'复用记录_{侧}.json').read_text(encoding='utf-8'))
        assert 日志['调用']['第三遍'] == 0, 日志


def test_标题双日期复用后真实Word分别更新(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '期初2024-01-01，期末2024-12-31，金额单位：人民币元'
    a = _带说明Word(tmp_path, '原件', 说明); 金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    原 = _有依据模型('元', '元', 说明)
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        if '表' in 答 and Path(目录).name.startswith('A_'):
            答['表'][0]['标题'] = {'原文': 说明, '日期': [
                {'片段': '2024-01-01', '期间角色': '本期期初'},
                {'片段': '2024-12-31', '期间角色': '本期期末'}]}
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    _跑一份(tmp_path / '首次', a, _B单位(tmp_path, '一.xlsx', '元', 100, 110), 金, 库)
    出 = _跑一份(tmp_path / '第二份', a, _B单位(tmp_path, '二.xlsx', '元', 901, 911), 金, 库)
    assert 读回Word(出['Word']).段落[0] == '期初2025-01-01，期末2025-12-31，金额单位：人民币元'
    assert 读回Word(出['Word']).表格()[1] == ['甲公司', '901', '911']
    日志 = json.loads((出['任务'] / '复用记录_A.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == 0, 日志


def test_坐标跨Z列AA至AD复用真实Word(tmp_path, monkeypatch):
    from openpyxl.utils import column_index_from_string, get_column_letter
    from 测试.单位样本 import 单位作用
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'; a = _带说明Word(tmp_path, '原件', 说明)
    原 = _有依据模型('元', '元', 说明); 金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    def 换列(c): return get_column_letter(column_index_from_string(c)+25)
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        if '列' in 答 and Path(目录).name.startswith('B_'):
            for c in 答['列']:
                c['列'] = 换列(c['列'])
                if c.get('期间实际'): c['期间实际']['位置'] = c['列']+'1'
            for r in 答['行']:
                if '槽位' in r: r['槽位'] = {换列(k): v for k,v in r['槽位'].items()}
            答['作用范围'] = 单位作用('AD1', '单位：人民币元', ['AB', 'AC'])
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    for 次, 数 in enumerate((100, 901)):
        b = _B单位(tmp_path, f'{次}.xlsx', '元', 数, 数+10)
        wb = load_workbook(b); wb.active.move_range('B1:E3', rows=0, cols=25); wb.save(b); wb.close()
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    assert 读回Word(出['Word']).表格()[1:] == [['甲公司', '901', '911'], ['乙公司', '902', '912']]
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == 0, 日志


def test_期间角色依据在第二层表头仍需逐列绑定(tmp_path):
    from openpyxl import Workbook
    p = tmp_path / '多层表头.xlsx'; wb = Workbook(); s = wb.active; s.title = 'sheet1'
    for row in [('项目', '2025-01-01', '2025-12-31', '单位：元'), ('', '期初余额', '期末余额', ''), ('甲', 1, 2, '')]: s.append(row)
    wb.save(p); wb.close()
    r = _记录(行数=3)
    r['行'][1] = {'行': 2, '角色': '表头'}
    for c, 字 in zip(r['列'][1:3], ['B', 'C']):
        c['角色依据'] = {'位置': 字+'2', '原文': '期初余额' if 字=='B' else '期末余额'}
    e = 建证据_精读('B', [r], 读快照(p), True)
    wb = load_workbook(p); wb.active['B2'], wb.active['C2'] = '期末余额', '期初余额'; wb.save(p); wb.close()
    out = 核对精读(e, 表视图(读快照(p), 'sheet1', (1,1,4,3), True))
    assert out['结果'] == '未命中'
    assert {x['表头相对'][0] for x in e['固定']['名单'] if x['类别']=='日期'} == {1}


@pytest.mark.parametrize('明确期间,日期位置', [(False, 'D1'), (True, 'D1'), (True, 'D99')])
def test_期间单日期只接受标明用途的独立期间(tmp_path, monkeypatch, 明确期间, 日期位置):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'; a = _带说明Word(tmp_path, '原件', 说明)
    金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    期间文 = ('报告期间2025-01-01至2025-12-31' if 明确期间
              else '公司成立于2010-01-01；担保到期于2025-12-31')
    原 = _有依据模型('元', '元', 说明)
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        if '列' in 答 and Path(目录).name.startswith('B_'):
            答['列'][1]['期间实际'] = None
            答['列'][2]['期间实际']['位置'] = 日期位置
            答['列'].append({'列': 'F', '路径': '说明', '角色': '说明列'})
            if 明确期间:
                答['列'][2]['期间依据'] = {'来源': '表内', '位置': 'F1', '原文': 期间文,
                    '起日片段': '2025-01-01', '止日片段': '2025-12-31'}
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    for 次, 数 in enumerate((100, 901)):
        b = _B单位(tmp_path, f'{次}.xlsx', '元', 数, 数+10)
        wb = load_workbook(b); wb.active['C1'] = '期初余额'; wb.active['F1'] = 期间文
        wb.save(b); wb.close()
        if 日期位置 == 'D99':
            with pytest.raises(RuntimeError, match='期间实际位置越界'):
                _跑一份(tmp_path / str(次), a, b, 金, 库)
            return
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    可确认 = 明确期间 and 日期位置 == 'D1'
    assert 读回Word(出['Word']).表格()[1] == (['甲公司', '901', '911'] if 可确认 else ['甲公司', '100', '110'])
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == (0 if 可确认 else 1), 日志


def test_入库坏组不发布而独立好组真实Word仍复用(tmp_path, monkeypatch):
    from 测试.test_真实桥接集成 import _三组样本, _三组金标准, _三组模型, _造B
    monkeypatch.setenv('审计附注_复用', '1')
    a = _三组样本(tmp_path); b = _造B(tmp_path); 金 = _三组金标准(tmp_path); 库 = 安全记录目录()
    名们 = ('货币资金', '应收账款', '存货')
    新a = tmp_path / '各组标题明确.docx'
    with ZipFile(a) as zin, ZipFile(新a, 'w') as zout:
        for p in zin.infolist():
            data = zin.read(p.filename)
            if p.filename == 'word/document.xml':
                tree = E.fromstring(data)
                for 表, 名 in zip(tree.findall('.//'+W+'body/'+W+'tbl'), 名们):
                    表.find('.//'+W+'t').text = 名+'（单位：人民币元）'
                data = E.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            zout.writestr(p, data)
    a = 新a
    wb = load_workbook(b)
    for 顶, 名 in zip((1, 4, 7), 名们):
        wb.active[f'A{顶}'] = 名+'（单位：人民币元）'
        wb.active[f'B{顶}'], wb.active[f'C{顶}'] = '2025-01-01', '2025-12-31'
    wb.active['B4'], wb.active['C4'] = '2025-12-31', '2025-06-30'
    wb.save(b); wb.close()
    原 = _三组模型()
    def 调用(配置, 指令, 材料, 目录):
        答 = json.loads(原(配置, 指令, 材料, 目录))
        if '划分' in 答 and Path(目录).name.startswith('A_'):
            答['划分'] = [x for 顶, 末, 科 in ((2,6,'货币资金'), (8,12,'应收账款'), (14,16,'存货'))
                         for x in ({'行范围':[顶-1,顶-1], '科目':'非表格'}, {'行范围':[顶,末], '科目':科})]
        if '列' in 答 and Path(目录).name.startswith('A_'):
            for c in 答['列']: c['期间实际'] = None  # A 是角色词表头，没有日期格。
            if Path(目录).name.endswith('001'):
                答['例外'] = [{'范围': 'C3', '处理': '未确认的例外'}]
        if '作用范围' in 答:
            i = int(Path(目录).name[-3:])-1
            答['作用范围']['单位依据'][0]['原文'] = 名们[i]+'（单位：人民币元）'
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 调用)
    for 次 in range(2):
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    日志 = json.loads((出['任务'] / '复用记录_A.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == 1, 日志
    assert 读回Word(出['Word']).表格(0)[1] == ['甲公司', '100', '110']
    assert 读回Word(出['Word']).表格(2)[1] == ['甲公司', '700', '710']
    from 核心.共用记录簿 import 记录簿
    项们 = 记录簿(库).快照()['解释'].values()
    assert not any(x['侧']=='A' and x['步骤']=='第三遍' and x['解释']['科目']=='货币资金' for x in 项们)


def test_名单数字形式固定标签改变必须重认(tmp_path):
    from 核心.复用证据 import 建证据_分科目, 核对分科目
    p = _表(tmp_path, '固定编号', [('1001', 1, 2)])
    记 = _记录(行数=2); 记['行'][1]['序号'] = None
    e = 建证据_分科目('B', [{'科目': '货币资金', '行范围': [1, 2]}], 读快照(p), True, 本侧=[记])
    wb = load_workbook(p); wb.active['A2'] = '1002'; wb.save(p); wb.close()
    assert 核对分科目(e, 读快照(p), True)['结果'] == '未命中'


def test_可见内容无预设表区仍从正常任务识别内部空白(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '0')
    说明 = '金额单位：人民币元'; a = _带说明Word(tmp_path, '原件', 说明)
    b = _B单位(tmp_path, '无表区.xlsx', '元', 901, 911)
    wb = load_workbook(b); s = wb.active; s['C3'] = None
    s['M88'] = '不可见辅助'; s.column_dimensions['M'].hidden = True
    wb.save(b); wb.close()
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '元', 说明))
    出 = _跑一份(tmp_path / '任务', a, b, _单组金标准(tmp_path))
    assert 读回Word(出['Word']).表格()[1] == ['甲公司', '901', '911']
    assert 读回Word(出['Word']).表格()[2] == ['乙公司', '', '912']
    格们 = [json.loads(x) for x in (出['任务']/'B合集.jsonl').read_text(encoding='utf-8').splitlines()]
    assert all('88' not in x['格'] for x in 格们)


def test_入库配对冲突即使两侧工作表同名仍不得发布(tmp_path, monkeypatch):
    from 核心.流程 import 运行任务
    from 核心.共用记录簿 import 记录簿
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'; a = _带说明Word(tmp_path, '原件', 说明)
    b = _B单位(tmp_path, '二.xlsx', '万元', 901, 911)
    wb = load_workbook(b); wb.active.title = 'sheet1'; wb.save(b); wb.close()
    monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '万元', 说明))
    库 = 安全记录目录()
    out = 运行任务({'word': str(a), 'b表': str(b), 'b工作表':'sheet1',
                  '金标准':str(_单组金标准(tmp_path)), '识别记录目录':str(库)}, tmp_path/'任务', {'测试':True})
    assert 读回Word(out['产物']['Word新副本']).表格() == 读回Word(a).表格()
    assert not any(x['步骤']=='第三遍' for x in 记录簿(库).快照()['解释'].values())


def test_单位材料保留证据字段和索引不让模型猜位置():
    from 核心.识别 import _表上下文
    上文 = {'tables': [{'range_name':'TA_001', 'title':'金额单位说明',
         'unit_text_candidates':['注册资本100万元', '本表金额单位：人民币元'],
         'preceding_blocks':[{'kind':'table', 'table_id':'前表'}, {'kind':'paragraph', 'text':'本表金额单位：人民币元'}]}]}
    材料 = _表上下文(上文, 'TA_001')
    assert 'unit_text_candidates' in 材料 and 'preceding_blocks' in 材料
    出 = json.loads(材料)
    assert 出['unit_text_candidates'][1] == '本表金额单位：人民币元'
    assert 出['preceding_blocks'][0]['table_id'] == '前表'


def test_期间首认期初晚于期末仍整组保留真实Word(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '0')
    说明 = '金额单位：人民币元'
    a = _带说明Word(tmp_path, '原件', 说明)
    b = _B单位(tmp_path, '矛盾日期.xlsx', '元', 901, 911)
    wb = load_workbook(b)
    wb.active['C1'], wb.active['D1'] = '2025-12-31', '2025-01-01'
    wb.save(b); wb.close()
    基 = _有依据模型('元', '元', 说明)
    def 模型(配置, 指令, 材料, 目录):
        答 = json.loads(基(配置, 指令, 材料, 目录))
        if '列' in 答 and Path(目录).name.startswith('B_'):
            答['列'][1]['期间实际']['值'] = '2025-12-31'
            答['列'][2]['期间实际']['值'] = '2025-01-01'
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 模型)
    出 = _跑一份(tmp_path / '任务', a, b, _单组金标准(tmp_path))
    assert 读回Word(出['Word']).表格() == 读回Word(a).表格()
    assert 读回Word(出['Word']).段落 == 读回Word(a).段落


def test_普通表头仅改合计仍全复用且真实Word合计为本次值(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    说明 = '金额单位：人民币元'
    原 = _带说明Word(tmp_path, '原件', 说明)
    a = tmp_path / '带合计.docx'
    with ZipFile(原) as zin, ZipFile(a, 'w') as zout:
        for p in zin.infolist():
            data = zin.read(p.filename)
            if p.filename == 'word/document.xml':
                tree = E.fromstring(data)
                tree.findall('.//' + W + 'tbl/' + W + 'tr')[2].find(W + 'tc').find('.//' + W + 't').text = '合计'
                data = E.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
            zout.writestr(p, data)
    基 = _有依据模型('元', '元', 说明)
    def 模型(配置, 指令, 材料, 目录):
        答 = json.loads(基(配置, 指令, 材料, 目录))
        if '列' in 答:
            答['行'][-1].update(名字='合计', 角色='合计', 序号=None)
            if Path(目录).name.startswith('B_'):
                for c in 答['列'][1:3]: c['期间实际'] = None
        return json.dumps(答, ensure_ascii=False)
    monkeypatch.setattr('核心.识别.调用识别', 模型)
    金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    for 次, 合计 in enumerate((111, 1234)):
        b = _B单位(tmp_path, f'{次}.xlsx', '元', 100, 110)
        wb = load_workbook(b); s = wb.active
        s['C1'], s['D1'], s['B3'], s['D3'] = '期初余额', '期末余额', '合计', 合计
        wb.save(b); wb.close()
        出 = _跑一份(tmp_path / str(次), a, b, 金, 库)
    assert 读回Word(出['Word']).表格()[1] == ['甲公司', '100', '110']
    assert 读回Word(出['Word']).表格()[2] == ['合计', '101', '1234']
    日志 = json.loads((出['任务'] / '复用记录_B.json').read_text(encoding='utf-8'))
    assert [日志['调用'][步] for 步 in ('第一遍', '第二遍', '第三遍')] == [0, 0, 0], 日志


def test_单位依据改变后新记录能被第三份任务复用(tmp_path, monkeypatch):
    monkeypatch.setenv('审计附注_复用', '1')
    金 = _单组金标准(tmp_path); 库 = 安全记录目录()
    for 次, 资本 in enumerate((100, 200, 200)):
        说明 = f'注册资本{资本}万元，本表金额单位：人民币元'
        a = _带说明Word(tmp_path, f'第{次}份', 说明)
        monkeypatch.setattr('核心.识别.调用识别', _有依据模型('元', '元', 说明))
        出 = _跑一份(tmp_path / str(次), a, _B单位(tmp_path, f'{次}.xlsx', '元', 901, 911), 金, 库)
    assert 读回Word(出['Word']).表格()[1] == ['甲公司', '901', '911']
    日志 = json.loads((出['任务'] / '复用记录_A.json').read_text(encoding='utf-8'))
    assert 日志['调用']['第三遍'] == 0, 日志
