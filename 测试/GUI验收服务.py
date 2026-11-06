# coding: utf-8
"""隔离的真实GUI验收服务；仅固定模型答复、数据目录及自动化文件选择路径。

文件选择的确定路径仅用于自动化；原生Windows窗口另行实测。结果接口保持产品实现。
"""
import argparse
import json
import os
import sys
import threading
import time
from pathlib import Path

根 = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(根))
from openpyxl import load_workbook
from 测试.组合样本 import 造Word, 造Excel, 造金标准
from 测试.test_Excel公式 import _模型, _造A
from 测试.test_复核五项修复 import _标题任务


def 写JSON(p, 值):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(值, ensure_ascii=False, indent=2), encoding="utf-8")


def 准备(证据):
    输入目录 = 证据 / "输入"
    输入目录.mkdir(parents=True, exist_ok=True)
    金 = 造金标准(输入目录 / "金标准.jsonl")
    A文 = 造Word(输入目录 / "待填报告.docx", 表内单位=True, 表前单位=False)
    B文 = 造Word(输入目录 / "来源报告.docx", 期初=901, 期末=102, 名单=("本次甲", "本次乙"), 表内单位=True, 表前单位=False)
    A表 = _造A(输入目录 / "待填报表.xlsx")
    wb = load_workbook(A表)
    wb.create_sheet("另表未选")["A1"] = "未选内容保持原样"
    wb.create_sheet("隐藏工作表")["A1"] = "隐藏内容保持原样"
    wb["隐藏工作表"].sheet_state = "hidden"
    wb.save(A表); wb.close()
    B表 = 造Excel(输入目录 / "本次数据.xlsx", {"余额": [["项目（单位：人民币元）","2025-01-01","2025-12-31"],["本次甲",901,102],["本次乙",1001,202],["合计",1902,304]]}, 起列=2)
    场景 = []
    for a in (A表, A文):
        for b in (B表, B文):
            场景.append({"名称": a.suffix + "←" + b.suffix,
                        "输入": {"A":{"路径":str(a),"工作表":["资金"] if a.suffix=='.xlsx' else []},
                                "B":{"路径":str(b),"工作表":["余额"] if b.suffix=='.xlsx' else []},
                                "金标准":str(金),"复用":False}})
    标题入, 标题模型 = _标题任务(输入目录, "界面标题", 记录=证据/"识别记录")
    运行 = 证据 / "任务记录"
    for n in range(100):
        目录 = 运行 / f"任务_性能样本_{n:03}"
        目录.mkdir(parents=True, exist_ok=True)
        是大表 = n == 99
        写JSON(目录/"任务说明.json", {"类型":"自动搬运","执行次数":1,"创建时间":"2026-11-06T12:00:00",
             "原始输入":{"A":{"路径":"性能样本_100张业务表_10000条明细.docx" if 是大表 else f"合成历史{n}.docx"},"B":{"路径":"合成来源.xlsx"}}})
        索引 = {"步骤":"done","产物":{}}
        if 是大表:
            明细, 结构 = [], []
            for 组 in range(100):
                结构.append({"侧":"A","读取部分":"P001","表组":f"G{组}","科目":f"性能科目{组}","表名":"余额构成","工作表":f"工作表{组}","范围":"B1:D101"})
                for 行 in range(100):
                    明细.append({"类型":"格子","task_id":目录.name,"执行次数":1,"侧":"A","读取部分":"P001","表组":f"G{组}","科目":f"性能科目{组}","表名":"余额构成","a格":f"工作表{组}!C{行+2}","b格":f"来源{组}!C{行+2}","原值":行,"原值状态":"值","来源值":行+1000,"来源值状态":"值","实际值":行+1000,"实际值已取得":True,"最终动作":"已写入","行路径":[f"名单{行:03} — 较长的往来单位名称用于检查显示"],"列路径":["期末余额"],"维度值":{"单位":"元"},"原因引用":[]})
            for 名, 条们 in (("交付明细.jsonl",明细),("结构清单.jsonl",结构)):
                (目录/名).write_text(''.join(json.dumps(x,ensure_ascii=False)+'\n' for x in 条们),encoding='utf-8')
            索引['产物']['交付明细']=str(目录/'交付明细.jsonl')
        写JSON(目录/"接续索引.json",索引)
    from 服务 import web, 界面设置
    import 核心.识别
    web.运行根 = 运行
    web.应用.state.配置路径 = 证据 / "测试模型配置.json"
    界面设置.设置路径 = 证据 / "界面设置.json"
    写JSON(界面设置.设置路径, {"金标准":str(金),"识别记录目录":str(证据/'识别记录'),"复用":False,"字号":14})
    web.读取模型配置 = lambda p: {"url":"固定答复","model":"GUI验收固定答复"}
    os.environ.pop('审计附注_复用',None)
    os.environ['审计附注_识别记录目录']=str(证据/'识别记录')
    常规模型 = _模型()
    def 模型(配置, 指令, 材料, 目录):
        任务 = next((p for p in Path(目录).parents if p.name.startswith('任务_')),None)
        说明 = json.loads((任务/'任务说明.json').read_text(encoding='utf-8')) if 任务 else {}
        用标题 = (说明.get('原始输入') or {}).get('金标准') == 标题入['金标准']
        with (证据/'实际模型调用.jsonl').open('a',encoding='utf-8') as f:
            f.write(json.dumps({'任务':str(任务),'目录':str(目录),'步骤':Path(目录).name},ensure_ascii=False)+'\n')
        return (标题模型 if 用标题 else 常规模型)(配置, 指令, 材料, 目录)
    核心.识别.调用识别 = 模型
    return {"场景":场景,"标题场景":标题入,"性能任务":"任务_性能样本_099","目录":str(运行)}


if __name__ == '__main__':
    解析=argparse.ArgumentParser();解析.add_argument('--目录',required=True);解析.add_argument('--窗口',action='store_true')
    参数=解析.parse_args();证据=Path(参数.目录).resolve();证据.mkdir(parents=True,exist_ok=True)
    清单=准备(证据)
    from 服务 import web
    if 参数.窗口:
        写JSON(证据/'样本清单.json',清单)
        from 桌面窗口 import 启动桌面
        启动桌面()
    else:
        from 桌面窗口 import 本机服务
        服务=本机服务()
        选择队列=[]
        # 设计T9允许自动化先用确定路径，原生选择另验；财务执行与结果接口不替换。
        web._选文件=lambda 类型='文件': 选择队列.pop(0) if 选择队列 else ''
        @web.应用.post('/验收/选择队列')
        def 设置队列(请求:dict):
            选择队列[:]=请求['路径'];return {'ok':True}
        @web.应用.post('/验收/清展示缓存')
        def 清展示缓存():
            from 服务.结果展示 import _缓存
            _缓存.clear()
            return {'ok': True}
        @web.应用.post('/验收/停止')
        def 停止():
            from 服务.任务服务 import 正在执行
            if 正在执行(web):
                return {'ok':False,'错误':'任务尚未安全结束'}
            服务.服务器.should_exit=True;return {'ok':True}
        服务.启动()
        写JSON(证据/'服务就绪.json',{**清单,'地址':服务.地址,'令牌':服务.令牌})
        try:
            服务.线程.join()
        finally:
            服务.停止()
