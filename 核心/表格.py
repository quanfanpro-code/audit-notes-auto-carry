# coding: utf-8
"""Excel 快照读取：抄自老项目 附注更新/表格.py（read_workbook 及依赖，逐字保留）。"""
from __future__ import annotations
import ast
import copy
import datetime
import hashlib
import json
import posixpath
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries, quote_sheetname

_RECALCULATED = {}
_CURRENCY = {"RMB":"CNY", "人民币":"CNY", "人民币元":"CNY", "美元":"USD", "港元":"HKD", "港币":"HKD", "欧元":"EUR"}
_REFERENCE = re.compile(r"^(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?(?P<a>\$?[A-Z]{1,3}\$?[1-9]\d*)(?::(?P<b>\$?[A-Z]{1,3}\$?[1-9]\d*))?$",re.I)

class WorkbookError(ValueError):
    """不能安全读取或更新。"""

def _close_workbook(workbook):
    """Workbook.close 不释放 keep_vba 的内存归档，须在缓冲区回收前关闭。"""
    if workbook is None:return
    try:workbook.close()
    finally:
        archive=getattr(workbook,"vba_archive",None)
        if archive is not None:archive.close()

def file_hash(path):
    h=hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda:f.read(1024*1024),b""): h.update(block)
    return h.hexdigest()

def _definitions(w):
    yield from w.defined_names.values()
    for s in w: yield from s.defined_names.values()

def _destinations(d):
    try:return list(d.destinations)
    except (TypeError,ValueError,AttributeError):return []

def _external(w):
    result=["外部工作簿链接"] if getattr(w,"_external_links",[]) else []
    for s in w:
        for c in s._cells.values():
            if c.data_type=="f" and (re.search(r"\[[^]]+\][^+*/(),;]*!",str(c.value)) or "|" in str(c.value) or re.search(r"\b(?:WEBSERVICE|RTD|CUBE\w*)\s*\(",str(c.value),re.I)):
                result.append(f"{s.title}!{c.coordinate}")
    return result

def _word_omissions(context, workbook):
    """桥接声明的空位在 Word 中没有单元格，不能作为财务空白识别。"""
    tables=context.get("tables",[]) if isinstance(context,dict) else context
    if not isinstance(tables,(list,tuple)):return {}
    result={}
    for table in tables:
        if not isinstance(table,dict) or "structural_omissions" not in table:continue
        entries=table["structural_omissions"]
        if not isinstance(entries,list):raise WorkbookError("Word 结构省略信息必须为列表")
        for item in entries:
            if not isinstance(item,dict):raise WorkbookError("Word 结构省略信息无效")
            sheet=item.get("sheet");area=str(item.get("range","")).replace("$","").upper()
            reason=item.get("reason")
            if not isinstance(sheet,str) or sheet not in workbook.sheetnames:
                raise WorkbookError("Word 结构省略信息引用未知工作表")
            if table.get("sheet") not in (None,sheet):
                raise WorkbookError("Word 结构省略位置与所属表格不一致")
            if not isinstance(reason,str) or not reason.strip():
                raise WorkbookError("Word 结构省略信息缺少原因")
            try:left,top,right,bottom=range_boundaries(area)
            except (ValueError,TypeError):raise WorkbookError("Word 结构省略范围无效") from None
            if not all(type(v) is int for v in (left,top,right,bottom)) or not (1<=left<=right<=16384 and 1<=top<=bottom<=1048576):
                raise WorkbookError("Word 结构省略范围无效")
            record={"sheet":sheet,"range":area,"reason":reason}
            records=result.setdefault(sheet,[])
            if record not in records:records.append(record)
    return result

def _word_table_ranges(context, workbook, names):
    """Word 表边界只是载体事实；当前命名区域优先于提取时的位置。"""
    tables=context.get("tables",[]) if isinstance(context,dict) else context
    if not isinstance(tables,(list,tuple)):return None
    tables=[t for t in tables if isinstance(t,dict) and t.get("table_id") and t.get("range_name") and t.get("sheet")]
    if not tables:return None
    result={}
    defined_names={d.name for d in _definitions(workbook)}
    for table in tables:
        sheet=table["sheet"];name=table["range_name"]
        if sheet not in workbook.sheetnames:raise WorkbookError("Word 表格范围引用未知工作表")
        if name in defined_names:
            areas={e["range"] for e in names.get(name,[]) if e["sheet"]==sheet}
            if len(areas)!=1:raise WorkbookError("Word 表格命名区域缺失、歧义或与工作表不一致")
            area=next(iter(areas)).replace("$","").upper()
            try:left,top,right,bottom=range_boundaries(area)
            except (ValueError,TypeError):raise WorkbookError("Word 表格命名区域不是有效矩形") from None
        else:
            top=table.get("first_row");left=table.get("first_column")
            rows=table.get("row_count");cols=table.get("column_count")
            if not all(type(v) is int and v>0 for v in (top,left,rows,cols)):
                raise WorkbookError("Word 表格缺少当前命名区域及有效的原始位置")
            right=left+cols-1;bottom=top+rows-1
        if not all(type(v) is int for v in (left,top,right,bottom)) or not (1<=left<=right<=16384 and 1<=top<=bottom<=1048576):
            raise WorkbookError("Word 表格物理范围无效")
        if (right-left+1)*(bottom-top+1)>500000:raise WorkbookError("Word 表格物理范围超过50万格，请核实位置关联")
        area=f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}"
        record={"table_id":table["table_id"],"range_name":name,"range":area}
        if record not in result.setdefault(sheet,[]):result[sheet].append(record)
    return result


def _empty_formula_cached_strings(path):
    """openpyxl把已保存的空字符串缓存也读成None；只按原包的明确类型恢复。"""
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    relns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    result={}
    with ZipFile(path) as archive:
        relationships={r.get('Id'):r for r in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        sheets=ET.fromstring(archive.read('xl/workbook.xml')).find(ns+'sheets')
        for sheet in sheets:
            relation=relationships[sheet.get(relns+'id')]
            if not relation.get('Type','').endswith('/worksheet'):continue
            target=relation.get('Target','')
            member=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
            xml=ET.fromstring(archive.read(member))
            addresses=set()
            for cell in xml.iter(ns+'c'):
                value=cell.find(ns+'v')
                if cell.get('t')=='str' and cell.find(ns+'f') is not None and value is not None and value.text in (None,''):
                    addresses.add(cell.get('r'))
            result[sheet.get('name')]=addresses
    return result


def read_workbook(path,context=None):
    path=Path(path).resolve()
    before=file_hash(path)
    w=load_workbook(path,data_only=False,keep_vba=path.suffix.lower()==".xlsm",keep_links=True)
    cache=None
    out={"path":str(path),"sha256":before,"sheets":[],"names":{},"context":context if context is not None else []}
    try:
        omissions_by_sheet=_word_omissions(out["context"],w)
        cache=load_workbook(path,data_only=True,keep_links=True)
        empty_strings=_empty_formula_cached_strings(path)
        for d in _definitions(w):
            for sheet,address in _destinations(d):out["names"].setdefault(d.name,[]).append({"sheet":sheet,"range":address})
        word_ranges=_word_table_ranges(out["context"],w,out["names"])
        for s in w:
            physical_boxes=[range_boundaries(item["range"]) for item in (word_ranges or {}).get(s.title,[])]
            omissions=omissions_by_sheet.get(s.title,[])
            omission_boxes=[range_boundaries(item["range"]) for item in omissions]
            def omitted(c):return any(l<=c.column<=r and t<=c.row<=b for l,t,r,b in omission_boxes)
            for m in s.merged_cells.ranges:
                if any(l<=m.max_col and r>=m.min_col and t<=m.max_row and b>=m.min_row for l,t,r,b in omission_boxes):
                    raise WorkbookError(f"{s.title}!{m} 与 Word 结构省略范围相交")
            for c in s._cells.values():
                if (c.value is not None or c.data_type=="f") and omitted(c):
                    raise WorkbookError(f"{s.title}!{c.coordinate} 在 Word 结构省略位置存在实际值或公式")
            anchors=[(c.row,c.column) for c in s._cells.values() if not isinstance(c,MergedCell) and c.value is not None]
            boxes=[]
            for entries in out["names"].values():
                for entry in entries:
                    if entry["sheet"]==s.title:
                        try:
                            box=range_boundaries(entry["range"])
                            if all(isinstance(x,int) for x in box):boxes.append(box)
                        except ValueError:pass
            for m in s.merged_cells.ranges:
                if s.cell(m.min_row,m.min_col).value is not None:boxes.append((m.min_col,m.min_row,m.max_col,m.max_row))
            max_row=max([r for r,c in anchors]+[b[3] for b in boxes+physical_boxes]+[1])
            max_col=max([c for r,c in anchors]+[b[2] for b in boxes+physical_boxes]+[1])
            hidden=[(d.min or column_index_from_string(k),d.max or column_index_from_string(k)) for k,d in s.column_dimensions.items() if d.hidden]
            def visible(c):return not isinstance(c,MergedCell) and not omitted(c)
            cells={}
            for c in list(s._cells.values()):
                if not visible(c):continue
                if c.value is None and (not c.has_style or c.row>max_row or c.column>max_col):continue
                formula=str(c.value) if c.data_type=="f" else None
                cells[c.coordinate]={"row":c.row,"column":c.column,"value":c.value,"formula":formula,"cached_value":cache[s.title][c.coordinate].value if formula else c.value,"number_format":c.number_format,"style_id":c.style_id}
                if formula and cells[c.coordinate]['cached_value'] is None and c.coordinate in empty_strings.get(s.title,set()):
                    cells[c.coordinate]['cached_value']=''
            # 仅展开确定的表格范围，不按污染后的 worksheet.max_row 扫描。
            for left,top,right,bottom in boxes:
                if (right-left+1)*(bottom-top+1)>500000:raise WorkbookError("命名业务区域超过50万格，请核实是否包含无用空白格式")
                for row in range(top,bottom+1):
                    for col in range(left,right+1):
                        c=s.cell(row,col)
                        if visible(c) and c.coordinate not in cells:
                            cells[c.coordinate]={"row":row,"column":col,"value":c.value,"formula":None,"cached_value":c.value,"number_format":c.number_format,"style_id":c.style_id}
            for c in cells.values():
                c["hidden_row"]=bool(s.row_dimensions.get(c["row"]) and s.row_dimensions[c["row"]].hidden)
                c["hidden_column"]=any(lo<=c["column"]<=hi for lo,hi in hidden)
                c["hidden_sheet"]=s.sheet_state!="visible"
                c["hidden"]=c["hidden_row"] or c["hidden_column"] or c["hidden_sheet"]
            out["sheets"].append({"name":s.title,"hidden":s.sheet_state!="visible","hidden_rows":[r for r,d in s.row_dimensions.items() if d.hidden],"hidden_columns":hidden,"max_row":max_row,"max_column":max_col,"merges":[str(m) for m in s.merged_cells.ranges],"cells":cells})
            if omissions:out["sheets"][-1]["structural_omissions"]=copy.deepcopy(omissions)
            if word_ranges is not None:out["sheets"][-1]["word_table_ranges"]=copy.deepcopy(word_ranges.get(s.title,[]))
        out["external_references"]=_external(w)
        if out["sha256"] in _RECALCULATED:out.update(recalculated=True,recalculation=_RECALCULATED[out["sha256"]])
        if file_hash(path)!=before:raise WorkbookError("来源文件在读取期间发生变化，请重新读取")
        return out
    finally:
        try:_close_workbook(cache)
        finally:_close_workbook(w)




# —— 新计划接口别名 ——
def 读快照(xlsx路径, 上下文=None):
    return read_workbook(xlsx路径, 上下文)
