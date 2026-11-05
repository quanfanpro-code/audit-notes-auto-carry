# coding: utf-8
"""只呈现文件明确记录的网格和合并，不判断表头角色或业务身份。

原生Excel两端用「可见表格文字」：只呈现本次可见集合里的位置；隐藏行列、打印区域外一律标成不可见，
跨边界合并的起点不可见时只留结构标记，不补取隐藏起点的正文。
"""
import json
from openpyxl.utils.cell import range_boundaries, coordinate_to_tuple, get_column_letter


# 与已验证的文字识别要求逐字一致，日常不读取外部技能文件。
TABLE_STRUCTURE_READING = ' 二维网格保留原行列和明确合并覆盖；先读清本格所在行列及受哪些原表头约束，再选择金标准并给出本格身份。并入位置不是独立空白，不凭空白补表头，不把原文顺序当作已经确定的表头关系。表头角色及作用范围由你依据完整原文判断，程序没有预选。'


def grid_text(rows, columns, values, merges, coordinate):
    covered = {}
    anchors = {}
    for top, left, bottom, right in merges:
        anchor = coordinate(top, left)
        anchors.setdefault((top, left), []).append(coordinate(top, left) + ':' + coordinate(bottom, right))
        for row in range(top, bottom + 1):
            for column in range(left, right + 1):
                if (row, column) != (top, left):
                    covered.setdefault((row, column), []).append(anchor)
    lines = ['| 行号 | ' + ' | '.join(columns.values()) + ' |',
             '| --- | ' + ' | '.join('---' for _ in columns) + ' |']
    for row in rows:
        fields = []
        for column in columns:
            at = (row, column)
            if at in covered:
                text = '并入' + '、'.join(covered[at])
                if at in values and values[at] not in (None, ''):
                    text += '；本位置另有原记录：' + json.dumps(values[at], ensure_ascii=False, default=str)
            elif at not in values:
                text = '未记录此位置的内容'
            else:
                text = coordinate(row, column) + ' ' + json.dumps(values[at], ensure_ascii=False, default=str)
                if values[at] in (None, ''):
                    text += '〔原格空白〕'
                if at in anchors:
                    text += '〔合并范围：' + '、'.join(anchors[at]) + '〕'
            fields.append(text.replace('|', '\\|').replace('\r', '\\r').replace('\n', '\\n'))
        lines.append('| ' + str(row) + ' | ' + ' | '.join(fields) + ' |')
    return '\n'.join(lines)


def workbook_table_ranges(snapshot, areas):
    """读取已有物理表区，不判断内容含义。"""
    tables = []
    for area in areas:
        candidates = []
        left, top, right, bottom = range_boundaries(area['range'])
        for name, named in snapshot.get('names', {}).items():
            for part in named:
                if part['sheet'] != area['sheet']:
                    continue
                l, t, r, b = range_boundaries(part['range'])
                if l <= left and r >= right and t <= top and b >= bottom:
                    candidates.append((name, part['sheet'], part['range']))
        tables.extend(candidates or [('', area['sheet'], area['range'])])
    return list(dict.fromkeys(tables))


def workbook_table_text(snapshot, areas):
    """优先呈现既有物理表区；没有表区记录时使用明确阅读范围。"""
    tables = workbook_table_ranges(snapshot, areas)
    bounds = [range_boundaries(span) for _, _, span in tables]
    lines = ['以下为原文件二维网格。并入表示明确的合并覆盖，不是独立空白；原格空白与未记录分开列示。']
    for index, (name, sheet_name, span) in enumerate(tables):
        left, top, right, bottom = bounds[index]
        containing = [j for j, (_, sheet, _) in enumerate(tables) if sheet == sheet_name
                      and bounds[j][0] <= left and bounds[j][1] <= top
                      and bounds[j][2] >= right and bounds[j][3] >= bottom]
        parent = max(containing, key=lambda j: ((bounds[j][2]-bounds[j][0]+1)*(bounds[j][3]-bounds[j][1]+1), -j))
        if parent != index:
            lines.append('原表区：' + name + '；工作表：' + sheet_name + '；范围：' + span +
                         '；原内容已列在同工作表网格 ' + tables[parent][2] + ' 中。')
            continue
        sheet = next(s for s in snapshot['sheets'] if s['name'] == sheet_name)
        values = {coordinate_to_tuple(address): (cell.get('cached_value') if cell.get('formula') else cell.get('value'))
                  for address, cell in sheet['cells'].items()
                  if top <= cell['row'] <= bottom and left <= cell['column'] <= right}
        merges = []
        for merged in sheet['merges']:
            l, t, r, b = range_boundaries(merged)
            if l <= right and r >= left and t <= bottom and b >= top:
                merges.append((t, l, b, r))
                if (t, l) not in values and (cell := sheet['cells'].get(get_column_letter(l) + str(t))):
                    lines.append('范围外合并起点原文：' + sheet_name + '!' + get_column_letter(l) + str(t) +
                                 ' ' + json.dumps(cell.get('value'), ensure_ascii=False, default=str))
        lines.append('\n工作表：' + sheet_name + '；原表区：' + name + '；范围：' + span)
        linked = [{key: item[key] for key in ('table_id', 'range_name', 'title', 'link_error') if item.get(key) is not None}
                  for item in snapshot.get('context', [])
                  if name and isinstance(item, dict) and item.get('range_name') == name and item.get('sheet') == sheet_name]
        if linked:
            lines.append('原表号及标题：' + json.dumps(linked, ensure_ascii=False))
        lines.append(grid_text(range(top, bottom + 1),
                               {c: get_column_letter(c) for c in range(left, right + 1)}, values, merges,
                               lambda r, c: get_column_letter(c) + str(r)))
    return '\n'.join(lines)


# —— 新计划接口 ——
def 表格转文字(快照, 范围, 用可见范围=False):
    """范围内表格转保结构文字（复用 workbook_table_text）。范围={"sheet","range"}。

    用可见范围=True 时只呈现原生Excel本次可见集合里的位置：隐藏行列、打印区域外的内容
    一律标成不可见，隐藏起点的合并正文不补取。
    """
    if 用可见范围:
        return 可见表格文字(快照, 范围["sheet"], 范围["range"])
    return workbook_table_text(快照, [范围])


def 可见表格文字(快照, 工作表, 范围):
    """按本次可见集合呈现网格：每个位置要么可见（含原格空白），要么明确写不可见。"""
    from .表格 import 取可见集

    可见 = 取可见集(快照, 工作表)
    if 可见 is None:
        raise RuntimeError(f"工作表 {工作表} 没有可见范围记录，不能按可见范围出材料")
    if 可见.空():
        raise RuntimeError(f"工作表 {工作表} 没有可见内容：{可见.原因 or '可见范围为空白'}")
    左, 顶, 右, 底 = range_boundaries(str(范围))
    表 = next(s for s in 快照["sheets"] if s["name"] == 工作表)
    格 = 表.get("cells") or {}
    values = {}
    for cell in 格.values():
        if 顶 <= cell["row"] <= 底 and 左 <= cell["column"] <= 右 and 可见.含(cell["row"], cell["column"]):
            values[(cell["row"], cell["column"])] = cell

    并入, 锚范围 = {}, {}
    for 合并 in 表.get("merges") or []:
        m左, m顶, m右, m底 = range_boundaries(合并)
        if m右 < 左 or m左 > 右 or m底 < 顶 or m顶 > 底:
            continue
        锚 = (m顶, m左)
        成员 = [(r, c) for r in range(max(m顶, 顶), min(m底, 底) + 1)
                for c in range(max(m左, 左), min(m右, 右) + 1)]
        if 可见.含(*锚):
            锚范围[锚] = 合并
            for m in 成员:
                if m != 锚 and 可见.含(*m):
                    并入[m] = 合并
        else:
            for m in 成员:
                if 可见.含(*m):
                    并入[m] = ""                    # 起点不可见：只留结构标记，不取正文

    行们 = [r for r in 可见.行们() if 顶 <= r <= 底]
    列们 = [c for c in range(左, 右 + 1)]
    lines = ["以下为本工作表本次可见单元格；不可见表示隐藏行列或打印区域外，未参与识别。",
             '| 行号 | ' + ' | '.join(get_column_letter(c) for c in 列们) + ' |',
             '| --- | ' + ' | '.join('---' for _ in 列们) + ' |']
    for row in 行们:
        fields = []
        for column in 列们:
            at = (row, column)
            坐标 = get_column_letter(column) + str(row)
            if not 可见.含(row, column):
                fields.append('不可见（隐藏或打印区域外）')
                continue
            if at in 并入:
                if 并入[at]:
                    text = '并入' + 并入[at]
                else:
                    text = '合并从属（起点不可见，不取隐藏起点文字）'
                cell = values.get(at) or {}
                if cell.get('value') not in (None, '') or cell.get('formula'):
                    text += '；本位置另有原记录：' + json.dumps(
                        cell.get('value'), ensure_ascii=False, default=str)
                fields.append(text)
                continue
            cell = values.get(at)
            if cell is None:
                fields.append('未记录此位置的内容')
                continue
            if cell.get('formula'):
                缓存 = cell.get('cached_value')
                if 缓存 is None:
                    text = 坐标 + ' 〔公式无保存结果〕'
                elif 缓存 == '':
                    text = 坐标 + ' ' + json.dumps(cell.get('value'), ensure_ascii=False,
                                                 default=str) + '〔公式，缓存为空字符串〕'
                else:
                    text = 坐标 + ' ' + json.dumps(缓存, ensure_ascii=False, default=str)
            else:
                值 = cell.get('value')
                text = 坐标 + ' ' + json.dumps(值, ensure_ascii=False, default=str)
                if 值 in (None, ''):
                    text += '〔原格空白〕'
            if at in 锚范围:
                text += '〔合并范围：' + 锚范围[at] + '〕'
            fields.append(text)
        lines.append('| ' + str(row) + ' | ' + ' | '.join(
            x.replace('|', '\\|').replace('\r', '\\r').replace('\n', '\\n') for x in fields) + ' |')
    return '\n'.join(lines)
