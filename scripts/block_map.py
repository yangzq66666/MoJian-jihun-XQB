# -*- coding: utf-8 -*-
"""
解析本地「机混需求表.xlsx」的 sheet「项目需求_项目路径_爆款拆分」，
输出「运营 -> 项目块」的结构。

规则（用户约定）：
  - 运营名单独占一行，写在 A 列，形如「运营甲项目」
  - 该行下面直到下一个运营名前，都是这个运营的项目
  - 运营块之间留一个空行
"""
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ns = {'m': NS[1:-1]}
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
from config import XLSX, SHEET_NAME as SHEET


def load(path=XLSX, sheet=SHEET):
    z = zipfile.ZipFile(path)
    ss = [''.join(x.text or '' for x in si.iter(NS + 't'))
          for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', ns)]
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    tgt = None
    for s in ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', ns):
        if s.get('name') == sheet:
            tgt = rels[s.get(R + 'id')]
    if tgt is None:
        raise SystemExit('找不到工作表: ' + sheet)
    root = ET.fromstring(z.read('xl/' + tgt))

    def col2num(c):
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        return n

    def cellval(c):
        ty = c.get('t')
        v = c.find('m:v', ns)
        i = c.find('m:is', ns)
        if ty == 'inlineStr' and i is not None:
            return ''.join(x.text or '' for x in i.iter(NS + 't'))
        if v is None:
            return ''
        return ss[int(v.text)] if ty == 's' else v.text

    rows = {}
    for row in root.find('m:sheetData', ns).findall('m:row', ns):
        r = int(row.get('r'))
        cells = {}
        for c in row.findall('m:c', ns):
            col = re.match(r'[A-Z]+', c.get('r')).group()
            vv = cellval(c)
            if str(vv).strip():
                cells[col2num(col)] = vv
        rows[r] = cells
    return rows


def block_map(rows):
    """返回 [{name, header_row, start_row, end_row, projects:[...]}]"""
    blocks = []
    cur = None
    for r in sorted(rows):
        cells = rows[r]
        if r == 1:
            continue
        a = str(cells.get(1, '')).strip()
        others = [cells.get(i) for i in (2, 3, 4, 5)]
        has_other = any(str(x).strip() for x in others if x is not None)
        # 运营表头：A 列有值，B~E 全空，且不以「项目」以外的形态出现
        if a and not has_other:
            cur = {'name': a, 'header_row': r, 'start_row': r + 1,
                   'end_row': r, 'projects': []}
            blocks.append(cur)
            continue
        if not a:
            continue
        if cur is None:
            cur = {'name': '(表头前)', 'header_row': None, 'start_row': r,
                   'end_row': r, 'projects': []}
            blocks.append(cur)
        cur['projects'].append({
            'row': r,
            '项目': a,
            '数量': str(cells.get(2, '')).strip(),
            '路径': str(cells.get(3, '')).strip(),
            '部门': str(cells.get(4, '')).strip(),
            '抽取片段': str(cells.get(5, '')).strip(),
        })
        cur['end_row'] = max(cur['end_row'], r)
    return blocks


def fmt(blocks):
    out = []
    for b in blocks:
        out.append('')
        out.append('【%s】 表头行=%s  项目行=%s~%s  共%d项' %
                   (b['name'], b['header_row'], b['start_row'], b['end_row'], len(b['projects'])))
        for p in b['projects']:
            qty = p['数量'] or '-'
            num = ''
            m = re.search(r'0-魔剪-素材\\(\d+)-', p['路径'])
            if m:
                num = '编号前缀' + m.group(1)
            out.append('   行%-3d %-14s 数量=%-4s %s' % (p['row'], p['项目'], qty, num))
    return '\n'.join(out)


if __name__ == '__main__':
    config.require_xlsx()
    rows = load()
    blocks = block_map(rows)
    print(fmt(blocks))
    out_json = os.path.join(config.STATE_DIR, 'block_map.json')
    os.makedirs(config.STATE_DIR, exist_ok=True)
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(blocks, f, ensure_ascii=False, indent=2)
    print('\n已写出:', out_json)
