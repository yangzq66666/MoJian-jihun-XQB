# -*- coding: utf-8 -*-
"""
插入新项目后的自检：对比插入前/后的表，确认除新增行外没有动到任何数据。

用法：
  python check_insert.py <原表.xlsx> <新表.xlsx>
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ns = {'m': NS[1:-1]}
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import SHEET_NAME as SHEET


def read_sheet(path):
    z = zipfile.ZipFile(path)
    bad = z.testzip()
    ss = [''.join(t.text or '' for t in si.iter(NS + 't'))
          for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', ns)]
    rels = {r.get('Id'): r.get('Target')
            for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    tgt = None
    for s in ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', ns):
        if s.get('name') == SHEET:
            tgt = rels[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
    if tgt is None:
        raise SystemExit('找不到工作表: ' + SHEET)
    root = ET.fromstring(z.read('xl/' + tgt))

    def c2n(c):
        n = 0
        for ch in c:
            n = n * 26 + ord(ch) - 64
        return n

    out = []
    for row in root.find('m:sheetData', ns).findall('m:row', ns):
        r = int(row.get('r'))
        cells = {}
        for c in row.findall('m:c', ns):
            col = re.match(r'[A-Z]+', c.get('r')).group()
            ty = c.get('t')
            v = c.find('m:v', ns)
            i = c.find('m:is', ns)
            if ty == 'inlineStr' and i is not None:
                vv = ''.join(x.text or '' for x in i.iter(NS + 't'))
            elif v is None:
                vv = ''
            elif ty == 's':
                vv = ss[int(v.text)]
            else:
                vv = v.text
            if str(vv).strip():
                cells[c2n(col)] = str(vv).strip()
        mx = max(cells) if cells else 0
        out.append((r, tuple(cells.get(i, '') for i in range(1, mx + 1))))
    return out, bad


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return
    old, bad1 = read_sheet(sys.argv[1])
    new, bad2 = read_sheet(sys.argv[2])
    print('原表行数 %-4d 新表行数 %-4d' % (len(old), len(new)))
    print('zip 完整性: 原表 %s | 新表 %s'
          % ('OK' if bad1 is None else bad1, 'OK' if bad2 is None else bad2))

    old_seq = [x[1] for x in old]
    added = [x for x in new if x[1] not in old_seq]
    kept = [x[1] for x in new if x[1] in old_seq]
    print('\n新增行 %d 条:' % len(added))
    for r, cols in added:
        print('  第%-3d行 %s' % (r, ' | '.join(cols)))
    same = kept == old_seq
    print('\n原始数据是否逐行按原顺序保留:', '是' if same else '否')
    if not same:
        for i in range(max(len(old_seq), len(kept))):
            a = old_seq[i] if i < len(old_seq) else None
            b = kept[i] if i < len(kept) else None
            if a != b:
                print('  首个差异 @第%d条\n    原: %s\n    新: %s' % (i + 1, a, b))
                break


if __name__ == '__main__':
    main()
