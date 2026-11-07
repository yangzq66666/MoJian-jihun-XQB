# -*- coding: utf-8 -*-
"""
适配「没有运营块」的扁平版《机混需求表》：把新项目写进表里现成的空白行。

与 add_projects.py 的区别：不插行、不做行号平移，只填已有空行，
因此其它所有数据、行号、格式都不变（可直接过 check_insert.py 自检）。

用法：
  python append_flat_rows.py items.json            # 只预览
  python append_flat_rows.py items.json --apply   # 备份原表 -> 生成 _new_ -> 覆盖原表

items.json: [{"运营": "运营甲", "项目": "项目名", "数量": 30}, ...]
新行的 C 列（项目文件地址）留空，D/E 列沿用最后一个项目行的值。
"""
import datetime
import json
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config
from add_projects import resolve_sheet, parse_rows, row_cells, esc

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CELL_RE = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)


def style_of(raw):
    m = re.search(r'\bs="(\d+)"', raw)
    return ' s="%s"' % m.group(1) if m else ''


def shared_map(z):
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    return {str(i): ''.join(t.text or '' for t in si.iter(NS + 't'))
            for i, si in enumerate(root.findall(NS + 'si'))}


def a_text(raw, shared):
    v, t, _ = row_cells(raw).get('A', ('', '', ''))
    return str(shared.get(v, v) if t == 's' else v).strip()


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    config.require_xlsx()
    items = json.load(open(sys.argv[1], encoding='utf-8'))
    apply = '--apply' in sys.argv
    xlsx = config.XLSX

    z = zipfile.ZipFile(xlsx)
    sheet_path = resolve_sheet(z)
    xml = z.read(sheet_path).decode('utf-8')
    shared = shared_map(z)
    rows = parse_rows(xml)
    rowmap = {rn: raw for rn, _s, _e, raw in rows}

    occupied = [rn for rn, raw in rowmap.items() if a_text(raw, shared)]
    last_data = max(occupied)
    blanks = [rn for rn in sorted(rowmap) if rn > last_data and not a_text(rowmap[rn], shared)]
    if len(blanks) < len(items):
        raise SystemExit('空行不够：需要 %d 行，只有 %d 行' % (len(items), len(blanks)))

    # 模板行 = 最后一个有数据的项目行，样式与 D/E 值都沿用它
    tmpl_raw = rowmap[last_data]
    tmpl_cells = {}
    for tok in CELL_RE.findall(tmpl_raw):
        m = re.search(r'\br="([A-Z]+)\d+"', tok)
        if m:
            tmpl_cells[m.group(1)] = tok
    styles = {c: style_of(tmpl_cells.get(c, '')) for c in 'ABCDE'}

    print('模板行 %d，样式 %s' % (last_data, styles))
    print('将写入空行:', blanks[:len(items)])
    for rn, it in zip(blanks, items):
        print('  行%-3d %-26s %s' % (rn, it['项目'], it.get('数量', '')))
    if not apply:
        print('\n（预览模式，未写文件；加 --apply 落盘）')
        return

    changes = {}
    for rn, it in zip(blanks, items):
        cells = ['<c r="A%d"%s t="inlineStr"><is><t>%s</t></is></c>'
                 % (rn, styles['A'], esc(it['项目']))]
        qty = str(it.get('数量', '')).strip()
        cells.append('<c r="B%d"%s><v>%s</v></c>' % (rn, styles['B'], qty) if qty
                     else '<c r="B%d"%s/>' % (rn, styles['B']))
        cells.append('<c r="C%d"%s/>' % (rn, styles['C']))
        for col in ('D', 'E'):
            raw = tmpl_cells.get(col, '')
            if raw:
                cells.append(re.sub(r'(r="%s)\d+"' % col, r'\g<1>%d"' % rn, raw, count=1))
        m = re.match(r'(<row\b[^>]*>).*?(</row>)', rowmap[rn], re.S)
        if not m:
            raise SystemExit('行 %d 结构异常，无法替换' % rn)
        changes[rn] = m.group(1) + ''.join(cells) + m.group(2)

    for rn, new_raw in changes.items():
        xml = xml.replace(rowmap[rn], new_raw, 1)
    ET.fromstring(xml.encode('utf-8'))
    print('sheet XML 合法性校验: 通过')

    out = os.path.join(os.path.dirname(xlsx), '_new_机混需求表.xlsx')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info in z.infolist():
            zo.writestr(info, xml.encode('utf-8') if info.filename == sheet_path
                        else z.read(info.filename))
    z.close()
    with zipfile.ZipFile(out) as zz:
        assert zz.testzip() is None
    print('已生成:', out)

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    bakdir = os.path.join(config.STATE_DIR, 'backup')
    os.makedirs(bakdir, exist_ok=True)
    bak = os.path.join(bakdir, '机混需求表_备份_追加行前_%s.xlsx' % stamp)
    shutil.copy2(xlsx, bak)
    print('备份原表:', bak)
    shutil.copy2(out, xlsx)
    print('已覆盖原表:', xlsx)


if __name__ == '__main__':
    main()
