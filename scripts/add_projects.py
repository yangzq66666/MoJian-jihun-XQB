# -*- coding: utf-8 -*-
"""
往本地「机混需求表.xlsx」的「项目需求_项目路径_爆款拆分」页插入新项目。

规则（用户约定）：
  - 新项目要放到它所属运营的块里（运营名下方，直到下一个运营名前）
  - 块之间保留一个空行

用法：
  python add_projects.py items.json            # 只预览，不写文件
  python add_projects.py items.json --apply    # 生成到 _new_ 文件（不直接覆盖原表）

items.json 格式：
  [{"运营": "运营丙", "项目": "新品名", "数量": 30,
    "路径": "U:\\...（可省略，省略则按块编号自动生成，需人工确认）",
    "抽取片段": "1-1-1-1-1", "部门": "默认部门"}]
"""
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

XLSX = config.XLSX
SHEET_TARGET = 'xl/worksheets/sheet3.xml'  # 由 sheet 名解析得到，见 resolve_sheet()
SHEET_NAME = config.SHEET_NAME

TEMPLATE_STYLES = {'A': '8', 'B': '8', 'C': '8', 'D': '31', 'E': '22'}


def esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def resolve_sheet(z):
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    m = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(SHEET_NAME), wb)
    if not m:
        raise SystemExit('找不到工作表 %s' % SHEET_NAME)
    rid = m.group(1)
    m2 = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid), rels)
    target = m2.group(1).lstrip('/')
    if not target.startswith('xl/'):
        target = 'xl/' + target
    return target


def parse_rows(xml):
    """返回 [(row_no, start_idx, end_idx, raw)]"""
    out = []
    for m in re.finditer(r'<row\b[^>]*>.*?</row>', xml, re.S):
        raw = m.group(0)
        rn = int(re.search(r'\br="(\d+)"', raw).group(1))
        out.append((rn, m.start(), m.end(), raw))
    return out


CELL_RE = re.compile(r'<c\b[^>]*/>|<c\b[^>]*>.*?</c>', re.S)


def row_cells(raw):
    """解析一行里的单元格：{列字母: (值, 类型, 样式)}"""
    cells = {}
    for tok in CELL_RE.findall(raw):
        m = re.match(r'<c\b([^>]*?)(/?)>', tok)
        attrs = m.group(1)
        body = tok[m.end():]
        ref = re.search(r'\br="([A-Z]+)\d+"', attrs)
        if not ref:
            continue
        col = ref.group(1)
        style = re.search(r'\bs="(\d+)"', attrs)
        t = re.search(r'\bt="([^"]+)"', attrs)
        v = re.search(r'<v>(.*?)</v>', body, re.S)
        isv = re.search(r'<is>.*?<t[^>]*>(.*?)</t>.*?</is>', body, re.S)
        cells[col] = (v.group(1) if v else (isv.group(1) if isv else ''),
                      t.group(1) if t else '', style.group(1) if style else '')
    return cells


def find_blocks(xml, shared):
    """运营块：A 列有值、B~E 全空的行 = 运营表头"""
    def txt(cell):
        v, t, _ = cell
        return shared.get(v, v) if t == 's' else v

    rows = parse_rows(xml)
    blocks = []
    cur = None
    for rn, s, e, raw in rows:
        if rn == 1:
            continue
        cells = row_cells(raw)
        a = str(txt(cells.get('A', ('', '', '')))).strip()
        rest = [txt(cells.get(k, ('', '', ''))) for k in ('B', 'C', 'D', 'E')]
        if a and not any(str(x).strip() for x in rest):
            cur = {'name': a, 'header_row': rn, 'last_row': rn, 'projects': []}
            blocks.append(cur)
            continue
        if a:
            if cur is None:
                cur = {'name': '(表头前)', 'header_row': None, 'last_row': rn, 'projects': []}
                blocks.append(cur)
            cur['projects'].append({'row': rn, '项目': a,
                                    '数量': str(txt(cells.get('B', ('', '', '')))).strip(),
                                    '路径': str(txt(cells.get('C', ('', '', '')))).strip()})
            cur['last_row'] = rn
    return blocks


def build_cell(ref, value, style=None, numeric=False):
    st = ' s="%s"' % style if style else ''
    if numeric and str(value).strip() != '':
        return '<c r="%s"%s><v>%s</v></c>' % (ref, st, value)
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, st, esc(value))


def build_row(r, item, styles):
    attrs = 'r="%d" spans="1:5" ht="15" x14ac:dyDescent="0.25"' % r
    c = []
    c.append(build_cell('A%d' % r, item['项目'], styles.get('A', '8')))
    c.append(build_cell('B%d' % r, item.get('数量', ''), styles.get('B', '8'), numeric=True))
    c.append(build_cell('C%d' % r, item.get('路径', ''), styles.get('C', '8')))
    c.append(build_cell('D%d' % r, item.get('部门', '默认部门'), styles.get('D', '31')))
    c.append(build_cell('E%d' % r, item.get('抽取片段', '1-1-1-1-1'), styles.get('E', '22')))
    return '<row %s>%s</row>' % (attrs, ''.join(c))


def shift_rows(xml, from_row, styles_from_row=None):
    """把所有 r >= from_row 的行整体下移一行（行号 +1）"""
    rows = parse_rows(xml)
    pieces = []
    last = 0
    for rn, s, e, raw in rows:
        if rn < from_row:
            continue
        pieces.append(xml[last:s])
        new = re.sub(r'\br="%d"' % rn, 'r="%d"' % (rn + 1), raw, count=1)
        new = re.sub(r'(<c\b[^>]*\br="[A-Z]+)%d"' % rn, r'\g<1>%d"' % (rn + 1), new)
        pieces.append(new)
        last = e
    pieces.append(xml[last:])
    return ''.join(pieces)


def preview(blocks, items):
    print('%-8s %-16s %-6s %s' % ('运营', '项目', '数量', '插入位置/路径'))
    print('-' * 92)
    for it in items:
        blk = next((b for b in blocks if b['name'].startswith(it['运营'])), None)
        if blk is None:
            print('%-8s %-16s %-6s !! 找不到运营块，可用块: %s'
                  % (it['运营'], it['项目'], it.get('数量', ''), [b['name'] for b in blocks]))
            continue
        print('%-8s %-16s %-6s 插到「%s」块末尾（第 %d 行之后）  %s'
              % (it['运营'], it['项目'], it.get('数量', ''), blk['name'],
                 blk['last_row'], it.get('路径', '(自动生成)')))


def auto_paths(blocks, items):
    for it in items:
        if it.get('路径'):
            continue
        blk = next((b for b in blocks if b['name'].startswith(it['运营'])), None)
        if blk is None:
            continue
        nums = []
        prefix = None
        leaf = re.escape(os.path.basename(config.PATH_PREFIX))
        for p in blk['projects']:
            m = re.search(leaf + r'\\(\d+)-(\d+)', p['路径'])
            if m:
                prefix = int(m.group(1))
                nums.append(int(m.group(2)))
        if prefix is None:
            prefix = 1
        nxt = (max(nums) + 1) if nums else 1
        it['路径'] = config.PATH_PREFIX + r'\%d-%d-%s' % (prefix, nxt, it['项目'])
        it['_auto_path'] = True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    config.require_xlsx()
    items = json.load(open(sys.argv[1], encoding='utf-8'))
    apply = '--apply' in sys.argv
    out_path = sys.argv[sys.argv.index('--out') + 1] if '--out' in sys.argv else \
        os.path.join(os.path.dirname(XLSX), '_new_机混需求表.xlsx')

    z = zipfile.ZipFile(XLSX)
    sheet_path = resolve_sheet(z)
    xml = z.read(sheet_path).decode('utf-8')

    ss = []
    NSM = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
    for si in sst.findall(NSM + 'si'):
        ss.append(''.join(t.text or '' for t in si.iter(NSM + 't')))
    shared = {str(i): v for i, v in enumerate(ss)}

    blocks = find_blocks(xml, shared)
    auto_paths(blocks, items)
    preview(blocks, items)
    if not apply:
        print('\n（预览模式，未写文件；加 --apply 生成新表）')
        return

    # 逐个插入（从后往前插，避免行号偏移影响前面的块）
    plan = []
    for it in items:
        blk = next((b for b in blocks if b['name'].startswith(it['运营'])), None)
        if blk is None:
            print('跳过（找不到运营块）:', it)
            continue
        plan.append((blk['last_row'] + 1, it, blk))
    plan.sort(key=lambda x: -x[0])

    for insert_at, it, blk in plan:
        # 模板样式取该块最后一个项目行
        rows = {rn: raw for rn, s, e, raw in parse_rows(xml)}
        styles = dict(TEMPLATE_STYLES)
        if blk['last_row'] in rows:
            c = row_cells(rows[blk['last_row']])
            for k in 'ABCDE':
                if k in c and c[k][2]:
                    styles[k] = c[k][2]
        xml = shift_rows(xml, insert_at)
        new_row = build_row(insert_at, it, styles)
        anchor = re.search(r'<row\b[^>]*r="%d"[^>]*>' % (insert_at + 1), xml)
        xml = xml[:anchor.start()] + new_row + xml[anchor.start():]
        print('已插入: %s -> 第 %d 行' % (it['项目'], insert_at))

    # 修正 dimension
    max_r = max(rn for rn, s, e, raw in parse_rows(xml))
    xml = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>',
                 r'<dimension ref="A1:\g<1>%d"/>' % max_r, xml, count=1)

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in z.infolist():
            data = xml.encode('utf-8') if info.filename == sheet_path else z.read(info.filename)
            out.writestr(info, data)
    z.close()
    print('\n已生成:', out_path)


if __name__ == '__main__':
    main()
