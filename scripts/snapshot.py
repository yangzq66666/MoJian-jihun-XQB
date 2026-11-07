# -*- coding: utf-8 -*-
"""
从机混需求表.xlsx 抓取「项目需求_项目路径_爆款拆分」快照。
用法: python snapshot.py
输出: baseline.json（机器可读）+ baseline.md（人可读）
"""
import re, os, json, zipfile, shutil, datetime
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ns = {'m': NS[1:-1]}
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

SRC = config.require_xlsx()
HERE = config.STATE_DIR
os.makedirs(HERE, exist_ok=True)

# 1. 复制源文件做备份
bak = os.path.join(HERE, '机混需求表_备份.xlsx')
shutil.copy2(SRC, bak)
print('源表已备份 ->', bak)
print('源文件修改时间:', datetime.datetime.fromtimestamp(os.path.getmtime(SRC)).strftime('%Y-%m-%d %H:%M:%S'))

# 2. 解压解析
tmp = os.path.join(HERE, '_unzip')
shutil.rmtree(tmp, ignore_errors=True)
os.makedirs(tmp)
with zipfile.ZipFile(bak) as z:
    z.extractall(tmp)

# rId -> target
rels = {}
for r in ET.parse(os.path.join(tmp, 'xl/_rels/workbook.xml.rels')).getroot():
    rels[r.get('Id')] = r.get('Target')

# 找到爆款拆分 sheet
SHEETNAME = config.SHEET_NAME
target = None
for s in ET.parse(os.path.join(tmp, 'xl/workbook.xml')).getroot().find('m:sheets', ns):
    if s.get('name') == SHEETNAME:
        target = rels[s.get(R + 'id')]
if not target:
    raise SystemExit('未找到工作表: ' + SHEETNAME)
sheetfile = os.path.join(tmp, 'xl', target.replace('/', os.sep))
print('目标工作表:', SHEETNAME, '->', target)

# shared strings
ss = [''.join(x.text or '' for x in si.iter(NS + 't'))
      for si in ET.parse(os.path.join(tmp, 'xl/sharedStrings.xml')).getroot().findall('m:si', ns)]


def c2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def val(c):
    ty = c.get('t')
    v = c.find('m:v', ns)
    i = c.find('m:is', ns)
    if ty == 's' and v is not None:
        return ss[int(v.text)]
    if ty == 'inlineStr' and i is not None:
        return ''.join(x.text or '' for x in i.iter(NS + 't'))
    return v.text if v is not None else ''


root = ET.parse(sheetfile).getroot()
data = root.find('m:sheetData', ns)

rows = []
for row in data.findall('m:row', ns):
    r = int(row.get('r'))
    cells = {}
    for c in row.findall('m:c', ns):
        cells[c2n(re.match(r'[A-Z]+', c.get('r')).group())] = val(c)
    rows.append((r, cells))

# 表头（行1）
hdr = rows[0][1] if rows else {}
header = [hdr.get(i, '') for i in range(1, 7)]

# 有数量的行（跳过表头行1）
hits = []
for r, cells in rows:
    if r == 1:
        continue
    qty = str(cells.get(2, '')).strip()
    if qty and re.fullmatch(r'\d+', qty):
        hits.append({
            'excel_row': r,
            '项目': cells.get(1, '').strip(),
            '数量': int(qty),
            '项目文件地址': cells.get(3, '').strip(),
            '部门': cells.get(4, '').strip(),
            '抽取片段数量': cells.get(5, '').strip(),
        })

# 输出目录表
outdir_rows = {}
try:
    outroot = ET.parse(os.path.join(tmp, 'xl', rels[
        [s.get(R + 'id') for s in ET.parse(os.path.join(tmp, 'xl/workbook.xml')).getroot().find('m:sheets', ns)
         if s.get('name') == config.OUTPUT_DIR_SHEET][0]].replace('/', os.sep))).getroot()
    od = outroot.find('m:sheetData', ns)
    for row in od.findall('m:row', ns):
        for c in row.findall('m:c', ns):
            outdir_rows[c.get('r')] = val(c)
except Exception as e:
    print('输出目录解析跳过:', e)

NOW = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
result = {
    'snapshot_time': NOW,
    'source_file': SRC,
    'source_mtime': datetime.datetime.fromtimestamp(os.path.getmtime(SRC)).strftime('%Y-%m-%d %H:%M:%S'),
    '工作表': SHEETNAME,
    '表头': header,
    '输出目录_导出父级路径': outdir_rows.get('B1', ''),
    '命中行数': len(hits),
    'expected_total_videos': sum(h['数量'] for h in hits),
    '命中行': hits,
}

with open(os.path.join(HERE, 'baseline.json'), 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

# Markdown
lines = []
lines.append('# 爆款拆分 基准快照')
lines.append('')
lines.append('**抓取时间**：%s' % NOW)
lines.append('**源表格**：`%s`（修改于 %s）' % (SRC, result['source_mtime']))
lines.append('**工作表**：%s' % SHEETNAME)
lines.append('**导出父级路径**：`%s`' % result['输出目录_导出父级路径'])
lines.append('')
lines.append('> 本快照为核对基准。表格后续变化不会影响本文件。')
lines.append('')
lines.append('## 有数量的行（共 %d 条，合计期望视频 %d 个）' % (len(hits), result['expected_total_videos']))
lines.append('')
lines.append('| Excel行号 | 机混项目 | 数量 | 部门 | 抽取片段数量 | 项目文件地址 |')
lines.append('|---|---|---|---|---|---|')
for h in hits:
    lines.append('| %d | **%s** | %d | %s | %s | `%s` |' % (
        h['excel_row'], h['项目'], h['数量'], h['部门'], h['抽取片段数量'], h['项目文件地址']))
lines.append('')
lines.append('## 预估组数')
lines.append('')
lines.append('按「数量 = 组数 × 5 人」推算：')
lines.append('')
lines.append('| 机混项目 | 数量 | 预估组数(数量÷5) |')
lines.append('|---|---|---|')
for h in hits:
    lines.append('| %s | %d | %d |' % (h['项目'], h['数量'], h['数量'] // 5))
lines.append('')

with open(os.path.join(HERE, 'baseline.md'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

# 3. 按时间戳存档一份（baseline.* 每天会被覆盖，存档用于回查历史）
arc_dir = os.path.join(HERE, '基准存档')
os.makedirs(arc_dir, exist_ok=True)
stamp = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
shutil.copy2(os.path.join(HERE, 'baseline.json'), os.path.join(arc_dir, 'baseline_%s.json' % stamp))
shutil.copy2(os.path.join(HERE, 'baseline.md'), os.path.join(arc_dir, 'baseline_%s.md' % stamp))
print('已存档 -> 基准存档/baseline_%s.json/.md' % stamp)

shutil.rmtree(tmp, ignore_errors=True)
print()
print('=== 快照已保存 ===')
print('命中行数:', len(hits), ' 合计期望视频:', result['expected_total_videos'])
print('导出父级路径:', result['输出目录_导出父级路径'])
for h in hits:
    print('  行%-3d %-14s 数量=%-4d 预估%d组' % (h['excel_row'], h['项目'], h['数量'], h['数量'] // 5))
