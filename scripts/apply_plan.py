# -*- coding: utf-8 -*-
"""
按 plan.json 把「魔剪」数量写进本地《机混需求表》的爆款拆分页。

行为：
  1. 备份原表到 STATE_DIR/backup/
  2. 校验每个目标行的 A 列项目名，防止错位写入
  3. 按运营块清空旧数量（默认清全部块，符合"每天重填要清掉前一天所有数量"）
  4. 写入今天的数量
  5. 校验 XML 合法性后打包；加 --apply 才覆盖原表

用法：
    python apply_plan.py plan.json                      # 预演，不写文件
    python apply_plan.py plan.json --apply              # 写盘
    python apply_plan.py plan.json --clear 运营甲,运营乙     # 只清这两个运营块
    python apply_plan.py plan.json --clear none         # 不清旧值，只覆盖今天有的行
    python apply_plan.py plan.json --keep 20,47         # 不清这两行（用户当天明确要保留的）
"""
import argparse
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

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def style_of(cell):
    m = re.search(r'\bs="(\d+)"', cell)
    return ' s="%s"' % m.group(1) if m else ''


def cell_re(col, n):
    return re.compile(r'<c\b[^>]*\br="%s%d"[^>]*/>|<c\b[^>]*\br="%s%d"[^>]*>.*?</c>' % (col, n, col, n), re.S)


def set_cell(xml, col, n, value):
    m = cell_re(col, n).search(xml)
    if not m:
        return xml, False
    new = '<c r="%s%d"%s><v>%s</v></c>' % (col, n, style_of(m.group(0)), value)
    return xml[:m.start()] + new + xml[m.end():], True


def clear_cell(xml, col, n):
    m = cell_re(col, n).search(xml)
    if not m:
        return xml
    new = '<c r="%s%d"%s/>' % (col, n, style_of(m.group(0)))
    return xml[:m.start()] + new + xml[m.end():]


def read_a_text(xml, sst, n):
    m = cell_re('A', n).search(xml)
    if not m:
        return ''
    cell = m.group(0)
    im = re.search(r'<is>\s*<t[^>]*>(.*?)</t>\s*</is>', cell, re.S)
    if im:
        return im.group(1)
    vm = re.search(r'<v>(\d+)</v>', cell)
    return sst[int(vm.group(1))] if vm else ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('plan')
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--clear', default='all',
                    help="all（默认，清全部运营块）| none | 逗号分隔的运营名")
    ap.add_argument('--keep', default='', help='逗号分隔的 Excel 行号，这些行不清空')
    args = ap.parse_args()

    with open(args.plan, encoding='utf-8') as f:
        plan = json.load(f)
    xlsx = plan.get('xlsx') or config.XLSX
    config.require_xlsx(xlsx)

    targets = {i['local_row']: (i['qty'], i['local_project'], i['doc_project'])
               for i in plan['items'] if i.get('local_row')}
    pending_new = [i for i in plan['items'] if not i.get('local_row')]
    if pending_new:
        print('!! 以下 %d 项在本地表没有对应行，本脚本不处理，先用 add_projects.py 新建：' % len(pending_new))
        for i in pending_new:
            print('   - %s（运营 %s，%d）' % (i['doc_project'], i['ops'] or '?', i['qty']))
        print()

    # ---- 备份 ----
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    bakdir = os.path.join(config.STATE_DIR, 'backup')
    os.makedirs(bakdir, exist_ok=True)
    bak = os.path.join(bakdir, '机混需求表_备份_%s.xlsx' % stamp)
    shutil.copy2(xlsx, bak)
    print('备份:', bak)

    # ---- 解包 ----
    work = os.path.join(config.STATE_DIR, '_apply_tmp')
    shutil.rmtree(work, ignore_errors=True)
    with zipfile.ZipFile(xlsx) as z:
        z.extractall(work)

    sst = [''.join(t.text or '' for t in si.iter(NS + 't'))
           for si in ET.parse(os.path.join(work, 'xl/sharedStrings.xml')).getroot()]
    shp = os.path.join(work, 'xl/worksheets/sheet3.xml')
    xml = open(shp, encoding='utf-8').read()

    # ---- 校验 A 列 ----
    print('\n=== 写入前校验 ===')
    bad = False
    for n, (qty, local, doc) in sorted(targets.items()):
        actual = read_a_text(xml, sst, n).strip()
        ok = actual == local
        print('  行%-3d 期望[%s] 实际[%s] %s' % (n, local, actual, 'OK' if ok else '!!不匹配'))
        bad = bad or not ok
    if bad:
        raise SystemExit('A 列校验失败，中止（表格结构可能变了，先跑 block_map.py 核对行号）')

    # ---- 清理范围 ----
    import block_map as bm
    blocks = bm.block_map(bm.load(xlsx))
    keep = {int(x) for x in args.keep.split(',') if x.strip()}

    if args.clear == 'none':
        ranges = []
    elif args.clear == 'all':
        ranges = [(b['projects'][0]['row'], b['projects'][-1]['row'])
                  for b in blocks if b['projects']]
    else:
        want = [x.strip() for x in args.clear.split(',') if x.strip()]
        ranges = []
        for b in blocks:
            if any(b['name'].startswith(w) for w in want) and b['projects']:
                ranges.append((b['projects'][0]['row'], b['projects'][-1]['row']))

    cleared = 0
    for lo, hi in ranges:
        for n in range(lo, hi + 1):
            if n in keep:
                continue
            before = xml
            xml = clear_cell(xml, 'B', n)
            if xml != before:
                cleared += 1

    print('\n清理范围:', ranges or '(不清空)')
    print('清空 B 单元格: %d 个' % cleared)

    # ---- 写入 ----
    for n, (qty, local, doc) in sorted(targets.items()):
        xml, ok = set_cell(xml, 'B', n, qty)
        if not ok:
            raise SystemExit('!! 行 %d 缺少 B 单元格' % n)
    print('写入 %d 行，合计 %d 个视频' % (len(targets), sum(v[0] for v in targets.values())))

    open(shp, 'w', encoding='utf-8').write(xml)
    ET.fromstring(open(shp, 'rb').read())
    print('sheet3.xml 合法性校验: 通过')

    # ---- 打包 ----
    out = os.path.join(os.path.dirname(xlsx), '_new_机混需求表.xlsx')
    with zipfile.ZipFile(xlsx) as zin, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = open(shp, 'rb').read() if item.filename == 'xl/worksheets/sheet3.xml' \
                else zin.read(item.filename)
            zout.writestr(item, data)
    with zipfile.ZipFile(out) as z:
        assert z.testzip() is None
    shutil.rmtree(work, ignore_errors=True)

    if args.apply:
        shutil.copy2(out, xlsx)
        print('\n已写入原表:', xlsx)
        print('下一步：python snapshot.py 刷新核对基准')
    else:
        print('\n预演完成，未覆盖原表:', out, '（加 --apply 写盘）')


if __name__ == '__main__':
    main()
