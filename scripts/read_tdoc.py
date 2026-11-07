# -*- coding: utf-8 -*-
"""
读取腾讯文档《需求对接表》「—剪辑需求—」页指定日期的数据块，
按业务规则筛出「魔剪」需求，并映射到本地《机混需求表》的项目行，输出 plan.json。

用法：
    python read_tdoc.py                      # 取今天
    python read_tdoc.py --date YYYY-MM-DD    # 指定日期
    python read_tdoc.py --dump               # 附带打印今日块全部行（排查用）

输出 plan.json（默认写到 STATE_DIR），由 apply_plan.py 消费。
规则见 references/business-rules.md。
"""
import argparse
import datetime
import glob
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

PAGE_ROWS = 700          # 每次拉 700 行 × 27 列 = 18900 单元格，低于 20000 上限


# ---------------- 腾讯文档调用 ----------------

def find_tdoc_script():
    if config.TDOC_SCRIPT and os.path.isfile(config.TDOC_SCRIPT):
        return config.TDOC_SCRIPT
    hits = glob.glob(config.TDOC_PLUGIN_GLOB)
    if not hits:
        raise SystemExit(
            '找不到 tencentdocs.py。请安装 tencent-docs 插件，'
            '或用环境变量 MOJIAN_TDOC_SCRIPT 指定路径。')
    # 版本号大的排后面，取最新
    hits.sort(key=lambda p: [int(x) if x.isdigit() else 0 for x in re.split(r'[\\/.]', p)])
    return hits[-1]


def tdoc_call(script, service, tool, args):
    cmd = [sys.executable, script, 'tdoc_call', service, tool,
           json.dumps(args, ensure_ascii=False)]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    out = (r.stdout or '').strip()
    if not out:
        raise SystemExit('腾讯文档无输出（stderr: %s）' % (r.stderr or '')[:500])
    # 取最后一行合法 JSON
    for line in reversed(out.splitlines()):
        line = line.strip()
        if line.startswith('{'):
            obj = json.loads(line)
            if 'error' in obj:
                raise SystemExit('腾讯文档报错: %s' % obj['error'])
            return obj.get('result', {}).get('structuredContent') or {}
    raise SystemExit('腾讯文档返回非 JSON: %s' % out[:500])


def norm_date(txt):
    """把 '2026/11/6 星期五' 之类归一成 (2026, 11, 6)。"""
    m = re.search(r'(20\d\d)\s*[/\-年]\s*(\d{1,2})\s*[/\-月]\s*(\d{1,2})', txt or '')
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def fetch_block(script, date_str):
    """拉取包含 date_str 的日期块，返回 (start_row, end_row, grid)。"""
    info = tdoc_call(script, 'sheet-mcp', 'get_sheet_info', {'file_id': config.TDOC_FILE_ID})
    sheet = next(s for s in info['sheets'] if s['sheet_id'] == config.TDOC_SHEET_ID)
    last = sheet['row_count'] - 1

    want = tuple(int(x) for x in date_str.split('-'))

    end = last
    while end > 0:
        start = max(0, end - PAGE_ROWS + 1)
        sc = tdoc_call(script, 'sheet-mcp', 'get_cell_data', {
            'file_id': config.TDOC_FILE_ID, 'sheet_id': config.TDOC_SHEET_ID,
            'start_row': start, 'end_row': end,
            'start_col': 0, 'end_col': config.COL_MAX,
            'return_csv': False})
        grid = {}
        for c in sc.get('cells', []):
            v = c.get('string_value')
            if v is None:
                v = c.get('value')
            if v is not None and str(v).strip():
                grid[(c['row'], c['col'])] = str(v).replace('\n', ' ').strip()

        # 本页里的日期块头
        heads = sorted((r, v) for (r, col), v in grid.items()
                       if col == config.COL_DATE and norm_date(v))
        for i, (h, txt) in enumerate(heads):
            if norm_date(txt) != want:
                continue
            if i + 1 < len(heads):
                blk_end = heads[i + 1][0] - 1
            else:
                blk_end = max(r for (r, _c) in grid)   # 今天通常是最末块
            return h, blk_end, grid
        end = start - 1
    raise SystemExit('未找到日期块: %s（表里可能还没建）' % date_str)


# ---------------- 本地表映射 ----------------

def load_blocks():
    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, here)
    import block_map as bm
    return bm.block_map(bm.load())


def load_aliases():
    here = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(here, '..', 'references', 'aliases.json')
    if os.path.isfile(p):
        with open(p, encoding='utf-8') as f:
            return {k: v for k, v in json.load(f).items() if not k.startswith('_')}
    return {}


SUFFIX_RE = re.compile(r'[-－]?(金牛|蓝帽|鹿鞭|OTC|otc|新包装)$|（[^）]*）|\([^)]*\)')


def strip_suffix(name):
    n = name
    for _ in range(4):
        m = SUFFIX_RE.sub('', n).strip(' -－')
        if m == n or not m:
            break
        n = m
    return n or name


def match_local(doc_name, flat, aliases):
    """把腾讯文档项目名映射到本地表项目名，返回 (local_name, how, candidates)。"""
    if doc_name in aliases:
        return aliases[doc_name], 'alias', []
    base = strip_suffix(doc_name)
    if base in aliases:
        return aliases[base], 'alias(去后缀)', []
    if base in flat:
        return base, '精确', []
    if doc_name in flat:
        return doc_name, '精确(原名)', []

    # 唯一包含匹配：base 是本地项目名的子串，或反之
    cand = [k for k in flat if k and (base in k or k in base)]
    if len(cand) == 1:
        return cand[0], '唯一包含', []
    return None, '未匹配', cand


# ---------------- 主流程 ----------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--date', default=datetime.date.today().strftime('%Y-%m-%d'))
    ap.add_argument('--out', default=None)
    ap.add_argument('--dump', action='store_true')
    args = ap.parse_args()

    config.require_xlsx()          # 路径没配置时给出 setup.py 指引
    script = find_tdoc_script()
    print('腾讯文档脚本:', script)

    start, end, grid = fetch_block(script, args.date)
    print('日期块: %s  行范围 %d~%d' % (args.date, start, end))

    aliases = load_aliases()
    blocks = load_blocks()
    flat = {}
    for b in blocks:
        for p in b['projects']:
            flat.setdefault(p['项目'], {'row': p['row'], 'ops': b['name']})

    if args.dump:
        print('\n--- 日期块明细 ---')
        for r in range(start, end + 1):
            vals = [grid.get((r, c), '') for c in range(0, 12)]
            if any(vals):
                print(r, '|'.join(vals))

    items, skipped = [], []
    for r in range(start, end + 1):
        proj = grid.get((r, config.COL_PROJECT), '')
        note = grid.get((r, config.COL_NOTE), '')
        ops = grid.get((r, config.COL_OPS), '')
        if not proj or proj == '总数':
            continue

        if re.search(config.SKIP_RE, note):
            skipped.append({'row': r, 'doc_project': proj, 'note': note,
                            'reason': '备注明确不要魔剪'})
            continue

        m = re.search(config.MAGIC_RE, note)
        if not m:
            skipped.append({'row': r, 'doc_project': proj, 'note': note,
                            'reason': '备注无 N魔剪/N混剪/N机混'})
            continue

        qty = int(m.group(1))
        local, how, cand = match_local(proj, flat, aliases)
        # 别名可能指向本地表里不存在的行（换机器/表结构不同），此时按业务规则走「新建行」
        hit = flat.get(local) if local else None
        item = {
            'doc_row': r, 'doc_project': proj, 'note': note, 'qty': qty,
            'ops': ops, 'local_project': local,
            'match': how + ('(表中无此行)' if local and not hit else ''),
            'local_row': hit['row'] if hit else None,
            'action': 'write' if hit else 'new',
        }
        if not hit:
            item['candidates'] = cand
            if not ops:
                item['warn'] = 'E列运营为空，需按 C 列合并单元格确认归属'
        items.append(item)

    # 归属运营：C 列是合并单元格，只有块首行有值，向下继承
    cur = ''
    ops_map = {}
    for r in range(start, end + 1):
        v = grid.get((r, config.COL_OPS), '')
        if v:
            cur = v
        ops_map[r] = cur
    for it in items:
        it['ops'] = ops_map.get(it['doc_row'], it.get('ops', ''))

    plan = {
        'date': args.date,
        'doc': {'file_id': config.TDOC_FILE_ID, 'sheet_id': config.TDOC_SHEET_ID,
                'block_range': [start, end]},
        'xlsx': config.XLSX,
        'items': items,
        'skipped': skipped,
        'total_qty': sum(i['qty'] for i in items),
    }

    out = args.out or os.path.join(config.STATE_DIR, 'plan.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)

    print('\n=== 魔剪需求 %d 项 / %d 个视频 ===' % (len(items), plan['total_qty']))
    print('%-22s %-5s %-14s %-6s %-6s %s' % ('腾讯文档项目', '数量', '本地项目', '行', '运营', '动作'))
    for i in items:
        print('%-22s %-5d %-14s %-6s %-6s %s' % (
            i['doc_project'], i['qty'], i['local_project'] or '(需新建)',
            i['local_row'] or '-', i['ops'] or '?',
            i['action'] + ('' if i['action'] == 'write' else ' ←' + ','.join(i.get('candidates', [])))))
    print('\n跳过 %d 项（备注非魔剪）' % len(skipped))
    for s in skipped:
        if s['note']:
            print('  - %s：%s' % (s['doc_project'], s['note']))
    print('\n计划已写入:', out)


if __name__ == '__main__':
    main()
