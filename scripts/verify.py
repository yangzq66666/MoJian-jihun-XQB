# -*- coding: utf-8 -*-
"""
拿 baseline.json（基准快照）核对 U 盘导出目录下各项目的实际视频数。
兼容两种文件命名格式：
  流水号格式(10-16起): ...-自动化-爆款拆分--<流水号>.mp4        （编号必须连续不重复）
  组号+姓名格式(10-15及前/10-20又出现): ...-爆款拆分-<组号>-<姓名>.mp4
      规则：组数 = 数量 ÷ 5；每组必须恰好 5 人（剪辑师甲/剪辑师乙/剪辑师丙/剪辑师丁/剪辑师戊），
            同组号出现 5 次属正常，不算重复。
用法: python verify.py
"""
import os, re, json, datetime, sys

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import config

HERE = config.STATE_DIR
BASELINE = config.require_state_file('baseline.json')
VIDEO_EXT = config.VIDEO_EXT
PEOPLE = config.PEOPLE

with open(BASELINE, encoding='utf-8') as f:
    base = json.load(f)

# 导出父级路径 -> 候选可访问路径（盘符可能掉线，回退到 UNC）
parent = base.get('输出目录_导出父级路径', '')
rel = parent.split(':\\', 1)[1] if ':\\' in parent else parent.replace('\\', '/')
drive = parent[:1] if ':\\' in parent else 'U'
candidates = ['//?/%s:/%s' % (drive, rel.replace('\\', '/'))]
try:
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Network\%s' % drive) as k:
        remote = winreg.QueryValueEx(k, 'RemotePath')[0]
    candidates.append(remote.replace('\\', '/') + '/' + rel.replace('\\', '/'))
except Exception:
    pass
candidates.append(parent.replace('\\', '/'))

upath = None
for c in candidates:
    try:
        if os.path.isdir(c):
            upath = c
            break
    except Exception:
        continue
if upath is None:
    print('!! 目录不可访问（网络盘未连接）:', parent)
    print('   已尝试:', candidates)
    sys.exit(1)

print('基准快照抓取时间:', base['snapshot_time'])
print('源表格        :', base['source_file'])
print('核对目录      :', parent)
if upath != candidates[0]:
    print('(经由备用地址:', upath, ')')
print('=' * 96)

dirs = os.listdir(upath)
folders = [d for d in dirs if os.path.isdir(os.path.join(upath, d))]


def analyze(names, qty):
    """返回 dict(fmt, n, status, notes)"""
    n = len(names)
    serial, groups, unparsed, blank = [], {}, [], []
    fmt = None
    for nm in names:
        m1 = re.search(r'爆款拆分--(\d+)\.mp4$', nm)
        m2 = re.search(r'爆款拆分-(\d+)-([^\-]+?)\.mp4$', nm)
        mb = re.search(r'爆款拆分-\.mp4$', nm)   # 空编号文件（10-19/10-21 均出现）
        if m1:
            serial.append(int(m1.group(1)))
            fmt = fmt or '流水号'
        elif m2:
            groups.setdefault(int(m2.group(1)), set()).add(m2.group(2))
            fmt = fmt or '组号+姓名'
        elif mb:
            blank.append(nm)
            fmt = fmt or '流水号'
        else:
            unparsed.append(nm)
    notes = []

    if fmt == '组号+姓名':
        exp_g = qty // 5
        gnums = sorted(groups)
        miss_g = [g for g in range(1, exp_g + 1) if g not in groups]
        extra_g = [g for g in gnums if g > exp_g]
        for g in gnums:
            ppl = groups[g]
            if len(ppl) != 5:
                notes.append('组%d人数=%d(%s)' % (g, len(ppl), ','.join(sorted(ppl))))
            else:
                miss = [p for p in PEOPLE if p not in ppl]
                extra = [p for p in ppl if p not in PEOPLE]
                if miss or extra:
                    notes.append('组%d异常 缺%s 多%s' % (g, miss, extra))
        if n != qty:
            status = '⚠️ 多%d' % (n - qty) if n > qty else '⚠️ 少%d' % (qty - n)
            if extra_g:
                notes.append('超出组: %s' % extra_g)
            if miss_g:
                notes.append('缺组: %s' % miss_g)
        elif notes or extra_g or miss_g or unparsed:
            status = '⚠️ 结构异常'
            if extra_g:
                notes.append('超出组: %s' % extra_g)
            if miss_g:
                notes.append('缺组: %s' % miss_g)
            if unparsed:
                notes.append('无法解析 %d 个文件' % len(unparsed))
        else:
            status = '✅ 一致'
        return dict(fmt=fmt, n=n, status=status, notes=notes,
                    ngrp=len(gnums), nppl=len(set().union(*groups.values())) if groups else 0)

    if fmt == '流水号':
        uniq = sorted(set(serial))
        miss = [i for i in range(1, max(uniq) + 1) if i not in uniq] if uniq else []
        dup = len(serial) - len(uniq)
        if blank:
            notes.append('含 %d 个无编号文件(爆款拆分-.mp4，已知现象)' % len(blank))
        if n != qty:
            status = '⚠️ 多%d' % (n - qty) if n > qty else '⚠️ 少%d' % (qty - n)
        elif miss or dup or unparsed:
            status = '⚠️ 数量对但编号异常(缺%d/重%d)' % (len(miss), dup)
            if miss:
                notes.append('缺号: %s' % miss)
            if unparsed:
                notes.append('无法解析 %d 个文件' % len(unparsed))
        else:
            status = '✅ 一致'
        return dict(fmt=fmt, n=n, status=status, notes=notes,
                    ngrp=len(uniq), nppl=None)

    # 无法识别任何格式
    if n == 0:
        return dict(fmt='(空)', n=0, status='⚠️ 少%d' % qty, notes=['文件夹为空'],
                    ngrp=0, nppl=0)
    return dict(fmt='未识别', n=n, status='⚠️ 命名无法解析',
                notes=['示例: %s' % names[0]], ngrp=None, nppl=None)


rows = []
# 同音近形字归一化（按本机表实际用字调整下面这一行映射）
VARIANT = str.maketrans({'辩': '辨'})
for h in base['命中行']:
    proj, qty = h['项目'], h['数量']
    matched = [d for d in folders if proj in d]
    via_alias = False
    if not matched:
        # 文件夹名可能用了同音错字，归一化后再匹配
        matched = [d for d in folders if proj in d.translate(VARIANT)]
        via_alias = bool(matched)
    if not matched:
        rows.append(dict(h=h, folder=None, res=dict(fmt=None, n=None, status='❌ 未找到文件夹',
                                                    notes=[], ngrp=None, nppl=None)))
        continue
    for folder in matched:
        p = os.path.join(upath, folder)
        names = [x for x in os.listdir(p) if x.lower().endswith(VIDEO_EXT)]
        res = analyze(names, qty)
        if via_alias:
            res = dict(res)
            res['notes'] = res['notes'] + ['文件夹名用字与本地表不同（%s），内容为同一项目' % folder]
        rows.append(dict(h=h, folder=folder, res=res))

ok = sum(1 for r in rows if r['res']['status'].startswith('✅'))
print('%-14s %5s %6s  %-22s %6s %6s' % ('项目', '期望', '实际', '结论', '组数', '人数'))
print('-' * 96)
for r in rows:
    res = r['res']
    print('%-14s %5d %6s  %-22s %6s %6s' % (
        r['h']['项目'], r['h']['数量'],
        res['n'] if res['n'] is not None else '—', res['status'],
        res['ngrp'] if res['ngrp'] is not None else '—',
        res['nppl'] if res['nppl'] else '—'))
print('-' * 96)
print('一致 %d / 共 %d 条' % (ok, len(rows)))
print('期望合计 %d 个视频，实际合计 %s 个' % (
    sum(h['数量'] for h in base['命中行']),
    sum(r['res']['n'] for r in rows if r['res']['n'] is not None)))

print()
print('=== 明细 ===')
for r in rows:
    h, res = r['h'], r['res']
    print('【%s】%s' % (h['项目'], res['status']))
    print('   行号 %d | 部门 %s | 抽取片段 %s' % (h['excel_row'], h['部门'], h['抽取片段数量']))
    print('   文件夹: %s' % (r['folder'] or '(无)'))
    if res['fmt']:
        print('   命名格式: %s' % res['fmt'])
    for note in res['notes']:
        print('   ⚠️ %s' % note)
    print()

# 写出报告
out = ['# 视频数量核对报告', '']
out.append('核对时间: %s' % datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
out.append('基准快照: %s' % base['snapshot_time'])
out.append('核对目录: `%s`' % parent)
out.append('')
out.append('| 机混项目 | 期望数量 | 实际视频数 | 结论 | 组数 | 人数 |')
out.append('|---|---|---|---|---|---|')
for r in rows:
    res = r['res']
    out.append('| %s | %d | %s | %s | %s | %s |' % (
        r['h']['项目'], r['h']['数量'],
        res['n'] if res['n'] is not None else '—', res['status'],
        res['ngrp'] if res['ngrp'] is not None else '—',
        res['nppl'] if res['nppl'] else '—'))
out.append('')
out.append('**一致 %d / 共 %d 条**，期望合计 %d 个视频，实际合计 %s 个。' % (
    ok, len(rows), sum(h['数量'] for h in base['命中行']),
    sum(r['res']['n'] for r in rows if r['res']['n'] is not None)))
out.append('')
if any(not r['res']['status'].startswith('✅') for r in rows):
    out.append('## 差异明细')
    out.append('')
    for r in rows:
        res = r['res']
        if res['status'].startswith('✅'):
            continue
        out.append('- **%s**（期望 %d，实际 %s）：%s' % (
            r['h']['项目'], r['h']['数量'],
            res['n'] if res['n'] is not None else '—', res['status']))
        for note in res['notes']:
            out.append('  - %s' % note)
    out.append('')
with open(os.path.join(HERE, '核对报告.md'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))
print('报告已写入 核对报告.md')
