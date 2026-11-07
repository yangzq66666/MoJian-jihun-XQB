#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置向导 / 环境自检 —— 首次安装或换机器后跑一次。

它做四件事：
  1. 自动搜索本机《机混需求表》（含桌面快捷方式 .lnk 反解），搜到就让你选
  2. 从表里已有的「项目文件地址」反推素材根目录（0-魔剪-素材），不用手抄
  3. 从表里「输出目录」页 B1 读出成片导出目录并检查是否可访问
  4. 结果写进 ~/.workbuddy/mojian-jihun-xqb/config.json（技能目录之外，重装不丢）

用法：
    python setup.py                          # 交互式配置（推荐首次使用）
    python setup.py --check                  # 只看当前配置 + 自检，不改文件
    python setup.py --check --deep           # 顺带真连一次腾讯文档
    python setup.py --xlsx "D:\\xx\\机混需求表.xlsx" --yes    # 非交互
    python setup.py --reset                  # 重新配置（覆盖已有 config.json）
"""
import argparse
import datetime
import json
import os
import re
import string
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
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

TARGET_XLSX_NAME = '机混需求表.xlsx'
MAGIC_SEG = '0-魔剪-素材'
SKIP_DIRS = {
    'windows', 'program files', 'program files (x86)', '$recycle.bin',
    'system volume information', 'appdata', 'node_modules', '.git', '.svn',
    'perflogs', 'recovery', 'msocache', 'config.msi', 'windows.old',
    'anaconda3', 'miniconda3', 'site-packages', 'temp', 'tmp',
}

OK, WARN, BAD = '  [OK]  ', '  [!!]  ', '  [XX]  '


def say(msg=''):
    print(msg)


def ok(msg):
    print(OK + msg)


def warn(msg):
    print(WARN + msg)


def bad(msg):
    print(BAD + msg)


def ask(prompt, default=''):
    """读一行；空输入取默认值。"""
    tip = '%s\n  > ' % prompt if not default else '%s\n  [%s]\n  > ' % (prompt, default)
    try:
        s = input(tip).strip()
    except EOFError:
        return default
    if not s:
        return default
    if len(s) >= 2 and s[0] == s[-1] and s[0] in '"\'':
        s = s[1:-1]
    return s


# ------------------------------------------------------------------
# 自动探测
# ------------------------------------------------------------------
def iter_drives():
    for ch in string.ascii_uppercase:
        root = '%s:\\' % ch
        try:
            if os.path.isdir(root):
                yield root
        except Exception:
            continue


def _walk_find(root, filename, maxdepth=2):
    root = os.path.abspath(root).rstrip('\\/')
    base = root.count(os.sep)

    def onerr(e):
        pass

    for cur, dirs, files in os.walk(root, onerror=onerr):
        if cur.count(os.sep) - base >= maxdepth:
            dirs[:] = []
        dirs[:] = [d for d in dirs
                   if d.lower() not in SKIP_DIRS and not d.startswith(('$', '.'))]
        if filename in files:
            yield os.path.join(cur, filename)


def resolve_lnk(lnk):
    """从 .lnk 二进制里抠出目标路径（快捷方式里存的是 UTF-16LE 文本）。"""
    try:
        raw = open(lnk, 'rb').read()
    except Exception:
        return None
    for enc in ('utf-16-le', 'mbcs', 'latin-1'):
        try:
            text = raw.decode(enc, errors='ignore')
        except Exception:
            continue
        for m in re.finditer(r'([A-Za-z]:[\\/][^\x00\r\n]{2,200}?\.xlsx)', text):
            p = m.group(1).replace('/', '\\')
            if os.path.isfile(p):
                return os.path.abspath(p)
    return None


def find_xlsx_candidates():
    home = os.path.expanduser('~')
    seen, cand = set(), []

    def add(p):
        if not p or not os.path.isfile(p):
            return
        p = os.path.abspath(p)
        k = os.path.normcase(p)
        if k not in seen:
            seen.add(k)
            cand.append(p)

    # 1) 桌面快捷方式（用户通常把 lnk 放桌面）
    for d in (os.path.join(home, 'Desktop'), os.path.join(home, '桌面')):
        try:
            for f in os.listdir(d):
                if f.lower().endswith('.lnk') and '机混' in f:
                    add(resolve_lnk(os.path.join(d, f)))
        except Exception:
            pass

    # 2) 固定常见位置
    for p in (config.DEFAULTS['xlsx'],
              os.path.join(home, 'Desktop', TARGET_XLSX_NAME),
              os.path.join(home, '桌面', TARGET_XLSX_NAME),
              os.path.join(home, 'Documents', TARGET_XLSX_NAME),
              os.path.join(home, 'Downloads', TARGET_XLSX_NAME)):
        add(p)

    # 3) 各数据盘浅扫（C 盘不全盘扫，只扫用户目录）
    for root in iter_drives():
        if root.upper().startswith('C'):
            continue
        add(os.path.join(root, TARGET_XLSX_NAME))
        try:
            for p in _walk_find(root, TARGET_XLSX_NAME, 2):
                add(p)
        except Exception:
            pass
    try:
        for p in _walk_find(home, TARGET_XLSX_NAME, 3):
            add(p)
    except Exception:
        pass
    return cand


# ------------------------------------------------------------------
# 读 xlsx（不依赖 openpyxl）
# ------------------------------------------------------------------
def _col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def read_xlsx_info(path):
    info = {'sheets': [], 'rows': [], 'output_dir': '', 'sheet_found': False}
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        raise SystemExit('打不开这个 xlsx（可能不是 Excel 文件或已损坏）：%s' % e)
    with z:
        names = z.namelist()
        shared = []
        if 'xl/sharedStrings.xml' in names:
            try:
                sroot = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in sroot.findall(NS + 'si'):
                    shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
            except Exception:
                pass

        rels, targets = {}, {}
        if 'xl/_rels/workbook.xml.rels' in names:
            for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')):
                rels[r.get('Id')] = r.get('Target')
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        for s in wb.find(NS + 'sheets'):
            nm = s.get('name')
            info['sheets'].append(nm)
            targets[nm] = (rels.get(s.get(RID)) or '').lstrip('/').replace('\\', '/')

        def cells_of(sheet_name):
            tgt = targets.get(sheet_name)
            if not tgt:
                return {}
            p = 'xl/' + tgt
            if p not in names:
                return {}
            root = ET.fromstring(z.read(p))
            out = {}
            for c in root.iter(NS + 'c'):
                m = re.match(r'([A-Z]+)(\d+)$', c.get('r') or '')
                if not m:
                    continue
                t = c.get('t')
                v = ''
                if t == 'inlineStr':
                    v = ''.join(x.text or '' for x in c.iter(NS + 't'))
                else:
                    ve = c.find(NS + 'v')
                    if ve is not None and ve.text:
                        if t == 's':
                            try:
                                v = shared[int(ve.text)]
                            except Exception:
                                v = ''
                        else:
                            v = ve.text
                if v and v.strip():
                    out[(int(m.group(2)), _col_num(m.group(1)))] = v.strip()
            return out

        if config.SHEET_NAME in targets:
            info['sheet_found'] = True
            cells = cells_of(config.SHEET_NAME)
            maxrow = max((r for r, _ in cells), default=0)
            for r in range(1, maxrow + 1):
                a = cells.get((r, 1), '')
                if a:
                    info['rows'].append((r, a, cells.get((r, 2), ''), cells.get((r, 3), '')))
        if config.OUTPUT_DIR_SHEET in targets:
            info['output_dir'] = cells_of(config.OUTPUT_DIR_SHEET).get((1, 2), '')
    return info


def block_rows(rows):
    """运营块首行（A 列形如「运营甲项目」）；第 1 行是表头「机混项目」，跳过。"""
    return [r for r in rows if r[0] > 1 and r[1].endswith('项目')]


def derive_prefix(rows):
    """从「项目文件地址」列反推素材根目录。"""
    for _, _name, _qty, p in rows:
        if p and MAGIC_SEG in p:
            m = re.search(r'^(.*?%s)(?:[\\/]|$)' % re.escape(MAGIC_SEG), p)
            if m:
                return m.group(1)
    return ''


def find_tdoc_script():
    import glob
    if config.TDOC_SCRIPT and os.path.isfile(config.TDOC_SCRIPT):
        return config.TDOC_SCRIPT
    hits = glob.glob(config.TDOC_PLUGIN_GLOB)
    if not hits:
        return ''
    hits.sort(key=lambda p: [int(x) if x.isdigit() else 0 for x in re.split(r'[\\/.]', p)])
    return hits[-1]


def dir_writable(d):
    try:
        os.makedirs(d, exist_ok=True)
        probe = os.path.join(d, '.write_test_%d' % os.getpid())
        with open(probe, 'w') as f:
            f.write('ok')
        os.remove(probe)
        return True
    except Exception:
        return False


# ------------------------------------------------------------------
# 自检
# ------------------------------------------------------------------
def doctor(cfg, deep=False):
    """返回 (errors, warnings)。"""
    errors, warnings = [], []
    say('')
    say('-' * 68)
    say('环境自检')
    say('-' * 68)

    say('配置文件：%s' % config.CONFIG_FILE)
    if os.path.isfile(config.CONFIG_FILE):
        ok('配置文件存在')
    else:
        warn('配置文件不存在，当前用的是出厂默认值 / 环境变量')

    xlsx = cfg.get('xlsx') or ''
    say('')
    say('本地需求表：%s' % xlsx)
    info = None
    if xlsx and os.path.isfile(xlsx):
        ok('文件存在  (%.1f KB，修改于 %s)' % (
            os.path.getsize(xlsx) / 1024.0,
            datetime.datetime.fromtimestamp(os.path.getmtime(xlsx)).strftime('%Y-%m-%d %H:%M')))
        try:
            info = read_xlsx_info(xlsx)
        except SystemExit as e:
            bad('解析失败：%s' % e)
            errors.append('xlsx 解析失败')
    else:
        bad('文件不存在')
        errors.append('需求表不存在')

    sheet = cfg.get('sheet_name') or config.SHEET_NAME
    say('')
    say('工作表：%s' % sheet)
    if info is not None:
        if info['sheet_found']:
            blocks = block_rows(info['rows'])
            qty_rows = [r for r in info['rows'] if r[2] not in ('', None)]
            proj = [r for r in info['rows'] if r[0] > 1 and not r[1].endswith('项目')]
            ok('工作表存在：%d 个运营块 / %d 个项目行 / 当前有数量的 %d 行'
               % (len(blocks), len(proj), len(qty_rows)))
        else:
            bad('表里找不到这个工作表')
            say('      现有的工作表：%s' % '、'.join(info['sheets']))
            errors.append('工作表名不匹配')

    say('')
    state = cfg.get('state_dir') or os.path.join(
        os.path.dirname(cfg.get('xlsx') or '.'), 'mojian-runtime')
    say('运行状态目录：%s' % state)
    if dir_writable(state):
        ok('可写')
        for f in ('baseline.json', '魔剪需求台账.md'):
            p = os.path.join(state, f)
            if os.path.isfile(p):
                ok('  已存在 %s' % f)
            else:
                warn('  还没有 %s' % f)
    else:
        bad('无法创建 / 不可写')
        errors.append('状态目录不可写')

    say('')
    say('素材根目录：%s' % cfg.get('path_prefix'))
    if cfg.get('path_prefix') and os.path.isdir(cfg['path_prefix']):
        try:
            n = len([d for d in os.listdir(cfg['path_prefix'])
                     if os.path.isdir(os.path.join(cfg['path_prefix'], d))])
        except Exception:
            n = -1
        ok('可访问，%d 个子目录' % n if n >= 0 else '可访问')
    else:
        warn('访问不到（网络盘可能没连），新增项目时会用到，不影响读表写表')
        warnings.append('素材根目录不可访问')

    say('')
    say('成片导出目录（表内「%s」页 B1）：%s'
        % (config.OUTPUT_DIR_SHEET, (info or {}).get('output_dir') or '（读不到）'))
    out = (info or {}).get('output_dir') or ''
    if not out:
        warn('表里没读到导出路径，核对会失败')
        warnings.append('导出目录未知')
    elif os.path.isdir(out):
        ok('可访问')
    else:
        alt = out.replace('\\', '/')
        if os.path.isdir(alt):
            ok('可访问')
        else:
            warn('访问不到（网络盘掉线或路径已改），核对时会再尝试 UNC 回退')
            warnings.append('导出目录当前不可访问')

    say('')
    say('腾讯文档：file_id=%s  sheet_id=%s'
        % (cfg.get('tdoc_file_id'), cfg.get('tdoc_sheet_id')))
    ts = cfg.get('tdoc_script') or find_tdoc_script()
    if ts:
        ok('tencentdocs.py：%s' % ts)
    else:
        bad('找不到 tencentdocs.py，读表会失败（装 tencent-docs 插件，或配置 tdoc_script）')
        errors.append('缺少腾讯文档插件脚本')

    if deep and ts:
        say('')
        say('正在试连腾讯文档…')
        import subprocess
        cmd = [sys.executable, ts, 'tdoc_call', 'sheet-mcp', 'get_sheet_info',
               json.dumps({'file_id': cfg.get('tdoc_file_id')}, ensure_ascii=False)]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True,
                               encoding='utf-8', errors='replace', timeout=120)
            out_txt = (r.stdout or '').strip()
            if r.returncode == 0 and 'sheet_id' in out_txt:
                ok('接口连通，能列出子表')
            else:
                warn('接口无有效返回：%s' % (out_txt or r.stderr or '')[:200])
                warnings.append('腾讯文档接口未验证通过')
        except Exception as e:
            warn('试连失败：%s' % e)
            warnings.append('腾讯文档接口未验证通过')

    say('')
    say('-' * 68)
    if errors:
        bad('自检未通过：%d 个问题，%d 个提醒' % (len(errors), len(warnings)))
    elif warnings:
        ok('自检通过：0 个问题，%d 个提醒（不影响使用）' % len(warnings))
    else:
        ok('自检全部通过')
    say('-' * 68)
    return errors, warnings


# ------------------------------------------------------------------
# 交互配置
# ------------------------------------------------------------------
def pick_xlsx(non_interactive, preset=''):
    say('')
    say('[1/5] 本地《机混需求表》位置')
    if preset:
        return preset

    say('  正在搜索本机候选文件…')
    cands = find_xlsx_candidates()
    if cands:
        say('  找到 %d 个候选：' % len(cands))
        for i, p in enumerate(cands, 1):
            try:
                inf = read_xlsx_info(p)
                mark = '' if inf['sheet_found'] else '   ← 缺目标工作表'
                say('    %d) %s%s' % (i, p, mark))
            except SystemExit:
                say('    %d) %s   ← 打不开' % (i, p))
        say('')
        if non_interactive:
            return cands[0]
        while True:
            s = ask('选编号（回车=1），或直接粘贴完整路径')
            if not s:
                return cands[0]
            if s.isdigit() and 1 <= int(s) <= len(cands):
                return cands[int(s) - 1]
            if os.path.isfile(s):
                return s
            bad('没这个文件：%s' % s)
    else:
        warn('没搜到《机混需求表.xlsx》，请手动粘贴完整路径')
        say('    提示：也可以把桌面快捷方式的目标路径贴进来。')
        if non_interactive:
            raise SystemExit('非交互模式下必须用 --xlsx 指定需求表路径')
        while True:
            s = ask('需求表完整路径', config.DEFAULTS['xlsx'])
            if os.path.isfile(s):
                return s
            bad('没这个文件：%s' % s)


def validate_xlsx(path):
    """返回 (info, 是否可用)。工作表不对时给一次确认机会。"""
    info = read_xlsx_info(path)
    return info


def main():
    ap = argparse.ArgumentParser(description='机混流水线配置向导 / 环境自检')
    ap.add_argument('--check', action='store_true', help='只自检，不改配置')
    ap.add_argument('--deep', action='store_true', help='自检时真连一次腾讯文档')
    ap.add_argument('--reset', action='store_true', help='覆盖已有配置')
    ap.add_argument('--yes', '-y', action='store_true', help='非交互，全部用自动探测/默认值')
    ap.add_argument('--xlsx', help='需求表路径')
    ap.add_argument('--state-dir', help='运行状态目录')
    ap.add_argument('--path-prefix', help='素材根目录（含 0-魔剪-素材）')
    ap.add_argument('--sheet', help='工作表名')
    ap.add_argument('--tdoc-file', help='腾讯文档 file_id')
    ap.add_argument('--tdoc-sheet', help='腾讯文档 sheet_id')
    ap.add_argument('--tdoc-script', help='tencentdocs.py 路径')
    args = ap.parse_args()

    say('')
    say('=' * 68)
    say('  机混（魔剪）需求流水线 · 配置向导')
    say('=' * 68)

    cur = config.load_config_file()
    cfg = dict(config.DEFAULTS)
    cfg.update({k: v for k, v in cur.items() if v not in (None, '')})
    cfg.setdefault('state_dir', '')

    # ---- 只自检 ----
    if args.check:
        if not cur and not args.xlsx:
            warn('还没有配置文件，下面是「出厂默认值」的检测结果')
            warn('要配置请运行：python setup.py')
        errors, _ = doctor(cfg, deep=args.deep)
        sys.exit(1 if errors else 0)

    if cur and not args.reset and not args.yes and not args.xlsx:
        say('')
        say('已存在配置：%s' % config.CONFIG_FILE)
        say('  需求表：%s' % cur.get('xlsx'))
        s = ask('要重新配置吗？输入 y 继续，直接回车=保留现有配置并自检 (y/N)', 'n')
        if s.lower() not in ('y', 'yes', '是'):
            doctor(cfg, deep=args.deep)
            return

    # ---- 1) 需求表 ----
    xlsx = pick_xlsx(args.yes, args.xlsx or '')
    xlsx = os.path.abspath(xlsx)
    while True:
        try:
            info = read_xlsx_info(xlsx)
        except SystemExit as e:
            bad(str(e))
            if args.yes:
                raise
            xlsx = os.path.abspath(ask('换个路径', xlsx))
            continue
        if info['sheet_found']:
            blocks = block_rows(info['rows'])
            proj = [r for r in info['rows'] if r[0] > 1 and not r[1].endswith('项目')]
            ok('已读取：%d 个运营块 / %d 个项目行' % (len(blocks), len(proj)))
            break
        bad('这个表里没有工作表「%s」' % config.SHEET_NAME)
        say('    现有工作表：%s' % '、'.join(info['sheets']))
        if args.yes:
            warn('按 --yes 继续，稍后可用 --sheet 指定正确的工作表名')
            break
        s = ask('换一个文件？(y=重选 / n=就用它)', 'y')
        if s.lower() in ('n', 'no', '不'):
            break
        xlsx = os.path.abspath(ask('需求表路径', xlsx))

    sheet = args.sheet or (cfg.get('sheet_name') or config.SHEET_NAME)
    if not args.yes and args.sheet is None and info['sheets'] and sheet not in info['sheets']:
        sheet = info['sheets'][0]

    # ---- 2) 素材根目录 ----
    say('')
    say('[2/5] 素材根目录（新增项目时用来生成「项目文件地址」）')
    auto_prefix = args.path_prefix or derive_prefix(info['rows']) or cfg.get('path_prefix', '')
    if auto_prefix:
        ok('从表里已有项目自动推导：%s' % auto_prefix)
    else:
        warn('表里没有可参考的项目地址，只能手填')
    prefix = auto_prefix if (args.yes or auto_prefix) else ''
    if not args.yes:
        s = ask('确认素材根目录（回车=用上面的）', auto_prefix)
        prefix = s or auto_prefix
    if prefix:
        if os.path.isdir(prefix):
            ok('目录可访问')
        else:
            warn('目录当前访问不到（网络盘没连？），不影响读表写表，只影响新增项目写路径')

    # ---- 3) 状态目录 ----
    say('')
    say('[3/5] 运行状态目录（核对基准、报告、台账、备份都放这里）')
    default_state = (args.state_dir or cfg.get('state_dir')
                     or os.path.join(os.path.dirname(xlsx), 'mojian-runtime'))
    state_dir = default_state
    if not args.yes:
        s = ask('确认状态目录（回车=用默认）', default_state)
        state_dir = s or default_state
    state_dir = os.path.abspath(state_dir)
    if dir_writable(state_dir):
        ok('可写')
    else:
        bad('不可写，请换一个目录')
        if not args.yes:
            state_dir = os.path.abspath(ask('换个状态目录', state_dir))

    # ---- 4) 导出目录（只读展示，来自表内 B1）----
    say('')
    say('[4/5] 成片导出目录（读自表内「%s」页 B1）' % config.OUTPUT_DIR_SHEET)
    out = info['output_dir']
    if out:
        say('    %s' % out)
        if os.path.isdir(out):
            ok('当前可访问')
        else:
            warn('当前访问不到（网络盘掉线属正常），核对时自动回退 UNC')
    else:
        warn('表里没读到，核对前请确认「%s」页 B1 有值' % config.OUTPUT_DIR_SHEET)

    # ---- 5) 腾讯文档 ----
    say('')
    say('[5/5] 腾讯文档《需求对接表》')
    fid = args.tdoc_file or cfg.get('tdoc_file_id') or config.DEFAULTS['tdoc_file_id']
    sid = args.tdoc_sheet or cfg.get('tdoc_sheet_id') or config.DEFAULTS['tdoc_sheet_id']
    if not args.yes:
        fid = ask('file_id', fid) or fid
        sid = ask('sheet_id', sid) or sid
    ok('file_id=%s  sheet_id=%s' % (fid, sid))

    # 只存用户显式指定的；自动搜到的不写死，否则插件升级换版本号后路径就失效
    script = args.tdoc_script or ('' if args.reset else cfg.get('tdoc_script', ''))
    if script and not os.path.isfile(script):
        warn('配置里的 tdoc_script 已失效，改为自动搜索：%s' % script)
        script = ''
    found = script or find_tdoc_script()
    if found:
        ok('tencentdocs.py：%s' % found)
        if not script:
            say('       （自动搜索，未写死，插件升级后仍能找到）')
    else:
        warn('没找到 tencentdocs.py，读表前需要安装 tencent-docs 插件')

    # ---- 落盘 ----
    cfg = {
        'version': 1,
        'xlsx': xlsx,
        'sheet_name': sheet,
        'output_dir_sheet': cfg.get('output_dir_sheet') or config.OUTPUT_DIR_SHEET,
        'state_dir': state_dir,
        'path_prefix': prefix,
        'tdoc_file_id': fid,
        'tdoc_sheet_id': sid,
        'tdoc_script': script,
        'updated_at': datetime.datetime.now().isoformat(timespec='seconds'),
    }
    saved = config.save_config_file(cfg)
    say('')
    ok('配置已保存：%s' % saved)

    errors, _ = doctor(cfg, deep=args.deep)
    if errors:
        say('')
        say('还有问题没解决，按上面的 [XX] 提示处理后重新运行：python setup.py --check')
        sys.exit(1)
    say('')
    say('可以开始用了：python read_tdoc.py')


if __name__ == '__main__':
    main()
