# -*- coding: utf-8 -*-
"""
集中配置：路径、腾讯文档坐标、业务常量。

优先级（高 → 低）：
    1. 环境变量（临时覆盖 / 测试用）
    2. 用户配置文件  ~/.workbuddy/mojian-jihun-xqb/config.json   ← 由 setup.py 生成
    3. 本文件里的出厂默认值

换机器 / 换目录**不要改这个文件**，跑一次：
    python setup.py
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _env(name, default=None):
    v = os.environ.get(name)
    return v if v else default


# 用户配置文件放在技能目录之外，重装技能不会丢
CONFIG_FILE = _env('MOJIAN_CONFIG',
                   os.path.join(os.path.expanduser('~'), '.workbuddy',
                                'mojian-jihun-xqb', 'config.json'))

# ---- 出厂默认值（公开版全部留空，必须由 setup.py / config.json 提供）----
# 本仓库不含任何真实本机路径、腾讯文档 ID 或人员名单；首次使用先跑 setup.py。
DEFAULTS = {
    'xlsx': '',                 # 本机《机混需求表》绝对路径
    'sheet_name': '项目需求_项目路径_爆款拆分',
    'output_dir_sheet': '输出目录',
    'state_dir': '',            # 留空 = 需求表同盘的 mojian-runtime
    'path_prefix': '',          # 素材根目录，由 setup.py 从表内反推
    'tdoc_file_id': '',         # 自己的腾讯文档 file_id
    'tdoc_sheet_id': '',        # 自己的腾讯文档 sheet_id
    'tdoc_script': '',          # 留空 = 自动搜索 tencent-docs 插件
}


def load_config_file(path=None):
    """读用户配置；不存在或损坏时返回空 dict（不抛异常）。"""
    p = path or CONFIG_FILE
    try:
        if os.path.isfile(p):
            with open(p, encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


def save_config_file(data, path=None):
    p = path or CONFIG_FILE
    d = os.path.dirname(p)
    if d:
        os.makedirs(d, exist_ok=True)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write('\n')
    return p


_FILE = load_config_file()
IS_CONFIGURED = os.path.isfile(CONFIG_FILE)


def _get(key, env_name=None, fallback=None):
    if env_name and os.environ.get(env_name):
        return os.environ[env_name]
    v = _FILE.get(key)
    if v not in (None, ''):
        return v
    if fallback is not None:
        return fallback
    return DEFAULTS.get(key, '')


# ---- 本地需求表 ----
XLSX = _get('xlsx', 'MOJIAN_XLSX')
SHEET_NAME = _get('sheet_name', 'MOJIAN_SHEET')
OUTPUT_DIR_SHEET = _get('output_dir_sheet')      # 导出父级路径在该页 B1

# ---- 运行状态目录（基准 / 报告 / 存档 / 备份），默认与需求表同盘 ----
STATE_DIR = (_env('MOJIAN_STATE_DIR')
             or _FILE.get('state_dir')
             or os.path.join(os.path.dirname(XLSX) or HERE, 'mojian-runtime'))

# ---- 新增项目时的默认值 ----
PATH_PREFIX = _get('path_prefix', 'MOJIAN_PATH_PREFIX')
DEFAULT_DEPT = '默认部门'
DEFAULT_SEGMENT = '1-1-1-1-1'
# 运营块名前缀（本地表 A 列形如「运营甲项目」），新增项目时用于定位块
OPS_BLOCK_SUFFIX = '项目'

# ---- 腾讯文档《需求对接表》----
TDOC_FILE_ID = _get('tdoc_file_id', 'MOJIAN_TDOC_FILE')
TDOC_SHEET_ID = _get('tdoc_sheet_id', 'MOJIAN_TDOC_SHEET')
TDOC_SCRIPT = _get('tdoc_script', 'MOJIAN_TDOC_SCRIPT')
TDOC_PLUGIN_GLOB = os.path.join(
    os.path.expanduser('~'), '.workbuddy', 'plugins', 'cache',
    'workbuddy-builtin', 'tencent-docs-plugin', '*', 'skills', 'tencent-docs', 'tencentdocs.py')

# 列索引（0 基）：A日期 B项目 C运营 D账户ID E总数 F备注 G~L 六个人
COL_DATE, COL_PROJECT, COL_OPS, COL_ACCOUNT = 0, 1, 2, 3
COL_TOTAL, COL_NOTE = 4, 5
COL_PEOPLE_START, COL_PEOPLE_END = 6, 11
COL_MAX = 26

# 备注里识别魔剪数量：N魔剪 / N混剪 / N机混 三种写法都要认
MAGIC_RE = r'(\d+)\s*(?:魔剪|混剪|机混)'
# 明确「不要魔剪」的备注直接跳过
SKIP_RE = r'不要(?:魔剪|混剪|机混)'

# ---- 业务常量 ----
# 成片命名「组号+姓名」格式用到的剪辑师名单。公开版为占位值，
# 实际名单写进 config.json 的 "people" 字段（见 config.example.json）。
_people = _FILE.get('people')
PEOPLE = _people if isinstance(_people, list) and _people else [
    '剪辑师甲', '剪辑师乙', '剪辑师丙', '剪辑师丁', '剪辑师戊']
GROUP_SIZE = 5          # 数量 = 组数 × 人数

# ---- 核对 ----
VIDEO_EXT = ('.mp4', '.mov', '.avi', '.mkv')


# ------------------------------------------------------------------
# 前置检查：路径没配好时给出人话提示，而不是抛 FileNotFoundError
# ------------------------------------------------------------------
SETUP_HINT = """
这个路径来自出厂默认值，换机器后需要先指定本机实际位置。

请先在 scripts 目录运行一次配置向导（会帮你搜索候选文件）：

    python setup.py            # 交互式配置
    python setup.py --check    # 只看当前配置与自检结果

配置文件：{cfg}
""".strip()


def require_xlsx(path=None):
    """确认需求表可读；不可读则打印指引并退出（退出码 2）。"""
    p = path or XLSX
    if os.path.isfile(p):
        return p
    sys.stderr.write(
        '\n[未配置] 找不到本地《机混需求表》：\n    %s\n\n%s\n'
        % (p, SETUP_HINT.format(cfg=CONFIG_FILE)))
    sys.exit(2)


def require_state_file(name, hint=None):
    """确认状态目录里的某个文件存在（如 baseline.json）。"""
    p = os.path.join(STATE_DIR, name)
    if os.path.isfile(p):
        return p
    sys.stderr.write(
        '\n[缺少状态文件] %s\n    期望位置：%s\n%s\n'
        % (hint or name, p,
           '请先运行 snapshot.py 生成核对基准。' if name.startswith('baseline')
           else '请检查 STATE_DIR 配置（python setup.py --check）。'))
    sys.exit(2)


def describe():
    """给 setup.py --check / 排查用。"""
    return {
        'config_file': CONFIG_FILE,
        'config_exists': IS_CONFIGURED,
        'xlsx': XLSX,
        'xlsx_exists': os.path.isfile(XLSX),
        'sheet_name': SHEET_NAME,
        'state_dir': STATE_DIR,
        'path_prefix': PATH_PREFIX,
        'tdoc_file_id': TDOC_FILE_ID,
        'tdoc_sheet_id': TDOC_SHEET_ID,
        'tdoc_script': TDOC_SCRIPT,
    }
