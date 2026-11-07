---
name: mojian-jihun-xqb
description: 机混/魔剪需求的每日流水线——读取腾讯文档《需求对接表》「—剪辑需求—」页当天日期块，按备注里的「N魔剪/N混剪/N机混」筛出需求，写入本地《机混需求表》「项目需求_项目路径_爆款拆分」页，并核对 U 盘导出目录里当天成片的数量与编号。当用户说「今天有哪些项目需要魔剪/混剪」「把魔剪数量更新到机混需求表」「读取剪辑需求页」「核对」「记住/记一下（表格改过了）」，或提到机混需求表、爆款拆分、自动化魔剪成片时使用。
agent_created: true
version: 1.0.0
---

# 机混（魔剪）需求每日流水线

处理一条固定的每日链路：

```
腾讯文档「—剪辑需求—」当天块
      ↓ 按备注筛 N魔剪 / N混剪 / N机混
    plan.json（含本地行映射）
      ↓ 清空昨日数量 + 写入今日数量
本地《机混需求表》「项目需求_项目路径_爆款拆分」
      ↓ snapshot.py 固化基准
      ↓ verify.py 比对 U 盘成片
   核对报告
```

业务判定规则、映射别名、清理口径全部在 `references/business-rules.md`，
**动手前先读它**。改 XML 的坑在 `references/xlsx-xml-editing.md`，
腾讯文档接口细节在 `references/tdoc-api.md`。

## 开工前（每次任务都先确认一次）

```bash
cd <skill>/scripts && python setup.py --check
```

- 显示「自检全部通过」或只有 `[!!]` 提醒（网络盘掉线等）→ 直接干活。
- 出现 `[XX]` 或提示「配置文件不存在」→ **先带用户完成初始化**（见下面第 0 步），
  不要带着错误路径继续跑。用户不熟悉命令行也没关系：向他确认「本机《机混需求表》放在哪个盘 / 哪个文件夹」，
  然后用 `python setup.py --xlsx "<用户给的路径>" --yes` 落盘，再 `--check` 复核。

## 第 0 步：首次使用先初始化（每台机器只做一次）

脚本里**不写死任何本机路径**，所有位置来自配置文件
`~/.workbuddy/mojian-jihun-xqb/config.json`（在技能目录之外，重装技能不会丢）。

换机器 / 第一次用，先跑配置向导：

```bash
cd <skill>/scripts
python setup.py            # 交互式：自动搜索本机需求表 → 让你选 → 自检
python setup.py --check    # 只自检，不写配置
python setup.py --xlsx "D:\xx\机混需求表.xlsx" --yes   # 无人值守
```

向导会自动做这几件事，不用你手抄路径：

1. 搜索本机 `机混需求表.xlsx`（含桌面 `.lnk` 快捷方式反解、各数据盘浅扫），列成菜单让用户选；
2. 从表里已有项目的「项目文件地址」反推素材根目录（`…\0-魔剪-素材`）；
3. 从表内「输出目录」页 B1 读出成片导出目录，并检查是否可访问；
4. 自动找到 `tencentdocs.py`；
5. 最后逐项自检（表存在 / 工作表对得上 / 状态目录可写 / 网络盘是否在线）。

**判断是否已初始化**：`python setup.py --check`，或看 `config.json` 是否存在。
路径不对时其它脚本会直接提示「请先运行 setup.py」，不会抛 `FileNotFoundError`。

配置项与环境变量覆盖：

| 配置键 | 环境变量 | 说明 |
|---|---|---|
| `xlsx` | `MOJIAN_XLSX` | 本地《机混需求表》全路径 |
| `state_dir` | `MOJIAN_STATE_DIR` | 基准/报告/台账/备份目录，默认需求表同盘 `mojian-runtime` |
| `path_prefix` | `MOJIAN_PATH_PREFIX` | 素材根目录（新增项目写地址用） |
| `sheet_name` | `MOJIAN_SHEET` | 承接需求的工作表名 |
| `tdoc_file_id` / `tdoc_sheet_id` | `MOJIAN_TDOC_FILE` / `MOJIAN_TDOC_SHEET` | 腾讯文档坐标 |
| `tdoc_script` | `MOJIAN_TDOC_SCRIPT` | tencentdocs.py 路径，默认自动搜索 |
| — | `MOJIAN_CONFIG` | 换一份配置文件（测试用） |

成片导出目录**不在配置里**，它读自需求表「输出目录」页 B1——改导出位置改表即可。
U 盘是网络盘，掉线时 `verify.py` 会自动回退到映射盘的 UNC 远程路径（从注册表读 `RemotePath`）。

脚本用本机 Python 3.9+ 运行（`python setup.py`），不依赖 openpyxl，全部走 zip+XML。

## 任务一：读表 → 写表（最常用）

```bash
cd <skill>/scripts

# 1) 读今天的需求，自动映射到本地行，输出 plan.json
python read_tdoc.py                 # 加 --date YYYY-MM-DD 指定日期；加 --dump 看原始块

# 2) 若 plan 里有 action=new 的项目，先建行（一项一项来，别批量）
python add_projects.py items.json            # 先预览
python add_projects.py items.json --apply    # 生成 _new_ 后自行覆盖原表
python check_insert.py <原表备份> <新表>      # 自检

# 3) 写完新行后重跑 read_tdoc.py，行号会自动更新

# 4) 按计划写入数量（清空昨日 + 写入今日）
python apply_plan.py plan.json               # 预演
python apply_plan.py plan.json --apply       # 落盘

# 5) 刷新核对基准
python snapshot.py
```

`read_tdoc.py` 输出的表格要原样转述给用户：项目、数量、组数、运营、本地行号。
**必须单独说明跳过了哪些项目及原因**（备注非魔剪、数量不明、需用户确认的「拼接」等）。

`apply_plan.py` 的清理口径：

| 场景 | 参数 |
|---|---|
| 默认：清掉前一天全部数量（含手填） | 不加参数 |
| 用户只要求处理部分运营 | `--clear 运营甲,运营丁` |
| 用户明确说某项保留 | `--keep 20,47` |
| 只想覆盖今天有的行、不动其它 | `--clear none` |

## 任务二：核对

用户说「核对」时**只跑 `verify.py`，不要重跑 `snapshot.py`**，否则基准会被刷新，昨天的对照就没了。

```bash
python verify.py
```

输出：一致 N / 共 M 条 + 差异明细，并写 `STATE_DIR/核对报告.md`。

- 兼容两种命名格式：流水号（编号必须 1..N 连续）与 组号+姓名（每组恰好 5 人）。
- 每个文件夹通常有 1 个无编号文件，属正常，会在备注里标出。
- **只报差异，绝不擅自移动、删除、重命名任何成片文件。**

汇报时按「一致 / 多出 / 缺少 / 未找到文件夹」分组，缺少的给出去向判断
（空文件夹=没开始或目录选错；少 10 个左右=可能还在跑）。

## 任务三：记住（表格被人改过）

用户说「表格我改了，你记一下」「记住」时：

```bash
python snapshot.py
```

把当前表数量固化为核对基准，并按时间戳存档到 `STATE_DIR/基准存档/baseline_YYYY-MM-DD_HHMMSS.*`。
`baseline.json/md` 每天被覆盖，存档目录用于回查。

## 任务四：维护台账

每天写完表**必须**在 `STATE_DIR/魔剪需求台账.md` 顶部（`# 标题` 说明之后）追加一节，
倒序排（最新在上）：

```markdown
## YYYY-MM-DD（星期X）

**今日文档需求合计 N 项 / M 个视频**（基准即 N 行 / M）。
今日块：腾讯文档 <起>~<止> 行（A 列 `YYYY/M/D 星期X`）。

| 项目（腾讯文档原名） | 本地表项目名 | 数量 | 预估组数 | 运营 | 本地行 |
|---|---|---|---|---|---|
| 示例项目-金牛（30混剪） | 示例项目 | 30 | 6 | 运营甲 | <本地行号> |

**当日决策**：
- 跳过项与原因、新建行、用户口头追加/保留、备份名、基准存档名。
```

## 脚本清单

| 脚本 | 作用 |
|---|---|
| `setup.py` | **首次使用 / 换机器**：搜索本机需求表、生成 config.json、环境自检 |
| `read_tdoc.py` | 读腾讯文档当日块 → 筛魔剪 → 映射本地行 → `plan.json` |
| `apply_plan.py` | 按 plan 清空旧数量并写入新数量，自带 A 列校验与备份 |
| `block_map.py` | 打印/导出「运营 → 项目」块结构与真实行号 |
| `add_projects.py` | 按运营块插入新项目行（`items.json` 驱动） |
| `check_insert.py` | 插行自检：除新增行外原始数据必须逐行原序保留 |
| `snapshot.py` | 固化核对基准 `baseline.json/md` + 时间戳存档 |
| `verify.py` | 比对 U 盘成片数量与编号连续性 |

## 容易翻车的地方

- **路径是每台机器各自的**：换机器先 `python setup.py`。别把上一个人机器上的
  `C:\某目录\…` 直接抄进脚本或文档。
- **行号会平移**：插行后所有下方行号 +1，写数量前一定要重跑 `read_tdoc.py` 或 `block_map.py` 拿新行号。
  台账里的历史行号不要当现值用。
- **别自己发明规则**：规则只认用户明确说出口的话。别把某一天的处理结果固化成规则。
- **别批量插行**：同一运营块一次插两项会拿到相同编号，必须逐项 `--apply`。
- **形近字**：本地表有固定用字时，历史成片可能出现形近字，核对时做归一化。
- **改表前先备份**，脚本已内置，手工改动前也要自己备份。
