# 机混（魔剪）需求每日流水线 · WorkBuddy Skill

> 这是一个 **WorkBuddy Skill**。怎么装进 WorkBuddy、怎么发给别人、怎么同步更新，见 [SHARE.md](SHARE.md)。
> 仓库：https://github.com/yangzq66666/mojian-jihun-xqb-skill （Public）

把《需求对接表》里当天的「N魔剪 / N混剪 / N机混」需求，写进本地《机混需求表》，
再核对导出目录里当天成片的数量与编号。

## 公开版说明（先读这段）

本仓库是**脱敏后的公开版**：不含任何真实员工姓名、投放项目/品牌名、腾讯文档 ID、
本机路径或内部业务决策记录，全部换成了中性示例。

所以 **拿到就能装，但不能直接跑出正确结果**——首次使用必须按自己团队的情况配置：

| 要填什么 | 在哪填 |
|---|---|
| 本机《机混需求表》路径、素材根目录、运行状态目录 | 跑 `python setup.py`，生成 `~/.workbuddy/mojian-jihun-xqb/config.json` |
| 自己的腾讯文档 `file_id` / `sheet_id` | 上面那个 config.json，字段见 `config.example.json` |
| 项目名映射（腾讯文档原名 → 本地表项目名） | `references/aliases.json`（示例条目可全删） |
| 剪辑师名单（成片「组号+姓名」命名用） | config.json 的 `people` 字段 |
| 业务判定规则、运营块结构、清理口径 | `references/business-rules.md`，按本机表改写 |
| 渠道/包装后缀正则 | `scripts/read_tdoc.py` 的 `SUFFIX_RE` |

**配置文件在技能目录之外**，重装或更新 skill 不会丢。

## 三步上手

```bash
cd scripts
python setup.py            # 交互式：搜索本机需求表 → 生成 config.json
python setup.py --check    # 自检
```

自检通过后按 [SKILL.md](SKILL.md) 跑每日流水线。

## 依赖

- Python 3 + 标准库
- 读腾讯文档需要 WorkBuddy 的「腾讯文档」连接器；接口要点见 `references/tdoc-api.md`
- 处理 xlsx 走原生 XML，不需要 openpyxl 写入

## 目录

```
SKILL.md                     主入口（触发词 + 四个任务的工作流）
config.example.json          配置示例（无真实值）
references/aliases.json      项目名映射示例
references/business-rules.md 业务规则模板
references/tdoc-api.md       腾讯文档接口要点
references/xlsx-xml-editing.md 改 xlsx XML 的坑
scripts/                     10 个脚本，见 SKILL.md 的脚本清单
```
