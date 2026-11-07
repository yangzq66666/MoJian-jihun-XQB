# 机混（魔剪）需求 Skill · 分享与同步更新指南

## 〇、版本仓库

| 项 | 值 |
|---|---|
| 仓库 | `https://github.com/yangzq66666/mojian-jihun-xqb-skill`（**Public**） |
| git 地址 | `https://github.com/yangzq66666/mojian-jihun-xqb-skill.git` |
| 本地目录 | `~/.workbuddy/skills/mojian-jihun-xqb` |
| 分支 | `main` |
| 版本 | 见 `SKILL.md` frontmatter 的 `version` |

> ⚠️ 用户名是 `yangzq66666`（**5 个 6**），多打一个会报 `Repository not found`
> （GitHub 对私有/不存在的仓库都用 404 伪装，容易误判成认证失败）。
> 以后改完功能：`git add -A && git commit -m "..." && git push` 即可，无需重新配置。

## 一、打包（1 条命令）

`git archive` 只打**已提交**文件；`config.json`、运行产物、备份、缓存都被 `.gitignore` 排除，
导出的 zip 天然干净，不用手工挑文件。

```sh
cd ~/.workbuddy/skills/mojian-jihun-xqb
git archive --format=zip --prefix=mojian-jihun-xqb/ \
  -o "../mojian-jihun-xqb_v1.0.0.zip" HEAD
```

## 二、别人怎么装进 WorkBuddy

**方式 A（最稳）**：把 zip 发过去 → 对方在 WorkBuddy 左侧
「专家 · 技能 · 连接器 → 技能 → 添加技能 → 上传技能」拖进去。

**方式 B（自助）**：GitHub 仓库页 → 绿色 **Code → Download ZIP** → 把**原始 zip** 拖进「上传技能」。

> ⚠️ GitHub 自动 zip 顶层是 `mojian-jihun-xqb-skill-main/`，不是 `mojian-jihun-xqb/`。
> **不要解压后重新压缩**（多套一层会报「找不到 SKILL.md」）。

**方式 C（以后更新最省事）**：对方新建一段对话，把下面整段贴进去：

```
帮我把下面这个公开仓库的 Skill 安装到 WorkBuddy。核心要求：绝对不要创建或覆盖任何本机配置。
仓库：https://github.com/yangzq66666/mojian-jihun-xqb-skill
分支：main
目标目录：~/.workbuddy/skills/mojian-jihun-xqb/

步骤：
1. 若目标目录已存在且有 .git，改为 git pull 更新
2. 匿名下载 main 分支全部文件，平铺放进目标目录（不要多套一层 mojian-jihun-xqb-skill-main/）
3. 确认 SKILL.md 在目标目录根部，frontmatter 里 name 是 mojian-jihun-xqb
4. 跑 python scripts/setup.py --check 看缺哪些配置，并告诉我按 README「公开版说明」逐项填
5. 不要生成或覆盖 ~/.workbuddy/mojian-jihun-xqb/config.json（那是用户自己的配置）
6. 提醒我重启 WorkBuddy 才生效

若 GitHub 直连超时，别反复重试，告诉我改用手传 zip。
```

## 三、装完必做（公开版特有）

这个 skill 是**脱敏版**，所有业务数据都要自己填，详见 `README.md`「公开版说明」：

1. `cd scripts && python setup.py` 生成 `~/.workbuddy/mojian-jihun-xqb/config.json`
2. 按 `config.example.json` 补自己的腾讯文档 `file_id` / `sheet_id`、`people` 名单
3. 按本机表改 `references/aliases.json` 与 `references/business-rules.md`
4. `python setup.py --check` 自检
5. **重启 WorkBuddy** 刷新技能索引

> 配置在技能目录之外，任何更新都覆盖不到；不要删目录重建。

## 四、更新怎么同步

| 安装方式 | 你 push 新版后对方要做的 |
|---|---|
| 方式 C（git 目录） | 一句「更新 mojian-jihun-xqb skill」，自动 `git pull` |
| 方式 A / B（zip 导入） | 收新 zip 覆盖解压到原目录，**别删目录重建** |

## 五、公开版脱敏清单（已全部处理）

| 类别 | 处理 |
|---|---|
| 员工姓名、运营姓名 | 换成「剪辑师甲…」「运营甲…」占位 |
| 投放项目/品牌名（20 多个） | `aliases.json` 整份换成虚构示例；业务规则里的易错项表换成示例 |
| 腾讯文档 `file_id` / `sheet_id` | 从代码默认值与文档里清空，改 `YOUR_*` 占位，由使用者自填 |
| 本机路径、内部盘符、部门名 | 清空或换成 `C:\你的目录\…`；`DEFAULT_DEPT` 改中性 |
| 内部决策记录（某日谁确认过等） | 改成中性规则陈述 |
| 台账示例里的真实日期与数量 | 换成 `YYYY-MM-DD` / `N` / `M` 占位 |
| 形近字示例（疑似真实项目名） | 换成通用说明 |
| 密钥 / API Key / token | 原本就没有 |

**代码框架完整可用**（所有 `.py` 通过 `py_compile`），但业务数据需自填才能跑出正确结果。

## 六、每次改完功能必做

1. bump `SKILL.md` frontmatter 的 `version`
2. 在下面变更记录补一行
3. `git commit` → `git push` → 重新出包

## 变更记录

| 版本 | 日期 | 变更 |
|---|---|---|
| 1.0.0 | 2026-11-07 | 首次整理为可公开分发版：全量脱敏、补 `.gitignore` / `README.md` / `SHARE.md` / `config.example.json`，业务常量改为可配置 |
