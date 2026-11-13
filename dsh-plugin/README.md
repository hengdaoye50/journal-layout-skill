# DSH 插件：journal-layout

把这个仓库里的技能（**中文期刊投稿版式排版**）挂进 DeepSeek Harness 的技能目录。
它是一个 Host-only 的**零依赖**插件：注册一个 bundled skill provider，把仓库根的
`SKILL.md` 交给 Harness 的技能注册表，并把 `scripts/`、`references/` 作为可读资源暴露给 Agent。

## 安装

在 DSH 里对这个目录执行 `plugin_manager` 的 `install_bundle`（target 用**绝对路径**）：

```
D:\CODEX\journal-layout-skill\dsh-plugin
```

或者用命令行：

```bash
dsh plugin install /absolute/path/to/journal-layout-skill/dsh-plugin
```

安装后，新会话的技能目录里会出现 `journal-layout`；插件页里会显示为
**Journal layout skill / 期刊版式排版技能**。要移除就 `plugin_manager` 的 `remove_bundle`。

> 本地目录安装链接的是检出目录，所以插件读取的是仓库根的 `SKILL.md` 与 `scripts/`
> ——技能本体只维护一份，不用在插件里再复一份。

## 它做了什么

- `cordis.patch.yml`：插入一行 `journal-layout-skill`（对应 `lib/index.js`）。
- `lib/index.js`：`ctx.skills.registerProvider()` 注册名为 `journal-layout` 的 provider，
  技能内容 = `SKILL.md` 正文 + 一段运行时说明（`scripts/`、`references/` 的绝对路径，
  以及 Python/LibreOffice 依赖提醒）。
- `locale/*.json`、`icon.svg`：插件页的标题、描述与图标。

技能名固定为 `journal-layout`，排序权重用内置技能层（`BUNDLED_SKILL_RANK = 600`）。

## 与"拷进 ~/.claude/skills"的区别

| | 拷目录 | DSH 插件 |
|---|---|---|
| 生效范围 | 该目录被扫描的宿主（Claude Code、Codex、Harness 的 user root） | 当前 Harness profile，所有会话 |
| 更新 | 重新拷一次 | 改仓库文件即生效（链接的是检出目录） |
| 卸载 | 删目录 | `remove_bundle` |
| 依赖 | 需要宿主支持 skills | 需要 Harness 的技能注册表（`@deepseek-ai/dsh-skill`） |

## 已知限制

- 插件只分发技能（文档 + 脚本），不注册模型工具：排版动作仍由 Agent 按 `SKILL.md`
  的七步流程调用 `scripts/*.py` 完成。
- `resourceBase` 指向检出目录，所以**不要移动或删除**仓库位置；换位置要重新安装。
- 技能脚本需要 `python-docx`、`lxml`、`Pillow`；抽页预览另需 PyMuPDF 与 LibreOffice。
