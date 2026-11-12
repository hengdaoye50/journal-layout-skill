---
name: journal-layout
description: 把一篇中文论文（Markdown 或 docx 底稿）排版成国内政治学与公共管理四本顶刊（《政治学研究》《中国行政管理》《公共管理学报》《管理世界》）的投稿版式：页面与边距、标题层级、正文行距、页下注编号、三线表、图件与题注、双栏版式，并用结构断言校验而不是逐页出图。当用户说"排版""投稿版式""按某刊格式改""把论文排成期刊格式""改成参考文献格式/页下注格式""论文转Word"时使用。
---

# 期刊投稿版式排版

把同一份论文底稿派生出四个刊的投稿版本，并**证明**它符合该刊版式。
版式是数值问题，所以用数值断言验收，不靠逐页截图肉眼比对。

## 什么时候用

- 用户要把论文排成国内期刊投稿格式，尤其是上述四刊或同类社科期刊
- 用户已有 Word 稿，要统一字体字号行距、标题层级、页下注、三线表
- 用户要在多个刊的版式之间批量切换，共用同一份正文
- 用户抱怨"排完看着不对/页数变了/图题位置不对"

不适用：写论文内容、改论证、润色文字（那是写作，不是排版）。需要生成 HTML/PPT 的排版走别的 skill。

## 先问清一件事：要"投稿版式"还是"刊出版式"

这是最容易做错方向的判断，做错就全功尽弃。

| | 投稿版式 `--mode submission` | 刊出版式 `--mode published` |
|---|---|---|
| 像什么 | 交给编辑部的稿子：字大、单栏、行距松，好读好批注 | 已经印出来的那一页：小字密排、可能双栏 |
| 依据 | 投稿须知 + 已交付成品 | 已刊例文 PDF 实测 |
| 正文典型值 | 宋体 12pt / 固定行距 17–19pt | 9.4–11.7pt / 行距 15–19pt |
| 何时用 | 用户说"投稿""送审""按某某刊要求排版式" | 用户说"排成和期刊一样""仿样张""像刊出来的样子" |

两套参数都在 `scripts/` 里：`journal_layout.py` 的 `SPECS`（submission）与
`published_specs.py` 的 `PUBLISHED_SPECS`（published）。用 `--mode` 切换，字段结构完全一致。
用户没说清时**问一句**，不要默认。

## 三条硬原则

1. **版式参数必须有出处。** 每个数值来自三处之一：官方投稿须知、已刊例文实测、已交付成品反查。
   拿不准就留 `None`（脚本会跳过并提示"未确认"），绝不猜一个看起来合理的默认值。
   最容易犯的错是沿用上一个项目的习惯——比如《社会》要求图题在下，但很多刊并非如此，必须实测。
2. **只改版式，不改文字。** 除声明过的标签替换（如"摘要："→"内容提要："），
   排版脚本不得增删任何字符。跑完必须用 `--ref` 比对文本摘要，证明正文一字未动。
3. **改之前留档，改之后断言。** 批量改成品前先 `snapshot.py --backup`，
   改完用 `snapshot.py --diff` 证明只有声明过的 OOXML 部件变了。

## 七步流程

```bash
PY=python   # 需要 python-docx / lxml / Pillow；可选 PyMuPDF 仅用于抽页预览
S=<本 skill 目录>/scripts
```

### 1 建项目骨架与素材归档

```
01_四刊最终稿/        ← 交付物（每刊一个 docx）
02_投稿指南与例文/     ← 每刊一个子目录：投稿须知 docx + 已刊例文 pdf
03_历史稿与原稿/       ← 每批次改动前的留档，目录名带日期与阶段
04_论文图件/           ← 图的 SVG 源与 PNG 成品
05_版式检查与预览/     ← 校验报告、抽页预览
06_处理脚本与临时数据/  ← 一次性脚本与中间文件
```

素材不是可选项：**没有例文 pdf 就没有版式依据**。找不到例文时，明确告诉用户"这几项没法确认"，
而不是凭经验填。

### 2 测量例文，定参数

先看 `references/journal-specs.md`（四刊参数表，含实测值/须知原文/脚注之争）。
要新加一个刊时，用同一套方法测量，写进 `scripts/journal_layout.py` 的 `SPECS`：

```bash
$PY $S/journal_layout.py --list          # 看现有四刊参数
```

测量要点见 `references/journal-specs.md` 的"测量方法"一节：页面 mediabox → 边距；
正文字号 → 从字符 bbox 高度推；行距 → 相邻行基线差；单双栏 → 正文文本块 x 范围。

### 3 Markdown → 语义化底稿

底稿是"版式中性"的中间态：题名/标题层级/正文/图表题注/参考文献/页下注各就各位，
但不带任何刊的具体数值。同一份底稿派生四刊。

```bash
$PY $S/md2docx.py --md 全文.md --out 03_历史稿与原稿/底稿.docx \
    --figure-dir 04_论文图件 --figure-width-cm 14.8
```

- 支持 `# 题名` / `## 一级` / `### 二级` / `**图1 题**` / 管道表格 / `[^1]` 页下注
- 页下注在**这一步**就变成真正的 Word footnote（`scripts/footnotes.py` 写 OOXML）
- 表格默认输出三线表
- 已有 docx 底稿可跳过本步，直接进第 4 步

### 4 套用期刊版式

```bash
$PY $S/journal_layout.py --journal politics --mode submission \
    --src 03_历史稿与原稿/底稿.docx --out 01_四刊最终稿/论文_政治学研究.docx
$PY $S/journal_layout.py --journal china-admin --mode published --src ... --out ...   # 仿刊出样张（双栏）
```
`--journal` 可选 `politics` / `management` / `china-admin` / `public-admin`；
`--mode` 可选 `submission`（默认）/ `published`。先 `--dry-run` 看会动什么，`--list` 看参数。

### 5 结构校验（这是主验收手段）

```bash
$PY $S/verify_layout.py --docx 01_四刊最终稿/论文_政治学研究.docx \
    --journal politics --mode submission --ref 03_历史稿与原稿/底稿.docx --strict
```

一次断言十几项：页面/边距、每类段落的字体字号行距缩进、分栏、页下注编号与字体、
图宽是否超通栏、图题/表题位置、三线表竖线、正文文本是否被改动。
输出 `PASS` / `FAIL` 与不符合项清单；`--strict` 时不符合即退出码 1。

**不要用逐页出图代替这一步。** 三十页图既慢又会漏，且人眼对 0.5pt 的字号差不敏感。
只有"图挤到页底""表格跨页断裂"这类排版观感问题才需要看图，用第 6 步抽 2–3 页。

### 6 必要时抽页预览

```bash
$PY $S/render_preview.py --docx 成品.docx --pages 1 7 12   # 默认只出首页+末页
```
需要本机装了 LibreOffice。它会转 PDF 并把指定页拼成一张缩略图。
**注意**：`soffice` 必须带 `-env:UserInstallation=` 指向可写目录，否则 Windows 上直接崩。

### 7 图件（如论文有示意图）

```bash
$PY $S/figlib.py                    # 不是可执行入口，是被 import 的绘图原语库
$PY examples/figure1_demo.py        # 画的示例
$PY $S/embed_figure.py --docx 成品.docx --image 04_论文图件/图1.png --dry-run
$PY $S/embed_figure.py --docx 成品.docx --image 04_论文图件/图1.png --index 1
```

作图与嵌入的约定见 `references/figure-specs.md`。核心纪律：SVG 是**可编辑源**，PNG 是**嵌入件**，
两者同源生成；替换 docx 内嵌图时断言正文文本不变、只有 `document.xml` 与对应 media 变化。

## 验收门（交付前必须全绿）

| 检查 | 手段 | 判读 |
|---|---|---|
| 版式数值符合该刊 | `verify_layout.py --journal X --strict` | 必须 PASS |
| 正文没被排版改动 | 同上 `--ref 底稿.docx` | 文本完整性 ✓ |
| 只改了该改的部件 | `snapshot.py --diff 备份目录 --against 成品目录 --allow 允许前缀` | PASS |
| 未确认项已向用户说明 | verify 输出的"未核对项" | 逐条交代，不隐瞒 |
| 图与论证对应 | `references/figure-specs.md` 的核对表 | 每张图列"图中关系/正文依据/表达边界" |

## 常见偏差与修法

| 症状 | 原因 | 修法 |
|---|---|---|
| 中文字体没生效，变成等线/Calibri | 只设了 `ascii`/`hAnsi`，没设 `eastAsia` | 必须同时设 `w:rFonts/@w:eastAsia` |
| 校验报"字体未设置" | 格式继承自样式，run 上没有 `rPr` | 校验器已回落样式链；真缺才补 |
| 页下注没出现 | python-docx 不支持 footnotes | 用 `footnotes.py` 直接写 OOXML（需 rels + content-type） |
| 页下注编号是 1,2,3 不是 ①②③ | `sectPr/w:footnotePr` 没设 `numFmt` | 设 `decimalEnclosedCircle`；重排设 `eachPage`/`continuous` |
| 图题和现实出版物位置相反 | 沿用了别的刊/别的项目的习惯 | 实测例文后改 `SPECS[...]["figure"]["caption_position"]` |
| 双栏版式把表/图挤进半栏 | 没给表和图单独开 1 栏区段 | `span_tables` / `span_figures` 会插区段边界 |
| 表格有网格竖线 | 用了 `Table Grid` 样式 | 设 `tblBorders` 内框为 none，只留顶/底/表头线 |
| 排版后页数暴涨 | 行距设成了倍数而非固定值 | 中文期刊多为固定行距（如 19pt），不是 1.5 倍 |

## 脚本一览

| 脚本 | 作用 |
|---|---|
| `scripts/journal_layout.py` | 投稿版式参数表（`SPECS`）+ 套版算法；`--mode` 切换两套参数 |
| `scripts/published_specs.py` | 刊出版式参数表（`PUBLISHED_SPECS`，四刊例文 PDF 实测值） |
| `scripts/md2docx.py` | Markdown → 语义化 docx 底稿（含页下注、三线表、插图） |
| `scripts/footnotes.py` | 页下注注入/样式统一/导出（直接操作 OOXML） |
| `scripts/verify_layout.py` | 结构断言校验（主验收手段） |
| `scripts/render_preview.py` | 抽页预览（可选，仅看图感） |
| `scripts/snapshot.py` | 改动前留档、改动后部件级差异断言 |
| `scripts/figlib.py` | 学术示意图绘图原语（Pillow + SVG 双输出，600dpi） |
| `scripts/embed_figure.py` | 替换 docx 内嵌图（保文本、原子写） |

## 参考文档

- `references/journal-specs.md` — 四刊逐项参数（实测/须知/脚本三方核对）+ 测量方法
- `references/workflow.md` — 完整流程、目录约定、留档纪律、交付清单
- `references/figure-specs.md` — 图件规范：通栏宽、灰度、字号、题注、论证核对表
- `references/pitfalls.md` — 踩坑记录：LibreOffice 崩溃、中文路径权限、OOXML 细节
