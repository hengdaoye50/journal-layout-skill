# 示例：《有组织的发现》四刊版式排版

## 论文信息

- **作者与标题**：周黎安，《有组织的发现：中国特色的产业政策执行机制》
- **发表信息**：《中国社会科学》2026 年第 8 期
- **来源**：作者提供的已发表 PDF（本仓库仅收录排版流程的中间产物与成品）

## 这份示例证明了什么

这是本 skill 的**端到端检验**：一篇 2.2 万字的真实已刊论文，只做"提取 → 底稿 → 套版"
三步，就能同时产出四本期刊的两种版本，并全部通过结构断言校验。

| 产出 | 说明 | 校验 |
|---|---|---|
| `out/*_submission.docx` | 投稿版式（交编辑部的稿子：字大、单栏、行距松） | 4/4 PASS |
| `out/*_published.docx` | 刊出版式（仿已刊样张：小字密排，中国行政管理为双栏） | 4/4 PASS |
| `preview/*-p1.png` | 各刊刊出版式的首页渲染 | — |
| `preview/china-admin-p2.png` | 《中国行政管理》双栏正文页（首页为题名/摘要通栏页） | — |
| `preview/management-notes.png` | 《管理世界》文末「注释」节（该刊特有的注释制式） | — |
| `source.md` | 从 PDF 提取、结构化的 Markdown 底稿（排版输入） | 与 PDF 正文逐字对齐 |

规模：2.2 万字、5 个一级标题、15 个二级/三级小标题、**32 条页下注**。
页下注是该刊**每页重新编号**（①②③ 每页从 ① 起），提取时按页配对后再统一编号——
这一点如果按"全篇唯一编号"处理，会只认出 6 条且内容错位（实测踩过）。

《中国行政管理》的双栏区段**另起新页**：LibreOffice 在"页中连续分节 + 单栏转双栏 + 大量页下注"
时会死循环，`columns_new_page` 开关（该刊默认 `True`）是为了让 Word 与 LibreOffice 都能稳定出图。
若你只在 Word 里用，可把它设回 `False`，正文就会像例文那样从首页页中开始双栏。

四刊版式差异一览（全部来自已刊例文实测，不是凭经验估的）：

| | 政治学研究 | 管理世界 | 中国行政管理 | 公共管理学报 |
|---|---|---|---|---|
| 页面 | 210×285 mm | 216×291 mm | A4 | A4 |
| 正文（刊出） | 11.74 pt | 9.51 pt | 10.50 pt | 9.44–10.5 pt |
| 分栏 | 单栏 | 单栏 | **双栏** | 双栏 |
| 注释 | 页下注 ①②③，每页重排 | **文末「注释」节** | 脚注（说明性）+ 文后参考文献 | 页下注 ①②③ |
| 摘要标签 | **内容提要:**（半角冒号） | 摘要： | ［摘　要］ | 摘要： |
| 图题位置 | 未确认（例文无图） | 图下方 | 图下方 | 图下方 |
| 表格制式 | 网格表（带外框） | 网格表（列间竖线） | 未确认（例文无表） | 三线表 |

## 提取环节（PDF → Markdown）

`source.md` 由 PyMuPDF 按版式特征提取，不是手工抄的：

- 正文 12.0 pt、续行左缘 63.8 pt、段首缩进 87.8 pt（2 字）→ 用缩进判断段落边界
- 一级标题 15.7 pt、二级标题 12.0 pt 黑体 → 标题层级
- 页下注 10.5 pt，位于该页最后一条正文行之下，以 ①②… 或 `*` 起头
- 剔除页眉、页码、`〔责任编辑：…〕`，以及第 24 页的英文摘要（ABSTRACTS）页

已知取舍：**英文摘要页与作者简介未收入**（前者是刊物统一栏目，后者各刊处理方式不同，
本 skill 未对署名区做版式建模，`--byline` 只给中立居中样式）。

## 复现命令

命令在 **skill 根目录**执行；`<python>` 换成本机装有 `python-docx`、`lxml`、`Pillow` 的解释器。

```bash
# 1) Markdown → 底稿。《管理世界》要文末注释，其余三刊用页下注，所以建两份底稿
python scripts/md2docx.py --md examples/zhou-li-an/source.md \
    --out /tmp/底稿_footnote.docx --notes-mode footnote \
    --byline 周黎安 --affil "（北京大学光华管理学院）"

python scripts/md2docx.py --md examples/zhou-li-an/source.md \
    --out /tmp/底稿_endnote.docx --notes-mode endnote \
    --byline 周黎安 --affil "（北京大学光华管理学院）"

# 2) 套版：四刊 × 两模式（管理世界用 endnote 底稿）
for j in politics china-admin public-admin; do
  for m in submission published; do
    python scripts/journal_layout.py --journal $j --mode $m \
      --src /tmp/底稿_footnote.docx --out "examples/zhou-li-an/out/${j}_${m}.docx"
  done
done
for m in submission published; do
  python scripts/journal_layout.py --journal management --mode $m \
    --src /tmp/底稿_endnote.docx --out "examples/zhou-li-an/out/management_${m}.docx"
done

# 3) 逐个校验（版式数值 + 文字未被改动）
python scripts/verify_layout.py --docx examples/zhou-li-an/out/politics_published.docx \
    --journal politics --mode published --ref /tmp/底稿_footnote.docx --strict
```

## 免责声明

> **免责声明**
>
> - 本文的著作权归作者周黎安与《中国社会科学》编辑部所有；本仓库收录的排版产物仅为非商业、教育/技术演示用途，不构成对原文的再授权；
> - 不得用于商业用途，不得转载或再发布原文内容；
> - 排版产物中的版式参数来自各刊已刊例文实测，与《中国社会科学》无关，也不代表任何期刊的官方格式；
> - 如权利人要求，将立即删除相关内容；
> - 引用原文请以《中国社会科学》2026 年第 8 期正式发表版本为准。
