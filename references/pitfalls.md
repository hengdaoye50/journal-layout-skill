# 踩坑记录

每条都是真跑出来的，不是查文档抄的。按"症状 → 原因 → 修法"写，方便遇到时对号入座。

## 一、环境类

### LibreOffice 转 PDF 直接崩（exit 0xC0000409）
**症状**：`soffice --headless --convert-to pdf` 退出码 0xC0000409，不生成 PDF。
**原因**：soffice 尝试写用户配置目录失败。
**修法**：必须显式指定可写的 profile 目录：

```bash
soffice "-env:UserInstallation=file:///D:/tmp/lo-profile" --headless --norestore \
        --convert-to pdf --outdir out in.docx
```

`render_preview.py` 每次用 `tempfile.mkdtemp()` 新建一个 profile，避免互相锁。

### pwsh 在中文路径工作区全部失败
**症状**：`SetNamedSecurityInfoW failed (Win32 5): grantWrite(D:\CODEX\中文目录)`，
任何命令都起不来。
**原因**：该目录缺少 WRITE_OWNER 权限，沙箱无法完成写授权。
**修法**：用 `diagnose-windows-sandbox-acl` 技能的一条命令修复（会备份并打印回滚命令）。
不是代码问题，别在脚本里绕。

### PowerShell 里写内联多行 Python 会把中文搞坏
**症状**：中文变乱码、语法错误、字符串被转义。
**修法**：一律把代码写成 `.py` 文件再执行。文件开头加

```python
import sys
sys.stdout.reconfigure(encoding="utf-8")
```

注意**只包一次**。已经 reconfigure 过再包一层 `TextIOWrapper`，会因二次包装关闭 buffer
而报 `I/O on closed file`。

## 二、OOXML 类

### python-docx 不支持页下注
**症状**：`Document()` 没有 footnote 接口。
**修法**：直接操作 zip 里的 `word/footnotes.xml`。完整一套要改四处：
1. `word/document.xml`：把 `[n]` 的 run 换成 `w:rPr/w:rStyle=FootnoteReference` + `w:footnoteReference/@w:id`
2. `word/footnotes.xml`：新增 `w:footnote`（第一条 run 是 `w:footnoteRef`，正文 run 前要留一个空格）
3. `word/_rels/document.xml.rels`：加 `Type=.../footnotes` 的关系项
4. `[Content_Types].xml`：加 `Override PartName=/word/footnotes.xml`
另外 `word/footnotes.xml` 里必须有两个特殊项：`w:type="separator"`（id=-1）与
`w:type="continuationSeparator"`（id=0），否则 Word 打开会报错。

### 中文字体设了不生效
**症状**：字体名设成"宋体"，Word 里显示等线/Calibri。
**原因**：只设了 `w:ascii` / `w:hAnsi`，中文走 `w:eastAsia`。
**修法**：
```python
rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
rfonts.set(qn("w:eastAsia"), "宋体")
rfonts.set(qn("w:ascii"), "Times New Roman")
rfonts.set(qn("w:hAnsi"), "Times New Roman")
```
样式（Normal/Title/Heading）也要设，否则新建段落又丢。

### 图片宽度的 XML 位置容易找错
**症状**：想改图宽，搜 `w:extent` 找不到。
**原因**：图片尺寸在 **`wp:extent`**（wordprocessingDrawing 命名空间），不是 `w:extent`。
要同时改 `wp:extent` 和 `a:xfrm/a:ext`，否则 Word 里显示尺寸与文档流尺寸不一致。
EMU 换算：1 cm = 360000 EMU。

### 表格制式搞反了：以为"中文期刊一律三线表"
**症状**：把稿子的表格统一改成三线表，或统一改成网格表，结果与该刊例文不符。
**实测**（数例文 PDF 里表体的竖线段）：
《政治学研究》=带左右外框的网格表（4 竖线）；《管理世界》=横线贯通 + 列间竖线、
无左右外框；《公共管理学报》=标准三线表（0 竖线）；《中国行政管理》例文无表（未确认）。
**修法**：按刊配 `TABLE_STYLE`（`grid-full` / `grid` / `three-line` / `None`），
不要用全局默认值，也不要把"多数期刊用三线表"当规则。
技术上两种都能做：设 `w:tblBorders` 的 `insideH/insideV/left/right` 为 `single` 或 `none`，
三线表再给表头行加 `w:tcBorders/w:bottom`。

### 页下注编号不是圈码
**症状**：脚注显示 1,2,3 而不是 ①②③。
**修法**：在 `w:sectPr/w:footnotePr` 里设
`<w:numFmt w:val="decimalEnclosedCircle"/>`；逐页重编再加 `<w:numRestart w:val="eachPage"/>`。
未显式设置时 Word 默认 `decimal` + `continuous`——校验器按这个默认值判读。

## 三、判断类（最容易错，且错了看不出来）

### 图题在上还是在下：必须实测，不能沿用习惯
**症状**：排版完成，图题位置与已刊例文相反。
**原因**：从上一个项目/别的期刊抄了习惯。中文期刊两种都有：《社会》要求图题在图下方，
而本项目实测的《中国行政管理》《管理世界》《公共管理学报》也都是图题在图**下方**——
但项目里已交付的成品反而把图题放在了上方，属于历史遗留偏差。
**修法**：把例文 PDF 渲染出来看一眼（或用 fitz 抽图块 bbox 与题注行 y 坐标比较），
写进 `SPECS[...]["figure"]["caption_position"]`。`None` 表示未确认，脚本保持原值。

**测量陷阱**：比较"题注距上方最近图块的距离"与"距下方最近图块的距离"时，
标签很容易写反——"图块在题注上方"= 题注在图**下方**。第一次就写反了，
后来渲染页面看图才纠正过来。数值测量之后一定要抽一页眼看确认。

### 校验报"字体未设置"但其实是对的
**症状**：`verify_layout.py` 说 body 字体 `None`，但 Word 里显示就是宋体。
**原因**：格式写在段落样式上，run 上没有 `rPr`。
**修法**：校验器读"有效格式"——run → 段落样式 → base_style 链逐级回落，读到即止。

### 逐页渲染看图代替不了结构校验
**症状**：为了稳妥把三十页全导出 PNG 逐页看，一轮检查十几分钟，还是漏。
**原因**：字体差 0.5pt、行距不是固定值、页下注编号格式错——这些肉眼看不出来，
但都是数值问题，直接读 XML 就能断言。
**修法**：结构断言做主验收；看图只用于少数"视觉观感"问题，默认只出首页+末页。

### 段落分类只看正则会误判
**症状**：`## 引言` 这类不带"一、"的一级标题被判成正文。
**原因**：分类规则只用正则匹配"一、二、"体例。
**修法**：分类时**样式优先于文字**：先看 `paragraph.style.name` 是 Title/Heading 1/Heading 2，
再退回正则。底稿由 `md2docx.py` 生成时这一点尤其重要。

## 四、工程纪律类

### 改完不知道有没有改坏别的东西
**修法**：`snapshot.py` 逐部件比对 zip 内容。声明允许变化的部件前缀（如
`word/document.xml word/media/`），其余任何变化都判失败。这比"看起来没问题"可靠得多。

### 写盘中断留下半个 docx
**修法**：所有改写 docx 的脚本都先写 `xxx.tmp` 再 `tmp.replace(target)`。
中断时原文件完好，不会出现打不开的 docx。

### 字号换算（中文排版常用）
| 名称 | pt | 名称 | pt |
|---|---|---|---|
| 初号 | 42 | 小初 | 36 |
| 一号 | 26 | 小一 | 24 |
| 二号 | 22 | 小二 | 18 |
| 三号 | 16 | 小三 | 15 |
| 四号 | 14 | 小四 | 12 |
| 五号 | 10.5 | 小五 | 9 |
| 六号 | 7.5 | 小六 | 6.5 |

### python-docx 存盘会重写多个部件，"只改了一处"的断言必然失败
**症状**：只想改两个图题的先后顺序，改完做部件级差异断言，发现 `styles.xml`、
`settings.xml`、`[Content_Types].xml`、各种 `rels` 全都变了。
**原因**：`Document.save()` 会把整个 OPC 包重新序列化，即使内容等价，字节也不同。
**修法**：凡是"只允许改某几个部件"的操作（换图、挪题注、改表格边框），
不要用 python-docx 存盘，**直接对 zip 里的 `word/document.xml` 做 XML 手术**，
其余部件逐字节照抄。这样部件断言才成立，也才敢在成品稿上动手。
`embed_figure.py` 和按例文校正成品的脚本都是这么做的。
