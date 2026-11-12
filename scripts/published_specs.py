#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""刊出版式参数表（PUBLISHED_SPECS）—— 仿已刊 PDF 的真实样张，不是投稿稿。

与本脚本同目录 journal_layout.py 里的 SPECS 是**投稿版式**：
    给审稿人/编辑部看的稿子，大字号、单栏、行距松，目标是"好读、好批注"。
本文件的 PUBLISHED_SPECS 是**刊出版式**：
    按四刊已刊例文 PDF 实测出来的真实印刷参数，小字号、密排、可能双栏，
    目标是"排出来像已经印出来的那一页"。
两者都是合法目标，用 `--mode submission` / `--mode published` 切换，取决于用户要哪个；
本文件不替代 SPECS、也不改动 SPECS。

数据来源：references/journal-specs.md 的〔实测〕值（PyMuPDF 直接量四刊例文 PDF，
2026-11-12 执行；每刊节的〔实测〕行 + §1 速查表 + §7 差异清单）。取值纪律：
1) 只写该文件里带〔实测〕标记的数值，单位一律 pt；不取平均、不近似取整、不跨刊填空。
2) 该文件标〔未确认〕的字段、只有〔脚本〕或〔须知〕来源的数值、以及"只给了区间或
   两个互斥实测值"的字段，一律写 None —— 脚本对 None 的处理是**跳过并在报告里提示**，
   绝不猜一个默认值顶上去。
3) 字段名、嵌套结构、单位与 journal_layout.py 的 SPECS 完全一致。
   唯一实现层面的差异见 china-admin.column_space 的注释（该字段在 SPECS 里实际按 twips 写入）。
4) 字体字段写字体族名（宋体 / 楷体 / 黑体），对应实测 PDF 内嵌的方正系字体；
   只有"族名不足以表达版式特征"的例外才写成字体本身（见 china-admin.title）。
5) 每刊末尾的"最没把握"注释块，列的是建议 Lead 复核/人工拍板的字段，不是待办清单。

用法（不需要 import 本文件，journal_layout.py 会按 --mode 自动选表）：
    python journal_layout.py --list --mode published
    python journal_layout.py --journal politics --src 底稿.docx --out 样张.docx --mode published
"""
PUBLISHED_SPECS: dict[str, dict] = {
    # =====================================================================
    # 《政治学研究》：单栏，页面 210×285 mm（比 A4 矮 12 mm），正文 11.74 pt
    # 实测把握度：高（页面/边距/正文/行距/缩进/摘要/页下注编号三方吻合）
    # 最没把握的字段：
    #   1) title.bold —— 实测"未检出加粗"，但 PDF 字体的字形本身可能掩盖加粗（旧脚本值 True）；
    #   2) table_cell.line —— 实测是"横线间距 17–18 pt 一条"的区间，没有单值，按纪律留 None；
    #   3) margins.top —— 实测的是"正文首行 ink 顶 98.25"，Word 的 top_margin 是文本框顶、
    #      比 ink 顶高约 7 pt（journal-specs.md §2.1）；本文按 §7"按实测改"的结论写 98.25，
    #      若要求 Word 版面与 PDF 墨迹逐点重合，这里应减约 7 pt。
    # =====================================================================
    "politics": {
        "name": "政治学研究",
        "page": {"w": 595.276, "h": 807.874},          # 实测 210 × 285 mm，不是 A4
        "margins": {"top": 98.25, "bottom": 76, "left": 77.94, "right": 77.96},
        "body": {"east": "宋体", "size": 11.74, "line": 19.0, "first_indent": 23.9},
        # 实测未检出加粗；标题行距/段前段后未单独实测 → 留 None
        "title": {"east": "宋体", "size": 22.51, "bold": False, "align": "center",
                  "line": None, "before": None, "after": None},
        # 「内容提要:」+ 半角冒号（实测冒号占宽 5.9 pt，不是全角）；两种输入都覆盖
        "abstract": {"east": "楷体", "size": 11.74, "side_indent": 24,
                     "line": 19.0, "space_after_keywords": None,
                     "rename": {"摘要：": "内容提要:", "摘要:": "内容提要:"}},
        "h1": {"east": "黑体", "size": 14.75, "bold": None, "align": "center",
               "first_indent": None, "line": None, "before": 15, "after": 15},
        # 实测该刊没有「（一）」级：二级用「1. 2. 3.」随文起排，与正文同字体同字号、不另加粗
        "h2": {"east": "宋体", "size": 11.74, "bold": False, "align": "left",
               "first_indent": 23.9, "line": None, "before": None, "after": None},
        # 例文未出现三级标题 → 〔未确认〕
        "h3": {"east": None, "size": None, "bold": None, "first_indent": None,
               "line": None, "before": None, "after": None},
        # 只有表；表题实测黑体 10.29 pt 居中、在表上方（表题位置由 verify 单独校验）
        "caption": {"east": "黑体", "size": 10.29, "align": "center",
                    "first_indent": None, "before": None, "after": None},
        "table_cell": {"east": "宋体", "size": 8.83, "line": None},
        # 本刊最硬的一条：①②③ 每页重排，注文宋体 10.29 / 行距 19.0（三方一致）
        "footnote": {"num_fmt": "decimalEnclosedCircle", "num_restart": "eachPage",
                     "east": "宋体", "size": 10.29, "line": 19.0},
        # 例文无图，图题位置/图宽均〔未确认〕
        "figure": {"caption_position": None, "width": None},
        "columns": 1,
    },

    # =====================================================================
    # 《中国行政管理》：双栏，A4，正文 10.50 pt / 行距 18.8 / 缩进 21
    # 实测把握度：高（正文/行距/缩进/分栏/一级标题与须知+实测吻合）
    # 最没把握的字段：
    #   1) title.east「方正超粗宋」—— 实测字体，Windows 本机大概率没装，Word 会回退；
    #      要出稿观感需人工换成可用字体，属人工决策，不由本模块猜；
    #   2) abstract.side_indent —— 实测"左缩进 15.8 pt、右至版心右缘（=0）+ 悬挂"，
    #      而 SPECS 的 side_indent 是左右同值，表达不了，留 None 以免把右侧也缩进；
    #   3) margins 的镜像装订：内 56.68 / 外 51.00，SPECS 的 margins 只有四个标量、
    #      没有奇偶镜像位，这里按"奇数页"取（左=内、右=外）。
    # =====================================================================
    "china-admin": {
        "name": "中国行政管理",
        "page": {"w": 595.28, "h": 841.89},            # A4，实测
        "margins": {"top": 79.33, "bottom": 56.5, "left": 56.68, "right": 51.0},
        "body": {"east": "宋体", "size": 10.5, "line": 18.8, "first_indent": 21.0},
        # 题名实测方正超粗宋 26 pt；是否另加粗未测 → bold 留 None
        "title": {"east": "方正超粗宋", "size": 26.0, "bold": None, "align": "center",
                  "line": None, "before": None, "after": None},
        # 标签实测「［摘　要］」（全角方括号，括号内"摘""要"之间一个全角空格）
        "abstract": {"east": "楷体", "size": 10.5, "side_indent": None,
                     "line": 18.8, "space_after_keywords": None,
                     "rename": {"摘要：": "［摘　要］", "摘要:": "［摘　要］",
                                "关键词：": "［关键词］", "关键词:": "［关键词］"}},
        # 一/二/三级标题均与正文同字号 10.5，只是字体不同（黑体 / 楷体 / 宋体）
        "h1": {"east": "黑体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": None, "before": None, "after": None},
        "h2": {"east": "楷体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": None, "before": None, "after": None},
        "h3": {"east": "宋体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": None, "before": None, "after": None},
        # 图题实测黑体 8.00 pt 居中、在图下方；表题只有〔须知〕"在表上方"、字号〔未确认〕，
        # 本字段两类题注共用，只能按图题实测值填（底稿有表时属近似）。
        "caption": {"east": "黑体", "size": 8.0, "align": "center",
                    "first_indent": None, "before": None, "after": None},
        # 例文无表 → 表内字号/行距〔未确认〕（旧脚本的 9 pt / 13 pt 无实测支撑）
        "table_cell": {"east": None, "size": None, "line": None},
        # 例文无页下注，该刊从未落地过编号与字体参数 → 全 None（不猜、不改、校验跳过）
        "footnote": {"num_fmt": None, "num_restart": None,
                     "east": None, "size": None, "line": None},
        "figure": {"caption_position": "below", "width": None},
        "columns": 2,
        # 实测栏距 22.7 pt（8.0 mm）。注意：apply 里该字段被**直接写进 w:cols/@w:space**，
        # 该属性单位是 twips（1 pt = 20 twips），故这里填 22.7 × 20 = 454。
        # （journal-specs.md §3.2 写的"≈595 twips"与它自己给的 22.7 pt / 8.0 mm 不自洽：
        #   595 twips = 29.75 pt；此处以 22.7 pt 为准换算。）
        "column_space": 454,
        "columns_start_at": "一、",                    # 实测：正文从「一、」开始分栏
        "span_tables": None,                           # 例文无表 → 〔未确认〕
        "span_figures": True,                          # 实测：图 1 为通栏图
    },

    # =====================================================================
    # 《公共管理学报》：双栏，正文行距 15.0，字号三值不自洽（见下）
    # 实测把握度：中（页下注是全项目唯一真正落地验证过的部分）
    # 最没把握的字段：
    #   1) body.size —— 实测 span size 9.44 pt / 字面宽 10.18 pt / 字步进 10.56 pt 三值
    #      不自洽，文档明确标"〔未确认〕该刊名义字号"，故留 None；要出稿须人工定 10–10.5 pt
    #      并按"221.1 pt 栏宽排 21–22 字"验算；
    #   2) margins 的镜像装订（内 80.16 / 外 51.00），同 china-admin，按奇数页取；
    #   3) footnote：实测首页"说明性脚注"（黑体标签+宋体内容 7.80 / 行距 10.6）与页下
    #      引文注（楷体 7.80 / 行距 11.0）不同，本表只有一套字段，只能按引文注填。
    # =====================================================================
    "public-admin": {
        "name": "公共管理学报",
        "page": {"w": 595.0, "h": 842.0},              # 实测 209.9 × 297.0 mm
        "margins": {"top": 110.72, "bottom": 77.2, "left": 80.16, "right": 51.0},
        "body": {"east": "宋体", "size": None, "line": 15.0, "first_indent": 20.5},
        "title": {"east": "黑体", "size": 22.03, "bold": None, "align": "center",
                  "line": None, "before": None, "after": None},
        # 摘要标签实测就是「摘要：」（全角冒号），无需 rename
        "abstract": {"east": "楷体", "size": 8.64, "side_indent": None,
                     "line": None, "space_after_keywords": None, "rename": {}},
        # 页内标题用阿拉伯数字（「1 引言」「2.1」），字体均为方正黑体
        "h1": {"east": "黑体", "size": 11.02, "bold": None, "align": "left",
               "first_indent": 0, "line": None, "before": None, "after": None},
        "h2": {"east": "黑体", "size": 9.44, "bold": None, "align": "left",
               "first_indent": 0, "line": None, "before": None, "after": None},
        "h3": {"east": "黑体", "size": 9.44, "bold": None, "align": "left",
               "first_indent": 0, "line": None, "before": None, "after": None},
        # 图题/表题同为黑体 8.64；图题在图下、表题在表上（表题位置由 verify 校验）
        "caption": {"east": "黑体", "size": 8.64, "align": "center",
                    "first_indent": None, "before": None, "after": None},
        # 表内实测楷体 7.80 / 7.92 两个值并存，行距 14–18 pt 区间 → size/line 留 None
        "table_cell": {"east": "楷体", "size": None, "line": None},
        "footnote": {"num_fmt": "decimalEnclosedCircle", "num_restart": "eachPage",
                     "east": "楷体", "size": 7.8, "line": 11.0},
        "figure": {"caption_position": "below", "width": None},
        "columns": 2,
        "column_space": 446,        # 栏距实测 22.3 pt → w:cols/@w:space 用 twips：22.3×20
        # 分栏起点：实测只说"首页题名/作者/摘要/关键词通栏、正文双栏"，未给出具体起始标记 →
        # None。此时 _set_columns 会回落到默认标记「一、」（底稿用阿拉伯数字标题时找不到起点，
        # 会提示"未找到双栏起点"并保持单栏）。
        "columns_start_at": None,
        "span_tables": True,                           # 实测：表 1 为通栏三线表
        "span_figures": None,                          # 图仅"疑为通栏"，〔未确认〕
    },

    # =====================================================================
    # 《管理世界》：单栏（例文 PDF 实测；纸刊是否双栏〔未确认〕），正文 9.51 pt
    # 实测把握度：低—中（须知几乎不给版式，旧脚本值基本是《政治学研究》参数平移）
    # 最没把握的字段：
    #   1) footnote —— 实测是**文末「注释」节**（①②③连续编号、7.84 pt、行距 11.0），
    #      而脚本只有 Word 页下注通道，注的位置无法编码；注文字体又未实测，故字号/行距
    #      一并留 None（数值记在这里：7.84 pt / 行距 11.0），只保留编号制式。要真仿该刊，
    #      需另加"脚注转文末注释"的步骤；
    #   2) h2.first_indent —— 实测是"左缩进 14.5 pt"，而字段语义是首行缩进，不等价 → None；
    #   3) caption.align —— 实测"居中于图表本身，不是居中于版心"，Word 段落居中只能相对
    #      版心；只有图表通栏时两者才等价。
    # =====================================================================
    "management": {
        "name": "管理世界",
        "page": {"w": 612.283, "h": 824.882},          # 实测 216 × 291 mm
        "margins": {"top": 97.15, "bottom": 78.9, "left": 65.87, "right": 65.81},
        "body": {"east": "宋体", "size": 9.51, "line": 15.6, "first_indent": 19.4},
        # 题名实测黑体 31.36 pt、题名末带 `*` 注号（16.62 pt）指向首页基金注
        "title": {"east": "黑体", "size": 31.36, "bold": None, "align": "center",
                  "line": None, "before": None, "after": None},
        # 摘要正文实测楷体 9.51 / 行距 15.6；左缩进 4.6 pt 与首行悬挂混算，字段无法表达 → None
        "abstract": {"east": "楷体", "size": 9.51, "side_indent": None,
                     "line": 15.6, "space_after_keywords": None, "rename": {}},
        "h1": {"east": "黑体", "size": 13.72, "bold": None, "align": "center",
               "first_indent": None, "line": None, "before": None, "after": None},
        "h2": {"east": "黑体", "size": 9.51, "bold": None, "align": "left",
               "first_indent": None, "line": None, "before": None, "after": None},
        # 例文未出现「1.」级三级标题 → 〔未确认〕
        "h3": {"east": None, "size": None, "bold": None, "first_indent": None,
               "line": None, "before": None, "after": None},
        # 图题/表题实测同为楷体 7.84 pt；图题在图下、表题在表上
        "caption": {"east": "楷体", "size": 7.84, "align": "center",
                    "first_indent": None, "before": None, "after": None},
        "table_cell": {"east": "宋体", "size": 6.86, "line": 9.9},
        "footnote": {"num_fmt": "decimalEnclosedCircle", "num_restart": "continuous",
                     "east": None, "size": None, "line": None},
        "figure": {"caption_position": "below", "width": None},
        "columns": 1,
    },
}
