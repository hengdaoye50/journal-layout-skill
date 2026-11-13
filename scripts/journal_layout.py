#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按中文社科期刊版式规格重排一篇 docx。

用法：
    python journal_layout.py --journal politics  --src 底稿.docx --out 政治学研究.docx
    python journal_layout.py --list                     # 看四刊参数与差异
    python journal_layout.py --journal management --src a.docx --out b.docx --dry-run
    python journal_layout.py --journal politics --src 底稿.docx --out 样张.docx --mode published
        --mode submission（默认）= 投稿版式 SPECS，好读、字号大、单栏；
        --mode published        = 刊出版式 PUBLISHED_SPECS（见 published_specs.py），
                                  按例文 PDF 实测的密排/双栏样张。

设计原则：
1) 参数与算法分离。所有可调数值集中在 SPECS 里，改版式不碰代码。
2) 缺失即跳过。某刊某字段为 None 表示"该刊未确认，不做改动"，绝不猜一个默认值。
3) 只做版式，不改文字。除显式声明的 rename（如 摘要：→内容提要：）外，
   脚本不增删正文任何字符；跑完可配合 verify_layout.py 断言文本未变。
4) 幂等。同一份文件重复跑结果一致。

依赖：python-docx, lxml
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from published_specs import PUBLISHED_SPECS   # 刊出版式（例文 PDF 实测），与 SPECS 同结构

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
Q = lambda tag: "{%s}%s" % (W, tag)

# ---------------------------------------------------------------------------
# 版式规格表
# ---------------------------------------------------------------------------
# 单位：尺寸一律 pt（1cm = 28.3465pt）；行距用"固定值 pt"。
# 来源：各刊官方投稿须知 + 例会例文实测 + 本项目已交付成品反查，详见
#       references/journal-specs.md。改某个刊的参数前先读那份表。
SPECS: dict[str, dict] = {
    # 《政治学研究》
    "politics": {
        "name": "政治学研究",
        "page": {"w": 595.276, "h": 807.874},          # 210 × 285 mm（比 A4 短，不是 A4）
        "margins": {"top": 91, "bottom": 76, "left": 78, "right": 78},
        "body": {"east": "宋体", "size": 12, "line": 19, "first_indent": 24},
        "title": {"east": "宋体", "size": 23, "bold": True, "align": "center",
                  "line": 31, "before": 45, "after": 44},
        "abstract": {"east": "楷体", "size": 12, "side_indent": 24,
                     "line": None, "space_after_keywords": 28,
                     # 实测：该刊用"内容提要:"且冒号是半角（占宽 5.9pt，全角应为 11.7pt）
                     "rename": {"摘要：": "内容提要:", "摘要:": "内容提要:"}},
        "h1": {"east": "黑体", "size": 15, "bold": False, "align": "center",
               "first_indent": 0, "line": None, "before": 15, "after": 15},
        "h2": {"east": "黑体", "size": 12, "bold": None, "align": None,
               "first_indent": 24, "line": None, "before": 6, "after": 4},
        "h3": {"east": "宋体", "size": 12, "bold": True, "first_indent": 24,
               "line": None, "before": 3, "after": 3},
        "caption": {"east": "黑体", "size": 10.5, "align": "center",
                    "first_indent": 0, "before": 5, "after": 5},
        "table_cell": {"east": "宋体", "size": 10.5, "line": 15},
        "footnote": {"num_fmt": "decimalEnclosedCircle", "num_restart": "eachPage",
                     "east": "宋体", "size": 10, "line": 15},
        "figure": {"caption_position": None, "width": None},   # 只指图题；该刊例文全文无图，未确认
        "columns": 1,
    },
    # 《管理世界》
    "management": {
        "name": "管理世界",
        "page": {"w": 612.283, "h": 824.882},
        "margins": {"top": 66, "bottom": 58, "left": 54, "right": 54},
        "body": {"east": "宋体", "size": 12, "line": 17, "first_indent": 24},
        "title": {"east": "黑体", "size": 22, "bold": True, "align": "center",
                  "line": 31, "before": 25, "after": 25},
        "abstract": {"east": "宋体", "size": 12, "side_indent": 24,
                     "line": None, "space_after_keywords": 28, "rename": {}},
        "h1": {"east": "黑体", "size": 15, "bold": False, "align": "center",
               "first_indent": 0, "line": None, "before": 15, "after": 15},
        "h2": {"east": "黑体", "size": 12, "bold": None, "align": None,
               "first_indent": 24, "line": None, "before": 6, "after": 4},
        "h3": {"east": "宋体", "size": 12, "bold": True, "first_indent": 24,
               "line": None, "before": 3, "after": 3},
        "caption": {"east": "黑体", "size": 10.5, "align": "center",
                    "first_indent": 0, "before": 5, "after": 5},
        "table_cell": {"east": "宋体", "size": 10.5, "line": 15},
        "footnote": {"num_fmt": "decimalEnclosedCircle", "num_restart": "continuous",
                     "east": "宋体", "size": 10, "line": 15},
        "figure": {"caption_position": "below", "width": None},   # 例文实测：题注在图下方
        "columns": 1,
    },
    # 《中国行政管理》：双栏正文，题名/摘要通栏，表格与图通栏（1 栏区段）
    "china-admin": {
        "name": "中国行政管理",
        "page": {"w": 595.276, "h": 841.890},          # 21 × 29.7 cm
        "margins": {"top": 56.7, "bottom": 56.7, "left": 56.7, "right": 56.7},
        "body": {"east": "宋体", "size": 10.5, "line": 18.75, "first_indent": 21},
        "title": {"east": "黑体", "size": 18, "bold": True, "align": "center",
                  "line": 26, "before": 35, "after": 35},
        "abstract": {"east": "楷体", "size": 10.5, "side_indent": 0,
                     "line": 18.25, "space_after_keywords": None, "rename": {}},
        "h1": {"east": "黑体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": 18.75, "before": 0, "after": 0},
        "h2": {"east": "黑体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": 18.75, "before": 0, "after": 0},
        "h3": {"east": "黑体", "size": 10.5, "bold": None, "align": "left",
               "first_indent": 21, "line": 18.75, "before": 0, "after": 0},
        # 题注字号 10.5pt 继承自 Normal；字体未被该刊脚本改动（仍为宋体），故 east 留空
        "caption": {"size": 10.5, "align": "center",
                    "first_indent": 0, "before": 3, "after": 3},
        "table_cell": {"east": "宋体", "size": 9, "line": 13},
        # 页下注：该刊没有落地过编号格式与字体参数（原项目只做过公共管理学报的脚注），
        # 因此全部留空＝不猜、不改、该校验项跳过。
        "footnote": {"num_fmt": None, "num_restart": None,
                     "east": None, "size": None, "line": None},
        "figure": {"caption_position": "below", "width": None},   # 例文实测：题注在图下方
        "columns": 2,
        "column_space": 360,
        "columns_start_at": "一、",        # 从一级标题开始双栏；之前通栏
        "span_tables": True,
        "span_figures": True,
    },
    # 《公共管理学报》：页下注为楷体小字号；其余版式以 references/journal-specs.md 为准。
    # 注意：本刊的页面/正文参数在原始项目中未编码成脚本（只做了脚注调整），
    # 因此这里只填已确认的脚注参数，其余留 None = 不动。
    "public-admin": {
        "name": "公共管理学报",
        "page": {"w": None, "h": None},
        "margins": {"top": None, "bottom": None, "left": None, "right": None},
        "body": {"east": None, "size": None, "line": None, "first_indent": None},
        "title": {},
        "abstract": {"rename": {}},
        "h1": {}, "h2": {}, "h3": {},
        "caption": {},
        "table_cell": {},
        "footnote": {"num_fmt": None, "num_restart": None,
                     "east": "楷体", "size": 8, "line": 11},
        "figure": {"caption_position": "below", "width": None},   # 例文实测：题注在图下方
        "columns": None,
    },
}

ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}


def get_specs(mode: str = "submission") -> dict[str, dict]:
    """按 mode 选参数表：submission=投稿版式（SPECS），published=刊出版式（PUBLISHED_SPECS）。

    两张表键相同、字段结构相同，只是"给谁看"不同：前者要可读，后者要像已刊 PDF。
    """
    return PUBLISHED_SPECS if mode == "published" else SPECS


# 段落分类规则（顺序敏感：先判专有，再判标题层级）
RE_TITLE_FIRST = None  # 题名＝首个非空段落
RE_ABSTRACT = re.compile(r"^(摘要|内容提要)[：:]")
RE_KEYWORDS = re.compile(r"^关键词[：:]")
RE_H1 = re.compile(r"^[一二三四五六七八九十]+、")
RE_H2 = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")
RE_H3 = re.compile(r"^\d+[．.、](?!\d)")
RE_CAPTION = re.compile(r"^[图表]\s*\d")
RE_FIG_CAPTION = re.compile(r"^图\s*\d")      # 只匹配图题：表题的上下位置规则不同，不能一起挪
RE_REFERENCES = re.compile(r"^参考文献$")
RE_REF_ITEM = re.compile(r"^\[?\d+\]?\s")


def classify(text: str, index: int, style: str | None = None) -> str:
    """把一个段落的文字（＋既有样式）映射到版式类别。

    样式优先于文字：底稿由 md2docx 生成时，题名/一级/二级标题带 Title、Heading 1、
    Heading 2 样式，比正则更可靠（例如"引言"不带"一、"但仍然是一级标题）。
    返回 body/title/abstract/keywords/h1/h2/h3/caption/ref_head。
    """
    t = text.strip()
    if index == 0:
        return "title"
    if RE_REFERENCES.match(t):
        return "ref_head"
    s = (style or "").strip().lower()
    if s == "byline":
        return "byline"          # 署名/单位行，默认不改（四刊署名版式未实测）
    if s == "title":
        return "title"
    if s.startswith("heading 1"):
        return "h1"
    if s.startswith("heading 2"):
        return "h2"
    if s.startswith("heading 3"):
        return "h3"
    if len(t) < 70:
        if RE_H1.match(t):
            return "h1"
        if RE_H2.match(t):
            return "h2"
    if RE_ABSTRACT.match(t):
        return "abstract"
    if RE_KEYWORDS.match(t):
        return "keywords"
    if len(t) < 100 and RE_CAPTION.match(t):
        return "caption"
    if len(t) < 70 and RE_H3.match(t):
        return "h3"
    return "body"


# ---------------------------------------------------------------------------
# 基础设置
# ---------------------------------------------------------------------------
def _half_pt(size: float) -> float:
    """把字号落到 Word 的 0.5 pt 网格（w:sz 以半磅为单位）。

    docx 里字号只有 0.5 pt 的粒度，且 python-docx 是**截断**写入（31.36 → 31.0、
    7.84 → 7.5）。实测刊出版式里大量非整半磅字号（31.36 / 7.84 / 9.44 / 11.74…），
    截断会白丢最多 0.5 pt，取最近的半磅才是 Word 能给到的最接近值。
    对投稿版式无影响：SPECS 的字号本来就都是 0.5 pt 的整数倍。
    """
    return round(size * 2) / 2


def _set_east(run_or_style, east: str) -> None:
    """设置东亚字体（中文期刊排版的关键：eastAsia 必须单独指定，否则回落宋体/等线）。"""
    rpr = run_or_style._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), east)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")


def _set_page(section, spec: dict) -> list[str]:
    notes = []
    page, mg = spec.get("page") or {}, spec.get("margins") or {}
    from docx.shared import Pt
    if page.get("w") and page.get("h"):
        section.page_width, section.page_height = Pt(page["w"]), Pt(page["h"])
    for side, key in (("top", "top_margin"), ("bottom", "bottom_margin"),
                      ("left", "left_margin"), ("right", "right_margin")):
        val = mg.get(side)
        if val:
            setattr(section, key, Pt(val))
    if not (mg.get("left") or mg.get("right")):
        notes.append("页边距未确认，保持原值")
    return notes


def _set_footnote_numbering(section, spec: dict) -> list[str]:
    """页下注编号格式：圈码①②③、逐页重编/连续编号。写进 sectPr/w:footnotePr。"""
    fn = spec.get("footnote") or {}
    if not (fn.get("num_fmt") or fn.get("num_restart")):
        return ["页下注编号格式未确认，保持原值"]
    pr = section._sectPr.find(qn("w:footnotePr"))
    if pr is None:
        pr = OxmlElement("w:footnotePr")
        section._sectPr.insert(0, pr)
    for tag, val in (("numFmt", fn.get("num_fmt")), ("numRestart", fn.get("num_restart"))):
        if not val:
            continue
        el = pr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            pr.append(el)
        el.set(qn("w:val"), val)
    return []


def _apply_paragraph_format(par, spec: dict, kind: str) -> None:
    """把一个段落按 kind 套上 spec 中的字体/字号/行距/缩进/间距/对齐。"""
    from docx.shared import Pt
    pf = par.paragraph_format
    st = spec.get(kind) or {}
    body = spec.get("body") or {}

    # 行距：固定值。None 表示该刊未确认，沿用原值。
    line = st.get("line", body.get("line") if kind == "body" else None)
    if line:
        pf.line_spacing = Pt(line)
    pf.space_before = Pt(st.get("before", 0) or 0)
    pf.space_after = Pt(st.get("after", 0) or 0)
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)

    first = st.get("first_indent")
    if kind == "body" and first is None:
        first = body.get("first_indent")
    if first is not None:
        pf.first_line_indent = Pt(first)

    align = st.get("align")
    if kind == "body":
        par.alignment = ALIGN["justify"]
    elif align in ALIGN:
        par.alignment = ALIGN[align]

    if kind in ("abstract", "keywords"):
        side = st.get("side_indent")
        if side is not None:
            pf.left_indent = Pt(side)
            pf.right_indent = Pt(side)
            pf.first_line_indent = Pt(0)
        if kind == "keywords" and st.get("space_after_keywords") is not None:
            pf.space_after = Pt(st["space_after_keywords"])

    if kind in ("h1", "h2", "h3", "caption", "ref_head"):
        pf.keep_with_next = True

    # 字号/字体/粗体
    east = st.get("east") or (body.get("east") if kind == "body" else None)
    size = st.get("size") or (body.get("size") if kind == "body" else None)
    bold = st.get("bold")
    for run in par.runs:
        if east:
            _set_east(run, east)
        if size:
            run.font.size = Pt(_half_pt(size))
        if bold is not None:
            run.bold = bold


def _apply_rename(par, mapping: dict) -> None:
    """例：政治学研究要求"内容提要："而非"摘要："。只改标签，不动正文。"""
    if not mapping:
        return
    for run in par.runs:
        for old, new in mapping.items():
            if run.text.startswith(old):
                run.text = new + run.text[len(old):]
                return


def _format_tables(doc, spec: dict) -> int:
    from docx.shared import Pt
    cell = spec.get("table_cell") or {}
    if not cell.get("size") and not cell.get("east"):
        return 0
    n = 0
    for table in doc.tables:
        for row in table.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.first_line_indent = Pt(0)
                    if cell.get("line"):
                        p.paragraph_format.line_spacing = Pt(cell["line"])
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for run in p.runs:
                        if cell.get("east"):
                            _set_east(run, cell["east"])
                        if cell.get("size"):
                            run.font.size = Pt(_half_pt(cell["size"]))
        n += 1
    return n


def _relocate_figures(doc, spec: dict) -> list[str]:
    """按刊要求调整图与图题的先后，并保证图与题不跨页拆开。

    注意：四刊例文实测都是「题注在图上方」（与《社会》等期刊相反）。所以这里不是
    无条件模仿某个旧习惯，而是由 spec.figure.caption_position 决定，两个方向对称处理。
    keepNext 写在图片段落上（图要跟着它的题注走），而不是反过来。
    """
    notes = []
    want = (spec.get("figure") or {}).get("caption_position")
    if want not in ("below", "above"):
        return ["图题位置未确认，保持原值"]
    body = doc.element.body

    def drawing_of(node):
        return node is not None and node.tag == qn("w:p") and \
            node.find(".//" + qn("w:drawing")) is not None

    def keep_next(node):
        """给 node 段落加 keepNext（若没有）。"""
        ppr = node.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            node.insert(0, ppr)
        if ppr.find(qn("w:keepNext")) is None:
            etree.SubElement(ppr, qn("w:keepNext"))

    def drop_keep_next(node):
        ppr = node.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:keepNext")) is not None:
            ppr.remove(ppr.find(qn("w:keepNext")))

    for par in list(doc.paragraphs):
        t = par.text.strip()
        if not (len(t) < 100 and RE_FIG_CAPTION.match(t)):
            continue
        node = par._p
        prev, nxt = node.getprevious(), node.getnext()
        if want == "below":
            if drawing_of(prev):
                keep_next(prev)                      # 已在图下方，只需绑定不分页
            elif drawing_of(nxt):
                body.remove(nxt)                     # 图在题后 → 把图挪到题前
                node.addprevious(nxt)
                keep_next(nxt)
                notes.append(f"图题移到图下方：{t[:16]}")
        else:  # above
            if drawing_of(nxt):
                keep_next(node)
            elif drawing_of(prev):
                body.remove(node)                    # 题在图后 → 把题挪到图前
                prev.addprevious(node)
                keep_next(node)
                notes.append(f"图题移到图上方：{t[:16]}")
    return notes


def _set_columns(doc, spec: dict, journal: str = "") -> list[str]:
    """双栏版式：从指定一级标题起双栏，题名/摘要区通栏；表格/图各自成通栏区段。"""
    from copy import deepcopy
    cols = spec.get("columns")
    if not cols or cols == 1:
        return []
    body = doc.element.body
    base = deepcopy(body.sectPr)

    def make_sect(n: int):
        sec = deepcopy(base)
        typ = sec.find(qn("w:type"))
        if typ is None:
            typ = OxmlElement("w:type")
            sec.insert(0, typ)
        typ.set(qn("w:val"), "continuous")
        c = sec.find(qn("w:cols"))
        if c is None:
            c = OxmlElement("w:cols")
            sec.append(c)
        c.set(qn("w:num"), str(n))
        # 注意：w:cols/@w:space 的单位是 twips（1 pt = 20 twips），不是 pt。
        # 因此 spec 里的 column_space 一栏请直接写 twips（如 22.7pt → 454）。
        c.set(qn("w:space"), str(spec.get("column_space", 360)))
        return sec

    def attach_before(node, n: int) -> bool:
        """把 sectPr 挂到 node 之前那个段落的 pPr 上（不插入空段）。"""
        prev = node.getprevious()
        if prev is None or prev.tag != qn("w:p"):
            return False
        ppr = prev.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            prev.insert(0, ppr)
        if ppr.find(qn("w:sectPr")) is not None:
            return False
        ppr.append(make_sect(n))
        return True

    def boundary(node, n: int, after=False):
        # 优先"挂到相邻段落"，避免插入空段（LibreOffice 对空段边界会卡死）
        if not after and attach_before(node, n):
            return
        if after:
            nxt = node.getnext()
            if nxt is not None and nxt.tag == qn("w:p"):
                ppr = nxt.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr")
                    nxt.insert(0, ppr)
                if ppr.find(qn("w:sectPr")) is None:
                    ppr.append(make_sect(n))
                    return
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ppr.append(make_sect(n))
        p.append(ppr)
        (node.addnext if after else node.addprevious)(p)

    marker = spec.get("columns_start_at")
    start = None
    for node in body:
        if node.tag == qn("w:p") and "".join(node.itertext()).strip().startswith(marker or "一、"):
            start = node
            break
    if start is None:
        return [f"未找到双栏起点“{marker}”，跳过分栏"]
    boundary(start, 1)                       # 起点之前＝通栏
    if spec.get("columns_new_page", COLUMNS_NEW_PAGE.get(journal, False)):
        ppr = start.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            start.insert(0, ppr)
        if ppr.find(qn("w:pageBreakBefore")) is None:
            etree.SubElement(ppr, qn("w:pageBreakBefore"))
    body.replace(body.sectPr, make_sect(cols))  # 其后正文＝双栏
    if spec.get("span_tables"):
        for node in list(body):
            if node.tag == qn("w:tbl"):
                prev = node.getprevious()
                txt = "".join(prev.itertext()).strip() if prev is not None else ""
                head = prev if txt.startswith("表") else node
                boundary(head, cols)
                boundary(node, 1, after=True)
    if spec.get("span_figures"):
        for node in list(body):
            if node.tag == qn("w:p") and node.find(".//" + qn("w:drawing")) is not None:
                nxt = node.getnext()
                tail = node
                if nxt is not None and "".join(nxt.itertext()).strip().startswith("图"):
                    tail = nxt
                boundary(node, cols)
                boundary(tail, 1, after=True)

    # 收尾边界：把多栏属性挂到最后一个尚未带 sectPr 的段落上，body 级 sectPr 回到 1 栏。
    # 否则"文档最后一节是多栏"时，LibreOffice 在起始页只铺第一栏（实测踩到）。
    tail_p = None
    for node in reversed(list(body)):
        if node.tag != qn("w:p"):
            continue
        ppr = node.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            continue
        tail_p = node
        break
    if tail_p is not None:
        ppr = tail_p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            tail_p.insert(0, ppr)
        ppr.append(make_sect(cols))
        body.replace(body.sectPr, make_sect(1))
    return [f"已分栏：正文 {cols} 栏，表格/图通栏（已加收尾边界）"]


# ---------------------------------------------------------------------------
# 页下注样式（python-docx 不暴露 footnotes.xml，必须过一层 lxml）
# ---------------------------------------------------------------------------
def style_footnotes(docx_path: Path, spec: dict) -> int:
    from zipfile import ZipFile, ZIP_DEFLATED
    fn = spec.get("footnote") or {}
    if not (fn.get("east") and fn.get("size")):
        return 0
    with ZipFile(docx_path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    if "word/footnotes.xml" not in files:
        return 0
    root = etree.fromstring(files["word/footnotes.xml"])
    count = 0
    for note in root.findall(Q("footnote")):
        if note.get(Q("type")):
            continue
        count += 1
        for p in note.iter(Q("p")):
            ppr = p.find(Q("pPr"))
            if ppr is None:
                ppr = etree.Element(Q("pPr"))
                p.insert(0, ppr)
            sp = ppr.find(Q("spacing"))
            if sp is None:
                sp = etree.SubElement(ppr, Q("spacing"))
            sp.set(Q("before"), "0")
            sp.set(Q("after"), "0")
            if fn.get("line"):
                sp.set(Q("line"), str(int(round(fn["line"] * 20))))  # pt → 1/20 pt
                sp.set(Q("lineRule"), "exact")
            for run in p.iter(Q("r")):
                rpr = run.find(Q("rPr"))
                if rpr is None:
                    rpr = etree.Element(Q("rPr"))
                    run.insert(0, rpr)
                rf = rpr.find(Q("rFonts"))
                if rf is None:
                    rf = etree.SubElement(rpr, Q("rFonts"))
                for k, v in (("eastAsia", fn["east"]), ("ascii", "Times New Roman"),
                             ("hAnsi", "Times New Roman"), ("cs", "Times New Roman")):
                    rf.set(Q(k), v)
                sz = rpr.find(Q("sz"))
                if sz is None:
                    sz = etree.SubElement(rpr, Q("sz"))
                sz.set(Q("val"), str(int(round(fn["size"] * 2))))       # pt → 半磅
                szcs = rpr.find(Q("szCs"))
                if szcs is None:
                    szcs = etree.SubElement(rpr, Q("szCs"))
                szcs.set(Q("val"), str(int(round(fn["size"] * 2))))
                for t, v in (("b", "0"), ("i", "0")):
                    e = rpr.find(Q(t))
                    if e is None:
                        e = etree.SubElement(rpr, Q(t))
                    e.set(Q("val"), v)
                count += 0
    files["word/footnotes.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = docx_path.with_suffix(".layout.tmp")
    with ZipFile(tmp, "w", ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    tmp.replace(docx_path)
    return count


# 表格制式：**四刊各不相同，不能一刀切**。这是例文实测结论（PyMuPDF 读 PDF 矢量
# 绘图对象数竖线），与投稿/刊出模式无关，属于刊物固有习惯：
#   《政治学研究》带左右外框的完整网格表（4 条竖线）
#   《管理世界》横线逐行贯通 + 列间竖线，但无左右外框
#   《公共管理学报》标准三线表（0 竖线，顶线/表头线/底线粗细分明）
#   《中国行政管理》该篇例文全文无表 → 未确认，保持原样
# 双栏区段是否另起新页。None/缺省＝从页中开始（更贴近例文），
# True＝另起新页（LibreOffice 在"页中分栏 + 大量页下注"时会死循环，需此开关兜底）。
COLUMNS_NEW_PAGE = {"china-admin": True}

TABLE_STYLE = {
    "politics": "grid-full",
    "management": "grid",
    "china-admin": None,
    "public-admin": "three-line",
}

BORDER_PLANS = {
    # 边框名: 线宽（1/8 pt）；0 = 不要
    "three-line": {"top": 12, "bottom": 12, "left": 0, "right": 0,
                   "insideH": 0, "insideV": 0},
    "grid": {"top": 12, "bottom": 12, "left": 0, "right": 0,
             "insideH": 4, "insideV": 4},
    "grid-full": {"top": 12, "bottom": 12, "left": 4, "right": 4,
                  "insideH": 4, "insideV": 4},
}


def apply_table_style(doc, style: str | None) -> int:
    """把文档里的表格改成该刊的表格制式。style 为 None 时不动。"""
    if not style or style not in BORDER_PLANS:
        return 0
    plan = BORDER_PLANS[style]
    n = 0
    for table in doc.tables:
        tblPr = table._tbl.tblPr
        for old in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(old)
        borders = OxmlElement("w:tblBorders")
        for edge, sz in plan.items():
            el = OxmlElement("w:" + edge)
            el.set(qn("w:val"), "single" if sz else "none")
            el.set(qn("w:sz"), str(sz or 0))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
            borders.append(el)
        tblPr.append(borders)
        # 三线表要单独给表头加一条下线（网格表的行线由 insideH 负责）
        if style == "three-line" and len(table.rows):
            for cell in table.rows[0].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                for old in tcPr.findall(qn("w:tcBorders")):
                    tcPr.remove(old)
                tcb = OxmlElement("w:tcBorders")
                bot = OxmlElement("w:bottom")
                bot.set(qn("w:val"), "single")
                bot.set(qn("w:sz"), "6")
                bot.set(qn("w:space"), "0")
                bot.set(qn("w:color"), "000000")
                tcb.append(bot)
                tcPr.append(tcb)
        n += 1
    return n


# ---------------------------------------------------------------------------
# 文末注释（endnote）节
# ---------------------------------------------------------------------------
RE_NOTE_HEAD = re.compile(r"^注\s*释$")
RE_NOTE_ITEM = re.compile(r"^(?:\d+[.．、]|[①②③④⑤⑥⑦⑧⑨⑩])")
RE_END_AFTER = re.compile(r"^(参考文献|附录|[一二三四五六七八九十]+、)")


def find_notes_block(doc):
    """定位文末「注释」节：返回 (标题段, [条目段], [节点])；没有则 (None, [], [])。"""
    paras = doc.paragraphs
    head = None
    for i, par in enumerate(paras):
        if RE_NOTE_HEAD.match(par.text.strip()) and par.style.name.startswith("Heading"):
            head = i
            break
    if head is None:
        return None, [], []
    items = []
    for par in paras[head + 1:]:
        t = par.text.strip()
        if not t:
            continue
        if RE_END_AFTER.match(t) or par.style.name.startswith("Heading"):
            break
        if RE_NOTE_ITEM.match(t):
            items.append(par)
        else:
            break
    nodes = [paras[head]._p] + [p._p for p in items]
    return paras[head], items, nodes


# 注释制式：与表格制式同理，属于**刊物固有习惯**，与投稿/刊出模式无关。
# 例文实测：《管理世界》是文末「注释」节；其余三刊是页下注（中国行政管理是上标［n］
# 接文后参考文献，本工具按页下注通道处理）。
NOTES_STYLE = {
    "politics": "footnote",
    "management": "endnote",
    "china-admin": "footnote",
    "public-admin": "footnote",
}


def notes_mode(spec: dict, journal: str) -> str:
    """该刊的注释制式：spec.notes.mode 优先，否则查 NOTES_STYLE。"""
    return (spec.get("notes") or {}).get("mode") or NOTES_STYLE.get(journal) or "footnote"


def style_endnotes(doc, spec: dict, journal: str) -> list[str]:
    """按该刊要求给「注释」节上样式，并把该节移到「参考文献」之前。

    为什么需要：中文社科期刊的注释节在参考文献**之前**，而 md2docx 的 endnote
    通道只能尾追加。这里一次性把顺序和字号都摆正。
    """
    ns = spec.get("notes") or {}
    if notes_mode(spec, journal) != "endnote":
        return []
    head, items, nodes = find_notes_block(doc)
    if head is None:
        return ["未找到「注释」节（该刊要求文末注释，底稿请用 md2docx --notes-mode endnote）"]
    from docx.shared import Pt
    notes = []
    body = doc.element.body
    # 1) 顺序：注释节移到「参考文献」之前
    ref = None
    for par in doc.paragraphs:
        if par.text.strip() == "参考文献":
            ref = par
            break
    if ref is not None and ref._p.getprevious() not in nodes:
        first, last = nodes[0], nodes[-1]
        if body.index(first) > body.index(ref._p):
            anchor = ref._p
            for node in nodes:
                body.remove(node)
                anchor.addprevious(node)
            notes.append("注释节移到参考文献之前")
    # 2) 条目样式：字号/行距/字体/缩进
    for par in items:
        pf = par.paragraph_format
        if ns.get("line"):
            pf.line_spacing = Pt(ns["line"])
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        for run in par.runs:
            if ns.get("east"):
                _set_east(run, ns["east"])
            if ns.get("size"):
                run.font.size = Pt(_half_pt(ns["size"]))
    notes.append(f"注释节 {len(items)} 条已按该刊样式（{ns.get('size')}pt / 行距 {ns.get('line')}pt）")
    return notes


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def apply_spec(src: Path, dst: Path, journal: str, dry_run: bool = False,
               mode: str = "submission") -> dict:
    spec = get_specs(mode)[journal]
    doc = Document(str(src))
    report = {"journal": journal, "name": spec["name"], "mode": mode,
              "notes": [], "kinds": {}}

    for section in doc.sections:
        report["notes"] += _set_page(section, spec)
        report["notes"] += _set_footnote_numbering(section, spec)

    # Normal 样式兜底：一句话控制全文默认字体
    body = spec.get("body") or {}
    if body.get("east") and body.get("size"):
        st = doc.styles["Normal"]
        st.font.size = __import__("docx").shared.Pt(_half_pt(body["size"]))
        _set_east(st, body["east"])

    # 文末注释节先归位，再参与下面的段落分类（否则 "1. …" 会被当成三级标题）
    report["notes"] += style_endnotes(doc, spec, journal)
    note_head, note_items, _ = find_notes_block(doc)
    note_nodes = {id(p._p) for p in ([note_head] if note_head is not None else []) + note_items}

    nonempty = 0
    for idx, par in enumerate(doc.paragraphs):
        text = par.text.strip()
        if not text:
            continue
        if id(par._p) in note_nodes:
            kind = "note_head" if (note_head is not None and par._p is note_head._p) else "note_item"
        else:
            kind = classify(text, nonempty, par.style.name)
        nonempty += 1
        report["kinds"][kind] = report["kinds"].get(kind, 0) + 1
        if kind in ("abstract", "keywords"):
            _apply_rename(par, (spec.get("abstract") or {}).get("rename") or {})
        if kind in ("note_head", "note_item"):
            continue                      # 已在 style_endnotes 里处理，不套正文格式
        if kind == "byline" and not spec.get("byline"):
            continue                      # 该刊署名版式未确认 → 保持中立的居中楷体
        _apply_paragraph_format(par, spec, kind)

    report["tables"] = _format_tables(doc, spec)
    tstyle = spec.get("table_style") or TABLE_STYLE.get(journal)
    if tstyle:
        report["table_style"] = f"{tstyle}（{apply_table_style(doc, tstyle)} 个表）"
    report["notes"] += _relocate_figures(doc, spec)
    report["notes"] += _set_columns(doc, spec, journal)

    if dry_run:
        report["dry_run"] = True
        return report

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    report["footnotes"] = style_footnotes(dst, spec)
    return report


def print_specs(mode: str = "submission") -> None:
    table = get_specs(mode)
    label = "刊出版式 PUBLISHED_SPECS（仿已刊 PDF 样张）" if mode == "published" \
        else "投稿版式 SPECS（给审稿人看的稿子）"
    print(f"【{label}】")
    for key, spec in table.items():
        print(f"\n== {spec['name']}  [{key}] ==")
        pg, mg = spec.get("page") or {}, spec.get("margins") or {}
        print(f"  页面  {pg.get('w')} × {pg.get('h')} pt")
        print(f"  边距  上{mg.get('top')} 下{mg.get('bottom')} 左{mg.get('left')} 右{mg.get('right')} pt")
        b = spec.get("body") or {}
        print(f"  正文  {b.get('east')} {b.get('size')}pt / 固定行距 {b.get('line')}pt / 首行缩进 {b.get('first_indent')}pt")
        print(f"  分栏  {spec.get('columns') or 1}")
        fn = spec.get("footnote") or {}
        print(f"  页下注 {fn.get('east')} {fn.get('size')}pt / {fn.get('num_fmt')} / {fn.get('num_restart')}")


def main() -> int:
    ap = argparse.ArgumentParser(description="按中文社科期刊版式重排 docx")
    ap.add_argument("--journal", help="版式键：politics / management / china-admin / public-admin")
    ap.add_argument("--src")
    ap.add_argument("--out")
    ap.add_argument("--dry-run", action="store_true", help="只报告将改什么，不写文件")
    ap.add_argument("--list", action="store_true", help="打印四刊参数")
    ap.add_argument("--json", action="store_true", help="以 json 输出报告")
    ap.add_argument("--mode", choices=["submission", "published"], default="submission",
                    help="submission=投稿版式 SPECS（默认）；published=刊出版式 PUBLISHED_SPECS")
    a = ap.parse_args()
    table = get_specs(a.mode)
    if a.list:
        print_specs(a.mode)
        return 0
    if not (a.journal and a.src and a.out):
        ap.print_help()
        return 2
    if a.journal not in table:
        print(f"未知期刊键：{a.journal}，可选：{', '.join(table)}", file=sys.stderr)
        return 2
    report = apply_spec(Path(a.src), Path(a.out), a.journal, a.dry_run, a.mode)
    if a.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"[{report['name']}] {'DRY-RUN ' if report.get('dry_run') else ''}段落分类：{report['kinds']}")
        print(f"  表格 {report.get('tables', 0)} 个；页下注格式化 {report.get('footnotes', 0)} 条")
        for n in report["notes"]:
            print("  · " + n)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
