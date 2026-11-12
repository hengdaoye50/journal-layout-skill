#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Markdown 全稿 → Word 底稿（语义化），一步到位，含真正的页下注。

为什么有这一步：
    期刊排版不要直接在成品 docx 上改。先把 Markdown 原稿转成一份"语义干净、版式中性"
    的底稿：题名/标题层级/正文/图表题注/参考文献/页下注各就各位，再用 journal_layout.py
    套具体期刊的数值。这样同一份底稿可以派生出四个刊的版本，且改版式不会碰到文字。

用法：
    python md2docx.py --md 全文.md --out 底稿.docx --figure-dir 图件 --figure-width-cm 14.8
    python md2docx.py --md 全文.md --out 初稿.docx --no-footnotes     # 先看结构
    python md2docx.py --md 全文.md --out 底稿.docx --emit-notes notes.json
    python md2docx.py --md 全文.md --out 底稿.docx --notes-mode endnote   # 文末「注释」节

两条注释通道（默认仍是页下注，行为不变）：
    --notes-mode footnote  页下注：正文 [n] → footnoteReference，注文进 word/footnotes.xml；
    --notes-mode endnote   文末注：正文 [n] 原样保留，注文作为文档末尾「注释」节追加，
                           不生成 footnotes.xml（《管理世界》的制式）。

支持的 Markdown 子集（够写中文社科论文）：
    # 题名            ## 一级标题      ### 二级标题
    正文段落（首行缩进 2 字符，两端对齐）
    **粗体** *斜体*
    | 表格 | 三线表 |        [^1] 行内注号      [^1]: 注释正文（--notes-mode 决定进页下注还是文末注释）
    **图1 图题**        题注下一行自动插入 figure-dir 中的 图1*.png
    ## 参考文献          其下条目自动悬挂缩进
    --- 分隔线（忽略）  ```代码块（忽略）
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
from docx.shared import Pt, Cm, RGBColor

sys.path.insert(0, str(Path(__file__).parent))
import footnotes  # noqa: E402  同目录的页下注模块

MD_INLINE = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|\[\^\d+\])")
RE_FN_DEF = re.compile(r"^\[\^(\d+)\]:\s*(.+)$")
RE_CAPTION = re.compile(r"^\*\*(图|表)\s*(\d+)")
CHAT_NOISE = re.compile(r"^##\s*[🧑🤖]")


def set_east(run, east: str) -> None:
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), east)
    rf.set(qn("w:ascii"), "Times New Roman")
    rf.set(qn("w:hAnsi"), "Times New Roman")


def add_rich(par, text: str) -> None:
    """写入带 **粗** *斜* [^n] 的行内格式。"""
    for tok in MD_INLINE.split(text):
        if not tok:
            continue
        if re.fullmatch(r"\[\^\d+\]", tok):
            par.add_run("[" + tok[2:-1] + "]")
        elif tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            par.add_run(tok[2:-2]).bold = True
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            par.add_run(tok[1:-1]).italic = True
        else:
            par.add_run(tok)


def three_line_table(table) -> None:
    """三线表：只留顶线、表头下线、底线。中文社科期刊的通行要求。"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (("top", 12), ("left", 0), ("bottom", 12), ("right", 0),
                     ("insideH", 0), ("insideV", 0)):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single" if sz else "none")
        el.set(qn("w:sz"), str(sz or 0))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    if len(table.rows):
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcb = OxmlElement("w:tcBorders")
            bot = OxmlElement("w:bottom")
            bot.set(qn("w:val"), "single")
            bot.set(qn("w:sz"), "6")
            bot.set(qn("w:space"), "0")
            bot.set(qn("w:color"), "000000")
            tcb.append(bot)
            tcPr.append(tcb)
        trPr = table.rows[0]._tr.get_or_add_trPr()
        if trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))


def build(md_path: Path, out_path: Path, figure_dir: Path | None,
          figure_width_cm: float, with_footnotes: bool,
          fig_map: dict | None = None, notes_mode: str = "footnote") -> dict:
    raw = md_path.read_text(encoding="utf-8")
    lines = raw.splitlines()

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for side in ("top_margin", "bottom_margin"):
        setattr(sec, side, Cm(2.5))
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2.7))

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.first_line_indent = Cm(0.85)
    normal.paragraph_format.space_after = Pt(0)
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), "宋体")

    for name, size in (("Title", 18), ("Heading 1", 14), ("Heading 2", 12)):
        st = doc.styles[name]
        st.font.name = "Times New Roman"
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.first_line_indent = Cm(0)
        st.paragraph_format.space_before = Pt(6)
        st.paragraph_format.space_after = Pt(6)
        r = st.element.get_or_add_rPr()
        f = r.find(qn("w:rFonts"))
        if f is None:
            f = OxmlElement("w:rFonts")
            r.insert(0, f)
        f.set(qn("w:eastAsia"), "黑体")
        # python-docx 内置 Title 样式带一条下边框线，中文期刊题名下不该有这条线
        ppr = st.element.get_or_add_pPr()
        border = ppr.find(qn("w:pBdr"))
        if border is not None:
            ppr.remove(border)

    notes: dict[str, str] = {}
    stats = {"paras": 0, "tables": 0, "figures": 0, "notes": 0, "skipped": 0}
    in_refs = False
    in_code = False
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        i += 1
        if line.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if not line or line == "---" or CHAT_NOISE.match(line):
            stats["skipped"] += 1
            continue

        m = RE_FN_DEF.match(line)
        if m:
            notes[m.group(1)] = m.group(2).strip()
            continue

        if line.startswith("|"):
            block = [line]
            while i < n and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1
            rows = []
            for raw_row in block:
                cells = [c.strip() for c in raw_row.strip("|").split("|")]
                if cells and all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                    continue          # Markdown 表头分隔行 |---|:--:|
                rows.append(cells)
            if rows:
                table = doc.add_table(rows=0, cols=len(rows[0]))
                for row in rows:
                    cells = table.add_row().cells
                    for c, v in zip(cells, row):
                        p = c.paragraphs[0]
                        p.paragraph_format.first_line_indent = Pt(0)
                        p.paragraph_format.line_spacing = 1.0
                        add_rich(p, v)
                        for run in p.runs:
                            run.font.size = Pt(9)
                            set_east(run, "宋体")
                three_line_table(table)
                stats["tables"] += 1
            continue

        if RE_CAPTION.match(line):
            kind = RE_CAPTION.match(line).group(1)     # 图 / 表
            num = RE_CAPTION.match(line).group(2)
            par = doc.add_paragraph()
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            par.paragraph_format.first_line_indent = Pt(0)
            add_rich(par, line.replace("**", ""))
            for run in par.runs:
                run.bold = False
                set_east(run, "黑体")
                run.font.size = Pt(10.5)
            if kind == "表":
                # 表题必须在表上方。Markdown 里常把表题写在表格之后，这里顺手挪上去。
                prev = par._p.getprevious()
                if prev is not None and prev.tag == qn("w:tbl"):
                    doc.element.body.remove(par._p)
                    prev.addprevious(par._p)
                stats["paras"] += 1
                continue
            # 图：题注后插入图片
            img = None
            if fig_map and str(num) in fig_map:
                img = Path(fig_map[str(num)])
            elif figure_dir:
                for pat in (f"图{num}*", f"fig{num}*", f"figure{num}*", f"FIG{num}*"):
                    cands = sorted(p for p in figure_dir.glob(pat)
                                   if p.suffix.lower() in (".png", ".jpg", ".jpeg"))
                    if cands:
                        img = cands[0]
                        break
            if img and img.exists():
                pic = doc.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pic.paragraph_format.first_line_indent = Pt(0)
                pic.add_run().add_picture(str(img), width=Cm(figure_width_cm))
                stats["figures"] += 1
            else:
                print(f"  ! 未找到图{num} 的图片（figure-dir={figure_dir}），仅保留题注",
                      file=sys.stderr)
            stats["paras"] += 1
            continue

        if line.startswith("# ") and not line.startswith("## "):
            par = doc.add_paragraph(style="Title")
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_rich(par, line[2:])
            stats["paras"] += 1
            continue

        if line.startswith("## "):
            title = line[3:].strip()
            in_refs = title == "参考文献"
            par = doc.add_paragraph(style="Heading 1")
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_rich(par, title)
            stats["paras"] += 1
            continue

        if line.startswith("### "):
            par = doc.add_paragraph(style="Heading 2")
            add_rich(par, line[4:].strip())
            stats["paras"] += 1
            continue

        par = doc.add_paragraph()
        if in_refs:
            par.paragraph_format.first_line_indent = Cm(-0.74)
            par.paragraph_format.left_indent = Cm(0.74)
            par.paragraph_format.line_spacing = 1.15
        add_rich(par, line)
        stats["paras"] += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    stats["notes"] = len(notes)

    if with_footnotes and notes:
        if notes_mode == "endnote":
            rep = footnotes.add_endnotes(out_path, notes,
                                         east="宋体", size=10, line=15)
            stats["notes_endnote"] = rep["added"]
        else:
            rep = footnotes.inject(out_path, notes,
                                   {"east": "宋体", "size": 10, "line": 15})
            stats["notes_injected"] = rep["injected"]
            if rep["missing"]:
                stats["notes_missing"] = rep["missing"]
    return stats


def main() -> int:
    ap = argparse.ArgumentParser(description="Markdown → 语义化 Word 底稿（含页下注/文末注释）")
    ap.add_argument("--md", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--figure-dir", help="图片目录，按 图1*.png / fig1*.png 匹配")
    ap.add_argument("--fig-map", help='显式映射 json：{"1": "路径/图1.png"}')
    ap.add_argument("--figure-width-cm", type=float, default=14.8)
    ap.add_argument("--no-footnotes", action="store_true",
                    help="不注入注释，只留 [n] 标记（优先于 --notes-mode）")
    ap.add_argument("--notes-mode", choices=["footnote", "endnote"], default="footnote",
                    help="footnote=页下注（默认，写 footnotes.xml）；"
                         "endnote=文末「注释」节（正文 [n] 原样保留，不写 footnotes.xml）")
    ap.add_argument("--emit-notes", help="把解析到的注释另存为 json，便于核对")
    a = ap.parse_args()
    fig_map = json.loads(Path(a.fig_map).read_text(encoding="utf-8")) if a.fig_map else None
    md = Path(a.md)
    stats = build(md, Path(a.out), Path(a.figure_dir) if a.figure_dir else None,
                  a.figure_width_cm, not a.no_footnotes, fig_map, a.notes_mode)
    print(f"底稿已生成：{a.out}")
    if a.no_footnotes:
        note_desc = "按 --no-footnotes 只留 [n] 标记"
    elif a.notes_mode == "endnote":
        note_desc = f"文末 {stats.get('notes_endnote', 0)} 条"
    else:
        note_desc = f"注入 {stats.get('notes_injected', 0)}"
    print(f"  段落 {stats['paras']}  表格 {stats['tables']}  插图 {stats['figures']}  "
          f"注释 {stats['notes']}（{note_desc}）")
    if stats.get("notes_missing"):
        print(f"  ! 注释定义缺失的编号：{stats['notes_missing']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
