#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""版式结构校验：直接读 docx 的 XML 断言版式参数，不渲染页面。

为什么不用"逐页转 PNG 人眼看"：
    逐页出图既慢又只能靠肉眼判断，改一处版式就要重看三十页，还容易漏。
    版式是数值问题，就该用数值断言：页面尺寸、边距、每类段落的字体/字号/行距/
    缩进、页下注编号格式、图宽是否超出通栏、题注在图上还是图下、三线表有没有内框线。
    真正需要肉眼看的，只剩"图与表格的视觉观感"，那一步用 render_preview.py 抽 2-3 页即可。

用法：
    python verify_layout.py --docx 成品.docx --journal politics
    python verify_layout.py --docx 成品.docx --journal politics --ref 底稿.docx
    python verify_layout.py --docx 成品.docx --journal politics --json --strict

退出码：0 = 全部通过；1 = 有不符合项。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Length, Pt
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))
from journal_layout import (SPECS, TABLE_STYLE, classify, find_notes_block, get_specs,  # noqa: E402
                            notes_mode)
# 与排版脚本共用同一套分类规则、参数表与表格制式

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
Q = lambda t: "{%s}%s" % (W, t)
TOL = {"size": 0.3, "line": 0.6, "indent": 1.5, "margin": 1.0, "page": 1.0, "width": 2.0}


def _len_pt(v):
    """把 python-docx 的长度/倍数统一成 pt 数值或 ('multiple', 倍)。"""
    if v is None:
        return None
    if isinstance(v, Length):
        return round(v.pt, 2)
    return ("multiple", round(float(v), 2))


def _run_attr(run):
    """取 run 本身显式设置的 eastAsia 字体与字号。"""
    rpr = run._element.find(qn("w:rPr"))
    east = size = None
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            east = rf.get(qn("w:eastAsia"))
        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            try:
                size = round(int(sz.get(qn("w:val"))) / 2, 2)
            except (TypeError, ValueError):
                pass
    return east, size


def _inherit_chain(style):
    """沿样式的 base_style 链逐级向上（当前 → 上级 → …）。"""
    seen = 0
    while style is not None and seen < 12:
        yield style
        style = style.base_style
        seen += 1


def _style_font(style):
    """样式链里第一个显式声明的中文字体与字号。"""
    east = size = None
    for st in _inherit_chain(style):
        el = st.element
        rpr = el.find(qn("w:rPr"))
        if rpr is None:
            continue
        if east is None:
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None and rf.get(qn("w:eastAsia")):
                east = rf.get(qn("w:eastAsia"))
        if size is None:
            sz = rpr.find(qn("w:sz"))
            if sz is not None:
                try:
                    size = round(int(sz.get(qn("w:val"))) / 2, 2)
                except (TypeError, ValueError):
                    pass
        if east and size:
            break
    return east, size


def _style_pf(style):
    """样式链里的行距与首行缩进（pt / 倍数）；只在段内没有直接设置时才有意义。"""
    line = indent = None
    for st in _inherit_chain(style):
        pf = st.paragraph_format
        if line is None:
            line = _len_pt(pf.line_spacing)
        if indent is None and pf.first_line_indent is not None:
            indent = round(pf.first_line_indent.pt, 2)
        if line is not None and indent is not None:
            break
    return line, indent


def _effective(par):
    """段落的"有效格式"：run 直接设置 → 段落样式链（继承 Word 的样式优先级）。

    Word 里大量格式并不写进 run，而是落在样式上。只看 run/rPr 会把"继承了正确格式"
    的段落误判为未设置，所以这里必须回落到样式链。
    """
    east = size = None
    for run in par.runs:
        e, s = _run_attr(run)
        east, size = e or east, s or size
        if east and size:
            break
    if not east or not size:
        se, ss = _style_font(par.style)
        east, size = east or se, size or ss
    pf = par.paragraph_format
    line = _len_pt(pf.line_spacing)
    indent = round(pf.first_line_indent.pt, 2) if pf.first_line_indent is not None else None
    if line is None or indent is None:
        sl, si = _style_pf(par.style)
        line, indent = line if line is not None else sl, \
            indent if indent is not None else si
    return east, size, line, indent


def _modal(values):
    clean = [v for v in values if v is not None]
    if not clean:
        return None, 0
    best = max(set(clean), key=clean.count)
    return best, clean.count(best)


def _footnotes(path: Path):
    with ZipFile(path) as z:
        if "word/footnotes.xml" not in z.namelist():
            return None
        return etree.fromstring(z.read("word/footnotes.xml"))


def _all_text(doc) -> str:
    parts = []
    for t in doc.element.iter(qn("w:t")):
        parts.append(t.text or "")
    return "".join(parts)


def _text_excluding_notes(doc) -> str:
    """取正文文本但跳过文末「注释」节。

    为什么：endnote 制式的刊，注释节会被版式脚本从文末挪到参考文献之前；
    这是版式要求的顺序调整，不该被判成"正文被改动"。
    """
    body = doc.element.body
    nodes = list(body)
    start = None
    for i, node in enumerate(nodes):
        if node.tag != qn("w:p"):
            continue
        txt = "".join(t.text or "" for t in node.iter(qn("w:t"))).strip()
        if txt in ("注释", "注　释"):
            start = i
            break
    if start is None:
        return _all_text(doc)
    end = start + 1
    while end < len(nodes):
        node = nodes[end]
        txt = "".join(t.text or "" for t in node.iter(qn("w:t"))).strip()
        if node.tag == qn("w:tbl"):
            break
        if txt and not re.match(r"^(?:\d+[.．、]|[①②③④⑤⑥⑦⑧⑨⑩])", txt):
            break
        end += 1
    parts = []
    for i, node in enumerate(nodes):
        if start <= i < end:
            continue
        parts.append("".join(t.text or "" for t in node.iter(qn("w:t"))))
    return "".join(parts)


def _norm(text: str) -> str:
    text = re.sub(r"\[\d+\]", "", text)          # 去掉 [1] 注号
    return re.sub(r"\s+", "", text)


def verify(docx_path: Path, journal: str, ref_path: Path | None = None,
           strict_tables: bool = False, mode: str = "submission") -> dict:
    spec = get_specs(mode)[journal]
    doc = Document(str(docx_path))
    res = {"journal": journal, "name": spec["name"], "mode": mode,
           "checks": [], "failures": []}

    def add(item, expect, actual, ok, note="", warn_only=False):
        """warn_only=True 表示该项是"提示"而非"不符合"——期刊未确证的要求不该判死。"""
        status = "OK" if ok else ("提示" if warn_only else "错")
        res["checks"].append({"项": item, "应为": expect, "实测": actual,
                              "结果": status, "备注": note})
        if not ok and not warn_only:
            res["failures"].append(f"{item}：应为 {expect}，实测 {actual} {note}".strip())
        elif not ok:
            res.setdefault("warnings", []).append(f"{item}：{actual} {note}".strip())

    # 1 页面与边距（多节文档按"全部节一致"聚合成一项，避免双栏稿刷出几十行）
    pg, mg = spec.get("page") or {}, spec.get("margins") or {}
    secs = list(doc.sections)
    if pg.get("w"):
        dims = {(round(s.page_width.pt, 1), round(s.page_height.pt, 1)) for s in secs}
        ok = len(dims) == 1 and all(abs(w - pg["w"]) <= TOL["page"] and
                                    abs(h - pg["h"]) <= TOL["page"] for w, h in dims)
        add(f"页面（{len(secs)} 节）", f"{pg['w']}×{pg['h']}pt",
            "; ".join(f"{w}×{h}pt" for w, h in sorted(dims)), ok,
            "" if len(dims) == 1 else "（各节页面不一致）")
    for side, attr in (("top", "top_margin"), ("bottom", "bottom_margin"),
                       ("left", "left_margin"), ("right", "right_margin")):
        want = mg.get(side)
        if not want:
            continue
        vals = {round(getattr(s, attr).pt, 1) for s in secs}
        ok = len(vals) == 1 and abs(next(iter(vals)) - want) <= TOL["margin"]
        add(f"边距-{side}", f"{want}pt",
            ", ".join(f"{v}pt" for v in sorted(vals)), ok,
            "" if len(vals) == 1 else "（各节边距不一致）")

    # 1b 分栏（中国行政管理是双栏，其余单栏）
    want_cols = spec.get("columns")
    if want_cols:
        nums = []
        for sec in doc.sections:
            c = sec._sectPr.find(qn("w:cols"))
            n = 1
            if c is not None and c.get(qn("w:num")):
                try:
                    n = int(c.get(qn("w:num")))
                except ValueError:
                    n = 1
            nums.append(n)
        ok = (want_cols in nums) if want_cols > 1 else all(n == 1 for n in nums)
        add("分栏", f"{want_cols} 栏", "各节栏数 " + ",".join(map(str, nums)), ok)

    # 2 段落分类与逐类格式
    # 文末注释节单独校验（见第 6 节），不参与正文各类别的统计，否则 "1. …" 会被算成三级标题
    note_elems = set()
    if notes_mode(spec, journal) == "endnote":
        _h, _items, _nodes = find_notes_block(doc)
        note_elems = {id(n) for n in _nodes}
    buckets: dict[str, list] = {}
    nonempty = 0
    empty = 0
    for par in doc.paragraphs:
        text = par.text.strip()
        if not text:
            empty += 1
            continue
        if id(par._p) in note_elems:
            continue
        kind = classify(text, nonempty, par.style.name)
        nonempty += 1
        east, size, line, indent = _effective(par)
        buckets.setdefault(kind, []).append({
            "east": east, "size": size, "line": line, "indent": indent,
            "align": str(par.alignment),
        })

    for kind, spec_key in (("title", "title"), ("body", "body"), ("h1", "h1"),
                           ("h2", "h2"), ("h3", "h3"), ("caption", "caption")):
        rows = buckets.get(kind) or []
        st = spec.get(spec_key) or {}
        if not rows:
            continue
        east, n_e = _modal([r["east"] for r in rows])
        size, n_s = _modal([r["size"] for r in rows])
        line, n_l = _modal([r["line"] for r in rows])
        ind, n_i = _modal([r["indent"] for r in rows])
        want_east = st.get("east") or (spec["body"].get("east") if kind == "body" else None)
        want_size = st.get("size") or (spec["body"].get("size") if kind == "body" else None)
        want_line = st.get("line") or (spec["body"].get("line") if kind == "body" else None)
        want_ind = st.get("first_indent")
        if kind == "body" and want_ind is None:
            want_ind = spec["body"].get("first_indent")

        def cmp(name, want, got, tol, unit="pt", count=None, total=None):
            if want is None or got is None:
                return
            if isinstance(got, tuple):
                ok = got == want
                gv = f"{got[1]}倍"
                wv = f"{want}倍" if not isinstance(want, tuple) else f"{want[1]}倍"
            else:
                ok = abs(float(got) - float(want)) <= tol
                gv, wv = f"{got}{unit}", f"{want}{unit}"
            extra = f"（{count}/{total} 条一致）" if count is not None else ""
            add(f"{spec_key}-{name}", wv, gv + extra, ok)
            if not ok:
                res["failures"][-1] += f"，多数段落 {count}/{total} 条符合" if count else ""

        if want_east:
            add(f"{spec_key}-中文字体", want_east, str(east),
                east == want_east, f"（{n_e}/{len(rows)} 条一致）")
        cmp("字号", want_size, size, TOL["size"], count=n_s, total=len(rows))
        if want_line is not None:
            cmp("行距", want_line, line, TOL["line"], count=n_l, total=len(rows))
        if want_ind is not None:
            cmp("首行缩进", want_ind, ind, TOL["indent"], count=n_i, total=len(rows))

    # 3 页下注
    fn = spec.get("footnote") or {}
    if fn.get("num_fmt"):
        pr = doc.sections[0]._sectPr.find(qn("w:footnotePr"))
        got_fmt = got_rs = None
        if pr is not None:
            for tag in ("numFmt", "numRestart"):
                el = pr.find(qn("w:" + tag))
                if el is not None:
                    if tag == "numFmt":
                        got_fmt = el.get(qn("w:val"))
                    else:
                        got_rs = el.get(qn("w:val"))
        note = ""
        # Word 的默认值：编号格式 decimal、重排 continuous。未显式写元素＝用默认值。
        if got_rs is None:
            got_rs, note = "continuous", "（未显式设置，Word 默认 continuous）"
        if got_fmt is None:
            got_fmt, note = "decimal", "（未显式设置，Word 默认 decimal）"
        add("页下注编号格式", fn.get("num_fmt"), str(got_fmt), got_fmt == fn.get("num_fmt"))
        add("页下注编号重排", fn.get("num_restart"), str(got_rs),
            got_rs == fn.get("num_restart"), note)

    root = _footnotes(docx_path)
    if root is not None and fn.get("size"):
        notes = [n for n in root.findall(Q("footnote")) if not n.get(Q("type"))]
        sizes, fonts, lines = [], [], []
        for note in notes:
            for p in note.iter(Q("p")):
                sp = p.find(Q("pPr") + "/" + Q("spacing"))
                if sp is not None and sp.get(Q("line")):
                    lines.append(round(int(sp.get(Q("line"))) / 20, 2))
            for r in note.iter(Q("r")):
                rpr = r.find(Q("rPr"))
                if rpr is None:
                    continue
                rf = rpr.find(Q("rFonts"))
                if rf is not None and rf.get(qn("w:eastAsia")):
                    fonts.append(rf.get(qn("w:eastAsia")))
                sz = rpr.find(Q("sz"))
                if sz is not None:
                    sizes.append(round(int(sz.get(qn("w:val"))) / 2, 2))
        if notes:
            add("页下注条数", "≥1", str(len(notes)), True)
            s, _ = _modal(sizes)
            f, _ = _modal(fonts)
            l, _ = _modal(lines)
            if s is not None:
                add("页下注字号", f"{fn['size']}pt", f"{s}pt",
                    abs(s - fn["size"]) <= TOL["size"])
            if f is not None and fn.get("east"):
                add("页下注字体", fn["east"], f, f == fn["east"])
            if l is not None and fn.get("line"):
                add("页下注行距", f"{fn['line']}pt", f"{l}pt",
                    abs(l - fn["line"]) <= TOL["line"])

    # 4 图件：宽度不超通栏、题注位置（图题与表题要分开判：图题常在图下，表题必在表上）
    sec0 = doc.sections[0]
    content_w = sec0.page_width.pt - sec0.left_margin.pt - sec0.right_margin.pt
    want_pos = (spec.get("figure") or {}).get("caption_position")
    figs, fig_caps, tab_caps = [], [], []
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for par in doc.paragraphs:
        p = par._p
        ext = p.findall(".//" + WP + "extent")   # 注意是 wp:extent，不是 w:extent
        if ext:
            try:
                figs.append(round(int(ext[0].get("cx")) / 360000, 2))   # EMU → cm
            except (TypeError, ValueError):
                pass
        t = par.text.strip()
        if len(t) < 100:
            m = re.match(r"^(图|表)\s*\d", t)
            if m:
                (fig_caps if m.group(1) == "图" else tab_caps).append(par)
    if figs:
        over = [w for w in figs if w * 28.3465 > content_w + TOL["width"]]
        add("图宽（cm，通栏内）", f"≤{content_w/28.3465:.2f}",
            ", ".join(f"{w}" for w in figs), not over,
            "（超通栏会导致图被裁或顶出版心）" if over else "")
    if want_pos and fig_caps:
        def has_drawing(node):
            return node is not None and node.tag == qn("w:p") and \
                node.find(".//" + qn("w:drawing")) is not None
        wrong = 0
        for cap in fig_caps:
            # 题注"在图下方" ⇔ 它的前一个兄弟节点是图片段落；反之亦然。
            if want_pos == "below":
                wrong += 0 if has_drawing(cap._p.getprevious()) else 1
            else:
                wrong += 0 if has_drawing(cap._p.getnext()) else 1
        expect = "图题在图下方" if want_pos == "below" else "图题在图上方"
        add("图题位置", expect, f"{len(fig_caps)-wrong}/{len(fig_caps)} 符合", wrong == 0)
    if tab_caps:
        # 表题必须在表上方：题注段之后 3 个兄弟节点内应出现表格
        ok_cnt = 0
        for cap in tab_caps:
            node = cap._p.getnext()
            for _ in range(3):
                if node is None:
                    break
                if node.tag == qn("w:tbl"):
                    ok_cnt += 1
                    break
                node = node.getnext()
        add("表题位置", "表题在表上方", f"{ok_cnt}/{len(tab_caps)} 符合",
            ok_cnt == len(tab_caps))

    # 5 表格制式（例文实测，四刊各不相同，不能一刀切：
    #   《政治学研究》《管理世界》是网格表，《公共管理学报》是三线表，
    #   《中国行政管理》例文无表 → 未确认，跳过）
    style = spec.get("table_style") or TABLE_STYLE.get(journal)

    def has_vertical(table) -> bool:
        tblPr = table._tbl.tblPr
        borders = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        if borders is not None:
            for edge in ("insideV", "left", "right"):
                el = borders.find(qn("w:" + edge))
                if el is not None and el.get(qn("w:val")) not in (None, "none", "nil"):
                    return True
        for cell in table.rows[0].cells if len(table.rows) else []:
            tcPr = cell._tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            tcb = tcPr.find(qn("w:tcBorders"))
            if tcb is None:
                continue
            for edge in ("left", "right"):
                el = tcb.find(qn("w:" + edge))
                if el is not None and el.get(qn("w:val")) not in (None, "none", "nil"):
                    return True
        return False

    if doc.tables and style:
        n_grid = sum(1 for t in doc.tables if has_vertical(t))
        n_line = len(doc.tables) - n_grid
        if style == "three-line":
            add("表格制式", "三线表（无竖线）", f"{n_line}/{len(doc.tables)} 合规",
                n_grid == 0, "" if n_grid == 0 else "（有竖线，属网格表）",
                warn_only=not strict_tables)
        elif style == "grid":
            add("表格制式", "网格表（有竖线，同该刊例文）", f"{n_grid}/{len(doc.tables)} 合规",
                n_grid == len(doc.tables), "",
                warn_only=not strict_tables)

    # 6 注释制式（例文实测：《管理世界》是文末「注释」节，其余三刊是页下注）
    ns = spec.get("notes") or {}
    if notes_mode(spec, journal) == "endnote":
        head = None
        for par in doc.paragraphs:
            if par.text.strip() in ("注释", "注　释"):
                head = par
                break
        add("注释节", "文末「注释」标题", "有" if head is not None else "缺",
            head is not None)
        if head is not None:
            import re as _re
            items = []
            started = False
            for par in doc.paragraphs:
                if head is not None and par._p is head._p:
                    started = True
                    continue
                if not started:
                    continue
                t = par.text.strip()
                if not t:
                    continue
                if _re.match(r"^(参考文献|附录|[一二三四五六七八九十]+、)", t):
                    break
                if _re.match(r"^(?:\d+[.．、]|[①②③④⑤⑥⑦⑧⑨⑩])", t):
                    items.append(par)
            add("注释条数", "≥1", str(len(items)), bool(items))
            if items and ns.get("size"):
                sizes = []
                for par in items:
                    for run in par.runs:
                        _e, s = _run_attr(run)
                        if s:
                            sizes.append(s)
                            break
                got, _n = _modal(sizes)
                if got is not None:
                    add("注释字号", f"{ns['size']}pt", f"{got}pt",
                        abs(got - ns["size"]) <= TOL["size"])
            # 顺序：注释应在参考文献之前
            refs = next((p for p in doc.paragraphs if p.text.strip() == "参考文献"), None)
            if refs is not None and head is not None:
                order_ok = doc.element.body.index(head._p) < doc.element.body.index(refs._p)
                add("注释节位置", "在参考文献之前", "之前" if order_ok else "之后", order_ok)
        if _footnotes(docx_path) is not None:
            res.setdefault("warnings", []).append("注释制式：该刊为文末注释，但文档仍含页下注部件")

    # 7 文本完整性
    if ref_path:
        ref_doc = Document(str(ref_path))
        if notes_mode(spec, journal) == "endnote":
            # 文末注释节会被版式脚本挪到参考文献之前，"顺序变了"不等于"文字改了"，
            # 所以两侧都先剔掉注释节再比。
            ref_text = _norm(_text_excluding_notes(ref_doc))
            got_text = _norm(_text_excluding_notes(doc))
        else:
            ref_text = _norm(_all_text(ref_doc))
            got_text = _norm(_all_text(doc))
        if ref_text == got_text:
            add("文本完整性", "与底稿完全一致", "一致", True)
        else:
            # 定位第一处差异
            i = next((k for k in range(min(len(ref_text), len(got_text)))
                      if ref_text[k] != got_text[k]), min(len(ref_text), len(got_text)))
            add("文本完整性", "与底稿完全一致",
                f"第 {i} 字起不同：底稿「{ref_text[i:i+18]}」→ 成品「{got_text[i:i+18]}」",
                False)
    else:
        res["checks"].append({"项": "文本完整性", "应为": "（未提供 --ref，跳过）",
                              "实测": "", "结果": "跳过", "备注": ""})

    res["stats"] = {"段落": nonempty, "空段落": empty, "表格": len(doc.tables),
                    "图": len(figs), "分类": {k: len(v) for k, v in buckets.items()}}
    # 未确认参数：spec 里为 None 的字段＝该刊这项没有可靠依据，校验时跳过。
    # 把它们显式列出来，避免"PASS"被误读成"全部核对过"。
    unconfirmed = []

    def walk(node, prefix=""):
        for k, v in (node or {}).items():
            if k in ("columns_start_at", "column_space", "span_tables", "span_figures"):
                continue
            if isinstance(v, dict):
                walk(v, f"{prefix}{k}.")
            elif v is None:
                unconfirmed.append(f"{prefix}{k}")
    for key in ("page", "margins", "body", "title", "abstract", "h1", "h2", "h3",
                "caption", "table_cell", "footnote", "figure"):
        walk(spec.get(key) or {}, f"{key}.")
    res["unconfirmed"] = unconfirmed
    res["pass"] = not res["failures"]
    return res


def main() -> int:
    ap = argparse.ArgumentParser(description="docx 版式结构校验（数值断言，不渲染）")
    ap.add_argument("--docx", required=True)
    ap.add_argument("--journal", required=True, choices=list(SPECS))
    ap.add_argument("--ref", help="对照底稿，检查文字是否被版式脚本改动")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--strict", action="store_true", help="有不符合项时退出码 1")
    ap.add_argument("--strict-tables", action="store_true",
                    help="把三线表不合规升级为不符合项（默认只提示）")
    ap.add_argument("--mode", choices=["submission", "published"], default="submission",
                    help="submission=投稿版式 SPECS（默认）；published=刊出版式 PUBLISHED_SPECS")
    a = ap.parse_args()
    res = verify(Path(a.docx), a.journal, Path(a.ref) if a.ref else None, a.strict_tables,
                 a.mode)
    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
        return 0 if res["pass"] or not a.strict else 1
    print(f"\n== 结构校验：{res['name']}"
          f"{'（刊出版式）' if res.get('mode') == 'published' else '（投稿版式）'}"
          f" ==  {Path(a.docx).name}")
    width = max(len(c["项"]) for c in res["checks"]) if res["checks"] else 8
    for c in res["checks"]:
        mark = {"OK": "✓", "错": "✗", "提示": "!", "跳过": "·"}.get(c["结果"], "?")
        print(f" {mark} {c['项']:<{width}}  应为 {c['应为']:<22} 实测 {c['实测']} {c['备注']}")
    if res.get("warnings"):
        print(f"\n提示 {len(res['warnings'])} 条（不影响结论）：")
        for w in res["warnings"]:
            print("  · " + w)
    if res["failures"]:
        print(f"\n不符合项 {len(res['failures'])} 条：")
        for f in res["failures"]:
            print("  - " + f)
        print("\n结论：FAIL")
    else:
        extra = ""
        if res.get("unconfirmed"):
            extra = f"；另有 {len(res['unconfirmed'])} 项该刊参数未确认，已跳过"
        print(f"\n结论：PASS（{res['stats']['段落']} 段 / {res['stats']['表格']} 表 / "
              f"{res['stats']['图']} 图{extra}）")
        if res.get("unconfirmed"):
            print("  未核对项：" + "、".join(res["unconfirmed"][:10])
                  + ("…" if len(res["unconfirmed"]) > 10 else ""))
    return 0 if res["pass"] else (1 if a.strict else 0)


if __name__ == "__main__":
    raise SystemExit(main())
