#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把正文里的 [1] [2] 标记变成真正的 Word 页下注（footnote），并统一页下注样式。

为什么需要这个脚本：
    python-docx 不支持页下注（它底层没有 footnotes.xml 的 API）。中文社科期刊
    普遍要求"页下注、圈码编号、逐页重编"，而 Markdown/Word 底稿里的注释通常是
    文末"注释"列表。本脚本直接操作 OOXML（zipfile + lxml）完成转换。

用法：
    python footnotes.py --docx 底稿.docx --notes notes.json      # 注入
    python footnotes.py --docx 初稿.docx --style --east 宋体 --size 10 --line 15
    python footnotes.py --docx 成品.docx --extract              # 读出现有页下注

notes.json 形如 {"1": "政务服务改革是……", "2": "便民服务事项通常……"}

安全设计：
    · 只在 run 文本"整个等于 [n]"时替换，正文里偶然出现的 [1] 不会被误伤；
    · 写盘先落 .tmp 再 replace，失败不留半成品；
    · 已存在 footnotes.xml 时做合并（追加编号），不覆盖原有注释。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
Q = lambda tag: "{%s}%s" % (W, tag)
FN_CT = ("application/vnd.openxmlformats-officedocument"
         ".wordprocessingml.footnotes+xml")


def _read(path: Path) -> dict:
    with ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write(path: Path, files: dict) -> None:
    tmp = path.with_suffix(".fn.tmp")
    with ZipFile(tmp, "w", ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    tmp.replace(path)


def _ensure_footnotes_part(files: dict) -> etree._Element:
    """确保 footnotes.xml、关系项与 content-type 都在；返回可写的根元素。"""
    if "word/footnotes.xml" not in files:
        root = etree.Element(Q("footnotes"), nsmap={"w": W})
        for fid, kind in ((-1, "separator"), (0, "continuationSeparator")):
            fn = etree.SubElement(root, Q("footnote"))
            fn.set(Q("type"), kind)
            fn.set(Q("id"), str(fid))
            p = etree.SubElement(fn, Q("p"))
            r = etree.SubElement(p, Q("r"))
            etree.SubElement(r, Q(kind))
        files["word/footnotes.xml"] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True)
        # 关系
        rels = etree.fromstring(files["word/_rels/document.xml.rels"])
        used = {x.get("Id") for x in rels}
        i = 1
        while "rId%d" % i in used:
            i += 1
        rel = etree.SubElement(rels, "{%s}Relationship" % PKG_REL)
        rel.set("Id", "rId%d" % i)
        rel.set("Type", R + "/footnotes")
        rel.set("Target", "footnotes.xml")
        files["word/_rels/document.xml.rels"] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        # content type
        ct = etree.fromstring(files["[Content_Types].xml"])
        if not any(o.get("PartName") == "/word/footnotes.xml"
                   for o in ct.findall("{%s}Override" % CT)):
            o = etree.SubElement(ct, "{%s}Override" % CT)
            o.set("PartName", "/word/footnotes.xml")
            o.set("ContentType", FN_CT)
            files["[Content_Types].xml"] = etree.tostring(
                ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    return etree.fromstring(files["word/footnotes.xml"])


def _make_footnote(nid: int, text: str, east: str | None = None,
                   size: float | None = None, line: float | None = None) -> etree._Element:
    fn = etree.Element(Q("footnote"))
    fn.set(Q("id"), str(nid))
    p = etree.SubElement(fn, Q("p"))
    ppr = etree.SubElement(p, Q("pPr"))
    etree.SubElement(ppr, Q("pStyle")).set(Q("val"), "FootnoteText")
    if line:
        sp = etree.SubElement(ppr, Q("spacing"))
        sp.set(Q("line"), str(int(round(line * 20))))
        sp.set(Q("lineRule"), "exact")
        sp.set(Q("before"), "0")
        sp.set(Q("after"), "0")
    r1 = etree.SubElement(p, Q("r"))
    rp = etree.SubElement(r1, Q("rPr"))
    etree.SubElement(rp, Q("rStyle")).set(Q("val"), "FootnoteReference")
    etree.SubElement(r1, Q("footnoteRef"))
    r2 = etree.SubElement(p, Q("r"))
    if east or size:
        rp2 = etree.SubElement(r2, Q("rPr"))
        if east:
            rf = etree.SubElement(rp2, Q("rFonts"))
            rf.set(Q("eastAsia"), east)
            rf.set(Q("ascii"), "Times New Roman")
            rf.set(Q("hAnsi"), "Times New Roman")
        if size:
            etree.SubElement(rp2, Q("sz")).set(Q("val"), str(int(round(size * 2))))
            etree.SubElement(rp2, Q("szCs")).set(Q("val"), str(int(round(size * 2))))
    t = etree.SubElement(r2, Q("t"))
    t.set(XML_SPACE, "preserve")
    t.text = " " + text
    return fn


def inject(docx_path: Path, notes: dict, style: dict | None = None,
           marker: str = r"\[(\d+)\]", verbose: bool = True) -> dict:
    """把正文中整段等于 [n] 的 run 替换为页下注引用，并写入 notes[n]。"""
    style = style or {}
    files = _read(docx_path)
    doc = etree.fromstring(files["word/document.xml"])
    foot = _ensure_footnotes_part(files)
    taken = {int(x.get(Q("id"))) for x in foot.findall(Q("footnote"))
             if x.get(Q("id"), "").lstrip("-").isdigit()}
    next_id = max([i for i in taken if i > 0] or [0]) + 1

    pat = re.compile(marker)
    replaced, missing = [], []
    id_of: dict[int, int] = {}
    for run in doc.iter(Q("r")):
        txt = "".join(t.text or "" for t in run.findall(Q("t")))
        m = pat.fullmatch(txt.strip()) if txt else None
        if not m:
            continue
        num = int(m.group(1))
        if str(num) not in notes and num not in notes:
            missing.append(num)
            continue
        nid = id_of.get(num)
        if nid is None:
            nid = next_id
            next_id += 1
            id_of[num] = nid
            foot.append(_make_footnote(nid, str(notes.get(str(num), notes.get(num))),
                                       style.get("east"), style.get("size"),
                                       style.get("line")))
        for child in list(run):
            run.remove(child)
        rp = etree.SubElement(run, Q("rPr"))
        etree.SubElement(rp, Q("rStyle")).set(Q("val"), "FootnoteReference")
        etree.SubElement(run, Q("footnoteReference")).set(Q("id"), str(nid))
        replaced.append(num)

    files["word/document.xml"] = etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/footnotes.xml"] = etree.tostring(
        foot, xml_declaration=True, encoding="UTF-8", standalone=True)
    _write(docx_path, files)
    rep = {"injected": len(replaced), "markers": sorted(set(replaced)),
           "missing": sorted(set(missing))}
    if verbose:
        print(f"注入页下注 {rep['injected']} 处：{rep['markers']}")
        if rep["missing"]:
            print(f"  ! 正文有标记但 notes 里没有条目：{rep['missing']}")
    return rep


def style_all(docx_path: Path, east: str, size: float, line: float | None = None) -> int:
    """统一已有页下注的字体/字号/行距。"""
    files = _read(docx_path)
    if "word/footnotes.xml" not in files:
        print("该文档没有页下注部分", file=sys.stderr)
        return 0
    root = etree.fromstring(files["word/footnotes.xml"])
    n = 0
    for note in root.findall(Q("footnote")):
        if note.get(Q("type")):
            continue
        n += 1
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
            if line:
                sp.set(Q("line"), str(int(round(line * 20))))
                sp.set(Q("lineRule"), "exact")
            for run in p.iter(Q("r")):
                rpr = run.find(Q("rPr"))
                if rpr is None:
                    rpr = etree.Element(Q("rPr"))
                    run.insert(0, rpr)
                rf = rpr.find(Q("rFonts"))
                if rf is None:
                    rf = etree.SubElement(rpr, Q("rFonts"))
                for k, v in (("eastAsia", east), ("ascii", "Times New Roman"),
                             ("hAnsi", "Times New Roman"), ("cs", "Times New Roman")):
                    rf.set(Q(k), v)
                for tag, val in (("sz", int(round(size * 2))),
                                 ("szCs", int(round(size * 2)))):
                    e = rpr.find(Q(tag))
                    if e is None:
                        e = etree.SubElement(rpr, Q(tag))
                    e.set(Q("val"), str(val))
                for tag in ("b", "i"):
                    e = rpr.find(Q(tag))
                    if e is None:
                        e = etree.SubElement(rpr, Q(tag))
                    e.set(Q("val"), "0")
    files["word/footnotes.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _write(docx_path, files)
    print(f"页下注样式已统一：{east} {size}pt" + (f" / 固定行距 {line}pt" if line else "")
          + f"，共 {n} 条")
    return n


def extract(docx_path: Path) -> list[str]:
    """读出现有页下注文本（第一条是编号 run，因此跳过 footnoteRef 所在的 run）。"""
    files = _read(docx_path)
    if "word/footnotes.xml" not in files:
        return []
    root = etree.fromstring(files["word/footnotes.xml"])
    out = []
    for note in root.findall(Q("footnote")):
        if note.get(Q("type")):
            continue
        parts = []
        for p in note.iter(Q("p")):
            for r in p.iter(Q("r")):
                if r.find(Q("footnoteRef")) is not None:
                    continue
                parts.append("".join(t.text or "" for t in r.findall(Q("t"))))
        out.append("".join(parts).strip())
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Word 页下注注入/样式/导出")
    ap.add_argument("--docx", required=True)
    ap.add_argument("--notes", help="注释 json：{\"1\": \"...\"}")
    ap.add_argument("--marker", default=r"\[(\d+)\]", help="正文标记正则，默认 [n]")
    ap.add_argument("--style", action="store_true", help="统一现有页下注样式")
    ap.add_argument("--east", default="宋体")
    ap.add_argument("--size", type=float, default=10)
    ap.add_argument("--line", type=float, default=None)
    ap.add_argument("--extract", action="store_true", help="打印现有页下注")
    a = ap.parse_args()
    p = Path(a.docx)
    if a.extract:
        for i, t in enumerate(extract(p), 1):
            print(f"[{i}] {t}")
        return 0
    if a.style:
        style_all(p, a.east, a.size, a.line)
        return 0
    if not a.notes:
        ap.print_help()
        return 2
    notes = json.loads(Path(a.notes).read_text(encoding="utf-8"))
    inject(p, notes, {"east": a.east, "size": a.size, "line": a.line}, a.marker)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
