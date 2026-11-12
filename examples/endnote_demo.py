#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""演示「文末注释」通道：md2docx.py --notes-mode endnote 产出一份带「注释」节的 docx。

背景：四刊里《管理世界》真刊是文末「注释」节（连续编号），其余三刊是页下注。
      md2docx 默认 --notes-mode footnote（写 word/footnotes.xml）；endnote 模式下
      正文里的 [1] [2] 标记原样保留，注文改成文档末尾的普通段落节，不生成 footnotes.xml。

跑法（在 skill 根目录或任意位置都行，路径按 __file__ 解析）：
    <bundled-python> examples/endnote_demo.py

产物（都落在 examples/out/，每次覆盖重写）：
    endnote_demo.md     极小的示例 Markdown：两段正文 + 两条 [^n]: 注释定义
    endnote_demo.docx   md2docx --notes-mode endnote 的产物

断言（失败即抛错、退出码非 0）：
    1) 文档里有「注释」标题段落，且用的是 Heading 1 样式；
    2) 末尾有两条编号注释：`1. …`、`2. …`；
    3) 正文里的 [1] [2] 标记还在，且没有被换成 footnoteReference；
    4) 包里没有 word/footnotes.xml，footnotes.extract() 读不到页下注。
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

from docx import Document
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
OUT = ROOT / "examples" / "out"
sys.path.insert(0, str(SCRIPTS))

import footnotes  # noqa: E402  同目录的注释模块
import md2docx  # noqa: E402

MD = """# 文末注释演示稿

正文第一段：这一句带第一个注号[^1]，用来确认 endnote 模式下正文标记原样保留。

正文第二段：这一句带第二个注号[^2]，再补一句普通正文，凑成两段正文。

[^1]: 第一条注释：说明数据来源与样本区间。
[^2]: 第二条注释：说明变量口径与缺失值处理。
"""


def run_cli(md: Path, out: Path) -> None:
    """走真实的 argparse + main() 路径：python md2docx.py ... --notes-mode endnote。"""
    argv = sys.argv
    sys.argv = ["md2docx.py", "--md", str(md), "--out", str(out),
                "--notes-mode", "endnote"]
    try:
        rc = md2docx.main()
    finally:
        sys.argv = argv
    assert rc == 0, f"md2docx.main() 退出码 {rc}"


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    md = OUT / "endnote_demo.md"
    docx = OUT / "endnote_demo.docx"
    md.write_text(MD, encoding="utf-8")
    print(f"[1/2] 生成示例 Markdown：{md.relative_to(ROOT)}")
    run_cli(md, docx)
    print(f"[2/2] 校验 {docx.relative_to(ROOT)}")

    doc = Document(str(docx))
    paras = doc.paragraphs
    texts = [p.text for p in paras]

    # 1) 「注释」标题 + Heading 1 样式
    heads = [i for i, p in enumerate(paras)
             if p.text.strip() == "注释" and p.style.name.lower() == "heading 1"]
    assert heads, "没有找到 Heading 1 样式的「注释」标题段落"
    head_i = heads[0]

    # 2) 两条编号注释（1. … / 2. …），且在标题之后、位于文档末尾
    note_idx = [i for i, t in enumerate(texts) if re.fullmatch(r"[12]\. .+", t.strip())]
    assert len(note_idx) == 2, f"文末注释条数应为 2，实际 {len(note_idx)}：{[texts[i] for i in note_idx]}"
    assert head_i < note_idx[0] < note_idx[1], "注释段落不在「注释」标题之后"
    tail = [p for p in paras[note_idx[-1] + 1:] if p.text.strip()]
    assert not tail, f"注释节之后还有正文段落：{[p.text for p in tail]}"
    assert texts[note_idx[0]] == "1. 第一条注释：说明数据来源与样本区间。", texts[note_idx[0]]
    assert texts[note_idx[1]] == "2. 第二条注释：说明变量口径与缺失值处理。", texts[note_idx[1]]
    # 注释段落格式：宋体 10 pt / 固定行距 15 pt（md2docx 传的默认参数）
    note_par = paras[note_idx[0]]
    assert note_par.runs, "注释段落没有 run"
    assert note_par.runs[0].font.size.pt == 10, note_par.runs[0].font.size
    assert note_par.paragraph_format.line_spacing.pt == 15, note_par.paragraph_format.line_spacing

    # 3) 正文 [1] [2] 标记还在，且没有 footnoteReference
    body = "\n".join(texts[:head_i])
    assert "[1]" in body and "[2]" in body, f"正文标记丢失：{body}"
    with zipfile.ZipFile(docx) as z:
        assert "word/footnotes.xml" not in z.namelist(), "endnote 模式不应生成 word/footnotes.xml"
        root = etree.fromstring(z.read("word/document.xml"))
    refs = [e for e in root.iter(footnotes.Q("footnoteReference"))]
    assert not refs, f"正文里出现了 {len(refs)} 个 footnoteReference（endnote 应保留普通文本）"

    # 4) 页下注读取为空
    assert footnotes.extract(docx) == [], "footnotes.extract() 应读不到页下注"

    print("PASS：文末「注释」节 2 条，正文 [1][2] 保留，无 footnotes.xml / footnoteReference")
    print(f"  产物：{docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
