#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""抽页预览：把 docx 渲染成 PDF，只抽指定页拼成一张缩略图。

定位很明确——它是 verify_layout.py 的补充，不是主校验手段。
版式对不对由结构断言判断；只有"图挤到页底""表格跨页断裂"这类视觉问题才需要看图，
而这类问题通常出现在含图表的那几页。所以默认只出 2 页（首页 + 最后一页），
用 --pages 指定可疑页，用 --all 才全量出图。

用法：
    python render_preview.py --docx 成品.docx --pages 1 7 12 --out preview/
    python render_preview.py --docx 成品.docx --all --dpi 110

前置：本机已装 LibreOffice。PDF 转图片需要 PyMuPDF（import fitz）。
     LibreOffice 必须带 -env:UserInstallation 指向可写目录，否则 soffice 会直接崩
     （exit 0xC0000409）。这一点在 references/pitfalls.md 里有记录。
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

SOFFICE_CANDIDATES = [
    r"D:\libreoffice\program\soffice.exe",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice() -> str | None:
    for c in SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    from shutil import which
    return which("soffice") or which("libreoffice")


def to_pdf(docx: Path, outdir: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise SystemExit("未找到 LibreOffice（soffice），无法转 PDF。可跳过本步骤。")
    # profile 必须建在工作区（可写且不受沙箱限制），且与输出目录分开：
    # 放在系统 TEMP 下 soffice 会因 profile 不可访问而 exit 77 / 0xC0000409；
    # 与 PDF 输出目录混在一起时偶发 libpng write error。用完即删。
    profile = outdir.parent / (outdir.name + ".lo-profile")
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, f"-env:UserInstallation=file:///{profile.resolve().as_posix()}",
           "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", str(outdir), str(docx)]
    # soffice 在中文 Windows 上的输出不是 UTF-8，必须显式给编码，否则 reader 线程
    # 会抛 UnicodeDecodeError，把真实错误掩盖成 NoneType。
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=180)
    pdf = outdir / (docx.stem + ".pdf")
    if r.returncode != 0 or not pdf.exists():
        raise SystemExit(f"LibreOffice 转换失败（exit {r.returncode}）\n"
                         f"stdout: {(r.stdout or '')[-500:]}\n"
                         f"stderr: {(r.stderr or '')[-500:]}")
    return pdf


RENDER_HELPER = r'''
import sys
import fitz
pdf, outdir, dpi, pages = sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4]
doc = fitz.open(pdf)
total = doc.page_count
sel = list(range(1, total + 1)) if pages == "all" else [int(p) for p in pages.split(",") if p]
for pno in sel:
    if not 1 <= pno <= total:
        continue
    pix = doc.load_page(pno - 1).get_pixmap(dpi=dpi)
    pix.save(f"{outdir}/page-{pno:02d}.png")
    print(f"rendered {pno}")
print(f"total {total}")
'''


def _find_fitz_python() -> str | None:
    """当前解释器没有 PyMuPDF 时，找一个装了的解释器来代渲染。"""
    import shutil
    for cand in (sys.executable, shutil.which("python"), shutil.which("python3"),
                 shutil.which("py")):
        if not cand:
            continue
        try:
            r = subprocess.run([cand, "-c", "import fitz"], capture_output=True, timeout=30)
            if r.returncode == 0:
                return cand
        except Exception:
            continue
    return None


def render(pdf: Path, pages: list[int] | None, dpi: int, outdir: Path) -> list[Path]:
    try:
        import fitz  # PyMuPDF
    except ImportError:
        helper_py = _find_fitz_python()
        if not helper_py:
            raise SystemExit("渲染 PDF 需要 PyMuPDF：pip install pymupdf\n"
                             f"（PDF 已生成，可直接用阅读器看：{pdf}）")
        script = outdir / "_render_pdf.py"
        script.write_text(RENDER_HELPER, encoding="utf-8")
        sel = "all" if pages is None else ",".join(str(p) for p in pages)
        r = subprocess.run([helper_py, str(script), str(pdf.resolve()), str(outdir.resolve()),
                            str(dpi), sel], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=300)
        if r.returncode != 0:
            raise SystemExit(f"代渲染失败：{(r.stderr or '')[-400:]}")
        print((r.stdout or "").strip())
        return sorted(outdir.glob("page-*.png"))
    doc = fitz.open(pdf)
    total = doc.page_count
    if pages is None:
        pages = sorted({1, total})
    out = []
    for pno in pages:
        if not 1 <= pno <= total:
            continue
        pix = doc.load_page(pno - 1).get_pixmap(dpi=dpi)
        f = outdir / f"page-{pno:02d}.png"
        pix.save(str(f))
        out.append(f)
    doc.close()
    print(f"共 {total} 页；已渲染 {[p for p in pages if 1 <= p <= total]}")
    return out


def montage(images: list[Path], out: Path, gap: int = 12) -> Path | None:
    """把几页并排拼成一张，方便一眼扫完。"""
    if len(images) < 2:
        return images[0] if images else None
    from PIL import Image
    ims = [Image.open(p) for p in images]
    h = max(i.height for i in ims)
    w = sum(i.width for i in ims) + gap * (len(ims) - 1)
    canvas = Image.new("RGB", (w, h), "white")
    x = 0
    for im in ims:
        canvas.paste(im, (x, 0))
        x += im.width + gap
    canvas.thumbnail((2400, 2400))
    canvas.save(out)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="docx → PDF → 抽样页面图")
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out", default="preview")
    ap.add_argument("--pages", type=int, nargs="*", help="要渲染的页码，如 1 7 12")
    ap.add_argument("--all", action="store_true", help="全部页面（一般不需要）")
    ap.add_argument("--dpi", type=int, default=120)
    a = ap.parse_args()
    docx = Path(a.docx)
    outdir = Path(a.out)
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = to_pdf(docx, outdir)
    if a.all:
        import fitz
        pages = list(range(1, fitz.open(pdf).page_count + 1))
    else:
        pages = a.pages
    imgs = render(pdf, pages, a.dpi, outdir)
    sheet = montage(imgs, outdir / "contact-sheet.png")
    print(f"PDF: {pdf}")
    if sheet:
        print(f"缩略总览: {sheet}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
