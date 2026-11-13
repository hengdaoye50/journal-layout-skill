#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""留档与差异断言：批量改成品前先备份，改完证明"只有预期部件变了"。

这个项目最贵的一条纪律是：
    每批次改版式之前，把当时的成品整份留档到 03_历史稿与原稿/<日期>_<阶段>前/；
    改完之后不要靠"看着差不多"，而是逐部件比对 zip 内容，证明只有声明过的部件
    发生变化。曾经因为没做这一步，一次"只换图片"的操作实际动了正文。

用法：
    # 改之前留档
    python snapshot.py --backup "01_四刊最终稿/*.docx" --label 版式调整
    # 改之后比对（before 是留档，after 是当前）
    python snapshot.py --diff 03_历史稿与原稿/2026-01-01_改动前 --against 01_四刊最终稿
    # 只允许 document.xml 与 media 变化
    python snapshot.py --diff ... --allow word/document.xml word/media/
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path
from zipfile import ZipFile

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

DEFAULT_ARCHIVE = "03_历史稿与原稿"


def backup(patterns: list[str], label: str, archive: Path = Path(DEFAULT_ARCHIVE),
           root: Path = Path(".")) -> Path:
    dst = archive / f"{date.today().isoformat()}_{label}前"
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for pat in patterns:
        for f in root.glob(pat):
            if f.is_file() and f.suffix == ".docx":
                target = dst / f.name
                if target.exists():
                    print(f"  已存在，跳过：{f.name}")
                    continue
                shutil.copy2(f, target)
                n += 1
                print(f"  留档 {f.name}")
    print(f"共留档 {n} 个文件 → {dst}")
    return dst


def _parts(p: Path) -> dict:
    with ZipFile(p) as z:
        return {n: z.read(n) for n in z.namelist()}


def diff(before_dir: Path, after_dir: Path, allow: list[str]) -> int:
    """逐文件、逐部件比对。allow 里的前缀视为允许变化。"""
    bad = 0
    names = sorted({f.name for f in before_dir.glob("*.docx")})
    if not names:
        print(f"备份目录里没有 docx：{before_dir}")
        return 2
    for name in names:
        a, b = before_dir / name, after_dir / name
        if not b.exists():
            print(f"[缺失] {name}：当前目录没有该文件")
            bad += 1
            continue
        pa, pb = _parts(a), _parts(b)
        changed = [n for n in pa if n in pb and pa[n] != pb[n]]
        added = [n for n in pb if n not in pa]
        removed = [n for n in pa if n not in pb]
        unexpected = [n for n in changed + added + removed
                      if not any(n == x or n.startswith(x) for x in allow)]
        head = f"[{name}] 变化部件 {len(changed + added + removed)} 个"
        if unexpected:
            print(f"{head}\n  ✗ 未声明变化：{unexpected}")
            bad += 1
        else:
            print(f"{head}\n  ✓ 全部在允许范围内：{changed + added + removed or '无'}")
    print(f"\n结论：{'PASS' if bad == 0 else 'FAIL'}（{bad} 个文件有未声明变化）")
    return 0 if bad == 0 else 1


def main() -> int:
    ap = argparse.ArgumentParser(description="成品留档与 docx 部件级差异断言")
    ap.add_argument("--backup", nargs="+", metavar="GLOB", help="要留档的文件通配")
    ap.add_argument("--label", default="改动", help="阶段名，如 版式调整")
    ap.add_argument("--archive", default=DEFAULT_ARCHIVE)
    ap.add_argument("--diff", metavar="BEFORE_DIR")
    ap.add_argument("--against", default=".", metavar="AFTER_DIR")
    ap.add_argument("--allow", nargs="*", default=[],
                    help="允许变化的部件前缀，如 word/document.xml word/media/")
    a = ap.parse_args()
    if a.backup:
        backup(a.backup, a.label, Path(a.archive))
        return 0
    if a.diff:
        return diff(Path(a.diff), Path(a.against), a.allow)
    ap.print_help()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
