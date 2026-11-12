# -*- coding: utf-8 -*-
"""embed_figure.py —— 把 PNG 图件替换进 .docx 的内嵌图（只改图字节 + cx/cy，正文零改动）

用途：期刊版式 docx 里图已排好位置，论文返修后只需换图，不想重排整篇文档时使用。
做法：zipfile 解包 -> lxml 定位 wp:inline/wp:anchor -> 读 a:blip 的 r:embed ->
rels 里找 media 目标 -> 替换该 media 字节 -> 同步 wp:extent 与 a:xfrm/a:ext 的
cx/cy（按新图纵横比重算，绝不拉伸）-> 先写 .tmp，重新打开校验后再 replace。

保护性断言（每次运行都做，--dry-run 也做内存校验）：
  1) 替换前后 word/document.xml 的 //w:t 文本序列完全一致；
  2) 压缩包内除 word/document.xml 与被替换的 word/media/*.png 外无其他部件变化；
  3) 新包能重新打开，且 media 字节、cx/cy 与预期一致；
  4) 只在全部校验通过后才落盘（先写 .tmp 再 replace），失败不产生坏文件。

环境：只用标准库（zipfile/struct/argparse）+ lxml；不需要 python-docx，也不读图。

命令行：
  python embed_figure.py --docx A.docx --image fig.png --dry-run
  python embed_figure.py --docx A.docx --image fig.png --index 1 --width-cm 16
  python embed_figure.py --docx A.docx --image fig.png --out B.docx
"""

from __future__ import annotations

import argparse
import posixpath
import shutil
import struct
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as E

sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):  # 捕获到 GBK 控制台时中文错误信息不乱码
    sys.stderr.reconfigure(encoding="utf-8")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
EMU_PER_CM = 360000
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


class EmbedError(RuntimeError):
    """用户可读的错误（参数/文档结构/断言失败）。"""


# ---------------------------------------------------------------------------
# 输入检查
# ---------------------------------------------------------------------------


def read_png_size(image_path: Path) -> tuple[int, int]:
    """从 PNG IHDR 读像素宽高（不依赖 Pillow）。"""
    data = image_path.read_bytes()
    if len(data) < 24 or data[:8] != PNG_SIGNATURE:
        raise EmbedError(f"不是有效的 PNG 文件：{image_path}")
    width, height = struct.unpack(">II", data[16:24])
    if width <= 0 or height <= 0:
        raise EmbedError(f"PNG 尺寸非法：{width}x{height}（{image_path}）")
    return width, height


def parse_parts(docx: Path) -> tuple[list[str], dict[str, bytes]]:
    """按原顺序读全部部件，返回 (顺序, {部件名: 字节})。"""
    if not docx.is_file():
        raise EmbedError(f"docx 不存在：{docx}")
    try:
        with ZipFile(docx) as z:
            order = z.namelist()
            parts = {name: z.read(name) for name in order}
    except Exception as exc:  # BadZipFile / OSError
        raise EmbedError(f"无法作为 zip/pptx/docx 打开：{docx}（{exc}）")
    for required in ("word/document.xml", "word/_rels/document.xml.rels"):
        if required not in parts:
            raise EmbedError(f"{docx.name} 缺少必需部件 {required}，可能不是 Word 文档")
    return order, parts


def find_drawings(doc) -> list:
    """按文档顺序取全部图片容器（inline 与 anchor）。"""
    return doc.xpath("//wp:inline | //wp:anchor", namespaces=NS)


def drawing_info(drawing, rels, index: int) -> dict:
    """解析单张图：rId、media 部件、容器名、当前 cx/cy。"""
    tag = E.QName(drawing).localname
    rids = drawing.xpath(".//a:blip/@r:embed", namespaces=NS)
    if not rids:
        raise EmbedError(f"第 {index + 1} 张图（{tag}）没有 a:blip r:embed，"
                         f"不是常规内嵌位图（可能是图表/外部链接图），本工具不处理")
    rid = rids[0]
    target = rels.get(rid)
    if not target:
        raise EmbedError(f"关系 {rid} 在 word/_rels/document.xml.rels 中不存在")
    if target.lower().startswith("http") or target.startswith("../"):
        raise EmbedError(f"关系 {rid} 指向外部：{target}，本工具只处理包内图片")
    target_clean = target[1:] if target.startswith("/") else target
    media = posixpath.normpath(posixpath.join("word", target_clean))
    if media.startswith("../") or media.startswith("word/.."):
        raise EmbedError(f"关系 {rid} 的 Target 越界：{target}")
    extents = drawing.xpath("./wp:extent", namespaces=NS)
    if not extents:
        raise EmbedError(f"{media} 的图片容器缺少 wp:extent，无法确定显示尺寸")
    cx, cy = int(extents[0].get("cx")), int(extents[0].get("cy"))
    names = drawing.xpath(".//wp:docPr/@name", namespaces=NS)
    descrs = drawing.xpath(".//wp:docPr/@descr", namespaces=NS)
    return {
        "index": index,
        "tag": tag,
        "rid": rid,
        "media": media,
        "cx": cx,
        "cy": cy,
        "name": names[0] if names else "",
        "descr": descrs[0] if descrs else "",
        "drawing": drawing,
    }


def collect_drawings(doc, parts) -> list[dict]:
    """附带 rels 的图片清单。"""
    rels_root = E.fromstring(parts["word/_rels/document.xml.rels"])
    rels = {x.get("Id"): x.get("Target") for x in rels_root}
    drawings = find_drawings(doc)
    if not drawings:
        raise EmbedError("document.xml 中没有找到任何内嵌图片（wp:inline / wp:anchor）")
    infos = []
    for i, drawing in enumerate(drawings):
        info = drawing_info(drawing, rels, i)
        if info["media"] not in parts:
            raise EmbedError(f"关系指向的 media 部件不在包里：{info['media']}")
        infos.append(info)
    return infos


# ---------------------------------------------------------------------------
# 尺寸换算
# ---------------------------------------------------------------------------


def cm_to_emu(cm: float) -> int:
    return int(round(float(cm) * EMU_PER_CM))


def emu_to_cm(emu: int) -> float:
    return emu / EMU_PER_CM


def fmt_cm(emu: int) -> str:
    return f"{emu_to_cm(emu):.2f}cm"


def choose_size(old_cx: int, old_cy: int, img_w: int, img_h: int,
                width_cm: float | None, height_cm: float | None) -> tuple[int, int, list[str]]:
    """决定新的 cx/cy。只给 width 时按图像纵横比重算高度；都不给则保留原宽。"""
    notes: list[str] = []
    new_cx = cm_to_emu(width_cm) if width_cm is not None else old_cx
    if width_cm is not None and height_cm is None:
        new_cy = int(round(new_cx * img_h / img_w))
    elif height_cm is not None:
        new_cy = cm_to_emu(height_cm)
        if width_cm is not None:
            old_aspect = old_cx / old_cy
            new_aspect = new_cx / new_cy
            if abs(new_aspect - old_aspect) / old_aspect > 0.01:
                notes.append("警告：同时指定 width/height 会改变显示纵横比（可能拉伸图像）")
    else:
        new_cy = int(round(new_cx * img_h / img_w))
    old_aspect = old_cx / old_cy
    new_aspect = new_cx / new_cy
    if abs(new_aspect - old_aspect) / old_aspect > 0.01:
        notes.append("提示：新显示纵横比与原框不同，按「保宽、重算高」处理，图像本身不拉伸")
    if width_cm is None and height_cm is None:
        notes.append("未指定 --width-cm，保留文档现有宽度，仅按新图纵横比重算高度")
    return new_cx, new_cy, notes


# ---------------------------------------------------------------------------
# 替换与校验
# ---------------------------------------------------------------------------


def build_replacement(parts: dict[str, bytes], info: dict,
                      image_bytes: bytes, new_cx: int, new_cy: int) -> dict[str, bytes]:
    """在内存里生成新包部件（不落盘）。"""
    doc = E.fromstring(parts["word/document.xml"])
    texts_before = doc.xpath("//w:t/text()", namespaces=NS)
    drawings = find_drawings(doc)
    target = drawings[info["index"]]
    extents = target.xpath("./wp:extent | .//a:xfrm/a:ext", namespaces=NS)
    if not extents:
        raise EmbedError("目标图片容器里没有 wp:extent / a:xfrm/a:ext，无法同步尺寸")
    for ext in extents:
        ext.set("cx", str(new_cx))
        ext.set("cy", str(new_cy))
    texts_after = doc.xpath("//w:t/text()", namespaces=NS)
    if texts_before != texts_after:
        raise EmbedError("内部错误：改尺寸时动到了正文文本，已中止")
    new_parts = dict(parts)
    new_parts["word/document.xml"] = E.tostring(
        doc, encoding="UTF-8", xml_declaration=True, standalone=True
    )
    new_parts[info["media"]] = image_bytes
    return new_parts


def verify_in_memory(parts: dict[str, bytes], new_parts: dict[str, bytes], info: dict,
                     image_bytes: bytes, new_cx: int, new_cy: int) -> None:
    """在内存里校验：文本不变、只改两个部件、新图与尺寸正确。"""
    # 1) 文本序列完全一致
    old_doc = E.fromstring(parts["word/document.xml"])
    new_doc = E.fromstring(new_parts["word/document.xml"])
    old_texts = old_doc.xpath("//w:t/text()", namespaces=NS)
    new_texts = new_doc.xpath("//w:t/text()", namespaces=NS)
    if old_texts != new_texts:
        raise EmbedError(f"正文文本发生变化（{len(old_texts)} 段 -> {len(new_texts)} 段），已中止")

    # 2) 仅 document.xml 与目标 media 变化
    changed = {name for name, blob in new_parts.items() if parts.get(name) != blob}
    expected = {"word/document.xml", info["media"]}
    extra = changed - expected
    if extra:
        raise EmbedError(f"除 document.xml 与 {info['media']} 外还有部件被改动：{sorted(extra)}")
    if "word/document.xml" not in changed:
        raise EmbedError("document.xml 未发生变化，尺寸同步失败")

    # 3) media 字节与 cx/cy 正确
    if new_parts[info["media"]] != image_bytes:
        raise EmbedError("新 media 字节与输入图片不一致")
    drawings = find_drawings(new_doc)
    target = drawings[info["index"]]
    extents = target.xpath("./wp:extent", namespaces=NS)
    if int(extents[0].get("cx")) != new_cx or int(extents[0].get("cy")) != new_cy:
        raise EmbedError("wp:extent 的 cx/cy 未按预期写入")


def write_parts(docx: Path, order: list[str], new_parts: dict[str, bytes],
                destination: Path | None = None, backup: bool = False) -> Path:
    """先写 .tmp，重新打开校验后 replace 到目标。"""
    target = destination or docx
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(target.name + ".tmp")
    try:
        with ZipFile(tmp, "w", ZIP_DEFLATED) as z:
            for name in order:
                z.writestr(name, new_parts[name])
        with ZipFile(tmp) as z:
            names = z.namelist()
            if names != order:
                raise EmbedError(f"重打包后部件顺序/数量变化（{len(names)} -> {len(order)}）")
            bad = z.testzip()
            if bad is not None:
                raise EmbedError(f"重打包后 zip 校验失败：{bad}")
            for name in ("word/document.xml",):
                z.read(name)
        if backup and destination is None:
            backup_path = docx.with_name(docx.stem + ".bak" + docx.suffix)
            shutil.copy2(docx, backup_path)
            print(f"已备份原文件：{backup_path}")
        tmp.replace(target)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    return target


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="把 PNG 替换进 docx 内嵌图：只改图字节与 cx/cy，正文文本零改动。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python embed_figure.py --docx A.docx --image fig.png --dry-run\n"
               "  python embed_figure.py --docx A.docx --image fig.png --index 2 --width-cm 16\n"
               "  python embed_figure.py --docx A.docx --image fig.png --out B.docx",
    )
    p.add_argument("--docx", required=True, type=Path, help="目标 .docx（默认原地替换）")
    p.add_argument("--image", required=True, type=Path, help="新图 PNG")
    p.add_argument("--width-cm", type=float, default=None,
                   help="新显示宽度（cm），默认保留文档现有宽度；给定时高度按图像纵横比重算")
    p.add_argument("--height-cm", type=float, default=None,
                   help="新显示高度（cm），一般不用；与 --width-cm 同用会改变显示纵横比")
    p.add_argument("--index", type=int, default=None,
                   help="第几张图（1 起，默认 1）")
    p.add_argument("--occurrence", type=int, default=None,
                   help="第几张图（0 起，兼容旧脚本；与 --index 二选一）")
    p.add_argument("--dry-run", action="store_true", help="只报告将改哪张图、尺寸如何，不写任何文件")
    p.add_argument("--out", type=Path, default=None, help="写到新文件而不是原地替换")
    p.add_argument("--backup", action="store_true", help="原地替换前备份为 <stem>.bak.docx")
    return p


def main(argv=None) -> int:
    args = build_arg_parser().parse_args(argv)

    try:
        if args.index is not None and args.occurrence is not None:
            raise EmbedError("--index 与 --occurrence 只能给一个")
        index = args.index if args.index is not None else (args.occurrence + 1 if args.occurrence is not None else 1)

        image_path = args.image
        img_w, img_h = read_png_size(image_path)
        image_bytes = image_path.read_bytes()
        order, parts = parse_parts(args.docx)
        doc = E.fromstring(parts["word/document.xml"])
        infos = collect_drawings(doc, parts)

        print(f"docx: {args.docx}")
        print(f"图片: {image_path.name}  {img_w}x{img_h}px  纵横比 {img_w / img_h:.3f}")
        print(f"文档内嵌图共 {len(infos)} 张：")
        for info in infos:
            print(f"  [{info['index'] + 1}] {info['tag']:6s} {info['media']:28s} "
                  f"{fmt_cm(info['cx'])} x {fmt_cm(info['cy'])}  name={info['name']!r}")
        if not (1 <= index <= len(infos)):
            raise EmbedError(f"--index={index} 越界，文档只有 {len(infos)} 张内嵌图（1..{len(infos)}）")
        info = infos[index - 1]

        shared = sum(1 for other in infos if other["media"] == info["media"])
        if shared > 1:
            print(f"警告：{info['media']} 被 {shared} 张图共用，替换后其余引用同一 media 的图也会变成新图")

        new_cx, new_cy, notes = choose_size(
            info["cx"], info["cy"], img_w, img_h, args.width_cm, args.height_cm
        )
        print()
        print(f"将替换第 {index} 张：{info['media']}（{info['rid']}，{info['tag']}，name={info['name']!r}）")
        print(f"  当前显示尺寸: {fmt_cm(info['cx'])} x {fmt_cm(info['cy'])}  "
              f"({info['cx']} x {info['cy']} EMU)")
        print(f"  新显示尺寸  : {fmt_cm(new_cx)} x {fmt_cm(new_cy)}  ({new_cx} x {new_cy} EMU)")
        for note in notes:
            print(f"  {note}")

        new_parts = build_replacement(parts, info, image_bytes, new_cx, new_cy)
        verify_in_memory(parts, new_parts, info, image_bytes, new_cx, new_cy)
        changed = [name for name in new_parts if parts.get(name) != new_parts[name]]
        print(f"  校验通过：正文 //w:t {len(doc.xpath('//w:t/text()', namespaces=NS))} 段文本不变；"
              f"将变化的部件：{', '.join(changed)}")

        if args.dry_run:
            print("[dry-run] 只做只读检查与内存校验，未写入任何文件。")
            return 0

        if not changed:
            print("内容无变化，跳过写入。")
            return 0

        written = write_parts(args.docx, order, new_parts, args.out, args.backup)
        print(f"已写入：{written}")
        return 0
    except EmbedError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - 兜底，保留堆栈便于排查
        print(f"未预期的错误：{exc!r}", file=sys.stderr)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
