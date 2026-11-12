# -*- coding: utf-8 -*-
"""figlib.py —— 学术示意图绘制原语库（Pillow 绘制 + 同步生成 SVG 双产物）

用途：为中文期刊论文绘制黑白/灰度机制示意图。一次调用同时得到两份产物：
  (1) 可编辑的 SVG 矢量源（投稿/返修时改文字用）；(2) 600dpi 的 PNG 嵌入件。
设计约束：仅依赖 Python 标准库 + Pillow；逻辑坐标系统一（画布按 1200x770 逻辑像素
设计，scale=3 超采样后输出 3600x2310 像素、dpi=600），Pillow 与 SVG 共用同一组坐标
与颜色，改一个数两边同步。

三行用法：
    from figlib import Canvas
    c = Canvas(1200, 770)                              # 逻辑像素；输出 3600x2310 @600dpi
    c.rect(30, 20, 400, 90); c.text(230, 65, "示例标题", 24, bold=True)
    c.save_both("examples/out", "figure1_demo")        # -> figure1_demo.png / figure1_demo.svg

只 import 不画图不会有任何副作用（模块级不创建画布、不读字体、不写文件）；布局代码
放在你自己的脚本里，从 figlib 取原语即可。更细的作图规范见 references/figure-specs.md。

字体约定：bold=True 用黑体（标题），bold=False 用宋体（正文）。自动探测
simhei.ttf / simsun.ttc / msyh.ttc，可用环境变量 FIGLIB_FONT_BOLD /
FIGLIB_FONT_REGULAR 指定 .ttf/.ttc 路径；探测不到时回退并告警（见 _resolve_font）。
"""

from __future__ import annotations

import math
import re
import sys
import warnings
from html import escape
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as exc:  # pragma: no cover - 环境问题直接说清楚
    raise SystemExit(
        "figlib 需要 Pillow：请在论文排版环境中安装 Pillow（pip install Pillow），"
        f"当前导入失败：{exc}"
    )

__version__ = "1.0"
__all__ = ["Canvas", "black_white", "font_warnings"]

# ---------------------------------------------------------------------------
# 字体探测：优先 Windows 常见中文字体，其次 macOS / Linux，最后回退默认字体
# ---------------------------------------------------------------------------

_FONT_WARNINGS: list[str] = []          # 探测过程中产生的告警（供验收脚本检查）
_RESOLVED: dict[bool, tuple] = {}       # bold -> (path|None, ttc_index, family, fallback)
_WARNED: set[str] = set()
_FONT_CACHE: dict[tuple[bool, int], "ImageFont.FreeTypeFont"] = {}

# (路径, .ttc 内部索引, SVG font-family 名, 是否回退字体)
_CANDIDATES: dict[bool, list[tuple[str, int, str, bool]]] = {
    True: [
        (r"C:\Windows\Fonts\simhei.ttf", 0, "SimHei", False),
        (r"C:\Windows\Fonts\msyhbd.ttc", 0, "Microsoft YaHei", False),
        (r"C:\Windows\Fonts\msyh.ttc", 0, "Microsoft YaHei", False),
        ("/System/Library/Fonts/PingFang.ttc", 0, "PingFang SC", False),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 0, "Noto Sans CJK SC", False),
        ("/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc", 0, "Noto Sans CJK SC", False),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 0, "DejaVu Sans", True),
    ],
    False: [
        (r"C:\Windows\Fonts\simsun.ttc", 0, "SimSun", False),
        (r"C:\Windows\Fonts\simfang.ttf", 0, "FangSong", False),
        (r"C:\Windows\Fonts\msyh.ttc", 0, "Microsoft YaHei", False),
        ("/System/Library/Fonts/PingFang.ttc", 0, "PingFang SC", False),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0, "Noto Sans CJK SC", False),
        ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0, "Noto Sans CJK SC", False),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 0, "DejaVu Sans", True),
    ],
}


def _warn_once(key: str, message: str) -> None:
    """同一类告警只提示一次，避免逐字符刷屏。"""
    if key in _WARNED:
        return
    _WARNED.add(key)
    _FONT_WARNINGS.append(message)
    warnings.warn(message, RuntimeWarning, stacklevel=3)


def _resolve_font(bold: bool) -> tuple:
    """探测字体；返回 (路径, ttc 索引, SVG 字体族名, 是否回退)。结果缓存。"""
    if bold in _RESOLVED:
        return _RESOLVED[bold]

    env_key = "FIGLIB_FONT_BOLD" if bold else "FIGLIB_FONT_REGULAR"
    import os

    env_path = os.environ.get(env_key)
    if env_path:
        p = Path(env_path).expanduser()
        if p.is_file():
            resolved = (str(p), 0, p.stem, False)
        else:
            _warn_once(f"env-{bold}", f"figlib 警告：环境变量 {env_key}={env_path} 指向的文件不存在，改用自动探测。")
            resolved = None
    else:
        resolved = None

    if resolved is None:
        for path, index, family, fallback in _CANDIDATES[bold]:
            if Path(path).is_file():
                resolved = (path, index, family, fallback)
                if fallback:
                    _warn_once(
                        f"fallback-{bold}",
                        f"figlib 警告：未找到可用的中文字体，回退到 {family}（缺少中文字形，"
                        f"中文可能显示为方框）。请安装 SimHei/SimSun 或设置 {env_key}。",
                    )
                break

    if resolved is None:
        _warn_once(
            f"none-{bold}",
            "figlib 警告：未找到任何可用字体，回退 Pillow 内置默认字体（不含中文字形，"
            "中文会显示为方框）。请设置 FIGLIB_FONT_REGULAR / FIGLIB_FONT_BOLD 指向中文字体。",
        )
        resolved = (None, 0, "sans-serif", True)

    _RESOLVED[bold] = resolved
    return resolved


def _load_font(bold: bool, size_px: int) -> "ImageFont.FreeTypeFont":
    """按逻辑字号取字体对象（真实字号 = size_px * scale 由调用方保证）。"""
    key = (bold, size_px)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    path, index, _family, _fallback = _resolve_font(bold)
    if path is None:
        try:  # Pillow >= 10.1 支持带字号的默认字体
            font = ImageFont.load_default(size=size_px)
        except TypeError:  # pragma: no cover - 老版本 Pillow
            font = ImageFont.load_default()
    else:
        try:
            font = ImageFont.truetype(path, size_px, index=index)
        except OSError as exc:  # 字体文件损坏 / 索引越界
            _warn_once(f"load-{bold}", f"figlib 警告：加载字体 {path} 失败（{exc}），回退默认字体。")
            try:
                font = ImageFont.load_default(size=size_px)
            except TypeError:  # pragma: no cover
                font = ImageFont.load_default()
    _FONT_CACHE[key] = font
    return font


def font_warnings() -> list[str]:
    """返回字体探测产生的告警列表（验收脚本可据此判断是否发生了回退）。"""
    return list(_FONT_WARNINGS)


# ---------------------------------------------------------------------------
# 颜色 / 虚线 / 文字锚点小工具
# ---------------------------------------------------------------------------

_HEX_RE = re.compile(r"#(?:[0-9a-fA-F]{3}){1,2}\b")
_ANCHOR_ALIASES = {
    "middle": "mm", "center": "mm", "left": "lm", "right": "rm",
    "top": "mt", "bottom": "mb",
}
_VALID_ANCHORS = {
    "la", "lm", "ls", "lt", "ma", "mm", "ms", "mt",
    "ra", "rm", "rs", "rt", "ld", "md", "rd",
}
_V_ANCHOR_TO_SVG_BASELINE = {"a": "central", "m": "central", "t": "hanging", "s": "alphabetic", "b": "alphabetic", "d": "alphabetic"}


def _anchor(value: str) -> str:
    """把 middle/left 之类的别名规整成 Pillow 锚点（两字符）。"""
    value = (value or "mm").strip()
    value = _ANCHOR_ALIASES.get(value, value)
    if value not in _VALID_ANCHORS:
        raise ValueError(
            f"不支持的锚点 {value!r}；可用 'mm' / 'lm' / 'rm' / 'mt' / 'mb' 或 middle/left/right/top/bottom"
        )
    return value


def _check_color(color: str) -> str:
    """颜色统一为 #rrggbb / 具名色；其余格式直接拒绝，避免 Pillow 与 SVG 解释不一致。"""
    color = str(color).strip()
    if color.startswith("#"):
        m = _HEX_RE.match(color)
        if not m:
            raise ValueError(f"颜色格式非法：{color!r}，请用 #rrggbb")
        val = color[1:]
        if len(val) == 3:
            val = "".join(ch * 2 for ch in val)
        return "#" + val.lower()
    if color.lower() == "white":
        return "#ffffff"
    if color.lower() == "black":
        return "#000000"
    if color.lower() == "none":
        return "none"
    raise ValueError(f"颜色格式非法：{color!r}，请用 #rrggbb 或 white/black/none")


def _crop_transparent(value: str) -> str | None:
    return None if value == "none" else value


# ---------------------------------------------------------------------------
# 画布
# ---------------------------------------------------------------------------


class Canvas:
    """逻辑坐标画布：Pillow 光栅 + 同步累积的 SVG 元素列表。

    所有原语坐标都用「逻辑像素」（画布宽 = width，例如 1200），内部按 scale 放大
    光栅绘制；SVG 写入逻辑坐标，viewBox 即 width x height。这样一份布局代码同时
    产出 600dpi 的 PNG（width*scale 像素）与可编辑 SVG。

    参数：
        width, height : 逻辑画布尺寸（px）。期刊通栏图建议 1200x770。
        scale         : 超采样倍数，默认 3（1200 -> 3600px 光栅）。
        dpi           : PNG 元数据 dpi，默认 600。
        width_mm      : SVG 的物理宽度（mm），默认 160（四刊通栏宽）。
        background    : 画布底色，默认 white。
    """

    def __init__(self, width: int, height: int, scale: int = 3, dpi: int = 600,
                 width_mm: float = 160.0, background: str = "#ffffff"):
        self.width = int(width)
        self.height = int(height)
        if self.width <= 0 or self.height <= 0:
            raise ValueError("画布宽高必须为正数")
        self.scale = max(1, int(scale))
        self.dpi = int(dpi)
        self.width_mm = float(width_mm)
        self.background = _check_color(background)
        self.image = Image.new("RGB", (self.width * self.scale, self.height * self.scale), self.background)
        self.draw = ImageDraw.Draw(self.image)
        height_mm = self.width_mm * self.height / self.width
        self._svg: list[str] = [
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.width_mm:g}mm" '
            f'height="{height_mm:.3f}mm" viewBox="0 0 {self.width} {self.height}">',
            f'<rect width="{self.width}" height="{self.height}" fill="{self.background}"/>',
        ]

    # -- 内部换算 ---------------------------------------------------------
    def _px(self, value: float) -> int:
        return int(round(value * self.scale))

    def _stroke(self, stroke_width: float) -> int:
        return max(1, int(round(stroke_width * self.scale)))

    @staticmethod
    def _dash_pattern(dash) -> tuple[float, float]:
        if dash is True:
            return 6.0, 5.0
        if isinstance(dash, (tuple, list)) and len(dash) == 2:
            on, off = float(dash[0]), float(dash[1])
            if on <= 0 or off < 0:
                raise ValueError(f"虚线参数非法：{dash!r}")
            return on, off
        raise ValueError(f"虚线参数非法：{dash!r}；用 False / True / (on, off)")

    def _draw_segment(self, p0: tuple[float, float], p1: tuple[float, float],
                      color: str, width: int, dash) -> None:
        x0, y0 = p0
        x1, y1 = p1
        if dash:
            on, off = self._dash_pattern(dash)
            on_s, off_s = on * self.scale, off * self.scale
            length = math.hypot(x1 - x0, y1 - y0)
            if length == 0:
                return
            ux, uy = (x1 - x0) / length, (y1 - y0) / length
            start = 0.0
            while start < length:
                end = min(start + on_s, length)
                self.draw.line(
                    [(x0 + ux * start, y0 + uy * start), (x0 + ux * end, y0 + uy * end)],
                    fill=color, width=width,
                )
                start = end + off_s
        else:
            self.draw.line([(x0, y0), (x1, y1)], fill=color, width=width)

    # -- 原语 -------------------------------------------------------------
    def rect(self, x: float, y: float, w: float, h: float, fill: str = "#ffffff",
             stroke: str = "#333333", stroke_width: float = 1.3, dash=False,
             radius: float = 0) -> None:
        """矩形。fill='none' 为透明；dash=True/(on,off) 为虚线边；radius>0 为圆角（不与虚线同用）。"""
        fill = _check_color(fill)
        stroke = _check_color(stroke)
        if w < 0 or h < 0:
            x, y, w, h = x + w, y + h, abs(w), abs(h)
        if radius and dash:
            raise ValueError("rect 的圆角与虚线边不能同时使用")
        x0, y0 = self._px(x), self._px(y)
        x1, y1 = self._px(x + w), self._px(y + h)
        sw_px = self._stroke(stroke_width)
        if radius:
            self.draw.rounded_rectangle(
                (x0, y0, x1, y1), radius=self._px(radius),
                fill=_crop_transparent(fill), outline=_crop_transparent(stroke), width=sw_px,
            )
            self._svg.append(
                f'<rect x="{x:g}" y="{y:g}" width="{w:g}" height="{h:g}" rx="{radius:g}" '
                f'fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width:g}"/>'
            )
        elif dash:
            self._draw_segment((x0, y0), (x1, y0), stroke, sw_px, dash)
            self._draw_segment((x1, y0), (x1, y1), stroke, sw_px, dash)
            self._draw_segment((x1, y1), (x0, y1), stroke, sw_px, dash)
            self._draw_segment((x0, y1), (x0, y0), stroke, sw_px, dash)
            on, off = self._dash_pattern(dash)
            self._svg.append(
                f'<rect x="{x:g}" y="{y:g}" width="{w:g}" height="{h:g}" fill="{fill}" '
                f'stroke="{stroke}" stroke-width="{stroke_width:g}" '
                f'stroke-dasharray="{on:g} {off:g}"/>'
            )
        else:
            self.draw.rectangle(
                (x0, y0, x1, y1), fill=_crop_transparent(fill),
                outline=_crop_transparent(stroke), width=sw_px,
            )
            self._svg.append(
                f'<rect x="{x:g}" y="{y:g}" width="{w:g}" height="{h:g}" '
                f'fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width:g}"/>'
            )

    def text(self, x: float, y: float, s: str, size: float = 21, bold: bool = False,
             color: str = "#181818", anchor: str = "mm", line_spacing: float = 1.35) -> tuple[float, float]:
        """单行/多行文字。尺寸为逻辑像素字号；返回整块文字的 (宽, 高) 逻辑尺寸。

        多行用 \\n 分隔，整体按锚点对齐，行距 = size * line_spacing。
        bold=True 用黑体（标题），False 用宋体（正文）。
        """
        anchor = _anchor(anchor)
        color = _check_color(color)
        lines = str(s).split("\n")
        size_px = max(1, int(round(size * self.scale)))
        font = _load_font(bold, size_px)
        lh = size * line_spacing if len(lines) > 1 else 0.0
        block_w = 0.0
        for line in lines:
            try:
                block_w = max(block_w, float(font.getlength(line)) / self.scale)
            except AttributeError:  # pragma: no cover - 极老 Pillow
                bbox = font.getbbox(line)
                block_w = max(block_w, float(bbox[2] - bbox[0]) / self.scale)
        ascent, descent = font.getmetrics()
        line_h = (ascent + descent) / self.scale
        block_h = line_h + lh * (len(lines) - 1)

        # 多行时：垂直居中锚点（mm/lm/rm）整块居中；其余锚点从 y 起向下排，
        # 避免 "lm" 这类锚点把整块文字顶到 y 上方。
        centered = anchor[1] == "m"
        for i, line in enumerate(lines):
            if centered:
                ly = y + (i - (len(lines) - 1) / 2.0) * lh
            else:
                ly = y + i * lh
            self.draw.text((x * self.scale, ly * self.scale), line, font=font,
                           fill=color, anchor=anchor)
            h_anchor, v_anchor = anchor[0], anchor[1]
            svg_anchor = {"l": "start", "m": "middle", "r": "end"}[h_anchor]
            baseline = _V_ANCHOR_TO_SVG_BASELINE[v_anchor]
            self._svg.append(
                f'<text x="{x:g}" y="{ly:g}" font-family="{escape(_svg_family(bold), quote=True)}" '
                f'font-size="{size:g}" text-anchor="{svg_anchor}" '
                f'dominant-baseline="{baseline}" fill="{color}">{escape(line)}</text>'
            )
        return block_w, block_h

    def text_size(self, s: str, size: float = 21, bold: bool = False) -> tuple[float, float]:
        """量文字尺寸，用于手工排版时算间距（不绘制）。"""
        size_px = max(1, int(round(size * self.scale)))
        font = _load_font(bold, size_px)
        width = 0.0
        for line in str(s).split("\n"):
            width = max(width, float(font.getlength(line)) / self.scale)
        ascent, descent = font.getmetrics()
        return width, (ascent + descent) / self.scale

    def line(self, points, arrow: bool = False, dash=False, width: float = 1.4,
             color: str = "#333333", arrow_size: float = 9.0) -> None:
        """折线；arrow=True 在最后一段末端画实心三角箭头；dash=True/(on,off) 为虚线。"""
        pts = [(float(px), float(py)) for px, py in points]
        if len(pts) < 2:
            raise ValueError("line 至少需要两个点")
        color = _check_color(color)
        sw_px = self._stroke(width)
        for p0, p1 in zip(pts, pts[1:]):
            self._draw_segment(
                (self._px(p0[0]), self._px(p0[1])), (self._px(p1[0]), self._px(p1[1])),
                color, sw_px, dash,
            )
        self._svg.append(
            '<polyline points="' + " ".join(f"{px:g},{py:g}" for px, py in pts) + '" fill="none" '
            f'stroke="{color}" stroke-width="{width:g}"'
            + (f' stroke-dasharray="{self._dash_pattern(dash)[0]:g} {self._dash_pattern(dash)[1]:g}"' if dash else "")
            + "/>"
        )
        if arrow:
            (x0, y0), (x1, y1) = pts[-2], pts[-1]
            angle = math.atan2(y1 - y0, x1 - x0)
            tip = (x1, y1)
            tail = [
                (x1 - arrow_size * math.cos(angle - 0.42), y1 - arrow_size * math.sin(angle - 0.42)),
                (x1 - arrow_size * math.cos(angle + 0.42), y1 - arrow_size * math.sin(angle + 0.42)),
            ]
            self.draw.polygon(
                [(self._px(px), self._px(py)) for px, py in (tip, tail[0], tail[1])], fill=color
            )
            self._svg.append(
                '<polygon points="' + " ".join(f"{px:.2f},{py:.2f}" for px, py in (tip, *tail))
                + f'" fill="{color}"/>'
            )

    def polygon(self, points, fill: str = "#ffffff", stroke: str = "#333333",
                stroke_width: float = 1.3) -> None:
        """闭合多边形（实线边；虚线边请用逐段 line 自行拼）。"""
        pts = [(float(px), float(py)) for px, py in points]
        if len(pts) < 3:
            raise ValueError("polygon 至少需要三个点")
        fill = _check_color(fill)
        stroke = _check_color(stroke)
        self.draw.polygon(
            [(self._px(px), self._px(py)) for px, py in pts],
            fill=_crop_transparent(fill), outline=_crop_transparent(stroke),
            width=self._stroke(stroke_width),
        )
        self._svg.append(
            '<polygon points="' + " ".join(f"{px:g},{py:g}" for px, py in pts)
            + f'" fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width:g}"/>'
        )

    # -- 语义化辅助 -------------------------------------------------------
    def box_with_header(self, x: float, y: float, w: float, h: float, title: str,
                        header_h: float = 40, title_size: float = 23,
                        body=None, body_size: float = 21,
                        fill: str = "#ffffff", header_fill: str = "#f1f1f1",
                        stroke: str = "#555555", stroke_width: float = 1.1,
                        header_dash=False, body_color: str = "#333333") -> None:
        """带标题栏的方框：外框 + 顶部浅色标题栏 + 可选正文行（居中）。"""
        self.rect(x, y, w, h, fill=fill, stroke=stroke, stroke_width=stroke_width)
        self.rect(x, y, w, header_h, fill=header_fill, stroke=stroke,
                  stroke_width=stroke_width, dash=header_dash)
        self.text(x + w / 2, y + header_h / 2, title, title_size, True, "#181818", "mm")
        if body:
            rows = [body] if isinstance(body, str) else list(body)
            spacing = body_size * 1.45
            avail_top = y + header_h
            avail_h = h - header_h
            first = avail_top + avail_h / 2 - spacing * (len(rows) - 1) / 2
            for i, row in enumerate(rows):
                self.text(x + w / 2, first + i * spacing, row, body_size, False, body_color, "mm")

    def node(self, x: float, y: float, title: str, lines=(), w: float = 280,
             h: float | None = None, title_h: float = 39, title_size: float = 23,
             line_size: float = 21, fill: str = "#ffffff", header_fill: str = "#f1f1f1",
             stroke: str = "#555555", stroke_width: float = 1.1) -> float:
        """结构图节点：标题栏 + 居中的若干行正文；h 省略时按行数自动定高。返回实际高度。"""
        rows = [lines] if isinstance(lines, str) else list(lines)
        spacing = round(line_size * 1.45)
        if h is None:
            h = title_h + 26 + spacing * max(len(rows) - 1, 0) + 30
        self.box_with_header(x, y, w, h, title, header_h=title_h, title_size=title_size,
                             body=None, fill=fill, header_fill=header_fill,
                             stroke=stroke, stroke_width=stroke_width)
        first = y + title_h + 26
        for i, row in enumerate(rows):
            self.text(x + w / 2, first + i * spacing, row, line_size, False, "#333333", "mm")
        return h

    def label_strip(self, x: float, y: float, w: float, h: float, title: str, subtitle: str = "",
                    fill: str = "#ededed", stroke: str = "#666666", stroke_width: float = 1.1,
                    title_size: float = 26, sub_size: float = 20) -> None:
        """左侧/旁侧的机制标签条：粗体机制名 + 小字说明（上下两行）。"""
        self.rect(x, y, w, h, fill=fill, stroke=stroke, stroke_width=stroke_width)
        offset = 0 if not subtitle else 20
        self.text(x + w / 2, y + h / 2 - offset, title, title_size, True, "#181818", "mm")
        if subtitle:
            self.text(x + w / 2, y + h / 2 + 22, subtitle, sub_size, False, "#333333", "mm")

    def legend_line(self, x: float, y: float, label: str, length: float = 48,
                    dash=False, arrow: bool = False, width: float = 1.4,
                    label_size: float = 20, color: str = "#333333") -> float:
        """底部图例的一条示例线 + 文字说明；返回文字右边界（方便下一条接排）。"""
        self.line([(x, y), (x + length, y)], arrow=arrow, dash=dash, width=width, color=color)
        tw, _th = self.text(x + length + 12, y, label, label_size, False, color, "lm")
        return x + length + 12 + tw

    # -- 输出 -------------------------------------------------------------
    def to_svg(self) -> str:
        """返回完整 SVG 文本（不落盘）。"""
        return "\n".join([*self._svg, "</svg>"])

    def save_png(self, path) -> Path:
        """保存 PNG，写入 dpi 元数据。"""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.image.save(path, dpi=(self.dpi, self.dpi))
        return path

    def save_svg(self, path) -> Path:
        """保存 SVG（可编辑源）。"""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.to_svg(), encoding="utf-8")
        return path

    def save_both(self, dirname, stem: str) -> tuple[Path, Path]:
        """同时保存 <stem>.png 与 <stem>.svg，返回 (png, svg)。"""
        dirname = Path(dirname)
        dirname.mkdir(parents=True, exist_ok=True)
        png = self.save_png(dirname / f"{stem}.png")
        svg = self.save_svg(dirname / f"{stem}.svg")
        return png, svg


def _svg_family(bold: bool) -> str:
    _path, _index, family, _fallback = _resolve_font(bold)
    return f"{family}, sans-serif" if bold else f"{family}, serif"


# ---------------------------------------------------------------------------
# 黑白化：按亮度四档重写颜色（245 / 180 / 85 / 25），投稿纯黑白件用
# ---------------------------------------------------------------------------

_BW_LEVELS = (245, 180, 85, 25)
_BW_THRESHOLDS = (210, 140, 90)


def _to_gray(hex_color: str) -> str:
    value = hex_color.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    r, g, b = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    lum = round(0.299 * r + 0.587 * g + 0.114 * b)
    if lum > _BW_THRESHOLDS[0]:
        level = _BW_LEVELS[0]
    elif lum > _BW_THRESHOLDS[1]:
        level = _BW_LEVELS[1]
    elif lum > _BW_THRESHOLDS[2]:
        level = _BW_LEVELS[2]
    else:
        level = _BW_LEVELS[3]
    return "#" + f"{level:02x}" * 3


def black_white(svg_path, out_path) -> Path:
    """把 SVG 里的彩色按亮度压成四档灰（245/180/85/25），产出纯黑白 SVG。

    与 06_处理脚本与临时数据/render_figures.cjs 的映射保持一致：
    亮度 >210 -> 245；>140 -> 180；>90 -> 85；其余 -> 25。
    PNG 的黑白化仍需栅格化器（如 sharp/Inkscape）；默认只交付 SVG 源，
    需要黑白 PNG 时用同一 SVG 重新导出即可。
    """
    svg_path, out_path = Path(svg_path), Path(out_path)
    if not svg_path.is_file():
        raise FileNotFoundError(f"SVG 不存在：{svg_path}")
    text = svg_path.read_text(encoding="utf-8")
    # 只替换属性里的 #hex，避免误伤 id / 文本内容（本库输出的颜色都在属性里）
    text = _HEX_RE.sub(lambda m: _to_gray(m.group(0)), text)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path
