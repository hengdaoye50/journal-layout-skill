# -*- coding: utf-8 -*-
"""figure1_demo.py —— 用 figlib 画一张两机制并列的机制示意图（《担责换权》图1 的可运行示例）

演示内容：
  1) Canvas(1200, 770, scale=3, dpi=600)：逻辑 1200x770 画布 -> 3600x2310 px / 600dpi PNG；
  2) 同一份布局代码同步产出可编辑 SVG（160mm 通栏宽，viewBox 1200x770）；
  3) 原语 rect / text / line（实线、虚线、箭头）与语义辅助 box_with_header / legend_line；
  4) black_white() 产出纯黑白投稿版 SVG（亮度四档 245/180/85/25）。

运行（任意目录均可，路径按本文件定位）：
    python examples/figure1_demo.py
产物：examples/out/figure1_demo.png / figure1_demo.svg / figure1_demo_bw.svg
"""

import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from figlib import Canvas, black_white, font_warnings  # noqa: E402

OUT = Path(__file__).resolve().parent / "out"
W, H = 1200, 770


def draw(c: Canvas) -> None:
    """图1 简化版：问题背景 -> 两机制并列 -> 条件支撑 -> 共同结果。坐标即逻辑像素。"""
    # ---- 1 问题背景条：只作背景，不画成自动触发改革行动的箭头 ----
    c.rect(30, 20, 1140, 83, fill="#f5f5f5", stroke="#c0c0c0", stroke_width=0.8)
    c.text(115, 61, "问题起点", 22, True)
    c.text(320, 61, "科层秘密", 24)
    c.line([(405, 61), (470, 61)], arrow=True)
    c.text(610, 61, "不确定的问责风险", 24)
    c.line([(752, 61), (817, 61)], arrow=True)
    c.text(985, 61, "跨部门合作障碍", 24)

    # ---- 2 两个机制并列：等重视觉权重，但因果角色不同 ----
    panels = [
        (30, "科层解密", "信息结构调整：问责风险的源头降低"),
        (760, "担责换权", "制度安排调整：成本收益结构重构"),
    ]
    for x, title, subtitle in panels:
        c.rect(x, 137, 410, 449, fill="#ffffff", stroke="#7b7b7b", stroke_width=1)
        c.rect(x, 137, 410, 86, fill="#ededed", stroke="#7b7b7b", stroke_width=1)
        c.text(x + 205, 163, title, 29, True)
        c.text(x + 205, 201, subtitle, 21)

    left_rows = [
        ("业务信息的制度性共享", "操作标准、判断规则、处理方式"),
        ("规则显性化与差错边界清晰", "部门独占转为跨部门可获取"),
        ("问责风险的源头降低", "风险可预见，量级可控制"),
    ]
    right_rows = [
        ("承接特定环节的问责风险", "换取相应业务权限"),
        ("启动阶段：降低权限让渡成本", "合作方接受权限调整"),
        ("持续阶段：业务依赖与知识共享", "由被动服从转向主动配合"),
    ]
    for x, rows in ((50, left_rows), (780, right_rows)):
        for i, (headline, detail) in enumerate(rows):
            y = 245 + i * 109
            c.rect(x, y, 370, 80, stroke="#aaaaaa", stroke_width=0.8)
            c.text(x + 185, y + 25, headline, 22, True)
            c.text(x + 185, y + 56, detail, 21)
            if i < 2:
                c.line([(x + 185, y + 80), (x + 185, y + 109)], arrow=True)

    # ---- 3 虚线条件支撑：指向担责换权整个机制区域，不强配为三个时序阶段 ----
    c.text(600, 271, "基础条件", 24, True)
    c.line([(440, 301), (760, 301)], arrow=True, dash=True)
    c.text(600, 354, "决策：可评估风险", 22)
    c.text(600, 400, "承担：可承接风险", 22)
    c.text(600, 446, "落地：可开展业务", 22)
    c.text(600, 497, "三重支撑", 21)

    # ---- 4 共同结果：括号汇聚，不表示任一机制单独充分 ----
    c.line([(235, 586), (235, 630), (965, 630), (965, 586)])
    c.line([(600, 630), (600, 664)], arrow=True)
    c.box_with_header(365, 664, 470, 57, "跨部门改革的实质推进", header_h=57, title_size=26,
                      fill="#ededed", header_fill="#ededed", stroke="#555555", stroke_width=1.2)

    # ---- 5 图例：实线=过程，虚线=条件支撑（图例在图内底部，题注在图外下方） ----
    c.legend_line(330, 751, "实线：机制过程及共同作用", length=48, dash=False, label_size=20)
    c.legend_line(690, 751, "虚线：条件支撑", length=48, dash=True, label_size=20)


def main() -> int:
    c = Canvas(W, H, scale=3, dpi=600)
    draw(c)
    png, svg = c.save_both(OUT, "figure1_demo")
    bw_svg = black_white(svg, OUT / "figure1_demo_bw.svg")

    # ---- 自检：PNG 宽度、dpi、SVG 物理宽度、黑白化结果 ----
    from PIL import Image

    with Image.open(png) as im:
        width, height = im.width, im.height
        dpi = tuple(round(float(v)) for v in im.info.get("dpi", (0, 0)))
    svg_head = svg.read_text(encoding="utf-8").splitlines()[0]
    svg_width = re.search(r'width="([^"]+)"', svg_head).group(1)

    bw_text = bw_svg.read_text(encoding="utf-8")
    colored = [m for m in re.findall(r"#[0-9a-fA-F]{6}", bw_text)
               if not (m[1:3] == m[3:5] == m[5:7])]

    print(f"PNG : {png}")
    print(f"SVG : {svg}")
    print(f"黑白: {bw_svg}")
    print(f"PNG 像素 {width}x{height}px，dpi={dpi}，逻辑画布 {W}x{H}，超采样 scale=3")
    print(f"SVG 物理宽度 {svg_width}，viewBox 0 0 {W} {H}；黑白 SVG 中残留彩色值 {len(colored)} 个")
    if font_warnings():
        print("字体告警：")
        for message in font_warnings():
            print("  -", message)
    else:
        print("字体：黑体/宋体探测正常（SimHei / SimSun）")

    assert width >= 1200, f"PNG 宽度不足 1200px：{width}"
    assert dpi == (600, 600), f"PNG dpi 不是 600：{dpi}"
    assert svg_width.endswith("mm") and float(svg_width[:-2]) == 160.0, f"SVG 物理宽度异常：{svg_width}"
    assert not colored, f"黑白 SVG 仍有彩色：{sorted(set(colored))[:5]}"
    print("自检通过：PNG>=1200px 且 dpi=600，SVG 160mm 可编辑源，黑白版无彩色。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
