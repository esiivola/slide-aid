#!/usr/bin/env python3
"""Render documentation workflow GIFs.

The demos are intentionally schematic: they show what a command changes without
PowerPoint selection chrome hiding the action. Each GIF shows the command inside
a simplified ribbon group, using the same icon files as the real add-in ribbon.
"""
from __future__ import annotations

import math
from pathlib import Path
from typing import Callable, Iterable

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "img"
ICONS = ROOT / "ribbon" / "images"

W, H = 960, 540
SLIDE = (38, 118, 922, 506)
RIBBON = (0, 0, W, 112)

NAVY = (35, 48, 70)
BLUE = (57, 107, 166)
LIGHT_BLUE = (223, 237, 254)
ORANGE = (237, 125, 49)
GREEN = (112, 173, 71)
RED = (192, 80, 77)
GREY = (120, 130, 145)
LIGHT_GREY = (235, 239, 245)
INK = (38, 48, 64)
MASTER = (211, 84, 74)
TARGET = (53, 118, 184)
YELLOW = (255, 205, 86)


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    names = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


FONT_11 = font(11)
FONT_13 = font(13)
FONT_15 = font(15)
FONT_16 = font(16)
FONT_18 = font(18)
FONT_20 = font(20, True)
FONT_24 = font(24, True)
FONT_30 = font(30, True)


def ease(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return t * t * (3 - 2 * t)


def lerp(a: float, b: float, t: float) -> float:
    return a + (b - a) * ease(t)


def mix(c1: tuple[int, int, int], c2: tuple[int, int, int], t: float) -> tuple[int, int, int]:
    return tuple(round(lerp(a, b, t)) for a, b in zip(c1, c2))


def text(
    d: ImageDraw.ImageDraw,
    xy: tuple[float, float],
    value: str,
    fill: tuple[int, int, int] = INK,
    anchor: str = "mm",
    fnt: ImageFont.ImageFont = FONT_15,
) -> None:
    d.text(xy, value, fill=fill, anchor=anchor, font=fnt)


def rounded(
    d: ImageDraw.ImageDraw,
    box: tuple[float, float, float, float],
    fill: tuple[int, int, int],
    outline: tuple[int, int, int] | None = None,
    width: int = 2,
    radius: int = 8,
) -> None:
    d.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def frame_base() -> Image.Image:
    img = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, H], fill=(248, 250, 253))
    d.rectangle(RIBBON, fill=(246, 248, 252))
    d.rectangle([SLIDE[0], SLIDE[1], SLIDE[2], SLIDE[3]], fill=(255, 255, 255), outline=(205, 214, 226), width=2)
    d.line([0, 111, W, 111], fill=(206, 216, 229), width=2)
    return img


def icon_image(name: str, size: int = 34) -> Image.Image:
    path = ICONS / f"{name}.png"
    img = Image.open(path).convert("RGBA")
    return img.resize((size, size), Image.Resampling.LANCZOS)


RIBBON_GROUPS = {
    "Slide Aid": {
        "Wizards": [
            ("sa_elements", "My Elements"),
            ("sa_formats", "My Formats"),
            ("sa_agenda", "Agenda"),
            ("sa_painter", "Format Painter"),
            ("sa_similar", "Select Similar"),
        ],
        "Position": [
            ("sa_align_left", "Left"),
            ("sa_align_right", "Right"),
            ("sa_align_top", "Top"),
            ("sa_align_bottom", "Bottom"),
            ("sa_align_center", "Center"),
            ("sa_align_middle", "Middle"),
            ("sa_to_slide", "To Slide"),
            ("sa_dist_h", "Distribute H"),
            ("sa_dist_v", "Distribute V"),
            ("sa_swap", "Swap"),
            ("sa_dock", "Dock"),
            ("sa_stack", "Stack"),
            ("sa_matrix", "Matrix"),
            ("sa_matrix_custom", "Matrix..."),
            ("sa_place", "Place"),
            ("sa_spacing", "Spacing"),
            ("sa_golden", "Golden Canon"),
        ],
        "Size": [
            ("sa_magic", "Magic"),
            ("sa_width", "Width"),
            ("sa_height", "Height"),
            ("sa_size", "W + H"),
            ("sa_stretch", "Stretch"),
            ("sa_fill", "Fill Gap"),
            ("sa_slice", "Slice"),
            ("sa_multiply", "Multiply"),
        ],
        "Shape": [
            ("sa_chain", "Process Chain"),
            ("sa_angles", "Angles"),
            ("sa_blockarrow", "Block Arrows"),
            ("sa_roundrect", "Rounded Rect."),
            ("sa_table", "Snap Table"),
        ],
        "Color": [
            ("sa_fillcolor", "Fill"),
            ("sa_linecolor", "Line"),
            ("sa_fontcolor", "Font"),
            ("sa_pick", "Pick Master"),
            ("sa_torgb", "To RGB"),
            ("sa_totheme", "To Theme"),
            ("sa_info", "Color Info"),
        ],
        "Text": [
            ("sa_margins", "Margins"),
            ("sa_fit", "Fit Text"),
            ("sa_wrap", "Wrap"),
            ("sa_split", "Split"),
            ("sa_merge", "Merge"),
            ("sa_txt_more", "More"),
        ],
        "View & Expert": [
            ("sa_hide", "Hide"),
            ("sa_unhide", "Unhide"),
            ("sa_masterobj", "Master Obj."),
            ("sa_cleanup", "Clean-up"),
            ("sa_paste", "Paste Slides"),
            ("sa_lang", "Language"),
            ("sa_shortcut", "Shortcuts"),
        ],
    },
    "Chart Aid": {
        "Charts": [
            ("sa_ch_col", "Column"),
            ("sa_ch_bar", "Bar"),
            ("sa_ch_stk", "Stacked"),
            ("sa_ch_sbr", "Stacked Bar"),
            ("sa_ch_pct", "100%"),
            ("sa_ch_wf", "Waterfall"),
            ("sa_ch_mek", "Mekko"),
            ("sa_ch_line", "Line"),
            ("sa_ch_area", "Area"),
            ("sa_ch_pie", "Pie"),
            ("sa_ch_don", "Doughnut"),
            ("sa_ch_scat", "Scatter"),
            ("sa_ch_gantt", "Gantt"),
        ],
        "Data": [
            ("sa_ch_edit", "Edit Data"),
            ("sa_ch_rebuild", "Rebuild"),
            ("sa_ch_help", "Layouts"),
            ("sa_ch_samples", "Samples"),
        ],
        "Style": [
            ("sa_ch_colors", "Themes"),
            ("sa_ch_colors", "Customize"),
            ("sa_ch_restyle", "Restyle All"),
            ("sa_ch_recolor", "Recolor"),
        ],
        "Annotations": [
            ("sa_ch_diff", "Difference"),
            ("sa_ch_pctdiff", "% Diff"),
            ("sa_ch_cagr", "CAGR"),
            ("sa_ch_vline", "Value Line"),
            ("sa_ch_avg", "Average"),
        ],
        "Elements": [
            ("sa_ch_harvey", "Harvey"),
            ("sa_ch_check", "Checkbox"),
            ("sa_ch_cycle", "Cycle"),
        ],
    },
}


ICON_CONTEXT = {
    icon: (tab, group)
    for tab, groups in RIBBON_GROUPS.items()
    for group, buttons in groups.items()
    for icon, _ in buttons
}


def ribbon_window(buttons: list[tuple[str, str]], active_icon: str, max_buttons: int = 7) -> list[tuple[str, str]]:
    idx = next((i for i, item in enumerate(buttons) if item[0] == active_icon), 0)
    start = max(0, min(idx - max_buttons // 2, len(buttons) - max_buttons))
    return buttons[start : start + max_buttons]


def short_label(label: str, width: int = 11) -> str:
    return label if len(label) <= width else label[: width - 1] + "."


def draw_ribbon(
    img: Image.Image,
    icon: str,
    label: str,
    click: float = 0.0,
) -> tuple[int, int]:
    d = ImageDraw.Draw(img)
    tab, group = ICON_CONTEXT[icon]
    d.rectangle([0, 0, W, 28], fill=(235, 239, 246))
    tabs = [("Slide Aid", 42, 142), ("Chart Aid", 146, 246)]
    for name, x1, x2 in tabs:
        active = name == tab
        d.rounded_rectangle([x1, 5, x2, 30], radius=6, fill=(255, 255, 255) if active else (235, 239, 246), outline=(196, 207, 222))
        text(d, ((x1 + x2) / 2, 17), name, fill=NAVY if active else GREY, fnt=FONT_15)

    all_buttons = RIBBON_GROUPS[tab][group]
    buttons = ribbon_window(all_buttons, icon)
    button_w = 84
    group_w = button_w * len(buttons) + 24
    group_x = max(280, min(620, W - group_w - 36))
    y1, y2 = 34, 104
    d.rounded_rectangle([group_x, y1, group_x + group_w, y2], radius=7, fill=(255, 255, 255), outline=(204, 214, 228), width=2)
    text(d, (group_x + group_w / 2, 94), group, fill=(91, 104, 122), fnt=FONT_11)

    hotspot = (group_x + 42, 58)
    for i, (btn_icon, btn_label) in enumerate(buttons):
        bx = group_x + 12 + i * button_w
        active = btn_icon == icon
        if active:
            fill = LIGHT_BLUE if click > 0 else (241, 247, 255)
            outline = BLUE
            rounded(d, (bx + 7, 39, bx + 69, 84), fill=fill, outline=outline, width=3, radius=6)
            hotspot = (bx + 38, 57)
        icon_rgba = icon_image(btn_icon, 28)
        img.paste(icon_rgba, (bx + 24, 43), icon_rgba)
        text(d, (bx + 38, 78), short_label(label if active else btn_label), fill=NAVY if active else INK, fnt=FONT_11)

    if click > 0:
        r = 12 + 13 * click
        alpha_col = mix(YELLOW, (255, 255, 255), click)
        d.ellipse([hotspot[0] - r, hotspot[1] - r, hotspot[0] + r, hotspot[1] + r], outline=alpha_col, width=4)
    return hotspot


def draw_cursor(d: ImageDraw.ImageDraw, pos: tuple[float, float]) -> None:
    x, y = pos
    pts = [(x, y), (x, y + 44), (x + 11, y + 34), (x + 20, y + 55), (x + 31, y + 50), (x + 22, y + 30), (x + 39, y + 30)]
    d.polygon(pts, fill=(255, 255, 255), outline=(22, 28, 36))
    d.line([x + 1, y + 1, x + 1, y + 41, x + 11, y + 31], fill=(22, 28, 36), width=2)


def click_ring(d: ImageDraw.ImageDraw, pos: tuple[float, float], t: float) -> None:
    if t <= 0:
        return
    r = 10 + 16 * t
    col = mix(YELLOW, (255, 255, 255), t)
    d.ellipse([pos[0] - r, pos[1] - r, pos[0] + r, pos[1] + r], outline=col, width=4)


def shape(
    d: ImageDraw.ImageDraw,
    box: tuple[float, float, float, float],
    fill: tuple[int, int, int],
    label: str = "",
    selected: bool = False,
    master: bool = False,
    radius: int = 8,
) -> None:
    rounded(d, box, fill=fill, outline=(255, 255, 255), width=2, radius=radius)
    if label:
        text(d, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), label, fill=(255, 255, 255), fnt=FONT_20)
    if selected or master:
        col = MASTER if master else TARGET
        d.rectangle([box[0] - 7, box[1] - 7, box[2] + 7, box[3] + 7], outline=col, width=3)
        for px, py in [(box[0] - 7, box[1] - 7), (box[2] + 7, box[1] - 7), (box[0] - 7, box[3] + 7), (box[2] + 7, box[3] + 7)]:
            d.rectangle([px - 5, py - 5, px + 5, py + 5], fill=(255, 255, 255), outline=col, width=3)
    if master:
        rounded(d, (box[0], box[1] - 37, box[0] + 120, box[1] - 9), fill=(255, 255, 255), outline=MASTER, width=2, radius=4)
        text(d, (box[0] + 12, box[1] - 23), "MASTER", fill=MASTER, anchor="lm", fnt=FONT_16)


def arrow(d: ImageDraw.ImageDraw, start: tuple[float, float], end: tuple[float, float], fill: tuple[int, int, int] = BLUE, width: int = 4) -> None:
    d.line([start, end], fill=fill, width=width)
    ang = math.atan2(end[1] - start[1], end[0] - start[0])
    for off in (2.55, -2.55):
        p = (end[0] - 16 * math.cos(ang + off), end[1] - 16 * math.sin(ang + off))
        d.line([end, p], fill=fill, width=width)


def table(d: ImageDraw.ImageDraw, x: int, y: int, cols: int, rows: int, cw: int, rh: int) -> None:
    d.rectangle([x, y, x + cols * cw, y + rows * rh], fill=(255, 255, 255), outline=(132, 148, 169), width=2)
    for c in range(1, cols):
        d.line([x + c * cw, y, x + c * cw, y + rows * rh], fill=(204, 213, 225), width=2)
    for r in range(1, rows):
        d.line([x, y + r * rh, x + cols * cw, y + r * rh], fill=(204, 213, 225), width=2)


def bars(d: ImageDraw.ImageDraw, x: int, y: int, vals: Iterable[int], color: tuple[int, int, int] = BLUE) -> None:
    vals = list(vals)
    bw = 34
    for i, v in enumerate(vals):
        h = v * 2.0
        d.rectangle([x + i * 56, y - h, x + i * 56 + bw, y], fill=color)
    d.line([x - 12, y, x + len(vals) * 56, y], fill=GREY, width=2)


class Demo:
    def __init__(self, name: str, icon: str, label: str, detail: str, scene: Callable[[ImageDraw.ImageDraw, float], tuple[float, float] | None]):
        self.name = name
        self.icon = icon
        self.label = label
        self.detail = detail
        self.scene = scene


def normal_progress(idx: int, n: int) -> float:
    return max(0.0, min(1.0, (idx - 9) / (n - 11)))


def make_gif(demo: Demo, frames: int = 26) -> None:
    imgs = []
    hotspot = (720, 58)
    for i in range(frames):
        p = normal_progress(i, frames)
        click = 1.0 - abs(i - 8) / 3 if 5 <= i <= 11 else 0.0
        img = frame_base()
        d = ImageDraw.Draw(img)
        object_cursor = demo.scene(d, p)
        hotspot = draw_ribbon(img, demo.icon, demo.label, max(0.0, click))
        if i < 4 and object_cursor:
            pos = object_cursor
            click_ring(d, pos, 1.0 - i / 4)
        elif i < 10:
            pos = (lerp(360, hotspot[0], (i - 4) / 5), lerp(300, hotspot[1], (i - 4) / 5))
        else:
            pos = hotspot
        draw_cursor(d, (pos[0] - 2, pos[1] - 2))
        imgs.append(img.convert("P", palette=Image.Palette.ADAPTIVE, colors=128))
    path = OUT / demo.name
    imgs[0].save(path, save_all=True, append_images=imgs[1:], duration=85, loop=0, optimize=False, disposal=2)


def align_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_box = (640, 292, 790, 370)
    target_x = lerp(250, 640, p)
    shape(d, (target_x, 184, target_x + 120, 246), ORANGE, "Target", selected=True)
    shape(d, (420, 318, 540, 380), GREEN, "Target", selected=True)
    shape(d, master_box, NAVY, "Master", master=True)
    d.line([640, 160, 640, 406], fill=(211, 84, 74), width=2)
    text(d, (650, 150), "left edge", fill=MASTER, anchor="lm", fnt=FONT_13)
    return (305, 215)


def master_order_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    boxes = [(165, 195, 295, 260), (350, 195, 480, 260), (580, 305, 730, 380)]
    labels = ["1 target", "2 target", "3 master"]
    for j, box in enumerate(boxes):
        is_master = j == 2
        selected = p > 0.15 * j
        shape(d, box, NAVY if is_master else (ORANGE if j == 0 else GREEN), labels[j], selected=selected and not is_master, master=selected and is_master)
    arrow(d, (300, 270), (570, 315), fill=MASTER)
    text(d, (420, 430), "Last selected object becomes the Master", fill=MASTER, fnt=FONT_20)
    return (230, 228)


def single_slide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    x = lerp(180, 428, p)
    y = lerp(190, 287, p)
    shape(d, (x, y, x + 105, y + 70), BLUE, "One", selected=True)
    d.line([480, 132, 480, 492], fill=(220, 226, 236), width=2)
    d.line([54, 312, 906, 312], fill=(220, 226, 236), width=2)
    text(d, (480, 480), "single selection uses the slide as reference", fill=GREY, fnt=FONT_16)
    return (232, 224)


def to_slide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    dx, dy = lerp(0, 220, p), lerp(0, 65, p)
    for x, y, c in [(155, 190, ORANGE), (285, 250, GREEN), (210, 340, BLUE)]:
        shape(d, (x + dx, y + dy, x + dx + 92, y + dy + 52), c, selected=True)
    d.rectangle([330, 210, 630, 410], outline=(220, 226, 236), width=3)
    text(d, (480, 198), "slide target area", fill=GREY, fnt=FONT_13)
    return (198, 216)


def reusable_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    shape(d, (150, 210, 280, 285), ORANGE, "Saved", selected=True)
    d.rectangle([620, 170, 820, 430], fill=(247, 249, 252), outline=(196, 207, 222), width=2)
    text(d, (720, 194), "My Elements", fill=NAVY, fnt=FONT_18)
    x = lerp(165, 655, p)
    y = lerp(225, 232, p)
    shape(d, (x, y, x + 90, y + 54), ORANGE, "", selected=False)
    if p > 0.55:
        shape(d, (360, 330, 490, 405), ORANGE, "Inserted")
    return (214, 248)


def painter_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_col = NAVY
    target_col = mix((220, 224, 230), master_col, p)
    shape(d, (175, 205, 305, 285), target_col, "Target", selected=True)
    shape(d, (390, 205, 520, 285), target_col, "Target", selected=True)
    shape(d, (650, 315, 790, 395), master_col, "Master", master=True)
    return (240, 246)


def agenda_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    sections = ["Context", "Options", "Decision"]
    for i, s in enumerate(sections):
        y = 178 + i * 62
        d.rectangle([130, y, 315, y + 40], fill=(245, 247, 250), outline=(194, 205, 220), width=2)
        text(d, (150, y + 20), s, anchor="lm", fnt=FONT_16)
    for i, s in enumerate(sections):
        x = lerp(520, 440 + i * 75, p)
        y = lerp(170 + i * 5, 245, p)
        d.rectangle([x, y, x + 60, y + 85], fill=(255, 255, 255), outline=BLUE, width=2)
        text(d, (x + 30, y + 32), str(i + 1), fill=BLUE, fnt=FONT_24)
        text(d, (x + 30, y + 58), s, fill=INK, fnt=FONT_11)
    return (224, 198)


def distribute_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    xs = [150, lerp(292, 338, p), lerp(475, 526, p), 710]
    for i, x in enumerate(xs):
        shape(d, (x, 260, x + 92, 322), [ORANGE, GREEN, BLUE, NAVY][i], selected=True)
    for x in [242, 430, 618]:
        d.line([x, 236, x, 348], fill=(220, 226, 236), width=2)
    text(d, (480, 400), "outer objects stay fixed; inner gaps become equal", fill=GREY, fnt=FONT_16)
    return (196, 292)


def stack_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(150, 245), (360, 190), (610, 315)]
    ends = [(150, 245), (258, 245), (366, 245)]
    cols = [ORANGE, GREEN, BLUE]
    for i, (s, e) in enumerate(zip(starts, ends)):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        shape(d, (x, y, x + 100, y + 62), cols[i], selected=True)
    return (200, 276)


def spacing_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    xs = [155, lerp(330, 305, p), lerp(560, 455, p)]
    for x, c in zip(xs, [ORANGE, GREEN, BLUE]):
        shape(d, (x, 250, x + 105, 314), c, selected=True)
    arrow(d, (260, 336), (305, 336), fill=GREY, width=3)
    arrow(d, (410, 336), (455, 336), fill=GREY, width=3)
    text(d, (357, 358), "exact 0.5 cm gaps", fill=GREY, fnt=FONT_15)
    return (205, 282)


def matrix_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(160, 190), (405, 155), (620, 235), (275, 365), (520, 345), (720, 390)]
    ends = [(250, 200), (360, 200), (470, 200), (250, 292), (360, 292), (470, 292)]
    cols = [ORANGE, GREEN, BLUE, NAVY, RED, GREY]
    for s, e, c in zip(starts, ends, cols):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        shape(d, (x, y, x + 78, y + 56), c, selected=True, radius=6)
    return (196, 218)


def golden_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (210, 165, 750, 430)
    rounded(d, master, fill=(255, 255, 255), outline=MASTER, width=3, radius=4)
    y = lerp(320, 244, p)
    shape(d, (360, y, 600, y + 78), NAVY, "Content", selected=True)
    d.line([210, 244, 750, 244], fill=YELLOW, width=4)
    text(d, (760, 244), "golden-canon height", fill=GREY, anchor="lm", fnt=FONT_13)
    return (480, 360)


def magic_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    cx, cy = 420, 288
    scale = lerp(0.75, 1.25, p)
    for ox, oy, w, h, c in [(-150, -65, 105, 58, ORANGE), (-20, -75, 155, 70, GREEN), (-120, 30, 240, 48, BLUE)]:
        box = (cx + ox * scale, cy + oy * scale, cx + (ox + w) * scale, cy + (oy + h) * scale)
        shape(d, box, c, selected=True)
    text(d, (420, 420), "sizes and text scale together", fill=GREY, fnt=FONT_16)
    return (285, 232)


def match_size_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_box = (620, 285, 785, 370)
    targets = [((200, 190, 310, 250), (172, 178, 337, 263), ORANGE), ((365, 250, 455, 333), (327, 249, 492, 334), GREEN)]
    for start, end, col in targets:
        box = tuple(lerp(a, b, p) for a, b in zip(start, end))
        shape(d, box, col, "Target", selected=True)
    shape(d, master_box, NAVY, "Master", master=True)
    return (255, 220)


def dock_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (620, 245, 770, 335)
    x = lerp(260, 500, p)
    shape(d, (x, 252, x + 120, 328), ORANGE, "Target", selected=True)
    shape(d, master, NAVY, "Master", master=True)
    text(d, (560, 372), "target moves until it touches the Master", fill=GREY, fnt=FONT_15)
    return (315, 290)


def fill_gap_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (650, 245, 795, 335)
    left = 220
    right = lerp(360, 650, p)
    shape(d, (left, 258, right, 322), ORANGE, "Target", selected=True)
    shape(d, master, NAVY, "Master", master=True)
    text(d, (500, 372), "target extends across the gap to the Master", fill=GREY, fnt=FONT_15)
    return (292, 290)


def stretch_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (615, 205, 800, 385)
    left = 235
    right = lerp(420, 800, p)
    shape(d, (left, 260, right, 326), ORANGE, "Target", selected=True)
    rounded(d, master, fill=(255, 255, 255), outline=MASTER, width=3, radius=6)
    text(d, (615, 186), "MASTER", fill=MASTER, anchor="lm", fnt=FONT_16)
    d.line([800, 190, 800, 400], fill=MASTER, width=2)
    text(d, (510, 430), "right edge stretches to the Master's far edge", fill=GREY, fnt=FONT_15)
    return (326, 292)


def swap_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    a = (lerp(190, 610, p), 230)
    b = (lerp(610, 190, p), 330)
    shape(d, (a[0], a[1], a[0] + 130, a[1] + 68), ORANGE, "A", selected=True)
    shape(d, (b[0], b[1], b[0] + 130, b[1] + 68), BLUE, "B", selected=True)
    arrow(d, (340, 260), (590, 350), fill=GREY, width=3)
    arrow(d, (590, 350), (340, 260), fill=GREY, width=3)
    return (255, 266)


def place_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    d.rectangle([480, 130, 910, 494], outline=(220, 226, 236), width=3)
    text(d, (695, 148), "right half", fill=GREY, fnt=FONT_13)
    x, y = lerp(145, 520, p), lerp(230, 210, p)
    w, h = lerp(150, 330, p), lerp(95, 245, p)
    shape(d, (x, y, x + w, y + h), BLUE, "Selection", selected=True)
    return (220, 278)


def slice_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    x, y, w, h = 245, 180, 380, 225
    if p < 0.45:
        shape(d, (x, y, x + w, y + h), BLUE, "Original", selected=True)
    else:
        gap = 8
        for r in range(3):
            for c in range(4):
                bx = x + c * (w / 4)
                by = y + r * (h / 3)
                d.rectangle([bx + gap / 2, by + gap / 2, bx + w / 4 - gap / 2, by + h / 3 - gap / 2], fill=BLUE, outline=(255, 255, 255), width=2)
    text(d, (435, 438), "one shape becomes editable pieces", fill=GREY, fnt=FONT_16)
    return (430, 292)


def multiply_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    x, y, w, h = 310, 210, 100, 62
    shape(d, (x, y, x + w, y + h), BLUE, "Original", selected=True)
    if p > 0.25:
        q = min(1.0, (p - 0.25) / 0.75)
        for r in range(3):
            for c in range(4):
                if r == 0 and c == 0:
                    continue
                bx = lerp(x, x + c * 116, q)
                by = lerp(y, y + r * 76, q)
                shape(d, (bx, by, bx + w, by + h), BLUE)
    text(d, (470, 468), "copies keep the original size", fill=GREY, fnt=FONT_16)
    return (360, 240)


def chain_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(150, 215), (385, 185), (640, 295)]
    ends = [(170, 250), (355, 250), (540, 250)]
    for i, (s, e) in enumerate(zip(starts, ends)):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        d.polygon([(x, y), (x + 120, y), (x + 150, y + 42), (x + 120, y + 84), (x, y + 84), (x + 30, y + 42)], fill=[ORANGE, GREEN, BLUE][i], outline=(255, 255, 255))
    return (225, 256)


def snap_table_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    table(d, 220, 170, 4, 3, 120, 70)
    starts = [(260, 155), (430, 238), (625, 358)]
    ends = [(280, 205), (460, 275), (640, 345)]
    for s, e, col in zip(starts, ends, [ORANGE, GREEN, BLUE]):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        d.ellipse([x - 16, y - 16, x + 16, y + 16], fill=col, outline=(255, 255, 255), width=2)
    return (260, 155)


def colors_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    target_col = mix((220, 224, 230), ORANGE, p)
    shape(d, (190, 230, 335, 310), target_col, "Target", selected=True)
    shape(d, (610, 230, 755, 310), ORANGE, "Master", master=True)
    arrow(d, (585, 270), (350, 270), fill=ORANGE)
    return (260, 270)


def text_tools_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    box = (210, 205, 640, 345)
    rounded(d, box, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
    margin = lerp(12, 36, p)
    text(d, (box[0] + margin, box[1] + margin), "Text box margins", anchor="la", fnt=FONT_24, fill=NAVY)
    d.rectangle([box[0] + margin, box[1] + margin, box[2] - margin, box[3] - margin], outline=(220, 226, 236), width=2)
    return (420, 275)


def split_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    if p < 0.5:
        rounded(d, (220, 235, 690, 320), fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
        text(d, (455, 278), "First thought | second thought", fill=NAVY, fnt=FONT_24)
        d.line([463, 242, 463, 312], fill=RED, width=3)
    else:
        rounded(d, (190, 235, 400, 320), fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
        rounded(d, (470, 235, 725, 320), fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
        text(d, (295, 278), "First thought", fill=NAVY, fnt=FONT_20)
        text(d, (598, 278), "second thought", fill=NAVY, fnt=FONT_20)
    return (463, 276)


def fit_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    box = (lerp(205, 310, p), lerp(205, 245, p), lerp(710, 600, p), lerp(385, 325, p))
    rounded(d, box, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
    text(d, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), "Short label", fill=NAVY, fnt=FONT_30)
    return (458, 294)


def hide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    alpha_col = mix(ORANGE, (248, 250, 253), p)
    for i, (x, y) in enumerate([(180, 190), (310, 230), (460, 180), (590, 265), (720, 220)]):
        shape(d, (x, y, x + 95, y + 65), alpha_col if i % 2 == 0 else mix(BLUE, (248, 250, 253), p), selected=True)
    text(d, (450, 420), "crowded objects are hidden temporarily", fill=GREY, fnt=FONT_16)
    return (230, 222)


def paste_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    copied = (150, 185, 260, 250)
    shape(d, copied, ORANGE, "Copy", selected=True)
    for i in range(3):
        x = 450 + i * 125
        d.rectangle([x, 175, x + 88, 124 + 190], fill=(255, 255, 255), outline=(188, 200, 216), width=2)
        if p > i * 0.2:
            shape(d, (x + 16, 222, x + 72, 258), ORANGE)
    return (205, 220)


def chart_create_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    if p < 0.55:
        table(d, 185, 185, 4, 4, 82, 44)
        for c, val in enumerate(["Q1", "Q2", "Q3"]):
            text(d, (185 + (c + 1) * 82 + 41, 207), val, fnt=FONT_13)
    x = lerp(520, 440, p)
    bars(d, int(x), 390, [55, 85, 70], BLUE)
    return (315, 235)


def chart_rebuild_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    vals = [lerp(45, 60, p), lerp(70, 42, p), lerp(52, 90, p)]
    table(d, 165, 190, 3, 4, 78, 40)
    bars(d, 470, 390, vals, BLUE)
    text(d, (250, 382), "edited data", fill=GREY, fnt=FONT_15)
    return (535, 315)


def samples_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    names = ["Column", "Waterfall", "Mekko", "Line"]
    for i, name in enumerate(names):
        x = lerp(180, 210 + i * 150, max(0, min(1, p * 1.25 - i * 0.15)))
        y = 210 + (i % 2) * 120
        d.rectangle([x, y, x + 108, y + 78], fill=(255, 255, 255), outline=BLUE, width=2)
        text(d, (x + 54, y + 25), name, fnt=FONT_13)
        bars(d, int(x + 22), int(y + 62), [15, 24, 20], [ORANGE, GREEN, BLUE, NAVY][i])
    return (225, 250)


def style_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    col = mix(BLUE, GREEN, p)
    bars(d, 265, 390, [60, 82, 42, 74], col)
    for i, c in enumerate([BLUE, GREEN, ORANGE, RED]):
        d.rectangle([610, 200 + i * 36, 665, 224 + i * 36], fill=c, outline=(255, 255, 255), width=2)
    text(d, (695, 254), "theme swatches", fill=GREY, anchor="lm", fnt=FONT_15)
    return (340, 300)


def recolor_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    colors = [ORANGE, mix(BLUE, RED, p), GREEN]
    for i, c in enumerate(colors):
        bars(d, 300 + i * 46, 390, [40 + i * 15], c)
    text(d, (392, 420), "one selected series changes color", fill=GREY, fnt=FONT_15)
    return (382, 310)


def diff_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    bars(d, 260, 390, [45, 78, 62], BLUE)
    if p > 0.35:
        arrow(d, (277, 285), (333, 235), fill=RED, width=4)
        text(d, (306, 250), "+33", fill=RED, fnt=FONT_18)
    return (278, 300)


def chart_elements_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    cx, cy = 430, 285
    d.ellipse([cx - 52, cy - 52, cx + 52, cy + 52], fill=(255, 255, 255), outline=BLUE, width=4)
    if p > 0.25:
        d.pieslice([cx - 52, cy - 52, cx + 52, cy + 52], -90, -90 + 270 * p, fill=BLUE)
    d.ellipse([cx - 30, cy - 30, cx + 30, cy + 30], fill=(255, 255, 255))
    text(d, (430, 375), "editable status marker", fill=GREY, fnt=FONT_16)
    return (430, 285)


def cycle_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    states = [0.0, 0.25, 0.5, 0.75, 1.0]
    state = states[min(4, int(p * 4.99))]
    cx, cy = 430, 285
    d.ellipse([cx - 54, cy - 54, cx + 54, cy + 54], fill=(255, 255, 255), outline=BLUE, width=4)
    if state > 0:
        d.pieslice([cx - 54, cy - 54, cx + 54, cy + 54], -90, -90 + 360 * state, fill=BLUE)
    d.ellipse([cx - 31, cy - 31, cx + 31, cy + 31], fill=(255, 255, 255))
    text(d, (430, 375), f"{int(state * 100)}%", fill=NAVY, fnt=FONT_24)
    return (430, 285)


DEMOS = [
    Demo("demo-master-selection-order.gif", "sa_align_left", "Left", "Align to Master", master_order_scene),
    Demo("demo-single-object-slide-reference.gif", "sa_align_center", "Center", "Slide reference", single_slide_scene),
    Demo("demo-to-slide-reference.gif", "sa_to_slide", "To Slide", "Center on slide", to_slide_scene),
    Demo("demo-reusable-libraries.gif", "sa_elements", "My Elements", "Save and insert", reusable_scene),
    Demo("demo-format-painter-select-similar.gif", "sa_painter", "Format Painter", "Copy Master style", painter_scene),
    Demo("demo-agenda.gif", "sa_agenda", "Agenda", "Build section slides", agenda_scene),
    Demo("demo-align-to-master.gif", "sa_align_left", "Left", "Align to Master", align_scene),
    Demo("demo-distribute-horizontal.gif", "sa_dist_h", "Distribute H", "Equal horizontal gaps", distribute_scene),
    Demo("demo-stack-horizontal.gif", "sa_stack", "Stack", "Horizontally", stack_scene),
    Demo("demo-spacing-exact-gap.gif", "sa_spacing", "Spacing", "Horizontal gap", spacing_scene),
    Demo("demo-matrix-grid.gif", "sa_matrix", "Matrix", "Arrange as grid", matrix_scene),
    Demo("demo-golden-canon-placement.gif", "sa_golden", "Golden Canon", "Place in Master", golden_scene),
    Demo("demo-magic-resizer.gif", "sa_magic", "Magic Resizer", "Scale arrangement", magic_scene),
    Demo("demo-match-size.gif", "sa_size", "Width + Height", "Match Master size", match_size_scene),
    Demo("demo-dock-fill-stretch.gif", "sa_dock", "Dock", "Touch the Master", dock_scene),
    Demo("demo-dock.gif", "sa_dock", "Dock", "Touch the Master", dock_scene),
    Demo("demo-fill-gap.gif", "sa_fill", "Fill Gap", "Extend to Master", fill_gap_scene),
    Demo("demo-stretch.gif", "sa_stretch", "Stretch", "Extend to far edge", stretch_scene),
    Demo("demo-swap-positions.gif", "sa_swap", "Swap", "Exchange positions", swap_scene),
    Demo("demo-place-on-slide.gif", "sa_place", "Place on Slide", "Right half", place_scene),
    Demo("demo-slice-multiply.gif", "sa_slice", "Slice", "Split into pieces", slice_scene),
    Demo("demo-slice.gif", "sa_slice", "Slice", "Split into pieces", slice_scene),
    Demo("demo-multiply.gif", "sa_multiply", "Multiply", "Create copies", multiply_scene),
    Demo("demo-shape-helpers.gif", "sa_chain", "Process Chain", "Connect arrows", chain_scene),
    Demo("demo-snap-to-table.gif", "sa_table", "Snap to Table", "Center in cell", snap_table_scene),
    Demo("demo-color-tools.gif", "sa_pick", "Pick from Master", "Copy colors", colors_scene),
    Demo("demo-text-tools.gif", "sa_margins", "Set Margins", "Text box margins", text_tools_scene),
    Demo("demo-split-at-cursor.gif", "sa_split", "Split at Cursor", "Split text box", split_scene),
    Demo("demo-fit-to-text.gif", "sa_fit", "Fit to Text", "Resize text box", fit_scene),
    Demo("demo-view-cleanup.gif", "sa_hide", "Hide Objects", "Temporary cleanup", hide_scene),
    Demo("demo-paste-on-slides.gif", "sa_paste", "Paste on Slides", "Repeat paste", paste_scene),
    Demo("demo-chart-create-from-table.gif", "sa_ch_col", "Column", "Create from table", chart_create_scene),
    Demo("demo-chart-rebuild.gif", "sa_ch_rebuild", "Rebuild", "Update chart", chart_rebuild_scene),
    Demo("demo-chart-samples.gif", "sa_ch_samples", "Sample Slides", "Insert examples", samples_scene),
    Demo("demo-chart-annotations-style.gif", "sa_ch_restyle", "Restyle All", "Apply chart style", style_scene),
    Demo("demo-recolor-series-click.gif", "sa_ch_recolor", "Recolor Series", "Selected series", recolor_scene),
    Demo("demo-difference-arrow-clicks.gif", "sa_ch_diff", "Difference", "Compare two bars", diff_scene),
    Demo("demo-chart-elements.gif", "sa_ch_harvey", "Harvey Ball", "Insert element", chart_elements_scene),
    Demo("demo-cycle-state-clicks.gif", "sa_ch_cycle", "Cycle State", "Advance state", cycle_scene),
]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for demo in DEMOS:
        make_gif(demo)
    print(f"{len(DEMOS)} documentation GIFs rendered to {OUT}")


if __name__ == "__main__":
    main()
