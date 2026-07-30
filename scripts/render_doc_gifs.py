#!/usr/bin/env python3
"""Render documentation workflow GIFs.

The demos are intentionally schematic: they show what a command changes without
PowerPoint selection chrome hiding the action. Each GIF shows the command inside
a simplified ribbon group, using the same icon files as the real add-in ribbon.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from PIL import Image, ImageDraw, ImageFont
from workflow_contracts import golden_top, scale_box_about


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "img"
ICONS = ROOT / "shared" / "icons"

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
ACTIVE_SELECTION_TARGETS: list[tuple[float, float]] = []
SELECTION_COMPLETE = False


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
    fill: tuple[int, int, int] | None,
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


def selection_frame(
    d: ImageDraw.ImageDraw,
    box: tuple[float, float, float, float],
    *,
    selected: bool | str = False,
    master: bool = False,
) -> None:
    show_selection = selected == "force" or (
        bool(selected or master)
        and (
            SELECTION_COMPLETE
            or any(
                box[0] - 12 <= px <= box[2] + 12 and box[1] - 12 <= py <= box[3] + 12
                for px, py in ACTIVE_SELECTION_TARGETS
            )
        )
    )
    if not show_selection:
        return

    col = MASTER if master else TARGET
    d.rectangle([box[0] - 7, box[1] - 7, box[2] + 7, box[3] + 7], outline=col, width=3)
    for px, py in [(box[0] - 7, box[1] - 7), (box[2] + 7, box[1] - 7), (box[0] - 7, box[3] + 7), (box[2] + 7, box[3] + 7)]:
        d.rectangle([px - 5, py - 5, px + 5, py + 5], fill=(255, 255, 255), outline=col, width=3)


def shape(
    d: ImageDraw.ImageDraw,
    box: tuple[float, float, float, float],
    fill: tuple[int, int, int],
    label: str = "",
    selected: bool | str = False,
    master: bool = False,
    radius: int = 8,
) -> None:
    rounded(d, box, fill=fill, outline=(255, 255, 255), width=2, radius=radius)
    if label:
        text(d, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), label, fill=(255, 255, 255), fnt=FONT_20)
    selection_frame(d, box, selected=selected, master=master)
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


@dataclass(frozen=True)
class Action:
    kind: str
    label: str
    target: tuple[float, float] | str | None = None
    frames: int = 6
    end_result: float | None = None
    menu_items: tuple[str, ...] = ()
    menu_index: int = 0
    prompt_value: str = ""


@dataclass(frozen=True)
class BehaviorContract:
    reference: str
    result: str
    invariant: str


class Demo:
    def __init__(
        self,
        name: str,
        icon: str,
        label: str,
        detail: str,
        scene: Callable[[ImageDraw.ImageDraw, float], tuple[float, float] | None],
        clicks: list[tuple[float, float]] | None = None,
        actions: list[Action] | None = None,
        contract: BehaviorContract | None = None,
    ):
        self.name = name
        self.icon = icon
        self.label = label
        self.detail = detail
        self.scene = scene
        self.clicks = clicks
        self.actions = actions
        self.contract = contract


def default_actions(demo: Demo) -> list[Action]:
    clicks = demo.clicks
    if clicks is None:
        probe = frame_base()
        click = demo.scene(ImageDraw.Draw(probe), 0.0)
        clicks = [click] if click else []
    actions = [
        Action("select", f"Select object {i + 1}", target=pos)
        for i, pos in enumerate(clicks)
        if pos is not None
    ]
    actions.extend(
        [
            Action("ribbon", f"Click {demo.label}", target="ribbon"),
            Action("animate", demo.detail, frames=18, end_result=1.0),
            Action("hold", "Result", frames=8),
        ]
    )
    return actions


def draw_menu(
    d: ImageDraw.ImageDraw,
    ribbon_hotspot: tuple[int, int],
    items: tuple[str, ...],
    active: int,
) -> tuple[float, float]:
    width = max(220, max((len(item) for item in items), default=12) * 8 + 34)
    x1 = max(42, min(W - width - 42, ribbon_hotspot[0] - width / 2))
    y1 = 105
    height = 34 * len(items) + 16
    d.rectangle([x1, y1, x1 + width, y1 + height], fill=(255, 255, 255), outline=(155, 169, 188), width=2)
    for idx, item in enumerate(items):
        iy = y1 + 8 + idx * 34
        if idx == active:
            d.rectangle([x1 + 5, iy, x1 + width - 5, iy + 30], fill=LIGHT_BLUE, outline=BLUE, width=2)
        text(d, (x1 + 16, iy + 15), item, anchor="lm", fill=NAVY if idx == active else INK, fnt=FONT_15)
    return (x1 + width - 24, y1 + 23 + active * 34)


def draw_prompt(d: ImageDraw.ImageDraw, label: str, value: str) -> tuple[float, float]:
    box = (285, 205, 675, 370)
    d.rectangle(box, fill=(250, 251, 253), outline=(130, 145, 166), width=2)
    text(d, (310, 236), label, anchor="lm", fill=NAVY, fnt=FONT_18)
    d.rectangle([310, 265, 650, 310], fill=(255, 255, 255), outline=(145, 159, 179), width=2)
    text(d, (325, 287), value, anchor="lm", fill=INK, fnt=FONT_16)
    rounded(d, (548, 325, 650, 357), fill=LIGHT_BLUE, outline=BLUE, width=2, radius=5)
    text(d, (599, 341), "OK", fill=NAVY, fnt=FONT_15)
    return (599, 341)


def draw_step(d: ImageDraw.ImageDraw, label: str) -> None:
    rounded(d, (58, 464, 470, 492), fill=(248, 250, 253), outline=(210, 219, 231), width=1, radius=4)
    text(d, (72, 478), label, anchor="lm", fill=GREY, fnt=FONT_13)


def action_result(actions: list[Action], index: int, local: float) -> float:
    result = 0.0
    for action in actions[:index]:
        if action.end_result is not None:
            result = action.end_result
    action = actions[index]
    if action.kind == "animate" and action.end_result is not None:
        return lerp(result, action.end_result, local)
    return result


def resolve_target(
    action: Action,
    ribbon_hotspot: tuple[int, int],
    overlay_hotspot: tuple[float, float] | None,
) -> tuple[float, float]:
    if isinstance(action.target, tuple):
        return action.target
    if action.target == "ribbon":
        return ribbon_hotspot
    if action.target in {"menu", "prompt"} and overlay_hotspot is not None:
        return overlay_hotspot
    return ribbon_hotspot


def make_gif(demo: Demo) -> None:
    global ACTIVE_SELECTION_TARGETS, SELECTION_COMPLETE
    actions = demo.actions or default_actions(demo)
    total_frames = sum(action.frames for action in actions)
    frame_starts: list[int] = []
    cursor = 0
    for action in actions:
        frame_starts.append(cursor)
        cursor += action.frames

    imgs = []
    previous_target: tuple[float, float] = (360, 300)
    for i in range(total_frames):
        action_index = max(idx for idx, start in enumerate(frame_starts) if start <= i)
        action = actions[action_index]
        local_frame = i - frame_starts[action_index]
        local = local_frame / max(1, action.frames - 1)
        p = action_result(actions, action_index, local)
        ACTIVE_SELECTION_TARGETS = [
            prior.target
            for prior in actions[:action_index]
            if prior.kind == "select" and isinstance(prior.target, tuple)
        ]
        if action.kind == "select" and isinstance(action.target, tuple) and local >= 0.55:
            ACTIVE_SELECTION_TARGETS.append(action.target)
        SELECTION_COMPLETE = bool(ACTIVE_SELECTION_TARGETS) and action.kind != "select"
        img = frame_base()
        d = ImageDraw.Draw(img)
        demo.scene(d, p)
        ribbon_click = local if action.kind == "ribbon" and local >= 0.55 else 0.0
        hotspot = draw_ribbon(img, demo.icon, demo.label, ribbon_click)

        overlay_hotspot = None
        if action.kind == "menu":
            overlay_hotspot = draw_menu(d, hotspot, action.menu_items, action.menu_index)
        elif action.kind == "prompt":
            overlay_hotspot = draw_prompt(d, action.label, action.prompt_value)

        target = resolve_target(action, hotspot, overlay_hotspot)
        if local < 0.5:
            pos = (
                lerp(previous_target[0], target[0], local * 2),
                lerp(previous_target[1], target[1], local * 2),
            )
        else:
            pos = target
        if action.kind in {"select", "ribbon", "menu", "prompt"} and local >= 0.55:
            click_ring(d, target, 1.0 - (local - 0.55) / 0.45)

        draw_step(d, action.label)
        draw_cursor(d, (pos[0] - 2, pos[1] - 2))
        imgs.append(img.convert("P", palette=Image.Palette.ADAPTIVE, colors=128))
        if local_frame == action.frames - 1:
            previous_target = target

    path = OUT / demo.name
    imgs[0].save(path, save_all=True, append_images=imgs[1:], duration=85, loop=0, optimize=False, disposal=2)
    ACTIVE_SELECTION_TARGETS = []
    SELECTION_COMPLETE = False


def align_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_box = (640, 292, 790, 370)
    target_x = lerp(250, 640, p)
    second_target_x = lerp(420, 640, p)
    shape(d, (target_x, 184, target_x + 120, 246), ORANGE, "Target", selected=True)
    shape(d, (second_target_x, 404, second_target_x + 120, 466), GREEN, "Target", selected=True)
    shape(d, master_box, NAVY, "Master", master=True)
    return (250, 184)


def master_order_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_box = (580, 395, 730, 470)
    starts = [(165, 175, 295, 240), (350, 265, 480, 330)]
    labels = ["1 target", "2 target"]
    for j, start in enumerate(starts):
        end = (580, start[1], 710, start[3])
        box = tuple(lerp(a, b, p) for a, b in zip(start, end))
        shape(d, box, ORANGE if j == 0 else GREEN, labels[j], selected=True)
    shape(d, master_box, NAVY, "3 master", master=True)
    return (165, 195)


def single_slide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    x = lerp(180, 428, p)
    y = 190
    shape(d, (x, y, x + 105, y + 70), BLUE, "One", selected=True)
    d.line([480, 132, 480, 492], fill=(220, 226, 236), width=2)
    text(d, (492, 146), "slide centre line", anchor="lm", fill=GREY, fnt=FONT_13)
    return (180, 190)


def to_slide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    for x, y, c in [(155, 190, ORANGE), (285, 250, GREEN), (210, 340, BLUE)]:
        aligned_x = lerp(x, 434, p)
        shape(d, (aligned_x, y, aligned_x + 92, y + 52), c, selected=True)
    d.line([480, 140, 480, 480], fill=(232, 237, 244), width=2)
    return (155, 190)


def library_panel(d: ImageDraw.ImageDraw) -> None:
    d.rectangle([610, 150, 845, 456], fill=(247, 249, 252), outline=(196, 207, 222), width=2)
    text(d, (728, 176), "My Elements", fill=NAVY, fnt=FONT_18)
    d.line([630, 198, 825, 198], fill=(212, 220, 232), width=2)


def reusable_save_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    library_panel(d)
    shape(d, (150, 225, 280, 300), ORANGE, "KPI", selected=True)
    if p > 0.05:
        rounded(d, (640, 214, 800, 290), fill=(255, 255, 255), outline=BLUE, width=2, radius=5)
        shape(d, (658, 225, 748, 278), ORANGE, "KPI")
        text(d, (720, 316), "KPI card saved", fill=GREY, fnt=FONT_13)
    return (150, 225)


def reusable_insert_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    library_panel(d)
    rounded(d, (640, 214, 800, 290), fill=(255, 255, 255), outline=BLUE, width=2, radius=5)
    shape(d, (658, 225, 748, 278), ORANGE, "KPI")
    if p > 0.05:
        shape(d, (230, 322, 320, 375), ORANGE, "KPI", selected="force")
        text(d, (360, 410), "inserted at its saved slide coordinates", fill=GREY, fnt=FONT_15)
    return (700, 240)


def reusable_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    if p <= 0.5:
        reusable_save_scene(d, p * 2)
    else:
        reusable_insert_scene(d, (p - 0.5) * 2)
    return (150, 225)


def painter_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_col = NAVY
    target_col = mix((220, 224, 230), master_col, p)
    shape(d, (175, 205, 305, 285), target_col, "Target", selected=True)
    shape(d, (390, 205, 520, 285), target_col, "Target", selected=True)
    shape(d, (650, 315, 790, 395), master_col, "Master", master=True)
    return (175, 205)


def select_similar_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    candidates = [
        ((160, 195, 275, 265), ORANGE, True),
        ((340, 195, 455, 265), GREEN, False),
        ((520, 195, 635, 265), ORANGE, True),
        ((250, 335, 365, 405), BLUE, False),
    ]
    for box, color, matches in candidates:
        shape(d, box, color, selected="force" if p > 0.05 and matches else False)
    shape(d, (650, 335, 790, 415), ORANGE, "Master", master=True)
    text(d, (470, 447), "matching shape type + fill become selected", fill=GREY, fnt=FONT_15)
    return (650, 335)


def agenda_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    sections = ["Context", "Options", "Decision"]
    text(d, (205, 148), "PowerPoint sections", fill=NAVY, fnt=FONT_18)
    for i, s in enumerate(sections):
        y = 178 + i * 62
        d.rectangle([130, y, 315, y + 40], fill=(245, 247, 250), outline=(194, 205, 220), width=2)
        text(d, (150, y + 20), s, anchor="lm", fnt=FONT_16)
    if p > 0.05:
        arrow(d, (340, 260), (440, 260), fill=GREY, width=3)
    text(d, (645, 148), "generated slides", fill=NAVY, fnt=FONT_18)
    if p > 0.08:
        d.rectangle([450, 202, 540, 285], fill=(255, 255, 255), outline=BLUE, width=2)
        text(d, (495, 228), "Overview", fill=BLUE, fnt=FONT_13)
        for i, s in enumerate(sections):
            text(d, (495, 246 + i * 10), s, fill=INK, fnt=FONT_11)
    for i, s in enumerate(sections):
        threshold = 0.28 + i * 0.16
        if p > threshold:
            x = 560 + i * 95
            y = 302
            d.rectangle([x, y, x + 82, y + 90], fill=(255, 255, 255), outline=BLUE, width=2)
            text(d, (x + 41, y + 24), str(i + 1), fill=BLUE, fnt=FONT_20)
            text(d, (x + 41, y + 50), s, fill=INK, fnt=FONT_11)
            d.rectangle([x + 10, y + 64, x + 72, y + 76], fill=YELLOW)
    return (130, 178)


def distribute_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    xs = [150, lerp(292, 338, p), lerp(475, 526, p), 710]
    for i, x in enumerate(xs):
        shape(d, (x, 260, x + 92, 322), [ORANGE, GREEN, BLUE, NAVY][i], selected=True)
    return (150, 260)


def stack_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(150, 245), (360, 190), (610, 315)]
    ends = [(150, 245), (250, 245), (350, 245)]
    cols = [ORANGE, GREEN, BLUE]
    for i, (s, e) in enumerate(zip(starts, ends)):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        shape(d, (x, y, x + 100, y + 62), cols[i], f"{i + 1}", selected=True)
    text(d, (500, 405), "Stack uses selection order. Shape 1 stays fixed.", fill=GREY, fnt=FONT_16)
    return (150, 245)


def spacing_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    xs = [155, lerp(330, 305, p), lerp(560, 455, p)]
    for x, c in zip(xs, [ORANGE, GREEN, BLUE]):
        shape(d, (x, 250, x + 105, 314), c, selected=True)
    return (155, 250)


def matrix_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(160, 190), (405, 155), (620, 235), (275, 365), (520, 345), (720, 390)]
    ends = [(300, 220), (378, 220), (456, 220), (300, 276), (378, 276), (456, 276)]
    cols = [ORANGE, GREEN, BLUE, NAVY, RED, GREY]
    for i, (s, e, c) in enumerate(zip(starts, ends, cols)):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        shape(d, (x, y, x + 78, y + 56), c, str(i + 1), selected=True, radius=6)
    text(d, (600, 438), "row by row in selection order", fill=GREY, fnt=FONT_15)
    return (160, 190)


def golden_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (210, 165, 750, 430)
    rounded(d, master, fill=(255, 255, 255), outline=MASTER, width=3, radius=4)
    selection_frame(d, master, master=True)
    rounded(d, (master[0], master[1] - 37, master[0] + 120, master[1] - 9), fill=(255, 255, 255), outline=MASTER, width=2, radius=4)
    text(d, (master[0] + 12, master[1] - 23), "MASTER", fill=MASTER, anchor="lm", fnt=FONT_16)
    final_y = golden_top(master[1], master[3] - master[1], 78)
    y = lerp(320, final_y, p)
    shape(d, (360, y, 600, y + 78), NAVY, "Content", selected=True)
    d.line([210, final_y, 750, final_y], fill=YELLOW, width=4)
    return (360, 320)


def magic_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    cx, cy = 412.5, 289.5
    scale = lerp(1.0, 1.2, p)
    for ox, oy, w, h, c in [(-150, -65, 105, 58, ORANGE), (-20, -75, 155, 70, GREEN), (-120, 30, 240, 48, BLUE)]:
        box = scale_box_about((cx + ox, cy + oy, cx + ox + w, cy + oy + h), (cx, cy), scale)
        shape(d, box, c, selected=True)
    return (270, 223)


def match_size_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master_box = (620, 285, 785, 370)
    targets = [
        ((160, 180, 270, 240), (132.5, 167.5, 297.5, 252.5), ORANGE),
        ((390, 270, 480, 353), (352.5, 269, 517.5, 354), GREEN),
    ]
    for start, end, col in targets:
        box = tuple(lerp(a, b, p) for a, b in zip(start, end))
        shape(d, box, col, "Target", selected=True)
    shape(d, master_box, NAVY, "Master", master=True)
    return (160, 180)


def dock_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (620, 245, 770, 335)
    x = lerp(260, 500, p)
    shape(d, (x, 252, x + 120, 328), ORANGE, "Target", selected=True)
    shape(d, master, NAVY, "Master", master=True)
    return (260, 252)


def fill_gap_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (650, 245, 795, 335)
    left = 220
    right = lerp(360, 650, p)
    shape(d, (left, 258, right, 322), ORANGE, "Target", selected=True)
    shape(d, master, NAVY, "Master", master=True)
    return (220, 258)


def stretch_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    master = (615, 205, 800, 385)
    left = 235
    right = lerp(420, 800, p)
    shape(d, (left, 260, right, 326), ORANGE, "Target", selected=True)
    rounded(d, master, fill=None, outline=MASTER, width=3, radius=6)
    selection_frame(d, master, master=True)
    rounded(d, (master[0], master[1] - 37, master[0] + 120, master[1] - 9), fill=(255, 255, 255), outline=MASTER, width=2, radius=4)
    text(d, (master[0] + 12, master[1] - 23), "MASTER", fill=MASTER, anchor="lm", fnt=FONT_16)
    return (235, 260)


def swap_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    a = (lerp(190, 610, p), lerp(230, 330, p))
    b = (lerp(610, 190, p), lerp(330, 230, p))
    shape(d, (a[0], a[1], a[0] + 130, a[1] + 68), ORANGE, "A", selected=True)
    shape(d, (b[0], b[1], b[0] + 130, b[1] + 68), BLUE, "B", selected=True)
    return (190, 230)


def place_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    d.rectangle([480, 130, 910, 494], outline=(220, 226, 236), width=3)
    x, y = lerp(145, 520, p), lerp(230, 210, p)
    w, h = lerp(150, 330, p), lerp(95, 245, p)
    shape(d, (x, y, x + w, y + h), BLUE, "Selection", selected=True)
    return (145, 230)


def slice_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    x, y, w, h = 245, 180, 380, 225
    q = max(0.0, min(1.0, (p - 0.05) / 0.95))
    gap = lerp(0, 12, q)
    for r in range(3):
        for c in range(4):
            bx = x + c * (w / 4)
            by = y + r * (h / 3)
            cell = [bx + gap / 2, by + gap / 2, bx + w / 4 - gap / 2, by + h / 3 - gap / 2]
            d.rectangle(cell, fill=BLUE, outline=(255, 255, 255), width=2)
    if p < 0.05:
        shape(d, (x, y, x + w, y + h), BLUE, "Original", selected=True)
    elif q < 0.55:
        d.rectangle([x - 7, y - 7, x + w + 7, y + h + 7], outline=TARGET, width=3)
        text(d, (x + w / 2, y + h / 2), "Original", fill=(255, 255, 255), fnt=FONT_20)
    return (245, 180)


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
    return (310, 210)


def block_arrow(
    d: ImageDraw.ImageDraw,
    x: float,
    y: float,
    width: float,
    height: float,
    fill: tuple[int, int, int],
    selected: bool,
    master: bool = False,
) -> None:
    head = min(34, width * 0.23)
    d.polygon(
        [(x, y), (x + width - head, y), (x + width, y + height / 2), (x + width - head, y + height), (x, y + height)],
        fill=fill,
        outline=(255, 255, 255),
    )
    show_selection = SELECTION_COMPLETE or any(
        x - 12 <= px <= x + width + 12 and y - 12 <= py <= y + height + 12
        for px, py in ACTIVE_SELECTION_TARGETS
    )
    if show_selection:
        col = MASTER if master else TARGET
        d.rectangle([x - 6, y - 6, x + width + 6, y + height + 6], outline=col, width=3)
    if master:
        text(d, (x + width / 2, y - 20), "3 MASTER", fill=MASTER, fnt=FONT_13)


def chain_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    starts = [(150, 215, 145, 62), (385, 185, 115, 112), (640, 295, 150, 84)]
    ends = [(150, 250, 145, 84), (295, 250, 115, 84), (410, 250, 150, 84)]
    for i, (start, end) in enumerate(zip(starts, ends)):
        vals = [lerp(a, b, p) for a, b in zip(start, end)]
        block_arrow(d, *vals, [ORANGE, GREEN, BLUE][i], selected=i < 2, master=i == 2)
        if i < 2:
            text(d, (vals[0] + vals[2] / 2, vals[1] - 20), f"{i + 1} TARGET", fill=TARGET, fnt=FONT_13)
    text(d, (600, 420), "Master defines height and arrow geometry; chain closes left to right", fill=GREY, fnt=FONT_13)
    return (150, 215)


def snap_table_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    table(d, 220, 170, 4, 3, 120, 70)
    starts = [(250, 190), (430, 238), (625, 358)]
    ends = [(280, 205), (400, 205), (640, 345)]
    for s, e, col in zip(starts, ends, [ORANGE, GREEN, BLUE]):
        x, y = lerp(s[0], e[0], p), lerp(s[1], e[1], p)
        d.ellipse([x - 16, y - 16, x + 16, y + 16], fill=col, outline=(255, 255, 255), width=2)
    return (250, 190)


def colors_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    target_col = mix((220, 224, 230), ORANGE, p)
    shape(d, (190, 230, 335, 310), target_col, "Target", selected=True)
    shape(d, (610, 230, 755, 310), ORANGE, "Master", master=True)
    return (190, 230)


def text_tools_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    box = (210, 205, 640, 345)
    rounded(d, box, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
    margin = lerp(12, 36, p)
    text(d, (box[0] + margin, box[1] + margin), "Text box margins", anchor="la", fnt=FONT_24, fill=NAVY)
    d.rectangle([box[0] + margin, box[1] + margin, box[2] - margin, box[3] - margin], outline=(220, 226, 236), width=2)
    return (210, 205)


def split_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    q = max(0.0, min(1.0, (p - 0.2) / 0.8))
    first = (220, 205, 690, 285)
    second = (220, 300, 690, 380)
    if q < 0.08:
        rounded(d, first, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
        text(d, (455, 245), "First thought | second thought", fill=NAVY, fnt=FONT_24)
    else:
        rounded(d, first, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
        text(d, (455, 245), "First thought", fill=NAVY, fnt=FONT_20)
        if q > 0.45:
            rounded(d, second, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
            text(d, (455, 340), "second thought", fill=NAVY, fnt=FONT_20)
    if q < 0.08:
        d.line([463, 212, 463, 278], fill=RED, width=3)
    return (463, 245)


def fit_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    box = (205, 205, lerp(710, 465, p), lerp(385, 292, p))
    rounded(d, box, fill=(255, 255, 255), outline=BLUE, width=3, radius=5)
    text(d, ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), "Short label", fill=NAVY, fnt=FONT_30)
    return (205, 205)


def hide_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    for i, (x, y) in enumerate([(180, 190), (310, 230), (460, 180), (590, 265), (720, 220)]):
        base = ORANGE if i % 2 == 0 else BLUE
        fill = mix(base, (255, 255, 255), p)
        outline = mix((255, 255, 255), (255, 255, 255), p)
        if p < 0.94:
            rounded(d, (x, y, x + 95, y + 65), fill=fill, outline=outline, width=2, radius=8)
    return (180, 190)


def paste_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    rounded(d, (85, 145, 265, 188), fill=LIGHT_GREY, outline=(185, 198, 216), width=2, radius=5)
    text(d, (175, 166), "Clipboard: KPI card", fill=NAVY, fnt=FONT_15)
    text(d, (170, 215), "thumbnail pane", fill=GREY, fnt=FONT_13)
    for i in range(3):
        y = 238 + i * 76
        d.rectangle([105, y, 235, y + 64], fill=(255, 255, 255), outline=BLUE, width=3)
        text(d, (91, y + 32), str(i + 1), fill=GREY, fnt=FONT_13)
        if p > 0.05:
            shape(d, (152, y + 18, 207, y + 48), ORANGE)
    shape(d, (430, 235, 610, 335), ORANGE, "KPI card")
    text(d, (520, 370), "same clipboard object pasted on each selected slide", fill=GREY, fnt=FONT_15)
    return (170, 270)


def chart_create_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    table(d, 120, 185, 4, 4, 70, 44)
    for c, val in enumerate(["Q1", "Q2", "Q3"]):
        text(d, (120 + (c + 1) * 70 + 35, 207), val, fnt=FONT_13)
    text(d, (260, 382), "source table remains editable", fill=GREY, fnt=FONT_13)
    if p > 0.05:
        bars(d, 560, 390, [55, 85, 70], BLUE)
        text(d, (655, 178), "editable shape chart", fill=NAVY, fnt=FONT_15)
    return (185, 185)


def chart_rebuild_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    vals = [lerp(45, 60, p), lerp(70, 42, p), lerp(52, 90, p)]
    if p < 0.78:
        table(d, 165, 190, 3, 4, 78, 40)
        text(d, (282, 370), "temporary Edit Data table", fill=GREY, fnt=FONT_13)
    bars(d, 470, 390, vals, BLUE)
    text(d, (570, 178), "chart rebuilt in place", fill=NAVY, fnt=FONT_15)
    return (470, 300)


def samples_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    names = ["Column", "Waterfall", "Mekko", "Line"]
    for i, name in enumerate(names):
        threshold = 0.08 + i * 0.18
        if p <= threshold:
            continue
        x = 210 + i * 150
        y = 210 + (i % 2) * 120
        d.rectangle([x, y, x + 108, y + 78], fill=(255, 255, 255), outline=BLUE, width=2)
        text(d, (x + 54, y + 25), name, fnt=FONT_13)
        bars(d, int(x + 22), int(y + 62), [15, 24, 20], [ORANGE, GREEN, BLUE, NAVY][i])
    return (180, 210)


def style_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    col = mix(BLUE, GREEN, p)
    bars(d, 180, 350, [55, 80, 48], col)
    bars(d, 560, 350, [65, 42, 75], col)
    if p > 0.5:
        for x in (180, 560):
            d.line([x - 5, 368, x + 174, 368], fill=NAVY, width=2)
    text(d, (480, 430), "all Chart Aid charts adopt the current style", fill=GREY, fnt=FONT_15)
    return (265, 270)


def recolor_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    baseline = 400
    for i, (lower, upper) in enumerate([(45, 30), (62, 38), (52, 44), (70, 26)]):
        x = 285 + i * 90
        d.rectangle([x, baseline - lower, x + 48, baseline], fill=BLUE)
        upper_color = mix(ORANGE, RED, p)
        d.rectangle([x, baseline - lower - upper, x + 48, baseline - lower], fill=upper_color)
    d.line([270, baseline, 690, baseline], fill=GREY, width=2)
    if p < 0.05:
        d.rectangle([375, baseline - 62 - 38, 423, baseline - 62], outline=TARGET, width=4)
    text(d, (480, 445), "one selected segment recolors its full series", fill=GREY, fnt=FONT_15)
    return (399, 310)


def diff_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    bars(d, 260, 390, [45, 78, 62], BLUE)
    if p > 0.35:
        x, y1, y2 = 333, 300, 266
        d.line([294, y1, x, y1], fill=GREY, width=2)
        d.line([x, y2, 372, y2], fill=GREY, width=2)
        arrow(d, (x, y1), (x, y2), fill=RED, width=3)
        arrow(d, (x, y2), (x, y1), fill=RED, width=3)
        text(d, (341, (y1 + y2) / 2), "+17", fill=RED, anchor="lm", fnt=FONT_18)
    return (260, 300)


def chart_elements_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    cx, cy = 430, 285
    if p > 0.05:
        d.ellipse([cx - 52, cy - 52, cx + 52, cy + 52], fill=(255, 255, 255), outline=BLUE, width=4)
        d.pieslice([cx - 52, cy - 52, cx + 52, cy + 52], -90, 180, fill=BLUE)
        d.ellipse([cx - 30, cy - 30, cx + 30, cy + 30], fill=(255, 255, 255))
        text(d, (430, 375), "75%", fill=NAVY, fnt=FONT_24)
    return (378, 233)


def cycle_scene(d: ImageDraw.ImageDraw, p: float) -> tuple[float, float]:
    states = [0.0, 0.25, 0.5, 0.75, 1.0]
    state = states[min(4, int(p * 4.99))]
    cx, cy = 430, 285
    d.ellipse([cx - 54, cy - 54, cx + 54, cy + 54], fill=(255, 255, 255), outline=BLUE, width=4)
    if state > 0:
        d.pieslice([cx - 54, cy - 54, cx + 54, cy + 54], -90, -90 + 360 * state, fill=BLUE)
    d.ellipse([cx - 31, cy - 31, cx + 31, cy + 31], fill=(255, 255, 255))
    text(d, (430, 375), f"{int(state * 100)}%", fill=NAVY, fnt=FONT_24)
    return (376, 231)


def workflow(
    clicks: list[tuple[float, float]],
    command: str,
    *,
    selection_labels: list[str] | None = None,
    menu_items: tuple[str, ...] = (),
    menu_index: int = 0,
    prompt: tuple[str, str] | None = None,
    end_result: float = 1.0,
) -> list[Action]:
    labels = selection_labels or [f"Select object {i + 1}" for i in range(len(clicks))]
    actions = [Action("select", labels[i], target=pos) for i, pos in enumerate(clicks)]
    actions.append(Action("ribbon", f"Click {command}", target="ribbon"))
    if menu_items:
        actions.append(
            Action(
                "menu",
                f"Choose {menu_items[menu_index]}",
                target="menu",
                menu_items=menu_items,
                menu_index=menu_index,
            )
        )
    if prompt:
        actions.append(Action("prompt", prompt[0], target="prompt", prompt_value=prompt[1], frames=8))
    actions.extend(
        [
            Action("animate", "Apply command", frames=18, end_result=end_result),
            Action("hold", "Result", frames=8),
        ]
    )
    return actions


def grid_workflow(
    click: tuple[float, float],
    command: str,
    rows: str,
    columns: str,
    gap: str,
) -> list[Action]:
    return [
        Action("select", "Select the source shape", target=click),
        Action("ribbon", f"Click {command}", target="ribbon"),
        Action("prompt", "Number of rows", target="prompt", prompt_value=rows, frames=7),
        Action("prompt", "Number of columns", target="prompt", prompt_value=columns, frames=7),
        Action("prompt", "Gap between pieces (cm)" if command == "Slice" else "Gap between copies (cm)", target="prompt", prompt_value=gap, frames=7),
        Action("animate", "Apply command", frames=18, end_result=1.0),
        Action("hold", "Result", frames=8),
    ]


def contract(reference: str, result: str, invariant: str) -> BehaviorContract:
    return BehaviorContract(reference, result, invariant)


DEMOS = [
    Demo("demo-master-selection-order.gif", "sa_align_left", "Left", "Align to Master", master_order_scene, actions=workflow([(180, 205), (365, 295), (595, 405)], "Left", selection_labels=["Select target 1", "Select target 2", "Select Master last"]), contract=contract("master-last", "targets align left", "Master position is unchanged")),
    Demo("demo-single-object-slide-reference.gif", "sa_align_center", "Center", "Slide reference", single_slide_scene, actions=workflow([(180, 190)], "Center", selection_labels=["Select one object"]), contract=contract("slide-single", "object centers horizontally", "vertical position is unchanged")),
    Demo("demo-to-slide-reference.gif", "sa_to_slide", "To Slide", "Center on slide", to_slide_scene, actions=workflow([(155, 190), (285, 250), (210, 340)], "To Slide", selection_labels=["Select object 1", "Select object 2", "Select object 3"], menu_items=("Left", "Right", "Top", "Bottom", "Center", "Middle"), menu_index=4), contract=contract("slide-explicit", "each object centers horizontally", "each vertical position is unchanged")),
    Demo("demo-reusable-libraries.gif", "sa_elements", "My Elements", "Save and insert", reusable_scene, actions=[
        Action("select", "Select reusable KPI card", target=(150, 225)),
        Action("ribbon", "Open My Elements", target="ribbon"),
        Action("menu", "Choose Add Selection to My Elements...", target="menu", menu_items=("KPI card", "Add Selection to My Elements...", "Manage Library..."), menu_index=1),
        Action("prompt", "Name for this element", target="prompt", prompt_value="KPI card", frames=8),
        Action("animate", "Save to My Elements", frames=10, end_result=0.5),
        Action("ribbon", "Open My Elements again", target="ribbon"),
        Action("menu", "Choose KPI card", target="menu", menu_items=("KPI card", "Add Selection to My Elements...", "Manage Library..."), menu_index=0),
        Action("animate", "Insert saved element", frames=10, end_result=1.0),
        Action("hold", "Result", frames=8),
    ], contract=contract("library", "selection is saved then inserted", "inserted shapes remain editable")),
    Demo("demo-my-elements-save.gif", "sa_elements", "My Elements", "Save selection", reusable_save_scene, actions=workflow([(150, 225)], "My Elements", selection_labels=["Select reusable KPI card"], menu_items=("Add Selection to My Elements...", "Manage Library..."), menu_index=0, prompt=("Name for this element", "KPI card")), contract=contract("library", "selection becomes a named menu item", "source shapes remain on the slide")),
    Demo("demo-my-elements-insert.gif", "sa_elements", "My Elements", "Insert saved item", reusable_insert_scene, actions=workflow([], "My Elements", menu_items=("KPI card", "Add Selection to My Elements...", "Manage Library..."), menu_index=0), contract=contract("library", "saved item is pasted at stored coordinates", "library source remains stored")),
    Demo("demo-format-painter-select-similar.gif", "sa_painter", "Format Painter", "Copy Master style", painter_scene, actions=workflow([(175, 205), (390, 205), (650, 315)], "Format Painter", selection_labels=["Select target 1", "Select target 2", "Select Master last"]), contract=contract("master-last", "targets receive full Master formatting", "target geometry and Master remain unchanged")),
    Demo("demo-select-similar.gif", "sa_similar", "Select Similar", "Select matching shapes", select_similar_scene, actions=workflow([(650, 335)], "Select Similar", selection_labels=["Select the Master"], menu_items=("Same Shape Type", "Same Fill Color", "Same Type + Fill"), menu_index=2), contract=contract("master-last", "matching slide shapes become selected", "nonmatching shapes remain unselected")),
    Demo("demo-agenda.gif", "sa_agenda", "Agenda", "Build section slides", agenda_scene, actions=workflow([], "Agenda"), contract=contract("sections", "overview and one separator per nonempty section are inserted", "previous generated agenda slides are replaced")),
    Demo("demo-align-to-master.gif", "sa_align_left", "Left", "Align to Master", align_scene, actions=workflow([(250, 184), (420, 404), (640, 292)], "Left", selection_labels=["Select target 1", "Select target 2", "Select Master last"]), contract=contract("master-last", "target left edges equal Master left", "Master and target vertical positions are unchanged")),
    Demo("demo-distribute-horizontal.gif", "sa_dist_h", "Distribute H", "Equal horizontal gaps", distribute_scene, actions=workflow([(150, 260), (292, 260), (475, 260), (710, 260)], "Distribute H"), contract=contract("current-position-order", "horizontal gaps become equal", "outer envelope is unchanged")),
    Demo("demo-stack-horizontal.gif", "sa_stack", "Stack", "Horizontally", stack_scene, actions=workflow([(150, 245), (360, 190), (610, 315)], "Stack", selection_labels=["Select first: stays fixed", "Select second", "Select third"], menu_items=("Horizontally", "Vertically", "Horizontally + Gap...", "Vertically + Gap..."), menu_index=0), contract=contract("selection-order", "objects touch in selected sequence", "first selected object stays fixed")),
    Demo("demo-spacing-exact-gap.gif", "sa_spacing", "Spacing", "Horizontal gap", spacing_scene, actions=workflow([(155, 250), (330, 250), (560, 250)], "Spacing", menu_items=("Horizontal...", "Vertical..."), menu_index=0, prompt=("Horizontal spacing (cm)", "0.50")), contract=contract("current-position-order", "requested horizontal gap is applied", "object sizes are unchanged")),
    Demo("demo-matrix-grid.gif", "sa_matrix", "Matrix", "Arrange as grid", matrix_scene, actions=workflow([(160, 190), (405, 155), (620, 235), (275, 365), (520, 345), (720, 390)], "Matrix", selection_labels=[f"Select object {i}" for i in range(1, 7)]), contract=contract("selection-order", "objects fill a near-square grid row by row", "object sizes are unchanged")),
    Demo("demo-golden-canon-placement.gif", "sa_golden", "Golden Canon", "Place in Master", golden_scene, actions=workflow([(360, 320), (210, 165)], "Golden Canon", selection_labels=["Select target", "Select Master last"]), contract=contract("master-last", "bottom margin equals twice top margin", "Master is unchanged")),
    Demo("demo-magic-resizer.gif", "sa_magic", "Magic Resizer", "Scale arrangement", magic_scene, actions=workflow([(295, 218), (425, 208), (325, 312)], "Magic Resizer", prompt=("Resize to (% of current size)", "120")), contract=contract("selection-bounds", "sizes, positions, and text scale around selection center", "arrangement proportions are preserved")),
    Demo("demo-match-size.gif", "sa_size", "Width + Height", "Match Master size", match_size_scene, actions=workflow([(160, 180), (390, 270), (620, 285)], "Width + Height", selection_labels=["Select target 1", "Select target 2", "Select Master last"]), contract=contract("master-last", "target width and height match Master", "each target center and Master are unchanged")),
    Demo("demo-dock-fill-stretch.gif", "sa_dock", "Dock", "Touch the Master", dock_scene, actions=workflow([(260, 252), (620, 245)], "Dock", selection_labels=["Select target", "Select Master last"], menu_items=("Dock Left", "Dock Right", "Dock Up", "Dock Down"), menu_index=1), contract=contract("master-last", "target touches Master", "Master is unchanged")),
    Demo("demo-dock.gif", "sa_dock", "Dock", "Touch the Master", dock_scene, actions=workflow([(260, 252), (620, 245)], "Dock", selection_labels=["Select target", "Select Master last"], menu_items=("Dock Left", "Dock Right", "Dock Up", "Dock Down"), menu_index=1), contract=contract("master-last", "target touches Master", "Master is unchanged")),
    Demo("demo-fill-gap.gif", "sa_fill", "Fill Gap", "Extend to Master", fill_gap_scene, actions=workflow([(220, 258), (650, 245)], "Fill Gap", selection_labels=["Select target", "Select Master last"], menu_items=("Leftwards", "Rightwards", "Upwards", "Downwards"), menu_index=1), contract=contract("master-last", "near target edge extends to Master", "target far edge and Master are unchanged")),
    Demo("demo-stretch.gif", "sa_stretch", "Stretch", "Extend to far edge", stretch_scene, actions=workflow([(235, 260), (615, 205)], "Stretch", selection_labels=["Select target", "Select Master last"], menu_items=("Left", "Right", "Top", "Bottom"), menu_index=1), contract=contract("master-last", "target extends to Master's far edge", "target opposite edge and Master are unchanged")),
    Demo("demo-swap-positions.gif", "sa_swap", "Swap", "Exchange positions", swap_scene, actions=workflow([(190, 230), (610, 330)], "Swap", menu_items=("At Centers", "At Top-Left", "At Top-Right", "At Bottom-Left", "At Bottom-Right"), menu_index=0), contract=contract("selection-order", "objects exchange full two-dimensional center positions", "object sizes are unchanged")),
    Demo("demo-place-on-slide.gif", "sa_place", "Place on Slide", "Right half", place_scene, actions=workflow([(145, 230)], "Place on Slide", menu_items=("Left Half", "Right Half", "Top Half", "Bottom Half"), menu_index=1), contract=contract("slide-region", "single object fills selected region with margin", "object remains editable")),
    Demo("demo-slice-multiply.gif", "sa_slice", "Slice", "Split into pieces", slice_scene, actions=grid_workflow((245, 180), "Slice", "3", "4", "0.10"), contract=contract("selection", "one shape becomes an editable grid", "original footprint is preserved")),
    Demo("demo-slice.gif", "sa_slice", "Slice", "Split into pieces", slice_scene, actions=grid_workflow((245, 180), "Slice", "3", "4", "0.10"), contract=contract("selection", "one shape becomes an editable grid", "original footprint is preserved")),
    Demo("demo-multiply.gif", "sa_multiply", "Multiply", "Create copies", multiply_scene, actions=grid_workflow((310, 210), "Multiply", "3", "4", "0.20"), contract=contract("selection", "copies fill the requested grid", "original shape remains first")),
    Demo("demo-shape-helpers.gif", "sa_chain", "Process Chain", "Connect arrows", chain_scene, actions=workflow([(150, 215), (385, 185), (640, 295)], "Process Chain", selection_labels=["Select target arrow 1", "Select target arrow 2", "Select Master arrow last"]), contract=contract("master-last-and-position-order", "targets adopt Master geometry and all arrows close left to right", "Master defines style but may move horizontally in the chain")),
    Demo("demo-snap-to-table.gif", "sa_table", "Snap to Table", "Center in cell", snap_table_scene, actions=workflow([(250, 190), (430, 238), (625, 358)], "Snap to Table", menu_items=("Center in Cell", "Left in Cell...", "Right in Cell..."), menu_index=0), contract=contract("table-cell", "objects center in the cells containing their centers", "objects outside the table do not move")),
    Demo("demo-color-tools.gif", "sa_pick", "Pick from Master", "Copy colors", colors_scene, actions=workflow([(190, 230), (610, 230)], "Pick from Master", selection_labels=["Select target", "Select Master last"], menu_items=("Fill + Line + Font", "Fill only", "Line only", "Font only"), menu_index=0), contract=contract("master-last", "target receives Master colors", "Master and target geometry are unchanged")),
    Demo("demo-text-tools.gif", "sa_margins", "Set Margins", "Text box margins", text_tools_scene, actions=workflow([(210, 205)], "Set Margins", prompt=("Internal margin, all sides (cm)", "0.30")), contract=contract("selection", "all four text margins receive one value", "shape bounds are unchanged")),
    Demo("demo-split-at-cursor.gif", "sa_split", "Split at Cursor", "Split text box", split_scene, actions=workflow([(463, 245)], "Split at Cursor", selection_labels=["Place the text cursor at the split"]), contract=contract("text-cursor", "original keeps first text and duplicate below keeps second", "formatting and horizontal position are preserved")),
    Demo("demo-fit-to-text.gif", "sa_fit", "Fit to Text", "Resize text box", fit_scene, actions=workflow([(205, 205)], "Fit to Text"), contract=contract("selection", "shape autosizes to contained text", "text and formatting are unchanged")),
    Demo("demo-view-cleanup.gif", "sa_hide", "Hide Objects", "Temporary cleanup", hide_scene, actions=workflow([(180, 190), (310, 230), (460, 180), (590, 265), (720, 220)], "Hide Objects"), contract=contract("selection", "selected objects become temporarily hidden", "positions and layer order are retained")),
    Demo("demo-paste-on-slides.gif", "sa_paste", "Paste on Slides", "Repeat paste", paste_scene, actions=workflow([(170, 270), (170, 346), (170, 422)], "Paste on Slides", selection_labels=["Select slide thumbnail 1", "Cmd-click slide thumbnail 2", "Cmd-click slide thumbnail 3"]), contract=contract("clipboard-slides", "clipboard content is pasted on every selected slide", "clipboard object and slide selection remain available")),
    Demo("demo-chart-create-from-table.gif", "sa_ch_col", "Column", "Create from table", chart_create_scene, actions=workflow([(185, 185)], "Column", selection_labels=["Select the source table"]), contract=contract("chart-data", "shape-based chart is created from table data", "normal source table remains")),
    Demo("demo-chart-rebuild.gif", "sa_ch_rebuild", "Rebuild", "Update chart", chart_rebuild_scene, actions=workflow([(165, 190), (470, 300)], "Rebuild", selection_labels=["Select temporary Edit Data table", "Select chart"]), contract=contract("chart-data", "chart is rebuilt in place from edited values", "temporary Edit Data table is removed")),
    Demo("demo-chart-samples.gif", "sa_ch_samples", "Sample Slides", "Insert examples", samples_scene, actions=workflow([], "Sample Slides"), contract=contract("presentation", "one live example slide per chart type is inserted", "existing slides remain")),
    Demo("demo-chart-annotations-style.gif", "sa_ch_restyle", "Restyle All", "Apply chart style", style_scene, actions=workflow([], "Restyle All"), contract=contract("presentation-charts", "all Chart Aid charts adopt current style", "data, position, size, and manual recolors are retained")),
    Demo("demo-recolor-series-click.gif", "sa_ch_recolor", "Recolor Series", "Selected series", recolor_scene, actions=workflow([(399, 310)], "Recolor Series", selection_labels=["Select one segment in the series"], prompt=("Choose series color", "Red")), contract=contract("chart-series", "every element in selected series receives the color", "other series and chart geometry are unchanged")),
    Demo("demo-difference-arrow-clicks.gif", "sa_ch_diff", "Difference", "Compare two bars", diff_scene, actions=workflow([(277, 300), (389, 266)], "Difference", selection_labels=["Select first bar", "Cmd-click second bar"]), contract=contract("chart-elements", "vertical double-arrow, helper lines, and data-derived label are added", "bars and chart data are unchanged")),
    Demo("demo-chart-elements.gif", "sa_ch_harvey", "Harvey Ball", "Insert element", chart_elements_scene, actions=workflow([], "Harvey Ball", prompt=("Completion percentage", "75")), contract=contract("slide", "one fixed-percentage Harvey ball is inserted", "element remains editable")),
    Demo("demo-cycle-state-clicks.gif", "sa_ch_cycle", "Cycle State", "Advance state", cycle_scene, actions=[
        Action("select", "Select the Harvey ball", target=(376, 231)),
        Action("ribbon", "Click Cycle State: 0% to 25%", target="ribbon"),
        Action("animate", "Advance one state", frames=5, end_result=0.25),
        Action("hold", "25%", frames=5),
        Action("ribbon", "Click Cycle State: 25% to 50%", target="ribbon"),
        Action("animate", "Advance one state", frames=5, end_result=0.5),
        Action("hold", "50%", frames=5),
        Action("ribbon", "Click Cycle State: 50% to 75%", target="ribbon"),
        Action("animate", "Advance one state", frames=5, end_result=0.75),
        Action("hold", "75%", frames=5),
        Action("ribbon", "Click Cycle State: 75% to 100%", target="ribbon"),
        Action("animate", "Advance one state", frames=5, end_result=1.0),
        Action("hold", "100%", frames=8),
    ], contract=contract("chart-element", "each command click advances exactly one state", "selected element stays in place")),
]


def validate_demos() -> None:
    names = [demo.name for demo in DEMOS]
    if len(names) != len(set(names)):
        raise ValueError("documentation GIF names must be unique")
    for demo in DEMOS:
        if demo.icon not in ICON_CONTEXT:
            raise ValueError(f"{demo.name}: unknown Ribbon icon {demo.icon}")
        if demo.contract is None:
            raise ValueError(f"{demo.name}: missing behavior contract")
        actions = demo.actions or default_actions(demo)
        if not any(action.kind == "ribbon" for action in actions):
            raise ValueError(f"{demo.name}: workflow never clicks the Ribbon")
        results = [action.end_result for action in actions if action.end_result is not None]
        if not results or results[-1] != 1.0:
            raise ValueError(f"{demo.name}: workflow does not reach its final result")
        for action in actions:
            if action.kind == "menu":
                if not action.menu_items or not 0 <= action.menu_index < len(action.menu_items):
                    raise ValueError(f"{demo.name}: invalid menu action")


def make_ribbon_tour() -> None:
    frames: list[Image.Image] = []
    durations: list[int] = []
    for tab, groups in RIBBON_GROUPS.items():
        for group, buttons in groups.items():
            for page_start in range(0, len(buttons), 8):
                page = buttons[page_start : page_start + 8]
                img = frame_base()
                draw_ribbon(img, page[0][0], page[0][1])
                d = ImageDraw.Draw(img)
                text(d, (480, 150), f"{tab} - {group}", fill=NAVY, fnt=FONT_24)
                if len(buttons) > 8:
                    text(d, (480, 176), f"Controls {page_start + 1}-{page_start + len(page)} of {len(buttons)}", fill=GREY, fnt=FONT_13)
                for i, (icon, label) in enumerate(page):
                    col = i % 4
                    row = i // 4
                    x = 150 + col * 190
                    y = 215 + row * 120
                    rounded(d, (x, y, x + 155, y + 88), fill=(255, 255, 255), outline=(202, 212, 226), width=2, radius=6)
                    icon_rgba = icon_image(icon, 42)
                    img.paste(icon_rgba, (x + 12, y + 16), icon_rgba)
                    text(d, (x + 66, y + 37), label, anchor="lm", fill=INK, fnt=FONT_15)
                text(
                    d,
                    (480, 480),
                    "Generated from apps/powerpoint/ribbon/customUI14.xml and shared/icons",
                    fill=GREY,
                    fnt=FONT_11,
                )
                frames.append(img.convert("P", palette=Image.Palette.ADAPTIVE, colors=128))
                durations.append(1250)
    path = OUT / "ribbon-tour.gif"
    frames[0].save(
        path,
        save_all=True,
        append_images=frames[1:],
        duration=durations,
        loop=0,
        optimize=False,
        disposal=2,
    )


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    validate_demos()
    for demo in DEMOS:
        make_gif(demo)
    make_ribbon_tour()
    print(f"{len(DEMOS)} workflow GIFs and the Ribbon tour rendered to {OUT}")


if __name__ == "__main__":
    main()
