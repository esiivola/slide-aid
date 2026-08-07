#!/usr/bin/env python3
"""Generate the Slide Aid ribbon icons.

The visual language uses a dark blue Master shape, light blue target
shapes, and red action arrows.

Outputs 32x32 PNGs (drawn 4x supersampled) to shared/icons/.
Run:  python3 scripts/make_icons.py
"""
import math
from pathlib import Path
from PIL import Image, ImageDraw

S = 4                 # supersample factor
SZ = 32               # final icon size
C = SZ * S            # canvas size

MASTER = (31, 73, 125, 255)      # Slide Aid dark blue
SHAPE_F = (220, 230, 241, 255)   # light blue fill
SHAPE_L = (79, 129, 189, 255)    # medium blue line
RED = (192, 80, 77, 255)         # action red
GREY = (89, 89, 89, 255)

OUT = Path(__file__).resolve().parent.parent / "shared" / "icons"
OUT.mkdir(parents=True, exist_ok=True)

def canvas():
    img = Image.new("RGBA", (C, C), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)

def save(img, name):
    img.resize((SZ, SZ), Image.LANCZOS).save(OUT / f"{name}.png")

def x(v): return v * S  # coords in 32px space

def master(d, box):
    d.rectangle([x(box[0]), x(box[1]), x(box[2]), x(box[3])], fill=MASTER)

def shape(d, box):
    d.rectangle([x(box[0]), x(box[1]), x(box[2]), x(box[3])],
                fill=SHAPE_F, outline=SHAPE_L, width=S)

def arrow(d, p1, p2, color=RED, w=2):
    a = (x(p1[0]), x(p1[1])); b = (x(p2[0]), x(p2[1]))
    d.line([a, b], fill=color, width=w * S)
    ang = math.atan2(b[1] - a[1], b[0] - a[0])
    hl = 4.2 * S
    for da in (2.6, -2.6):
        d.line([b, (b[0] - hl * math.cos(ang + da), b[1] - hl * math.sin(ang + da))],
               fill=color, width=w * S)

def dashes(d, p1, p2, color=RED, w=1, dash=2.2, gap=1.6):
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    t = 0.0
    while t < L:
        e = min(t + dash, L)
        d.line([(x(p1[0] + ux * t), x(p1[1] + uy * t)),
                (x(p1[0] + ux * e), x(p1[1] + uy * e))], fill=color, width=w * S)
        t = e + gap

# ================= Position =================
# Align: dark Master bar = the reference edge, light shapes flush on it.

def align_edges():
    img, d = canvas()                      # LEFT
    master(d, (3, 2, 6, 30))
    shape(d, (7, 4, 19, 10)); shape(d, (7, 13, 25, 19)); shape(d, (7, 22, 14, 28))
    save(img, "sa_align_left")
    save(img.transpose(Image.FLIP_LEFT_RIGHT), "sa_align_right")

    img, d = canvas()                      # TOP
    master(d, (2, 3, 30, 6))
    shape(d, (4, 7, 10, 19)); shape(d, (13, 7, 19, 25)); shape(d, (22, 7, 28, 14))
    save(img, "sa_align_top")
    save(img.transpose(Image.FLIP_TOP_BOTTOM), "sa_align_bottom")

def align_center():
    img, d = canvas()                      # CENTER (horizontal centering)
    shape(d, (8, 3, 24, 9)); shape(d, (11, 12, 21, 18)); shape(d, (5, 21, 27, 27))
    d.line([x(16), x(1), x(16), x(31)], fill=RED, width=S)
    save(img, "sa_align_center")

    img, d = canvas()                      # MIDDLE (vertical centering)
    shape(d, (3, 8, 9, 24)); shape(d, (12, 11, 18, 21)); shape(d, (21, 5, 27, 27))
    d.line([x(1), x(16), x(31), x(16)], fill=RED, width=S)
    save(img, "sa_align_middle")

def to_slide():
    img, d = canvas()
    d.rectangle([x(2), x(4), x(30), x(28)], outline=GREY, width=S)
    shape(d, (11, 10, 20, 18))
    arrow(d, (10, 14), (4, 14))
    save(img, "sa_to_slide")

def dock():
    img, d = canvas()
    master(d, (3, 7, 11, 25))
    shape(d, (19, 10, 29, 22))
    arrow(d, (18, 16), (12.5, 16))
    save(img, "sa_dock")

def distribute():
    img, d = canvas()
    shape(d, (2, 10, 8, 22)); shape(d, (13, 10, 19, 22)); shape(d, (24, 10, 30, 22))
    d.line([x(9.5), x(16), x(11.5), x(16)], fill=RED, width=S)
    d.line([x(20.5), x(16), x(22.5), x(16)], fill=RED, width=S)
    save(img, "sa_dist_h")
    save(img.rotate(90), "sa_dist_v")

def swap():
    img, d = canvas()
    master(d, (3, 3, 13, 13))
    shape(d, (19, 19, 29, 29))
    arrow(d, (15, 8), (26, 15))
    arrow(d, (17, 24), (6, 17))
    save(img, "sa_swap")

def stack():
    img, d = canvas()
    master(d, (2, 10, 10, 22)); shape(d, (10, 10, 21, 22)); shape(d, (21, 10, 30, 22))
    save(img, "sa_stack")

def matrix():
    img, d = canvas()
    for r in range(2):
        for c in range(2):
            shape(d, (4 + c * 13, 5 + r * 13, 15 + c * 13, 14 + r * 13))
    save(img, "sa_matrix")

def matrix_custom():
    # like matrix, but with red "settings" dots (custom columns/gaps)
    img, d = canvas()
    for r in range(2):
        for c in range(2):
            shape(d, (3 + c * 12, 4 + r * 12, 13 + c * 12, 12 + r * 12))
    for i in range(3):
        cx = 12 + i * 7
        d.ellipse([x(cx), x(27), x(cx + 3), x(30)], fill=RED)
    save(img, "sa_matrix_custom")

def place():
    # slide with highlighted left half (position presets)
    img, d = canvas()
    d.rectangle([x(2), x(4), x(30), x(28)], outline=GREY, width=S)
    master(d, (4, 6, 15, 26))
    dashes(d, (16, 5), (16, 27), color=RED, w=1)
    save(img, "sa_place")

def text_more():
    # text lines + red plus (more text tools)
    img, d = canvas()
    for i, (l, r) in enumerate([(4, 26), (4, 22), (4, 24)]):
        d.line([x(l), x(7 + i * 6), x(r), x(7 + i * 6)], fill=SHAPE_L, width=2 * S)
    d.line([x(23), x(21), x(23), x(29)], fill=RED, width=2 * S)
    d.line([x(19), x(25), x(27), x(25)], fill=RED, width=2 * S)
    save(img, "sa_txt_more")

def spacing():
    img, d = canvas()
    shape(d, (2, 8, 10, 24)); shape(d, (22, 8, 30, 24))
    d.line([x(11.5), x(10), x(11.5), x(22)], fill=RED, width=S)
    d.line([x(20.5), x(10), x(20.5), x(22)], fill=RED, width=S)
    arrow(d, (14.5, 16), (19, 16), w=1)
    arrow(d, (17.5, 16), (13, 16), w=1)
    save(img, "sa_spacing")

def golden():
    img, d = canvas()
    d.rectangle([x(5), x(2), x(27), x(30)], outline=MASTER, width=2 * S)
    shape(d, (9, 8, 23, 16))
    save(img, "sa_golden")

# ================= Size =================

def magic():
    img, d = canvas()
    shape(d, (8, 10, 20, 22))
    arrow(d, (20, 10), (28, 2))
    arrow(d, (8, 22), (2, 28))
    save(img, "sa_magic")

def size_w():
    img, d = canvas()
    shape(d, (3, 9, 29, 23))
    arrow(d, (16, 16), (26, 16)); arrow(d, (16, 16), (6, 16))
    save(img, "sa_width")
    save(img.rotate(90), "sa_height")

def size_wh():
    img, d = canvas()
    shape(d, (3, 5, 29, 27))
    arrow(d, (16, 16), (26, 7)); arrow(d, (16, 16), (6, 25))
    save(img, "sa_size")

def stretch():
    img, d = canvas()
    master(d, (24, 4, 30, 28))
    shape(d, (3, 10, 15, 22))
    arrow(d, (16, 16), (23, 16))
    save(img, "sa_stretch")

def fill_gap():
    img, d = canvas()
    master(d, (24, 4, 30, 28))
    shape(d, (3, 10, 15, 22))
    d.rectangle([x(15.5), x(10), x(23), x(22)], fill=(220, 230, 241, 140),
                outline=RED, width=S)
    save(img, "sa_fill")

def slice_():
    img, d = canvas()
    shape(d, (4, 6, 28, 26))
    d.line([x(16), x(6), x(16), x(26)], fill=RED, width=S)
    d.line([x(4), x(16), x(28), x(16)], fill=RED, width=S)
    save(img, "sa_slice")

def multiply():
    img, d = canvas()
    shape(d, (12, 12, 28, 28)); shape(d, (8, 8, 24, 24)); master(d, (4, 4, 20, 20))
    save(img, "sa_multiply")

# ================= Shape =================

def chevron(d, x0, y0, w, h, fill, outline):
    pts = [(x(x0), x(y0)), (x(x0 + w - 4), x(y0)), (x(x0 + w), x(y0 + h / 2)),
           (x(x0 + w - 4), x(y0 + h)), (x(x0), x(y0 + h)), (x(x0 + 4), x(y0 + h / 2))]
    d.polygon(pts, fill=fill, outline=outline, width=S)

def process_chain():
    img, d = canvas()
    chevron(d, 1, 10, 11, 12, SHAPE_F, SHAPE_L)
    chevron(d, 11, 10, 11, 12, SHAPE_F, SHAPE_L)
    chevron(d, 21, 10, 11, 12, MASTER, MASTER)
    save(img, "sa_chain")

def angles():
    img, d = canvas()
    master(d, (3, 20, 15, 28))
    sq = Image.new("RGBA", (C, C), (0, 0, 0, 0))
    ds = ImageDraw.Draw(sq)
    shape(ds, (17, 5, 29, 13))
    sq = sq.rotate(25, center=(x(23), x(9)), resample=Image.BICUBIC)
    img.alpha_composite(sq)
    d = ImageDraw.Draw(img)
    d.arc([x(4), x(6), x(22), x(24)], start=245, end=340, fill=RED, width=2 * S)
    save(img, "sa_angles")

def block_arrow():
    img, d = canvas()
    pts = [(3, 12), (18, 12), (18, 6), (29, 16), (18, 26), (18, 20), (3, 20)]
    d.polygon([(x(a), x(b)) for a, b in pts], fill=SHAPE_F, outline=SHAPE_L, width=S)
    save(img, "sa_blockarrow")

def rounded_rect():
    img, d = canvas()
    d.rounded_rectangle([x(4), x(7), x(28), x(25)], radius=x(6),
                        fill=SHAPE_F, outline=SHAPE_L, width=S)
    d.arc([x(4), x(7), x(16), x(19)], start=180, end=270, fill=RED, width=S)
    save(img, "sa_roundrect")

def table_snap():
    img, d = canvas()
    d.rectangle([x(2), x(4), x(30), x(28)], outline=GREY, width=S)
    d.line([x(16), x(4), x(16), x(28)], fill=GREY, width=S)
    d.line([x(2), x(16), x(30), x(16)], fill=GREY, width=S)
    d.ellipse([x(19.5), x(19.5), x(26.5), x(26.5)], fill=MASTER)
    save(img, "sa_table")

# ================= Color =================

def color_bar(d, color):
    d.rectangle([x(4), x(24), x(28), x(29)], fill=color)

def fill_color():
    img, d = canvas()
    d.rectangle([x(7), x(4), x(25), x(20)], fill=MASTER)
    color_bar(d, RED)
    save(img, "sa_fillcolor")

def line_color():
    img, d = canvas()
    d.rectangle([x(7), x(5), x(25), x(19)], outline=MASTER, width=2 * S)
    color_bar(d, RED)
    save(img, "sa_linecolor")

def font_color():
    img, d = canvas()
    d.line([x(9), x(20), x(16), x(4)], fill=GREY, width=2 * S)
    d.line([x(23), x(20), x(16), x(4)], fill=GREY, width=2 * S)
    d.line([x(12), x(14), x(20), x(14)], fill=GREY, width=2 * S)
    color_bar(d, RED)
    save(img, "sa_fontcolor")

def to_rgb():
    img, d = canvas()
    d.rectangle([x(3), x(9), x(11), x(23)], fill=(192, 80, 77, 255))
    d.rectangle([x(12), x(9), x(20), x(23)], fill=(155, 187, 89, 255))
    d.rectangle([x(21), x(9), x(29), x(23)], fill=(79, 129, 189, 255))
    save(img, "sa_torgb")

def to_theme():
    img, d = canvas()
    for i, c in enumerate([MASTER, SHAPE_L, (141, 179, 226, 255), SHAPE_F]):
        d.rectangle([x(3 + i * 7), x(9), x(9 + i * 7), x(23)], fill=c)
    save(img, "sa_totheme")

def pick():
    img, d = canvas()
    # eyedropper: bulb top-right, wide barrel tapering to tip bottom-left
    d.line([x(21), x(11), x(11), x(21)], fill=GREY, width=5 * S)   # barrel
    d.polygon([(x(13), x(17)), (x(15), x(19)), (x(7), x(25))], fill=GREY)  # tip
    d.ellipse([x(17), x(3), x(28), x(14)], fill=GREY)              # bulb
    master(d, (2, 22, 10, 30))                                     # color swatch at tip
    save(img, "sa_pick")

def info():
    img, d = canvas()
    d.ellipse([x(5), x(5), x(27), x(27)], outline=MASTER, width=2 * S)
    d.ellipse([x(14.4), x(9), x(17.6), x(12.2)], fill=MASTER)
    d.line([x(16), x(15), x(16), x(22)], fill=MASTER, width=3 * S)
    save(img, "sa_info")

# ================= Text =================

def txt_lines(d, x0, x1, ys, color=GREY):
    for Y in ys:
        d.line([x(x0), x(Y), x(x1), x(Y)], fill=color, width=S)

def split_text():
    img, d = canvas()
    shape(d, (2, 6, 14, 26)); txt_lines(d, 4, 12, (10, 14, 18, 22))
    shape(d, (18, 6, 30, 26)); txt_lines(d, 20, 28, (10, 14, 18, 22))
    d.line([x(16), x(3), x(16), x(29)], fill=RED, width=S)
    save(img, "sa_split")

def merge_text():
    img, d = canvas()
    shape(d, (2, 4, 12, 12)); txt_lines(d, 4, 10, (7, 9))
    shape(d, (2, 20, 12, 28)); txt_lines(d, 4, 10, (23, 25))
    shape(d, (19, 8, 29, 24)); txt_lines(d, 21, 27, (12, 15, 18, 21))
    arrow(d, (12.5, 8), (18, 13), w=1)
    arrow(d, (12.5, 24), (18, 19), w=1)
    save(img, "sa_merge")

def margins():
    img, d = canvas()
    shape(d, (3, 5, 29, 27))
    d.rectangle([x(8), x(10), x(24), x(22)], outline=RED, width=S)
    save(img, "sa_margins")

def fit_text():
    img, d = canvas()
    shape(d, (6, 8, 26, 24)); txt_lines(d, 9, 23, (13, 17, 20))
    arrow(d, (1, 16), (5, 16), w=1)
    arrow(d, (31, 16), (27, 16), w=1)
    save(img, "sa_fit")

def wrap_text():
    img, d = canvas()
    txt_lines(d, 4, 28, (7,)); txt_lines(d, 4, 22, (14,)); txt_lines(d, 4, 16, (21,))
    d.line([x(26), x(12), x(26), x(19)], fill=RED, width=S)
    arrow(d, (26, 19), (19, 21), w=1)
    save(img, "sa_wrap")

# ================= Wizards / productivity =================

def painter():
    img, d = canvas()
    d.line([x(7), x(25), x(15), x(17)], fill=GREY, width=3 * S)         # handle
    pts = [(15, 17), (20, 10), (26, 16), (21, 23)]                      # bristles
    d.polygon([(x(a), x(b)) for a, b in pts], fill=MASTER)
    shape(d, (20, 24, 29, 31))                                          # target shape
    save(img, "sa_painter")

def similar():
    img, d = canvas()
    shape(d, (5, 5, 13, 13)); shape(d, (19, 5, 27, 13)); shape(d, (5, 19, 13, 27))
    shape(d, (19, 19, 27, 27))
    for p1, p2 in [((2, 2), (30, 2)), ((30, 2), (30, 30)),
                   ((30, 30), (2, 30)), ((2, 30), (2, 2))]:
        dashes(d, p1, p2)
    save(img, "sa_similar")

def cleanup():
    img, d = canvas()
    d.line([x(22), x(3), x(14), x(17)], fill=GREY, width=2 * S)         # handle
    pts = [(9, 17), (17, 17), (20, 27), (6, 27)]                        # head
    d.polygon([(x(a), x(b)) for a, b in pts], fill=MASTER)
    for i in range(3):                                                  # dust
        d.line([x(23 + i * 2.4), x(21 + i * 2), x(25 + i * 2.4), x(19 + i * 2)],
               fill=RED, width=S)
    save(img, "sa_cleanup")

def paste_slides():
    img, d = canvas()
    d.rectangle([x(7), x(6), x(25), x(29)], fill=SHAPE_F, outline=GREY, width=S)
    d.rectangle([x(12), x(3), x(20), x(8)], fill=GREY)
    d.line([x(16), x(13), x(16), x(23)], fill=RED, width=2 * S)
    d.line([x(11), x(18), x(21), x(18)], fill=RED, width=2 * S)
    save(img, "sa_paste")

def language():
    img, d = canvas()
    txt_lines(d, 4, 18, (7, 12), color=GREY)
    d.line([x(7), x(19), x(13), x(26)], fill=RED, width=3 * S)
    d.line([x(13), x(26), x(27), x(8)], fill=RED, width=3 * S)
    save(img, "sa_lang")

def elements():
    img, d = canvas()
    shape(d, (4, 4, 14, 14)); shape(d, (18, 4, 28, 14)); shape(d, (4, 18, 14, 28))
    master(d, (18, 18, 28, 28))
    save(img, "sa_elements")

def formats():
    img, d = canvas()
    pts = [(4, 10), (18, 10), (26, 16), (18, 22), (4, 22)]              # tag
    d.polygon([(x(a), x(b)) for a, b in pts], fill=SHAPE_F, outline=SHAPE_L, width=S)
    d.ellipse([x(8), x(14), x(11.4), x(17.4)], fill=MASTER)
    save(img, "sa_formats")

def agenda():
    img, d = canvas()
    shape(d, (5, 5, 27, 10))
    master(d, (5, 14, 27, 19))
    shape(d, (5, 23, 27, 28))
    save(img, "sa_agenda")

def shortcut():
    img, d = canvas()
    d.rectangle([x(2), x(8), x(30), x(24)], outline=GREY, width=S)
    for i in range(5):
        d.rectangle([x(5 + i * 4.6), x(11), x(8 + i * 4.6), x(14)], fill=MASTER)
    d.rectangle([x(9), x(17), x(23), x(20)], fill=MASTER)               # space bar
    save(img, "sa_shortcut")

def master_objects():
    img, d = canvas()
    d.rectangle([x(2), x(4), x(30), x(28)], outline=GREY, width=S)
    master(d, (6, 8, 16, 15))
    for p1, p2 in [((19, 18), (27, 18)), ((27, 18), (27, 25)),
                   ((27, 25), (19, 25)), ((19, 25), (19, 18))]:
        dashes(d, p1, p2, color=SHAPE_L)
    save(img, "sa_masterobj")

# ================= Color swatches (menu item icons) =================
# Theme slots show the standard Office theme as a static preview; the
# applied color always follows the presentation's actual theme.
THEME_PREVIEW = {                       # ordinal -> hex (see ThemeIndexFromOrdinal)
    1: "000000", 2: "FFFFFF", 3: "44546A", 4: "E7E6E6",
    5: "4472C4", 6: "ED7D31", 7: "A5A5A5", 8: "FFC000",
    9: "5B9BD5", 10: "70AD47",
}
PALETTE = {                             # exact palette colors from modColors
    1: "1F497D", 2: "4F81BD", 3: "9BBB59", 4: "C0504D",
    5: "F79646", 6: "8064A2", 7: "595959", 8: "D9D9D9",
}

def hex_rgba(h):
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 255)

def swatch(name, hexcol):
    img, d = canvas()
    d.rounded_rectangle([x(3), x(3), x(29), x(29)], radius=x(4),
                        fill=hex_rgba(hexcol), outline=(160, 160, 160, 255), width=S)
    save(img, name)

def swatches():
    for i, h in THEME_PREVIEW.items():
        swatch(f"sa_sw_t{i}", h)
    for i, h in PALETTE.items():
        swatch(f"sa_sw_p{i}", h)

# ================= Chart Aid =================

def _bars(d, heights, colors=None, y_base=28, x0=4, bw=5, gap=2):
    for i, hh in enumerate(heights):
        col = (colors or [MASTER])[i % len(colors or [MASTER])]
        d.rectangle([x(x0 + i * (bw + gap)), x(y_base - hh),
                     x(x0 + i * (bw + gap) + bw), x(y_base)], fill=col)

def chart_column():
    img, d = canvas()
    _bars(d, [10, 16, 13, 22], [MASTER])
    d.line([x(3), x(28), x(30), x(28)], fill=GREY, width=S)
    save(img, "sa_ch_col")

def chart_bar():
    img, d = canvas()
    for i, ww in enumerate([12, 20, 16]):
        d.rectangle([x(4), x(5 + i * 8), x(4 + ww), x(10 + i * 8)], fill=MASTER)
    d.line([x(4), x(3), x(4), x(29)], fill=GREY, width=S)
    save(img, "sa_ch_bar")

def chart_stacked(pct):
    img, d = canvas()
    cols = [(10, 7), (14, 9), (11, 12)] if not pct else [(13, 13), (13, 13), (13, 13)]
    for i, (a, b) in enumerate(cols):
        xx = 5 + i * 9
        d.rectangle([x(xx), x(28 - a), x(xx + 6), x(28)], fill=MASTER)
        d.rectangle([x(xx), x(28 - a - b), x(xx + 6), x(28 - a)], fill=SHAPE_L)
    d.line([x(3), x(28), x(30), x(28)], fill=GREY, width=S)
    save(img, "sa_ch_pct" if pct else "sa_ch_stk")

def chart_sbar():
    img, d = canvas()
    for i, (a, b) in enumerate([(10, 8), (15, 6), (8, 11)]):
        yy = 5 + i * 8
        d.rectangle([x(4), x(yy), x(4 + a), x(yy + 5)], fill=MASTER)
        d.rectangle([x(4 + a), x(yy), x(4 + a + b), x(yy + 5)], fill=SHAPE_L)
    d.line([x(4), x(3), x(4), x(29)], fill=GREY, width=S)
    save(img, "sa_ch_sbr")

def chart_waterfall():
    img, d = canvas()
    d.rectangle([x(3), x(14), x(8), x(28)], fill=(155, 187, 89, 255))
    d.rectangle([x(10), x(8), x(15), x(14)], fill=(155, 187, 89, 255))
    d.rectangle([x(17), x(8), x(22), x(15)], fill=(192, 80, 77, 255))
    d.rectangle([x(24), x(15), x(29), x(28)], fill=(191, 191, 191, 255))
    d.line([x(8), x(14), x(10), x(14)], fill=GREY, width=S)
    d.line([x(15), x(8), x(17), x(8)], fill=GREY, width=S)
    d.line([x(22), x(15), x(24), x(15)], fill=GREY, width=S)
    save(img, "sa_ch_wf")

def chart_mekko():
    img, d = canvas()
    widths = [10, 6, 9]
    splits = [0.4, 0.65, 0.3]
    xx = 4
    for wdt, sp in zip(widths, splits):
        d.rectangle([x(xx), x(4), x(xx + wdt), x(4 + 24 * sp)], fill=MASTER)
        d.rectangle([x(xx), x(4 + 24 * sp), x(xx + wdt), x(28)], fill=SHAPE_L)
        xx += wdt + 1
    save(img, "sa_ch_mek")

def chart_line():
    img, d = canvas()
    pts = [(4, 24), (11, 14), (18, 18), (28, 6)]
    d.line([(x(a), x(b)) for a, b in pts], fill=MASTER, width=2 * S, joint="curve")
    for a, b in pts:
        d.ellipse([x(a - 1.6), x(b - 1.6), x(a + 1.6), x(b + 1.6)], fill=RED)
    d.line([x(3), x(28), x(30), x(28)], fill=GREY, width=S)
    save(img, "sa_ch_line")

def chart_area():
    img, d = canvas()
    d.polygon([(x(3), x(28)), (x(3), x(18)), (x(12), x(10)), (x(20), x(16)),
               (x(29), x(6)), (x(29), x(28))], fill=SHAPE_L)
    d.polygon([(x(3), x(28)), (x(3), x(23)), (x(12), x(19)), (x(20), x(23)),
               (x(29), x(16)), (x(29), x(28))], fill=MASTER)
    save(img, "sa_ch_area")

def chart_pie(doughnut):
    img, d = canvas()
    box = [x(5), x(5), x(27), x(27)]
    d.pieslice(box, start=-90, end=60, fill=MASTER)
    d.pieslice(box, start=60, end=160, fill=SHAPE_L)
    d.pieslice(box, start=160, end=270, fill=(141, 179, 226, 255))
    if doughnut:
        d.ellipse([x(11.5), x(11.5), x(20.5), x(20.5)], fill=(0, 0, 0, 0))
        # punch the hole by drawing background-transparent circle: draw white
        d.ellipse([x(11.5), x(11.5), x(20.5), x(20.5)], fill=(255, 255, 255, 255))
    save(img, "sa_ch_don" if doughnut else "sa_ch_pie")

def chart_scatter():
    img, d = canvas()
    d.line([x(4), x(3), x(4), x(28)], fill=GREY, width=S)
    d.line([x(4), x(28), x(29), x(28)], fill=GREY, width=S)
    for a, b, r in [(10, 20, 2), (14, 12, 3), (20, 16, 2), (24, 7, 3.6), (17, 23, 1.6)]:
        d.ellipse([x(a - r), x(b - r), x(a + r), x(b + r)], fill=MASTER)
    save(img, "sa_ch_scat")

def chart_gantt():
    img, d = canvas()
    d.line([x(3), x(5), x(29), x(5)], fill=GREY, width=S)
    for i, (a, b) in enumerate([(4, 14), (10, 22), (16, 28)]):
        d.rounded_rectangle([x(a), x(8 + i * 7), x(b), x(12 + i * 7)],
                            radius=x(2), fill=MASTER if i != 1 else SHAPE_L)
    save(img, "sa_ch_gantt")

def chart_edit():
    img, d = canvas()
    # mini table
    d.rectangle([x(3), x(6), x(15), x(26)], outline=GREY, width=S)
    d.line([x(3), x(13), x(15), x(13)], fill=GREY, width=S)
    d.line([x(3), x(19), x(15), x(19)], fill=GREY, width=S)
    d.line([x(9), x(6), x(9), x(26)], fill=GREY, width=S)
    arrow(d, (16, 16), (21, 16))
    _bars(d, [6, 10, 8], [MASTER], y_base=26, x0=22, bw=2.4, gap=1)
    save(img, "sa_ch_edit")

def chart_diff():
    img, d = canvas()
    d.rectangle([x(4), x(14), x(10), x(28)], fill=MASTER)
    d.rectangle([x(20), x(6), x(26), x(28)], fill=MASTER)
    arrow(d, (15, 15), (15, 7), w=1)
    arrow(d, (15, 7), (15, 15), w=1)
    dashes(d, (10, 14), (14, 14), color=GREY)
    dashes(d, (16, 6), (20, 6), color=GREY)
    save(img, "sa_ch_diff")

def chart_pctdiff():
    img, d = canvas()
    d.rectangle([x(4), x(16), x(10), x(28)], fill=MASTER)
    d.rectangle([x(20), x(8), x(26), x(28)], fill=MASTER)
    # % sign
    d.ellipse([x(11), x(3), x(15), x(7)], outline=RED, width=S)
    d.ellipse([x(17), x(9), x(21), x(13)], outline=RED, width=S)
    d.line([x(11), x(13), x(21), x(3)], fill=RED, width=S)
    save(img, "sa_ch_pctdiff")

def chart_cagr():
    img, d = canvas()
    d.rectangle([x(4), x(18), x(10), x(28)], fill=MASTER)
    d.rectangle([x(22), x(8), x(28), x(28)], fill=MASTER)
    arrow(d, (7, 14), (25, 4), w=2)
    save(img, "sa_ch_cagr")

def chart_vline():
    img, d = canvas()
    _bars(d, [10, 16, 8, 20], [MASTER])
    dashes(d, (2, 15), (30, 15), color=RED, w=1)
    save(img, "sa_ch_vline")

def chart_avg():
    img, d = canvas()
    _bars(d, [12, 18, 8, 15], [SHAPE_L])
    dashes(d, (2, 15), (30, 15), color=GREY, w=1)
    save(img, "sa_ch_avg")

def chart_harvey():
    img, d = canvas()
    d.ellipse([x(5), x(5), x(27), x(27)], outline=GREY, width=2 * S)
    d.pieslice([x(5), x(5), x(27), x(27)], start=-90, end=90, fill=GREY)
    save(img, "sa_ch_harvey")

def chart_check():
    img, d = canvas()
    d.rounded_rectangle([x(5), x(5), x(27), x(27)], radius=x(3),
                        outline=GREY, width=2 * S)
    d.line([x(10), x(16), x(15), x(21)], fill=(79, 129, 189, 255), width=3 * S)
    d.line([x(15), x(21), x(23), x(10)], fill=(79, 129, 189, 255), width=3 * S)
    save(img, "sa_ch_check")

def chart_cycle():
    img, d = canvas()
    d.arc([x(6), x(6), x(26), x(26)], start=-60, end=170, fill=GREY, width=2 * S)
    arrow(d, (8, 22), (10, 26), color=GREY, w=2)
    d.line([x(13), x(16), x(16), x(19)], fill=(79, 129, 189, 255), width=2 * S)
    d.line([x(16), x(19), x(21), x(12)], fill=(79, 129, 189, 255), width=2 * S)
    save(img, "sa_ch_cycle")

def chart_help():
    img, d = canvas()
    d.rectangle([x(4), x(6), x(20), x(26)], outline=GREY, width=S)
    d.line([x(4), x(12), x(20), x(12)], fill=GREY, width=S)
    d.line([x(12), x(6), x(12), x(26)], fill=GREY, width=S)
    d.ellipse([x(20), x(16), x(30), x(26)], fill=MASTER)
    d.line([x(24.4), x(18.4), x(25.6), x(19.6)], fill=(255, 255, 255, 255), width=S)
    d.line([x(25), x(20.5), x(25), x(23.5)], fill=(255, 255, 255, 255), width=2 * S)
    save(img, "sa_ch_help")

def chart_colors():
    img, d = canvas()
    _bars(d, [10, 16, 22], [MASTER, (155, 187, 89, 255), RED], x0=5, bw=6, gap=3)
    d.line([x(3), x(28), x(30), x(28)], fill=GREY, width=S)
    save(img, "sa_ch_colors")

def chart_rebuild():
    # bars + circular arrow (rebuild from table/data)
    img, d = canvas()
    _bars(d, [8, 12, 10], [MASTER], y_base=28, x0=4, bw=4, gap=2)
    d.arc([x(14), x(3), x(28), x(17)], start=300, end=200, fill=RED, width=2 * S)
    arrow(d, (16.5, 13), (15.5, 15.5), w=2)
    save(img, "sa_ch_rebuild")

def chart_restyle():
    # bars + brush stroke (restyle all charts)
    img, d = canvas()
    _bars(d, [7, 11, 9], [SHAPE_F], x0=3, bw=4.5, gap=2)
    d.polygon([x(18), x(24), x(26), x(10), x(29), x(12), x(22), x(26)], fill=RED)
    d.ellipse([x(16), x(24), x(21), x(29)], fill=MASTER)
    save(img, "sa_ch_restyle")

def chart_recolor():
    img, d = canvas()
    _bars(d, [12, 18], [SHAPE_F], x0=4, bw=6, gap=3)
    d.rectangle([x(22), x(6), x(28), x(28)], fill=RED)
    arrow(d, (14, 8), (21, 8), w=1)
    save(img, "sa_ch_recolor")

def chart_settings():
    # three sliders with knobs - the native Chart Settings panel
    img, d = canvas()
    rows = [9, 16, 23]
    knob = [21, 12, 25]
    for i, yy in enumerate(rows):
        d.line([x(5), x(yy), x(27), x(yy)], fill=GREY, width=S)
        kx = knob[i]
        d.ellipse([x(kx - 2.6), x(yy - 2.6), x(kx + 2.6), x(yy + 2.6)],
                  fill=MASTER, outline=SHAPE_L, width=S)
    save(img, "sa_ch_settings")

def chart_editcolors():
    # painter's palette with colour dabs - the native Edit Colors dialog
    img, d = canvas()
    d.ellipse([x(3), x(6), x(26), x(27)], fill=SHAPE_F, outline=SHAPE_L, width=S)
    dabs = [((9, 12), MASTER), ((15, 10), (155, 187, 89, 255)),
            ((19, 15), RED), ((11, 19), (255, 192, 0, 255))]
    for (cx, cy), col in dabs:
        d.ellipse([x(cx - 2.2), x(cy - 2.2), x(cx + 2.2), x(cy + 2.2)], fill=col)
    save(img, "sa_ch_editcolors")

# Color themes for Chart Aid (must match ThemeDef in modChartStyle.bas)
THEMES = [
    ("office",    "Office",        ["4472C4","ED7D31","A5A5A5","FFC000","5B9BD5","70AD47"]),
    ("nordic",    "Nordic Blue",   ["1F4E79","2E75B6","9DC3E6","BDD7EE","636363","D9D9D9"]),
    ("fjord",     "Fjord",         ["264653","2A9D8F","E9C46A","F4A261","E76F51","8AB17D"]),
    ("forest",    "Forest",        ["1B4332","2D6A4F","40916C","74C69D","B7E4C7","95D5B2"]),
    ("sunset",    "Sunset",        ["073B4C","118AB2","06D6A0","FFD166","EF476F","26547C"]),
    ("berry",     "Berry",         ["4A1942","893168","C05299","E29ACD","6F6F6F","CFCFCF"]),
    ("greyscale", "Greyscale",     ["212529","495057","6C757D","ADB5BD","CED4DA","DEE2E6"]),
    ("financial", "Financial",     ["00304D","006BA6","FFB81C","97999B","DA291C","63666A"]),
    ("vivid",     "Vivid",         ["3D348B","7678ED","F7B801","F18701","F35B04","5F0F40"]),
]

def theme_images():
    for key, _label, cols in THEMES:
        # wide strip for the gallery (six blocks)
        img = Image.new("RGBA", (96 * S, 20 * S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        for i, h in enumerate(cols):
            c = (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 255)
            d.rectangle([i * 16 * S, 0, (i + 1) * 16 * S - S, 20 * S], fill=c)
        img.resize((96, 20), Image.LANCZOS).save(OUT / f"sa_pal_{key}.png")
        # 16px mini icon for the menu fallback (2x3 grid)
        img2, d2 = canvas()
        for i, h in enumerate(cols):
            c = (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 255)
            gx, gy = i % 3, i // 3
            d2.rectangle([x(2 + gx * 10), x(4 + gy * 13), x(10 + gx * 10), x(15 + gy * 13)], fill=c)
        save(img2, f"sa_palm_{key}")

def chart_samples():
    img, d = canvas()
    d.rectangle([x(2), x(4), x(24), x(22)], outline=GREY, width=S)      # slide
    _bars(d, [5, 8, 6], [MASTER], y_base=19, x0=12, bw=2.6, gap=1)
    d.rectangle([x(4), x(8), x(10), x(19)], outline=GREY, width=S)      # mini table
    d.line([x(4), x(12), x(10), x(12)], fill=GREY, width=S)
    d.line([x(7), x(8), x(7), x(19)], fill=GREY, width=S)
    d.rectangle([x(20), x(18), x(30), x(28)], fill=SHAPE_F, outline=SHAPE_L, width=S)
    d.rectangle([x(22), x(20), x(32), x(30)], outline=SHAPE_L, width=S) # stacked slides
    save(img, "sa_ch_samples")

def charts():
    chart_colors(); chart_recolor(); chart_samples(); theme_images()
    chart_rebuild(); chart_restyle(); chart_settings(); chart_editcolors()
    chart_column(); chart_bar(); chart_stacked(False); chart_stacked(True)
    chart_sbar(); chart_waterfall(); chart_mekko(); chart_line(); chart_area()
    chart_pie(False); chart_pie(True); chart_scatter(); chart_gantt()
    chart_edit(); chart_diff(); chart_pctdiff(); chart_cagr(); chart_vline()
    chart_avg(); chart_harvey(); chart_check(); chart_cycle(); chart_help()

# ================= View =================

def eye(crossed):
    img, d = canvas()
    d.ellipse([x(4), x(10), x(28), x(24)], outline=GREY, width=2 * S)
    d.ellipse([x(12.5), x(13.5), x(19.5), x(20.5)], fill=MASTER)
    if crossed:
        d.line([x(5), x(27), x(27), x(5)], fill=RED, width=2 * S)
    return img

def hide_unhide():
    save(eye(True), "sa_hide")
    save(eye(False), "sa_unhide")

ALL = [align_edges, align_center, to_slide, dock, distribute, swap, stack,
       matrix, matrix_custom, place, text_more, spacing, golden, magic,
       size_w, size_wh, stretch, fill_gap,
       slice_, multiply, process_chain, angles, block_arrow, rounded_rect,
       table_snap, fill_color, line_color, font_color, to_rgb, to_theme,
       pick, info, split_text, merge_text, margins, fit_text, wrap_text,
       painter, similar, cleanup, paste_slides, language, elements, formats,
       agenda, shortcut, master_objects, hide_unhide, swatches, charts]

if __name__ == "__main__":
    for f in ALL:
        f()
    n = len(list(OUT.glob("*.png")))
    print(f"{n} icons written to {OUT}")
