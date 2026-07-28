#!/usr/bin/env python3
"""Render example images for docs/CHARTS.md using the same geometry
as the VBA chart builders (ported). Run: python3 tools/render_examples.py
"""
import math
from pathlib import Path
from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parent.parent / "docs" / "img"
OUT.mkdir(parents=True, exist_ok=True)

W, H = 480, 300
ACCENTS = [(68, 114, 196), (237, 125, 49), (165, 165, 165),
           (255, 192, 0), (91, 155, 213), (112, 173, 71)]
GREY = (89, 89, 89)
UP, DOWN, TOT = (155, 187, 89), (192, 80, 77), (191, 191, 191)

# Per-chart datasets (mirroring the Sample Slides in the add-in)
GRID = [["", "2023", "2024", "2025", "2026"],
        ["Europe", 42, 48, 55, 61],
        ["Americas", 35, 39, 46, 58],
        ["Asia", 18, 26, 37, 52]]
GRID_STK = [["", "2024", "2025", "2026"],
            ["Hardware", 50, 48, 45],
            ["Software", 25, 32, 41],
            ["Services", 12, 20, 31]]
GRID_PCT = [["", "2022", "2024", "2026"],
            ["Online", 20, 38, 57],
            ["Retail", 65, 48, 31],
            ["Partner", 15, 14, 12]]
GRID_LINE = [["", "Q1 25", "Q2 25", "Q3 25", "Q4 25", "Q1 26", "Q2 26"],
             ["Us", 42, 45, 49, 55, 62, 71],
             ["Competitor", 58, 57, 55, 54, 52, 51]]
GRID_AREA = [["", "2022", "2023", "2024", "2025", "2026"],
             ["Gen 1", 40, 32, 22, 12, 5],
             ["Gen 2", 8, 25, 38, 42, 40],
             ["Gen 3", 0, 3, 12, 28, 47]]
GRID_MEK = [["", "Europe", "Americas", "Asia"],
            ["Premium", 25, 20, 10],
            ["Standard", 20, 35, 15],
            ["Budget", 10, 25, 20]]

def canvas():
    img = Image.new("RGB", (W, H), (255, 255, 255))
    return img, ImageDraw.Draw(img)

def save(img, name):
    img.save(OUT / f"{name}.png")

def scale(vmin, vmax, top, ph):
    vmax = max(vmax, 0); vmin = min(vmin, 0)
    if vmax - vmin < 1e-6: vmax = vmin + 1
    ppu = ph / (vmax - vmin)
    return ppu, top + vmax * ppu

def txt(d, x, y, s, fill=(60, 60, 60), anchor="mm", size=12):
    d.text((x, y), str(s), fill=fill, anchor=anchor, font_size=size)

def bars(GRID, stacked=False, pct=False, name="column"):
    img, d = canvas()
    nser, ncat = len(GRID) - 1, len(GRID[0]) - 1
    pt, ph, pl, pw = 40, H - 80, 30, W - 60
    vmax = 0
    for c in range(1, ncat + 1):
        tot = sum(GRID[r][c] for r in range(1, nser + 1))
        vmax = max(vmax, tot if stacked else max(GRID[r][c] for r in range(1, nser + 1)))
    if pct: vmax = 100
    ppu, y0 = scale(0, vmax, pt, ph)
    slot = pw / ncat
    for r in range(1, nser + 1):                          # legend
        lx = pl + (r - 1) * 90
        d.rectangle([lx, 18, lx + 10, 28], fill=ACCENTS[r - 1])
        txt(d, lx + 16, 23, GRID[r][0], anchor="lm")
    for c in range(1, ncat + 1):
        cum = 0
        ctot = sum(GRID[r][c] for r in range(1, nser + 1))
        for r in range(1, nser + 1):
            v = GRID[r][c]
            if pct: v = v / ctot * 100
            if stacked or pct:
                bw = slot * 0.6
                bl = pl + (c - 1) * slot + (slot - bw) / 2
                bt, bh = y0 - (cum + v) * ppu, v * ppu
                cum += v
                lab = f"{v:.0f}%" if pct else f"{v:g}"
                col = ACCENTS[r - 1]
                d.rectangle([bl, bt, bl + bw, bt + bh], fill=col)
                txt(d, bl + bw / 2, bt + bh / 2, lab,
                    fill=(255, 255, 255) if sum(col) < 460 else (50, 50, 50))
            else:
                bsz = slot * 0.7 / nser
                bl = pl + (c - 1) * slot + slot * 0.15 + (r - 1) * bsz
                bt, bh = y0 - v * ppu, v * ppu
                d.rectangle([bl, bt, bl + bsz * 0.9, bt + bh], fill=ACCENTS[r - 1])
                txt(d, bl + bsz * 0.45, bt - 9, f"{v:g}")
        if stacked and not pct:
            txt(d, pl + (c - 1) * slot + slot / 2, y0 - cum * ppu - 10, f"{cum:g}")
        txt(d, pl + (c - 1) * slot + slot / 2, y0 + 12, GRID[0][c])
    d.line([pl, y0, pl + pw, y0], fill=GREY, width=2)
    save(img, name)

def waterfall():
    img, d = canvas()
    rows = [("Revenue", 120), ("COGS", -45), ("Gross profit", "="),
            ("Opex", -32), ("EBITDA", "="), ("D&A", -12), ("EBIT", "=")]
    cum = vmax = vmin = 0
    for _, v in rows:
        if v != "=": cum += v
        vmax, vmin = max(vmax, cum), min(vmin, cum)
    pt, ph, pl, pw = 30, H - 70, 30, W - 60
    ppu, y0 = scale(vmin, vmax, pt, ph)
    slot = pw / len(rows); bw = slot * 0.55
    cum = 0; prev = None
    for i, (lbl, v) in enumerate(rows):
        bl = pl + i * slot + (slot - bw) / 2
        if v == "=":
            v2 = cum; col = TOT
            bt, bh = (y0 - v2 * ppu, v2 * ppu) if v2 >= 0 else (y0, -v2 * ppu)
        else:
            v2 = v; col = UP if v >= 0 else DOWN
            bt = y0 - (cum + v) * ppu if v >= 0 else y0 - cum * ppu
            bh = abs(v) * ppu
            cum += v
        d.rectangle([bl, bt, bl + bw, bt + bh], fill=col)
        if prev is not None:
            d.line([prev_r, prev_y, bl, prev_y], fill=(150, 150, 150))
        prev, prev_r, prev_y = True, bl + bw, y0 - cum * ppu
        txt(d, bl + bw / 2, (bt - 10) if (v == "=" or v >= 0) else bt + bh + 10, f"{v2:g}")
        txt(d, bl + bw / 2, pt + ph + 12, lbl)
    d.line([pl, y0, pl + pw, y0], fill=GREY, width=2)
    save(img, "waterfall")

def mekko(GRID):
    img, d = canvas()
    nser, ncat = len(GRID) - 1, len(GRID[0]) - 1
    pt, ph, pl, pw = 40, H - 90, 30, W - 120
    tots = [sum(GRID[r][c] for r in range(1, nser + 1)) for c in range(1, ncat + 1)]
    grand = sum(tots); x = pl
    for ci, ct in enumerate(tots):
        xw = (pw - 2 * (ncat - 1)) * ct / grand
        cum = 0
        for r in range(1, nser + 1):
            v = GRID[r][ci + 1]
            sh = ph * v / ct
            st = pt + ph - (cum + v) / ct * ph
            d.rectangle([x, st, x + xw, st + sh], fill=ACCENTS[r - 1])
            txt(d, x + xw / 2, st + sh / 2, f"{v:g}", fill=(255, 255, 255))
            cum += v
        txt(d, x + xw / 2, pt - 10, f"{ct:g}")
        txt(d, x + xw / 2, pt + ph + 12, GRID[0][ci + 1])
        x += xw + 2
    for r in range(1, nser + 1):
        d.rectangle([x + 8, pt + (r - 1) * 20, x + 18, pt + (r - 1) * 20 + 10], fill=ACCENTS[r - 1])
        txt(d, x + 24, pt + (r - 1) * 20 + 5, GRID[r][0], anchor="lm")
    save(img, "mekko")

def line_area(GRID, area, name):
    img, d = canvas()
    nser, ncat = len(GRID) - 1, len(GRID[0]) - 1
    pt, ph, pl, pw = 40, H - 80, 30, W - 110
    vmax = max(sum(GRID[r][c] for r in range(1, nser + 1)) if area
               else max(GRID[r][c] for r in range(1, nser + 1))
               for c in range(1, ncat + 1))
    ppu, y0 = scale(0, vmax, pt, ph)
    slot = pw / ncat
    base = [0.0] * (ncat + 1)
    for r in range(1, nser + 1):
        pts = []
        for c in range(1, ncat + 1):
            v = GRID[r][c] + (base[c] if area else 0)
            pts.append((pl + (c - 1) * slot + slot / 2, y0 - v * ppu))
        if area:
            poly = pts + [(pl + (c - 1) * slot + slot / 2, y0 - base[c] * ppu)
                          for c in range(ncat, 0, -1)]
            d.polygon(poly, fill=ACCENTS[r - 1])
            for c in range(1, ncat + 1):
                base[c] += GRID[r][c]
        else:
            d.line(pts, fill=ACCENTS[r - 1], width=3)
            for (px, py), c in zip(pts, range(1, ncat + 1)):
                d.ellipse([px - 3, py - 3, px + 3, py + 3], fill=ACCENTS[r - 1])
                txt(d, px, py - 12, f"{GRID[r][c]:g}")
        txt(d, pts[-1][0] + 12, pts[-1][1], GRID[r][0], fill=ACCENTS[r - 1], anchor="lm")
    for c in range(1, ncat + 1):
        txt(d, pl + (c - 1) * slot + slot / 2, y0 + 12, GRID[0][c])
    d.line([pl, y0, pl + pw, y0], fill=GREY, width=2)
    save(img, name)

def pie(doughnut, name):
    img, d = canvas()
    rows = ([("Subscriptions", 55), ("Licenses", 25), ("Services", 20)] if doughnut
            else [("Personnel", 48), ("Facilities", 21), ("Marketing", 17), ("Other", 14)])
    tot = sum(v for _, v in rows)
    cx, cy, R = W // 2 - 40, H // 2, 100
    a0 = -90
    for i, (lbl, v) in enumerate(rows):
        sw = 360 * v / tot
        d.pieslice([cx - R, cy - R, cx + R, cy + R], a0, a0 + sw,
                   fill=ACCENTS[i], outline=(255, 255, 255), width=2)
        mid = math.radians(a0 + sw / 2)
        lr = R * (1.25 if doughnut else 0.62)
        txt(d, cx + math.cos(mid) * lr, cy + math.sin(mid) * lr,
            f"{lbl} {v}", fill=(50, 50, 50) if doughnut else (255, 255, 255))
        a0 += sw
    if doughnut:
        d.ellipse([cx - R * 0.45, cy - R * 0.45, cx + R * 0.45, cy + R * 0.45],
                  fill=(255, 255, 255))
    save(img, name)

def scatter():
    img, d = canvas()
    rows = [("Alpha", 32, 4, 120), ("Bravo", 18, 12, 80), ("Charlie", 9, 22, 40),
            ("Delta", 4, 28, 15), ("Echo", 25, -2, 95)]
    pl, pt, pw, ph = 50, 20, W - 90, H - 60
    xs = [r[1] for r in rows]; ys = [r[2] for r in rows]
    x0, x1 = min(xs) - 0.5, max(xs) + 0.5
    y0_, y1 = min(ys) - 1, max(ys) + 1
    d.line([pl, pt, pl, pt + ph], fill=GREY, width=2)
    d.line([pl, pt + ph, pl + pw, pt + ph], fill=GREY, width=2)
    szmax = max(r[3] for r in rows)
    for lbl, xv, yv, sz in rows:
        px = pl + (xv - x0) / (x1 - x0) * pw
        py = pt + ph - (yv - y0_) / (y1 - y0_) * ph
        r_ = 5 + 14 * math.sqrt(sz / szmax)
        d.ellipse([px - r_, py - r_, px + r_, py + r_], fill=ACCENTS[0] + (128,) if False else ACCENTS[0])
        txt(d, px + r_ + 4, py, lbl, anchor="lm")
    save(img, "scatter")

def gantt():
    img, d = canvas()
    rows = [("Discovery", 36, 38), ("Design", 38, 41), ("Build", 41, 46), ("Testing", 45, 48), ("Launch", 49, 49)]
    lab_w, pt = 90, 40
    pl, pw = 30 + lab_w, W - 60 - lab_w
    tmin, tmax = 36, 49
    ppt = pw / (tmax - tmin)
    row_h = 40
    d.line([pl, pt, pl + pw, pt], fill=GREY, width=2)
    for k in range(5):
        gv = tmin + (tmax - tmin) * k / 4
        gx = pl + (gv - tmin) * ppt
        d.line([gx, pt, gx, pt + row_h * len(rows)], fill=(220, 220, 220))
        txt(d, gx, pt - 12, f"{gv:g}")
    for i, (name, a, b) in enumerate(rows):
        yy = pt + i * row_h
        txt(d, pl - 8, yy + row_h / 2, name, anchor="rm")
        x1 = pl + (a - tmin) * ppt
        x2 = pl + (b - tmin) * ppt
        if x2 - x1 < 1:
            s = row_h * 0.28
            d.polygon([(x1, yy + row_h / 2 - s), (x1 + s, yy + row_h / 2),
                       (x1, yy + row_h / 2 + s), (x1 - s, yy + row_h / 2)], fill=GREY)
        else:
            d.rounded_rectangle([x1, yy + row_h * 0.25, x2, yy + row_h * 0.75],
                                radius=8, fill=ACCENTS[0])
    save(img, "gantt")

bars(GRID, name="column")
bars(GRID_STK, stacked=True, name="stacked")
bars(GRID_PCT, pct=True, stacked=True, name="pct")
waterfall()
mekko(GRID_MEK)
line_area(GRID_LINE, False, "line")
line_area(GRID_AREA, True, "area")
pie(False, "pie")
pie(True, "doughnut")
scatter()
gantt()
print(f"{len(list(OUT.glob('*.png')))} example images -> {OUT}")
