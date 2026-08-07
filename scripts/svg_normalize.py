#!/usr/bin/env python3
"""Normalize SVG path data to ABSOLUTE M/L/C/Z subpaths.

Shared by the IconAid build scripts and used to render both the web sidebar
previews and the editable PowerPoint freeforms from one representation, so they
match. Real icon paths use relative commands, H/V, S/Q/T and arcs; we convert
everything to absolute M/L/C/Z. Cubic/quadratic curves are preserved exactly;
arcs are approximated with cubic beziers finely enough to stay smooth. A
character-level scanner handles concatenated arc flags (e.g. "a5 5 0 013 4").
"""
from __future__ import annotations
import math, re

_NUM = re.compile(r'[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?')
_CMDS = "MmLlHhVvCcSsQqTtAaZz"

# Max sweep per cubic when flattening arcs. 90 degrees per cubic is already
# visually exact (~0.03% radial error); source cubics/quadratics are preserved
# exactly regardless, so this only affects circular-arc segments.
ARC_MAX_RAD = math.pi / 2      # 4 cubics per full circle
PRECISION = 2                  # 0.01 of a 24-unit viewBox = invisible at any icon size


class _Scanner:
    def __init__(self, d): self.d = d; self.i = 0; self.n = len(d)

    def _skip(self):
        while self.i < self.n and self.d[self.i] in ' ,\t\r\n':
            self.i += 1

    def at_end(self):
        self._skip(); return self.i >= self.n

    def cmd(self):
        self._skip()
        if self.i < self.n and self.d[self.i] in _CMDS:
            c = self.d[self.i]; self.i += 1; return c
        return None

    def number(self):
        self._skip()
        m = _NUM.match(self.d, self.i)
        if not m:
            raise ValueError(f"expected number at {self.i}: {self.d[self.i:self.i+12]!r}")
        self.i = m.end(); return float(m.group())

    def flag(self):
        self._skip()
        if self.i < self.n and self.d[self.i] in "01":
            v = int(self.d[self.i]); self.i += 1; return v
        return int(self.number())


def _arc_to_cubics(x1, y1, rx, ry, phi_deg, large, sweep, x2, y2):
    if rx == 0 or ry == 0 or (x1 == x2 and y1 == y2):
        return [((x1, y1), (x2, y2), (x2, y2))]
    phi = math.radians(phi_deg % 360)
    rx, ry = abs(rx), abs(ry)
    dx2, dy2 = (x1 - x2) / 2.0, (y1 - y2) / 2.0
    cosp, sinp = math.cos(phi), math.sin(phi)
    x1p = cosp * dx2 + sinp * dy2
    y1p = -sinp * dx2 + cosp * dy2
    lam = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry)
    if lam > 1:
        s = math.sqrt(lam); rx *= s; ry *= s
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    co = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        co = -co
    cxp = co * rx * y1p / ry
    cyp = -co * ry * x1p / rx
    cx = cosp * cxp - sinp * cyp + (x1 + x2) / 2.0
    cy = sinp * cxp + cosp * cyp + (y1 + y2) / 2.0

    def ang(ux, uy, vx, vy):
        dot = ux * vx + uy * vy
        ln = math.hypot(ux, uy) * math.hypot(vx, vy)
        a = math.acos(max(-1.0, min(1.0, dot / ln))) if ln else 0.0
        return -a if (ux * vy - uy * vx) < 0 else a

    theta1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dtheta = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dtheta > 0:
        dtheta -= 2 * math.pi
    if sweep and dtheta < 0:
        dtheta += 2 * math.pi

    nseg = max(1, int(math.ceil(abs(dtheta) / ARC_MAX_RAD - 1e-9)))
    delta = dtheta / nseg
    t = (4.0 / 3.0) * math.tan(delta / 4.0)
    out, th = [], theta1
    for _ in range(nseg):
        c1, s1 = math.cos(th), math.sin(th)
        c2, s2 = math.cos(th + delta), math.sin(th + delta)

        def pt(c, s):
            return (cosp * rx * c - sinp * ry * s + cx, sinp * rx * c + cosp * ry * s + cy)
        p1 = pt(c1, s1); p2 = pt(c2, s2)
        d1 = (-cosp * rx * s1 - sinp * ry * c1, -sinp * rx * s1 + cosp * ry * c1)
        d2 = (-cosp * rx * s2 - sinp * ry * c2, -sinp * rx * s2 + cosp * ry * c2)
        out.append(((p1[0] + t * d1[0], p1[1] + t * d1[1]),
                    (p2[0] - t * d2[0], p2[1] - t * d2[1]), p2))
        th += delta
    return out


def normalize(d: str):
    sc = _Scanner(d)
    cx = cy = sx = sy = 0.0
    pcx = pcy = None
    pqx = pqy = None
    cmd = None
    subpaths, cur = [], None
    while not sc.at_end():
        c = sc.cmd()
        if c is not None:
            cmd = c
        elif cmd is None or cmd in 'Zz':
            break
        rel = cmd.islower(); C = cmd.upper()
        if C == 'M':
            x = sc.number(); y = sc.number()
            if rel: x += cx; y += cy
            cx, cy = x, y; sx, sy = x, y
            cur = [('M', x, y)]; subpaths.append(cur)
            pcx = pqx = None
            cmd = 'l' if rel else 'L'
        elif C == 'L':
            x = sc.number(); y = sc.number()
            if rel: x += cx; y += cy
            cur.append(('L', x, y)); cx, cy = x, y; pcx = pqx = None
        elif C == 'H':
            x = sc.number()
            if rel: x += cx
            cur.append(('L', x, cy)); cx = x; pcx = pqx = None
        elif C == 'V':
            y = sc.number()
            if rel: y += cy
            cur.append(('L', cx, y)); cy = y; pcx = pqx = None
        elif C == 'C':
            x1 = sc.number(); y1 = sc.number(); x2 = sc.number(); y2 = sc.number(); x = sc.number(); y = sc.number()
            if rel: x1 += cx; y1 += cy; x2 += cx; y2 += cy; x += cx; y += cy
            cur.append(('C', x1, y1, x2, y2, x, y)); pcx, pcy = x2, y2; cx, cy = x, y; pqx = None
        elif C == 'S':
            x2 = sc.number(); y2 = sc.number(); x = sc.number(); y = sc.number()
            if rel: x2 += cx; y2 += cy; x += cx; y += cy
            x1, y1 = (cx, cy) if pcx is None else (2 * cx - pcx, 2 * cy - pcy)
            cur.append(('C', x1, y1, x2, y2, x, y)); pcx, pcy = x2, y2; cx, cy = x, y; pqx = None
        elif C == 'Q':
            qx = sc.number(); qy = sc.number(); x = sc.number(); y = sc.number()
            if rel: qx += cx; qy += cy; x += cx; y += cy
            cur.append(('C', cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy),
                        x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y))
            pqx, pqy = qx, qy; cx, cy = x, y; pcx = None
        elif C == 'T':
            x = sc.number(); y = sc.number()
            if rel: x += cx; y += cy
            qx, qy = (cx, cy) if pqx is None else (2 * cx - pqx, 2 * cy - pqy)
            cur.append(('C', cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy),
                        x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y))
            pqx, pqy = qx, qy; cx, cy = x, y; pcx = None
        elif C == 'A':
            rx = sc.number(); ry = sc.number(); rot = sc.number()
            large = sc.flag(); sweep = sc.flag()
            x = sc.number(); y = sc.number()
            if rel: x += cx; y += cy
            for cp1, cp2, end in _arc_to_cubics(cx, cy, rx, ry, rot, large, sweep, x, y):
                cur.append(('C', cp1[0], cp1[1], cp2[0], cp2[1], end[0], end[1]))
            cx, cy = x, y; pcx = pqx = None
        elif C == 'Z':
            if cur: cur.append(('Z',))
            cx, cy = sx, sy; pcx = pqx = None; cmd = None
    return [sp for sp in subpaths if len(sp) >= 2]


def _fmt(v: float) -> str:
    s = f"{v:.{PRECISION}f}".rstrip('0').rstrip('.')
    return '0' if s in ('', '-0') else s


def emit(subpath, scale=1.0) -> str:
    out = []
    for seg in subpath:
        if seg[0] == 'Z':
            out.append('Z')
        else:
            out.append(seg[0] + ' '.join(_fmt(v * scale) for v in seg[1:]))
    return ' '.join(out)


def normalize_field(d: str, scale=1.0):
    """One original 'd' -> list of absolute M/L/C/Z subpath strings (scaled)."""
    return [emit(sp, scale) for sp in normalize(d)]
