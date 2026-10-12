"""Reference implementation of the VBA preset geometry engine (VxPresetGeo in src/VecStamp.bas).

It evaluates the compact preset strings produced by build/make_presets.py exactly like the VBA
code does and turns them into SVG paths, so the engine can be checked without PowerPoint:

    python3 tests/preset_engine.py out.svg      # contact sheet of all presets (default adjustments)

The arc conversion matches GeoArcTo: DrawingML arcTo angles are visual angles; the parametric
angle on the ellipse is atan2(wR*sin(a), hR*cos(a)).
"""
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), 'build'))
import make_presets  # noqa: E402

DEG = math.pi / 180 / 60000


def builtin(tok, w, h):
    ss, ls = min(w, h), max(w, h)
    table = {'w': w, 'h': h, 'l': 0, 't': 0, 'r': w, 'b': h, 'hc': w / 2, 'vc': h / 2, 'ss': ss, 'ls': ls,
             'cd2': 10800000, 'cd4': 5400000, 'cd8': 2700000, 'k34': 16200000, 'k38': 8100000,
             'k58': 13500000, 'k78': 18900000}
    if tok in table:
        return table[tok]
    if tok.startswith('wd'):
        return w / int(tok[2:])
    if tok.startswith('hd'):
        return h / int(tok[2:])
    if tok.startswith('ssd'):
        return ss / int(tok[3:])
    raise KeyError(tok)


class Engine:
    def __init__(self, src, w, h, adjust=None):
        self.w, self.h = w, h
        sec = src.split('|')
        self.adj = [0.0]
        for i, d in enumerate([x for x in sec[0].split(',') if x]):
            ang = d.startswith('!')
            v = float(d.lstrip('!'))
            if adjust and i < len(adjust) and adjust[i] is not None:      # PowerPoint-style value
                a = adjust[i]
                if ang:
                    v = a if abs(a) > 400 else a * 60000
                    if v < 0 <= float(d.lstrip('!')):
                        v += 21600000
                else:
                    v = a if abs(a) > 10 else a * 100000
            self.adj.append(v)
        self.g = []
        for f in [x for x in sec[1].split(';') if x]:
            self.g.append(self.evaluate(f.split(',')))
        self.paths = [p for p in sec[2].split('~') if p]
        self.rect = [self.arg(t) for t in sec[3].split(',')]

    def arg(self, t):
        if t[0] == '$':
            return self.g[int(t[1:])]
        if t[0] == '@':
            return self.adj[int(t[1:])]
        if t[0] == '-' or t[0].isdigit():
            return float(t)
        return builtin(t, self.w, self.h)

    def evaluate(self, f):
        op, a = f[0], [self.arg(t) for t in f[1:]]
        x = a[0] if a else 0
        y = a[1] if len(a) > 1 else 0
        z = a[2] if len(a) > 2 else 0
        if op == 'm':
            return x * y / z if z else 0.0
        if op == 'a':
            return x + y - z
        if op == 'd':
            return (x + y) / z if z else 0.0
        if op == 'q':
            return y if x > 0 else z
        if op == 'b':
            return abs(x)
        if op == 't':
            return math.atan2(y, x) / DEG
        if op == 'c':
            return x * math.cos(math.atan2(z, y))
        if op == 'o':
            return x * math.cos(y * DEG)
        if op == 'x':
            return max(x, y)
        if op == 'n':
            return min(x, y)
        if op == 'h':
            return math.sqrt(x * x + y * y + z * z)
        if op == 'p':
            return x if y < x else (z if y > z else y)
        if op == 's':
            return x * math.sin(math.atan2(z, y))
        if op == 'i':
            return x * math.sin(y * DEG)
        if op == 'r':
            return math.sqrt(max(x, 0))
        if op == 'g':
            return x * math.tan(y * DEG)
        if op == 'v':
            return x
        raise ValueError(op)

    def svg_paths(self):
        """[(d, filled, stroked)] in shape coordinates (y down)."""
        out = []
        for p in self.paths:
            head, cmds = p.split(':', 1)
            fill, stroke, pw, ph = head.split(',')
            sx = self.w / float(pw) if float(pw) else 1.0
            sy = self.h / float(ph) if float(ph) else 1.0
            d, cu, cv = [], 0.0, 0.0
            for c in cmds.split(';'):
                t = c.split(',')
                k, v = t[0], [self.arg(x) for x in t[1:]]
                if k == 'M':
                    cu, cv = v[0] * sx, v[1] * sy
                    d.append(f'M{cu:.3f},{cv:.3f}')
                elif k == 'L':
                    cu, cv = v[0] * sx, v[1] * sy
                    d.append(f'L{cu:.3f},{cv:.3f}')
                elif k == 'C':
                    d.append('C' + ' '.join(f'{v[i] * sx:.3f},{v[i + 1] * sy:.3f}' for i in (0, 2, 4)))
                    cu, cv = v[4] * sx, v[5] * sy
                elif k == 'Q':
                    x1, y1, x, y = v[0] * sx, v[1] * sy, v[2] * sx, v[3] * sy
                    d.append(f'C{cu + 2 / 3 * (x1 - cu):.3f},{cv + 2 / 3 * (y1 - cv):.3f} '
                             f'{x + 2 / 3 * (x1 - x):.3f},{y + 2 / 3 * (y1 - y):.3f} {x:.3f},{y:.3f}')
                    cu, cv = x, y
                elif k == 'A':
                    wr, hr, st, sw = v[0] * sx, v[1] * sy, v[2] * DEG, v[3] * DEG
                    for (mx, my, ex, ey, rx, ry) in arc_pieces(cu, cv, wr, hr, st, sw):
                        # the VBA writes EllipticalArcTo rows (end + a point on the arc); here an SVG arc
                        d.append(f'A{rx:.3f},{ry:.3f} 0 0 {1 if sw > 0 else 0} {ex:.3f},{ey:.3f}')
                        cu, cv = ex, ey
                elif k == 'Z':
                    d.append('Z')
            out.append((' '.join(d), fill == 'f', stroke == '1'))
        return out


def arc_pieces(cu, cv, wr, hr, st, sw):
    """Pieces (mid x, mid y, end x, end y, rx, ry) of an arcTo, at most 90 degrees each - as GeoArcTo."""
    if wr <= 0 or hr <= 0 or abs(sw) < 1e-9:
        return []
    t0 = math.atan2(wr * math.sin(st), hr * math.cos(st))
    t1 = math.atan2(wr * math.sin(st + sw), hr * math.cos(st + sw))
    if abs(sw) >= 2 * math.pi - 1e-7:
        dt = math.copysign(2 * math.pi, sw)
    else:
        dt = t1 - t0
        if sw > 0:
            while dt <= 0:
                dt += 2 * math.pi
        else:
            while dt >= 0:
                dt -= 2 * math.pi
    cx, cy = cu - wr * math.cos(t0), cv - hr * math.sin(t0)
    n = int(abs(dt) / (math.pi / 2) - 1e-9) + 1
    pieces = []
    for i in range(n):
        a = t0 + dt * i / n
        b = t0 + dt * (i + 1) / n
        m = (a + b) / 2
        pieces.append((cx + wr * math.cos(m), cy + hr * math.sin(m), cx + wr * math.cos(b), cy + hr * math.sin(b), wr, hr))
    return pieces


def sheet(path):
    compiled = make_presets.compile_all()
    seen, cells = set(), []
    for ast, (prst, src) in sorted(compiled.items()):
        if prst in seen:
            continue
        seen.add(prst)
        cells.append((ast, prst, src))
    cols, cw, ch = 12, 150, 140
    rows = (len(cells) + cols - 1) // cols
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{cols * cw}" height="{rows * ch}" '
             f'font-family="sans-serif" font-size="10"><rect width="100%" height="100%" fill="#fff"/>']
    for i, (ast, prst, src) in enumerate(cells):
        x0, y0 = (i % cols) * cw + 20, (i // cols) * ch + 12
        e = Engine(src, 110, 80)
        parts.append(f'<g transform="translate({x0},{y0})">')
        for d, filled, stroked in e.svg_paths():
            parts.append(f'<path d="{d}" fill="{"#9cc3e6" if filled else "none"}" fill-rule="evenodd" '
                         f'stroke="{"#1f4e79" if stroked else "none"}" stroke-width="1.2"/>')
        parts.append(f'<text x="0" y="100">{ast} {prst}</text></g>')
    parts.append('</svg>')
    open(path, 'w').write(''.join(parts))
    return len(cells)


if __name__ == '__main__':
    print(sheet(sys.argv[1] if len(sys.argv) > 1 else 'presets.svg'), 'presets drawn')
