"""LibreOffice shapes -> VecStamp shape dicts for the native VSDX writer.

Geometry sources
* custom shapes (almost every shape drawn in Impress / Draw or imported from PowerPoint):
  the shape's own EnhancedCustomShape definition (equations, path, adjustment values)
  is evaluated here, so the Visio geometry matches LibreOffice exactly;
* lines, polylines, polygons, Bezier curves, connectors: their point lists;
* rectangles, ellipses and text frames: their frame;
* pictures, tables, charts, OLE objects and anything we cannot convert: an EMF picture.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import math
import os
import re
import tempfile

from vecstamp_vsdx import ARROWS, KAPPA, Path, rect_path

HMM = 2540.0                 # 1/100 mm per inch
PT_PER_HMM = 72.0 / 2540.0

# com.sun.star.drawing.EnhancedCustomShapeParameterType
P_NORMAL, P_EQUATION, P_ADJUSTMENT, P_LEFT, P_TOP, P_RIGHT, P_BOTTOM = range(7)
P_XSTRETCH, P_YSTRETCH, P_HASSTROKE, P_HASFILL, P_WIDTH, P_HEIGHT, P_LOGWIDTH, P_LOGHEIGHT = range(7, 15)
# com.sun.star.drawing.EnhancedCustomShapeSegmentCommand
(C_UNKNOWN, C_MOVETO, C_LINETO, C_CURVETO, C_CLOSESUBPATH, C_ENDSUBPATH, C_NOFILL, C_NOSTROKE,
 C_ANGLEELLIPSETO, C_ANGLEELLIPSE, C_ARCTO, C_ARC, C_CLOCKWISEARCTO, C_CLOCKWISEARC,
 C_ELLIPTICALQUADRANTX, C_ELLIPTICALQUADRANTY, C_QUADRATICCURVETO, C_ARCANGLETO) = range(18)
POINTS_PER_COMMAND = {C_MOVETO: 1, C_LINETO: 1, C_CURVETO: 3, C_QUADRATICCURVETO: 2, C_ANGLEELLIPSETO: 3,
                      C_ANGLEELLIPSE: 3, C_ARCTO: 4, C_ARC: 4, C_CLOCKWISEARCTO: 4, C_CLOCKWISEARC: 4,
                      C_ELLIPTICALQUADRANTX: 1, C_ELLIPTICALQUADRANTY: 1, C_ARCANGLETO: 2}

NATIVE_TYPES = ('CustomShape', 'RectangleShape', 'EllipseShape', 'TextShape', 'LineShape', 'PolyLineShape',
                'PolyPolygonShape', 'OpenBezierShape', 'ClosedBezierShape', 'PolyLinePathShape',
                'PolyPolygonPathShape', 'OpenFreeHandShape', 'ClosedFreeHandShape', 'ConnectorShape',
                'MeasureShape', 'TitleTextShape', 'OutlinerShape', 'SubTitleShape')


class Unsupported(Exception):
    pass


def _get(obj, name, default=None):
    try:
        info = obj.getPropertySetInfo()
        if info.hasPropertyByName(name):
            return obj.getPropertyValue(name)
    except Exception:
        pass
    return default


def _rgb(v):
    if v is None or v < 0:
        return (0, 0, 0)
    return ((v >> 16) & 255, (v >> 8) & 255, v & 255)


def _enum(v):
    return getattr(v, 'value', v)


# ------------------------------------------------------------------ custom shape evaluator
_TOKEN = re.compile(r'\s*(\?\d+|\$\d+|[A-Za-z_][A-Za-z_0-9]*|\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+|[-+*/(),])')


class _Formula:
    """Evaluates LibreOffice enhanced-geometry formulas (draw:formula syntax)."""

    FUNCS = {'abs': abs, 'sqrt': lambda x: math.sqrt(max(x, 0)), 'sin': math.sin, 'cos': math.cos,
             'tan': math.tan, 'atan': math.atan, 'atan2': math.atan2, 'min': min, 'max': max}

    def __init__(self, equations, adjust, env):
        self.src = [str(e) for e in equations]
        self.adjust = adjust
        self.env = env
        self.cache = {}
        self.busy = set()

    def equation(self, i):
        if i in self.cache:
            return self.cache[i]
        if i in self.busy or i >= len(self.src):
            return 0.0
        self.busy.add(i)
        try:
            v = self.evaluate(self.src[i])
        except Exception:
            v = 0.0
        self.busy.discard(i)
        self.cache[i] = v
        return v

    def evaluate(self, text):
        return float(_Parser(self, _TOKEN.findall(text)).parse())


class _Parser:
    """Recursive-descent parser for one formula (own token state, so nested ?n lookups are safe)."""

    def __init__(self, formula, toks):
        self.f = formula
        self.toks = toks
        self.pos = 0

    def parse(self):
        return self._expr()

    def _peek(self):
        return self.toks[self.pos] if self.pos < len(self.toks) else None

    def _take(self):
        t = self._peek()
        self.pos += 1
        return t

    def _expr(self):
        v = self._term()
        while self._peek() in ('+', '-'):
            op = self._take()
            r = self._term()
            v = v + r if op == '+' else v - r
        return v

    def _term(self):
        v = self._unary()
        while self._peek() in ('*', '/'):
            op = self._take()
            r = self._unary()
            v = v * r if op == '*' else (v / r if r else 0.0)
        return v

    def _unary(self):
        if self._peek() == '-':
            self._take()
            return -self._unary()
        if self._peek() == '+':
            self._take()
            return self._unary()
        return self._atom()

    def _args(self):
        self._take()                      # '('
        args = [self._expr()]
        while self._peek() == ',':
            self._take()
            args.append(self._expr())
        self._take()                      # ')'
        return args

    def _atom(self):
        t = self._take()
        if t is None:
            return 0.0
        if t == '(':
            v = self._expr()
            self._take()
            return v
        if t[0] == '?':
            return self.f.equation(int(t[1:]))
        if t[0] == '$':
            k = int(t[1:])
            return float(self.f.adjust[k]) if k < len(self.f.adjust) else 0.0
        if t[0].isdigit() or t[0] == '.':
            return float(t)
        name = t.lower()
        if name == 'if':
            a = self._args()
            return a[1] if a[0] > 0 else a[2]
        if name in _Formula.FUNCS:
            return float(_Formula.FUNCS[name](*self._args()))
        if name == 'pi':
            return math.pi
        return float(self.f.env.get(name, 0.0))


def _adjust_values(geo):
    out = []
    for a in geo.get('AdjustmentValues') or ():
        v = a.Value
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            out.append(0.0)
    return out


def custom_shape_paths(shape, w, h):
    """(paths, text_rect) of a custom shape in local screen coordinates (1/100 mm, y down).

    text_rect is (x0, y0, x1, y1) of the first text frame, or None. Raises Unsupported."""
    geo = {p.Name: p.Value for p in shape.getPropertyValue('CustomShapeGeometry')}
    path = {p.Name: p.Value for p in (geo.get('Path') or ())}
    coords = path.get('Coordinates') or ()
    segments = path.get('Segments') or ()
    if not coords:
        raise Unsupported('custom shape without path definition')
    vb = geo.get('ViewBox')
    if vb is not None and vb.Width > 0 and vb.Height > 0:
        left, top, vw, vh = float(vb.X), float(vb.Y), float(vb.Width), float(vb.Height)
        logw, logh = vw, vh
    else:
        left, top, vw, vh = 0.0, 0.0, float(w), float(h)
        logw, logh = float(w), float(h)
    env = {'left': left, 'top': top, 'right': left + vw, 'bottom': top + vh, 'width': vw, 'height': vh,
           'logwidth': logw, 'logheight': logh, 'xstretch': 0.0, 'ystretch': 0.0,
           'hasstroke': 1.0, 'hasfill': 1.0}
    f = _Formula(geo.get('Equations') or (), _adjust_values(geo), env)
    subviews = [(float(s.Width), float(s.Height)) for s in (path.get('SubViewSize') or ())]

    def value(par):
        t, v = par.Type, par.Value
        if t == P_NORMAL:
            try:
                return float(v)
            except (TypeError, ValueError):
                return 0.0
        if t == P_EQUATION:
            return f.equation(int(v))
        if t == P_ADJUSTMENT:
            k = int(v)
            return f.adjust[k] if k < len(f.adjust) else 0.0
        return {P_LEFT: left, P_TOP: top, P_RIGHT: left + vw, P_BOTTOM: top + vh, P_WIDTH: vw, P_HEIGHT: vh,
                P_LOGWIDTH: logw, P_LOGHEIGHT: logh, P_HASSTROKE: 1.0, P_HASFILL: 1.0}.get(t, 0.0)

    mirror_x, mirror_y = bool(geo.get('MirroredX')), bool(geo.get('MirroredY'))
    state = {'sub': 0}

    def scale():
        k = state['sub']
        if k < len(subviews) and subviews[k][0] > 0 and subviews[k][1] > 0:
            return w / subviews[k][0], h / subviews[k][1], 0.0, 0.0
        return w / vw, h / vh, left, top

    def pt(pair):
        sx, sy, ox, oy = scale()
        x = (value(pair.First) - ox) * sx
        y = (value(pair.Second) - oy) * sy
        if mirror_x:
            x = w - x
        if mirror_y:
            y = h - y
        return x, y

    def radii(pair):
        sx, sy, _, _ = scale()
        return abs(value(pair.First) * sx), abs(value(pair.Second) * sy)

    paths = []
    cur = {'path': None, 'nofill': False, 'noline': False, 'start': None, 'pt': (0.0, 0.0), 'group': []}

    def new_path(x, y):
        p = Path(closed=False)
        p.move(x, y)
        cur['path'] = p
        cur['group'].append(p)
        cur['start'] = cur['pt'] = (x, y)
        paths.append(p)

    def ensure_path():
        if cur['path'] is None:
            new_path(*cur['pt'])

    def line_to(x, y):
        ensure_path()
        cur['path'].line(x, y)
        cur['pt'] = (x, y)

    def cubic_to(x1, y1, x2, y2, x, y):
        ensure_path()
        cur['path'].cubic(x1, y1, x2, y2, x, y)
        cur['pt'] = (x, y)

    def arc(cx, cy, rx, ry, t0, sweep, connect):
        """Elliptic arc by parametric angles (radians, y down: positive sweep = clockwise on screen)."""
        x0, y0 = cx + rx * math.cos(t0), cy + ry * math.sin(t0)
        if connect:
            if cur['path'] is None:
                new_path(x0, y0)
            elif abs(cur['pt'][0] - x0) > 0.01 or abs(cur['pt'][1] - y0) > 0.01:
                line_to(x0, y0)
        else:
            new_path(x0, y0)
        n = max(1, int(math.ceil(abs(sweep) / (math.pi / 2) - 1e-9)))
        d = sweep / n
        k = 4.0 / 3.0 * math.tan(d / 4)
        t = t0
        for _ in range(n):
            t1 = t + d
            c0, s0, c1, s1 = math.cos(t), math.sin(t), math.cos(t1), math.sin(t1)
            cubic_to(cx + rx * (c0 - k * s0), cy + ry * (s0 + k * c0),
                     cx + rx * (c1 + k * s1), cy + ry * (s1 - k * c1),
                     cx + rx * c1, cy + ry * s1)
            t = t1

    def param_angle(rx, ry, visual):
        return math.atan2(rx * math.sin(visual), ry * math.cos(visual))

    ci = 0
    for seg in segments:
        cmd, count = int(seg.Command), int(seg.Count)
        need = POINTS_PER_COMMAND.get(cmd, 0)
        if cmd in (C_CLOSESUBPATH,):
            if cur['path'] is not None:
                cur['path'].closed = True
                cur['pt'] = cur['start']
                cur['path'] = None
            continue
        if cmd == C_ENDSUBPATH:
            for p in cur['group']:
                p.nofill = p.nofill or cur['nofill']
                p.noline = p.noline or cur['noline']
            cur.update(path=None, nofill=False, noline=False, group=[])
            state['sub'] += 1
            continue
        if cmd == C_NOFILL:
            cur['nofill'] = True
            continue
        if cmd == C_NOSTROKE:
            cur['noline'] = True
            continue
        if need == 0:
            continue
        reps = max(count, 1)
        quadrant_x = cmd == C_ELLIPTICALQUADRANTX
        for r in range(reps):
            pts = coords[ci:ci + need]
            ci += need
            if len(pts) < need:
                break
            if cmd == C_MOVETO:
                x, y = pt(pts[0])
                if r == 0:
                    new_path(x, y)
                else:
                    line_to(x, y)
            elif cmd == C_LINETO:
                line_to(*pt(pts[0]))
            elif cmd == C_CURVETO:
                (x1, y1), (x2, y2), (x, y) = pt(pts[0]), pt(pts[1]), pt(pts[2])
                cubic_to(x1, y1, x2, y2, x, y)
            elif cmd == C_QUADRATICCURVETO:
                (qx, qy), (x, y) = pt(pts[0]), pt(pts[1])
                px, py = cur['pt']
                cubic_to(px + 2 / 3 * (qx - px), py + 2 / 3 * (qy - py), x + 2 / 3 * (qx - x), y + 2 / 3 * (qy - y), x, y)
            elif cmd == C_ARCANGLETO:
                rx, ry = radii(pts[0])
                st, sw = math.radians(value(pts[1].First)), math.radians(value(pts[1].Second))
                if rx < 1e-6 or ry < 1e-6:
                    continue
                if mirror_x:
                    st, sw = math.pi - st, -sw
                if mirror_y:
                    st, sw = -st, -sw
                t0 = param_angle(rx, ry, st)
                t1 = param_angle(rx, ry, st + sw)
                sweep = t1 - t0
                if sw > 0 and sweep < 0:
                    sweep += 2 * math.pi
                elif sw < 0 and sweep > 0:
                    sweep -= 2 * math.pi
                if abs(sw) >= 2 * math.pi - 1e-9:
                    sweep = math.copysign(2 * math.pi, sw)
                px, py = cur['pt']
                cx, cy = px - rx * math.cos(t0), py - ry * math.sin(t0)
                arc(cx, cy, rx, ry, t0, sweep, True)
            elif cmd in (C_ANGLEELLIPSETO, C_ANGLEELLIPSE):
                (cx, cy) = pt(pts[0])
                rx, ry = radii(pts[1])
                a0, a1 = value(pts[2].First), value(pts[2].Second)
                sweep = -math.radians((a1 - a0) % 360 or 360)
                t0 = -math.radians(a0)
                if mirror_x:
                    t0, sweep = math.pi - t0, -sweep
                if mirror_y:
                    t0, sweep = -t0, -sweep
                arc(cx, cy, rx, ry, t0, sweep, cmd == C_ANGLEELLIPSETO and cur['path'] is not None)
            elif cmd in (C_ARCTO, C_ARC, C_CLOCKWISEARCTO, C_CLOCKWISEARC):
                (x1, y1), (x2, y2), (sx, sy), (ex, ey) = pt(pts[0]), pt(pts[1]), pt(pts[2]), pt(pts[3])
                cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
                rx, ry = abs(x2 - x1) / 2, abs(y2 - y1) / 2
                if rx < 1e-6 or ry < 1e-6:
                    continue
                t0 = math.atan2((sy - cy) / ry, (sx - cx) / rx)
                t1 = math.atan2((ey - cy) / ry, (ex - cx) / rx)
                clockwise = cmd in (C_CLOCKWISEARCTO, C_CLOCKWISEARC)
                if mirror_x != mirror_y:
                    clockwise = not clockwise
                sweep = t1 - t0
                if clockwise and sweep <= 0:
                    sweep += 2 * math.pi
                elif not clockwise and sweep >= 0:
                    sweep -= 2 * math.pi
                arc(cx, cy, rx, ry, t0, sweep, cmd in (C_ARCTO, C_CLOCKWISEARCTO))
            elif cmd in (C_ELLIPTICALQUADRANTX, C_ELLIPTICALQUADRANTY):
                x, y = pt(pts[0])
                px, py = cur['pt']
                if quadrant_x:
                    cubic_to(px + (x - px) * KAPPA, py, x, y - (y - py) * KAPPA, x, y)
                else:
                    cubic_to(px, py + (y - py) * KAPPA, x - (x - px) * KAPPA, y, x, y)
                quadrant_x = not quadrant_x
    for p in cur['group']:
        p.nofill = p.nofill or cur['nofill']
        p.noline = p.noline or cur['noline']
    if not paths:
        raise Unsupported('empty custom shape')
    text_rect = None
    frames = path.get('TextFrames') or ()
    if frames:
        state['sub'] = len(subviews)              # text frames use the main view box
        try:
            (ax, ay), (bx, by) = pt(frames[0].TopLeft), pt(frames[0].BottomRight)
            text_rect = (min(ax, bx), min(ay, by), max(ax, bx), max(ay, by))
            if text_rect[2] - text_rect[0] < 1 or text_rect[3] - text_rect[1] < 1:
                text_rect = None
        except Exception:
            text_rect = None
    return paths, text_rect


# ------------------------------------------------------------------ frames and styles
class Frame:
    """Unrotated frame of a shape: centre, size (1/100 mm) and rotation (radians, counter-clockwise)."""

    def __init__(self, shape):
        size = shape.getSize()
        self.w, self.h = float(size.Width), float(size.Height)
        m = _get(shape, 'Transformation')
        if m is not None:
            a, b, tx = m.Line1.Column1, m.Line1.Column2, m.Line1.Column3
            c, d, ty = m.Line2.Column1, m.Line2.Column2, m.Line2.Column3
            self.cx, self.cy = a * 0.5 + b * 0.5 + tx, c * 0.5 + d * 0.5 + ty
        else:
            pos = shape.getPosition()
            self.cx, self.cy = pos.X + self.w / 2, pos.Y + self.h / 2
        self.angle = math.radians((_get(shape, 'RotateAngle', 0) or 0) / 100.0)

    def to_local(self, x, y):
        """Absolute page point -> local screen coordinates of the unrotated frame."""
        dx, dy = x - self.cx, y - self.cy
        c, s = math.cos(self.angle), math.sin(self.angle)
        return self.w / 2 + dx * c - dy * s, self.h / 2 + dx * s + dy * c


def _to_visio(paths_screen, h):
    """Paths in local screen 1/100 mm (y down) -> Visio local inches (y up)."""
    out = []
    for p in paths_screen:
        q = Path(p.closed, p.nofill, p.noline)
        for s in p.segs:
            if s[0] in ('M', 'L'):
                q.segs.append((s[0], s[1] / HMM, (h - s[2]) / HMM))
            elif s[0] == 'C':
                q.segs.append(('C', s[1] / HMM, (h - s[2]) / HMM, s[3] / HMM, (h - s[4]) / HMM,
                               s[5] / HMM, (h - s[6]) / HMM))
            elif s[0] == 'E':
                q.segs.append(('E', s[1] / HMM, (h - s[2]) / HMM, s[3] / HMM, s[4] / HMM))
        out.append(q)
    return out


def _poly_paths(shape, frame, closed):
    """Point-list shapes (lines, polygons, Bezier curves, connectors)."""
    pp = _get(shape, 'PolyPolygonBezier')
    paths = []
    if pp is not None:
        for poly, flags in zip(pp.Coordinates, pp.Flags):
            pts = [frame.to_local(q.X, q.Y) for q in poly]
            fl = [_enum(f) for f in flags]
            if not pts:
                continue
            p = Path(closed=closed)
            p.move(*pts[0])
            i = 1
            while i < len(pts):
                if fl[i] == 'CONTROL' and i + 2 < len(pts) and fl[i + 1] == 'CONTROL':
                    (x1, y1), (x2, y2), (x, y) = pts[i], pts[i + 1], pts[i + 2]
                    p.cubic(x1, y1, x2, y2, x, y)
                    i += 3
                else:
                    p.line(*pts[i])
                    i += 1
            paths.append(p)
        return paths
    poly = _get(shape, 'PolyPolygon')
    for pts in poly or ():
        loc = [frame.to_local(q.X, q.Y) for q in pts]
        if loc:
            p = Path(closed=closed)
            p.move(*loc[0])
            for x, y in loc[1:]:
                p.line(x, y)
            paths.append(p)
    return paths


def _fill(shape):
    style = _enum(_get(shape, 'FillStyle', 'NONE'))
    if style in (None, 'NONE'):
        return None
    trans = (_get(shape, 'FillTransparence', 0) or 0) / 100.0
    color = _rgb(_get(shape, 'FillColor', 0xFFFFFF))
    fill = {'color': color, 'trans': trans}
    if style == 'GRADIENT':
        g = _get(shape, 'FillGradient')
        if g is not None:
            a, b = _rgb(g.StartColor), _rgb(g.EndColor)
            kind = _enum(g.Style)
            stops = [(0.0, a, trans), (0.5, b, trans), (1.0, a, trans)] if kind == 'AXIAL' else \
                [(0.0, a, trans), (1.0, b, trans)]
            fill.update(color=a, stops=stops, angle=math.radians(g.Angle / 10.0 + 90))
    return fill


def _arrow(name, width_hmm, line_hmm):
    if not name:
        return 0, 2
    n = name.lower()
    if 'open' in n or 'line' in n:
        kind = 'open'
    elif 'stealth' in n or 'concave' in n:
        kind = 'stealth'
    elif 'diamond' in n:
        kind = 'diamond'
    elif 'oval' in n or 'circle' in n:
        kind = 'oval'
    elif 'square' in n:
        kind = 'square'
    else:
        kind = 'triangle'
    ratio = (width_hmm or 0) / max(line_hmm or 35, 1)
    size = 1 if ratio < 3 else (2 if ratio < 5 else 3)
    return ARROWS[kind], size


def _line(shape):
    style = _enum(_get(shape, 'LineStyle', 'NONE'))
    if style in (None, 'NONE'):
        return None
    width = _get(shape, 'LineWidth', 0) or 0
    line = {'color': _rgb(_get(shape, 'LineColor', 0)), 'weight': max(width * PT_PER_HMM, 0.5),
            'trans': (_get(shape, 'LineTransparence', 0) or 0) / 100.0, 'pattern': 1}
    if style == 'DASH':
        dash = _get(shape, 'LineDash')
        line['pattern'] = 3 if dash is not None and not dash.Dashes else 2
    for end, key in (('Start', 'begin'), ('End', 'end')):
        a, size = _arrow(_get(shape, 'Line%sName' % end, ''), _get(shape, 'Line%sWidth' % end, 0), width)
        if a:
            line[key + '_arrow'], line[key + '_size'] = a, size
    return line


_ALIGN = {'LEFT': 0, 'RIGHT': 2, 'BLOCK': 3, 'CENTER': 1, 'STRETCH': 3}
_VALIGN = {'TOP': 0, 'CENTER': 1, 'BOTTOM': 2, 'BLOCK': 1}


def _text(shape):
    if not hasattr(shape, 'getText'):
        return None
    try:
        if not shape.getString():
            return None
        xtext = shape.getText()
    except Exception:
        return None
    paras = []
    en = xtext.createEnumeration()
    while en.hasMoreElements():
        para = en.nextElement()
        if not hasattr(para, 'createEnumeration'):
            continue
        spacing = _get(para, 'ParaLineSpacing')
        spline = -1.2
        if spacing is not None:
            mode = spacing.Mode
            if mode == 0:                       # PROP
                spline = -1.2 * spacing.Height / 100.0
            elif mode in (1, 3):                # MINIMUM / FIX (1/100 mm)
                spline = spacing.Height * PT_PER_HMM
        p = {'align': _ALIGN.get(_enum(_get(para, 'ParaAdjust', 'LEFT')), 0), 'spline': spline,
             'before': (_get(para, 'ParaTopMargin', 0) or 0) * PT_PER_HMM,
             'after': (_get(para, 'ParaBottomMargin', 0) or 0) * PT_PER_HMM, 'runs': []}
        pe = para.createEnumeration()
        while pe.hasMoreElements():
            run = pe.nextElement()
            s = run.getString()
            if not s:
                continue
            esc = _get(run, 'CharEscapement', 0) or 0
            p['runs'].append({
                'text': s, 'font': _get(run, 'CharFontName', 'Liberation Sans') or 'Liberation Sans',
                'asian': _get(run, 'CharFontNameAsian', '') or '',
                'size': float(_get(run, 'CharHeight', 18) or 18),
                'bold': (_get(run, 'CharWeight', 100) or 100) >= 150,
                'italic': _enum(_get(run, 'CharPosture', 'NONE')) not in ('NONE', None),
                'underline': (_get(run, 'CharUnderline', 0) or 0) != 0,
                'color': _rgb(_get(run, 'CharColor', -1)), 'pos': 1 if esc > 0 else (2 if esc < 0 else 0)})
        paras.append(p)
    if not any(p['runs'] for p in paras):
        return None
    hadj = _enum(_get(shape, 'TextHorizontalAdjust', 'BLOCK'))
    if hadj in ('CENTER', 'RIGHT') and all(p['align'] == 0 for p in paras):
        for p in paras:                         # text frame centred / right-aligned as a whole
            p['align'] = 1 if hadj == 'CENTER' else 2
    k = PT_PER_HMM
    nowrap = (_get(shape, 'TextWordWrap', True) is False) or bool(_get(shape, 'TextAutoGrowWidth', False))
    mode = _enum(_get(shape, 'WritingMode', 'LR_TB'))
    return {'paras': paras, 'valign': _VALIGN.get(_enum(_get(shape, 'TextVerticalAdjust', 'CENTER')), 1),
            'margins': ((_get(shape, 'TextLeftDistance', 0) or 0) * k, (_get(shape, 'TextRightDistance', 0) or 0) * k,
                        (_get(shape, 'TextUpperDistance', 0) or 0) * k, (_get(shape, 'TextLowerDistance', 0) or 0) * k),
            'nowrap': nowrap, 'vertical': mode == 'TB_RL'}


# ------------------------------------------------------------------ walker
class Collector:
    def __init__(self, ctx, export_emf):
        self.ctx = ctx
        self.export_emf = export_emf        # function(shape) -> EMF bytes
        self.native = 0
        self.pictures = 0

    def convert(self, shape, ox, oy):
        """shape -> dict; (ox, oy): parent origin in page 1/100 mm (left, bottom)."""
        stype = shape.getShapeType().rsplit('.', 1)[-1]
        name = _get(shape, 'Name', '') or ''
        if stype == 'GroupShape':
            pos, size = shape.getPosition(), shape.getSize()
            gx, gy = pos.X, pos.Y + size.Height
            kids = [self.convert(shape.getByIndex(i), gx, gy) for i in range(shape.getCount())]
            return {'kind': 'group', 'name': name, 'x': (pos.X - ox) / HMM, 'y': (oy - gy) / HMM,
                    'w': size.Width / HMM, 'h': size.Height / HMM, 'children': kids}
        if stype in NATIVE_TYPES:
            try:
                d = self._native(shape, stype, ox, oy)
                self.native += 1
                d['name'] = name
                return d
            except Unsupported:
                pass
            except Exception:
                pass
        return self._picture(shape, ox, oy, name)

    def _native(self, shape, stype, ox, oy):
        fr = Frame(shape)
        w, h = max(fr.w, 1.0), max(fr.h, 1.0)
        fill, line = _fill(shape), _line(shape)
        text_rect = None
        if stype == 'CustomShape':
            paths, text_rect = custom_shape_paths(shape, w, h)
        elif stype in ('RectangleShape', 'TextShape', 'TitleTextShape', 'OutlinerShape', 'SubTitleShape'):
            r = float(_get(shape, 'CornerRadius', 0) or 0)
            paths = [_rect_screen(w, h, r)]
        elif stype == 'EllipseShape':
            if _enum(_get(shape, 'CircleKind', 'FULL')) != 'FULL':
                raise Unsupported('ellipse segment')
            paths = [Path().ellipse(w / 2, h / 2, w / 2, h / 2)]
        else:
            closed = stype in ('PolyPolygonShape', 'ClosedBezierShape', 'PolyPolygonPathShape', 'ClosedFreeHandShape')
            paths = _poly_paths(shape, fr, closed)
            if not paths:
                raise Unsupported('no points')
            if not closed:
                fill = None
        text = _text(shape)
        if text and text_rect:
            x0, y0, x1, y1 = text_rect
            text['box'] = (x0 / HMM, (h - y1) / HMM, (x1 - x0) / HMM, (y1 - y0) / HMM)
        return {'kind': 'shape', 'x': (fr.cx - w / 2 - ox) / HMM, 'y': (oy - fr.cy - h / 2) / HMM,
                'w': w / HMM, 'h': h / HMM, 'angle': fr.angle, 'geom': _to_visio(paths, h), 'fill': fill,
                'line': line, 'text': text}

    def _picture(self, shape, ox, oy, name):
        data = self.export_emf(shape)
        r = shape.getPropertyValue('BoundRect')
        self.pictures += 1
        return {'kind': 'foreign', 'name': name or 'Picture', 'x': (r.X - ox) / HMM,
                'y': (oy - r.Y - r.Height) / HMM, 'w': max(r.Width, 1) / HMM, 'h': max(r.Height, 1) / HMM,
                'image': data}


def _rect_screen(w, h, r):
    """Rectangle in screen coordinates (rect_path builds y-up paths; the shape is symmetric)."""
    return rect_path(w, h, r)


def collect(ctx, shapes, rect, export_emf):
    """shapes (z-order, bottom first) and their bounding rect (x0, y0, x1, y1) in 1/100 mm.

    Returns (dicts, page width in, page height in, collector)."""
    x0, y0, x1, y1 = rect
    col = Collector(ctx, export_emf)
    dicts = [col.convert(s, x0, y1) for s in shapes]
    return dicts, (x1 - x0) / HMM, (y1 - y0) / HMM, col


def emf_exporter(ctx, graphic_export):
    """Returns export_emf(shape) using the engine's graphic_export(ctx, shapes, path, media)."""
    def export_emf(shape):
        fd, p = tempfile.mkstemp(suffix='.emf')
        os.close(fd)
        try:
            graphic_export(ctx, [shape], p, 'image/x-emf')
            return open(p, 'rb').read()
        finally:
            try:
                os.remove(p)
            except OSError:
                pass
    return export_emf

