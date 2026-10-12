"""Checks the Visio geometry the VBA preset engine writes, using LibreOffice as the Visio reader.

The geometry rows are produced exactly as VxPresetGeo / GeoArcTo / GeoEArc / GeoCubic do in
src/VecStamp.bas (MoveTo, LineTo, NURBSTo for Bezier curves, EllipticalArcTo for arcs, inches,
y upwards), packed into a .vsdx with the template parts from src/vsdx, converted to PNG by
LibreOffice and compared with the reference drawing of the same presets (tests/preset_engine.py)
rendered by cairosvg: every preset must cover the same pixels.

    python3 tests/test_vsdx_presets.py            # needs LibreOffice and cairosvg; run as a normal user
"""
import io
import math
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, 'build'))
import make_presets  # noqa: E402
import preset_engine as pe  # noqa: E402

CELL_W, CELL_H, COLS = 110.0, 80.0, 12       # points
GAP = 30.0


def num(v):
    s = ('%.6f' % v).rstrip('0').rstrip('.')
    return '0' if s in ('-0', '') else s


def cell(n, v):
    return f"<Cell N='{n}' V='{v}'/>"


class Geo:
    """Mirror of the VBA Geo* helpers (no flips: GeoXY is x = u / 72, y = (h - v) / 72)."""

    def __init__(self, h):
        self.h = h
        self.sections = []
        self.rows = []
        self.closed = False

    def xy(self, u, v):
        return u / 72.0, (self.h - v) / 72.0

    def begin(self, nofill, noline):
        self.rows, self.nofill, self.noline, self.closed = [], nofill, noline, False

    def row(self, t, cells):
        self.rows.append(f"<Row T='{t}' IX='{len(self.rows) + 1}'>" + ''.join(cell(k, num(v)) for k, v in cells) + '</Row>')

    def move(self, u, v):
        x, y = self.xy(u, v)
        self.row('MoveTo', [('X', x), ('Y', y)])
        self.x0, self.y0, self.lx, self.ly = x, y, x, y
        self.cu, self.cv, self.su, self.sv = u, v, u, v

    def line(self, u, v):
        x, y = self.xy(u, v)
        self.row('LineTo', [('X', x), ('Y', y)])
        self.lx, self.ly, self.cu, self.cv = x, y, u, v

    def cubic(self, u1, v1, u2, v2, u, v):
        x1, y1 = self.xy(u1, v1)
        x2, y2 = self.xy(u2, v2)
        x, y = self.xy(u, v)
        f = f'NURBS(1,3,1,1,{num(x1)},{num(y1)},0,1,{num(x2)},{num(y2)},0,1)'
        self.rows.append(f"<Row T='NURBSTo' IX='{len(self.rows) + 1}'>" + cell('X', num(x)) + cell('Y', num(y)) +
                         cell('A', 1) + cell('B', 1) + cell('C', 0) + cell('D', 1) + f"<Cell N='E' V='{f}' F='{f}'/></Row>")
        self.lx, self.ly, self.cu, self.cv = x, y, u, v

    def quad(self, u1, v1, u, v):
        self.cubic(self.cu + 2 / 3 * (u1 - self.cu), self.cv + 2 / 3 * (v1 - self.cv),
                   u + 2 / 3 * (u1 - u), v + 2 / 3 * (v1 - v), u, v)

    def arc(self, wr, hr, st, sw):
        for (mu, mv, eu, ev, rx, ry) in pe.arc_pieces(self.cu, self.cv, wr, hr, st * pe.DEG, sw * pe.DEG):
            mx, my = self.xy(mu, mv)
            x, y = self.xy(eu, ev)
            c, d = (0.0, rx / ry) if rx >= ry else (math.pi / 2, ry / rx)
            self.row('EllipticalArcTo', [('X', x), ('Y', y), ('A', mx), ('B', my), ('C', c), ('D', d)])
            self.lx, self.ly, self.cu, self.cv = x, y, eu, ev

    def close(self):
        if not self.rows:
            return
        if abs(self.lx - self.x0) > 1e-6 or abs(self.ly - self.y0) > 1e-6:
            self.row('LineTo', [('X', self.x0), ('Y', self.y0)])
            self.lx, self.ly = self.x0, self.y0
        self.cu, self.cv = self.su, self.sv
        self.closed = True

    def end(self):
        if not self.rows:
            return
        nf = 1 if (self.nofill or not self.closed) else 0
        self.sections.append(f"<Section N='Geometry' IX='{len(self.sections)}'>" + cell('NoFill', nf) +
                             cell('NoLine', 1 if self.noline else 0) + cell('NoShow', 0) + ''.join(self.rows) + '</Section>')


def preset_geometry(src, w, h):
    e = pe.Engine(src, w, h)
    g = Geo(h)
    for p in e.paths:
        head, cmds = p.split(':', 1)
        fill, stroke, pw, ph = head.split(',')
        sx = w / float(pw) if float(pw) else 1.0
        sy = h / float(ph) if float(ph) else 1.0
        g.begin(fill != 'f', stroke == '0')
        for c in cmds.split(';'):
            t = c.split(',')
            a = [e.arg(x) for x in t[1:]]
            if t[0] == 'M':
                g.move(a[0] * sx, a[1] * sy)
            elif t[0] == 'L':
                g.line(a[0] * sx, a[1] * sy)
            elif t[0] == 'C':
                g.cubic(a[0] * sx, a[1] * sy, a[2] * sx, a[3] * sy, a[4] * sx, a[5] * sy)
            elif t[0] == 'Q':
                g.quad(a[0] * sx, a[1] * sy, a[2] * sx, a[3] * sy)
            elif t[0] == 'A':
                g.arc(a[0] * sx, a[1] * sy, a[2], a[3])
            elif t[0] == 'Z':
                g.close()
        g.end()
    return ''.join(g.sections)


def build(presets, path):
    shapes, page_h = [], math.ceil(len(presets) / COLS) * (CELL_H + GAP) + GAP
    page_w = COLS * (CELL_W + GAP) + GAP
    for i, (_ast, _prst, src) in enumerate(presets):
        x = GAP + (i % COLS) * (CELL_W + GAP)
        top = GAP + (i // COLS) * (CELL_H + GAP)
        y = page_h - top - CELL_H
        w, h = CELL_W / 72, CELL_H / 72
        shapes.append(f"<Shape ID='{i + 1}' NameU='s{i}' Name='s{i}' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>"
                      + cell('PinX', num(x / 72 + w / 2)) + cell('PinY', num(y / 72 + h / 2)) + cell('Width', num(w))
                      + cell('Height', num(h)) + cell('LocPinX', num(w / 2)) + cell('LocPinY', num(h / 2)) + cell('Angle', 0)
                      + cell('FillForegnd', '#9CC3E6') + cell('FillPattern', 1) + cell('LineColor', '#1F4E79')
                      + "<Cell N='LineWeight' V='0.0166667' U='PT'/>" + cell('LinePattern', 1)
                      + preset_geometry(src, CELL_W, CELL_H) + '</Shape>')
    values = {'PW': num(page_w / 72), 'PH': num(page_h / 72), 'W': num(page_w / 72), 'H': num(page_h / 72),
              'CX': num(page_w / 144), 'CY': num(page_h / 144), 'LX': num(page_w / 144), 'LY': num(page_h / 144),
              'VERSION': 'test', 'TITLE': 'presets', 'CREATED': '2026-01-01T00:00:00Z', 'RELS': '',
              'SHAPES': ''.join(shapes)}
    vdir = os.path.join(ROOT, 'src', 'vsdx')
    entries = [l.strip().split('=', 1) for l in open(os.path.join(vdir, 'map.txt'), encoding='utf-8') if '=' in l]
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for src_name, dst in entries:
            txt = open(os.path.join(vdir, src_name), encoding='utf-8').read()
            for k in [k for k in values if k != 'SHAPES'] + ['SHAPES']:
                txt = txt.replace('{' + k + '}', values[k])
            z.writestr(dst, txt)
    return page_w, page_h


def reference_png(presets, page_w, page_h, scale):
    import cairosvg
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{page_w}" height="{page_h}"><rect width="100%" height="100%" fill="#fff"/>']
    for i, (_ast, _prst, src) in enumerate(presets):
        x = GAP + (i % COLS) * (CELL_W + GAP)
        y = GAP + (i // COLS) * (CELL_H + GAP)
        parts.append(f'<g transform="translate({x},{y})">')
        for d, filled, stroked in pe.Engine(src, CELL_W, CELL_H).svg_paths():
            parts.append(f'<path d="{d}" fill="{"#9cc3e6" if filled else "none"}" fill-rule="evenodd" '
                         f'stroke="{"#1f4e79" if stroked else "none"}" stroke-width="1.2"/>')
        parts.append('</g>')
    parts.append('</svg>')
    return cairosvg.svg2png(bytestring=''.join(parts).encode(), output_width=int(page_w * scale))


def soffice():
    for c in ('soffice', 'libreoffice', '/opt/libreoffice/program/soffice'):
        p = shutil.which(c) or (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def main():
    from PIL import Image, ImageChops
    exe = soffice()
    if not exe:
        print('SKIP: LibreOffice not found')
        return
    compiled = make_presets.compile_all()
    seen, presets = set(), []
    for ast, (prst, src) in sorted(compiled.items()):
        if prst not in seen:
            seen.add(prst)
            presets.append((ast, prst, src))
    tmp = tempfile.mkdtemp(prefix='vs-presets-')
    vsdx = os.path.join(tmp, 'presets.vsdx')
    page_w, page_h = build(presets, vsdx)
    prof = 'file://' + os.path.join(tmp, 'profile')
    subprocess.run([exe, '-env:UserInstallation=' + prof, '--headless', '--norestore', '--convert-to', 'png',
                    '--outdir', tmp, vsdx], check=True, capture_output=True, timeout=300)
    got = Image.open(os.path.join(tmp, 'presets.png')).convert('RGB')
    scale = got.width / page_w
    ref = Image.open(io.BytesIO(reference_png(presets, page_w, page_h, scale))).convert('RGB').resize(got.size)
    bad = []
    for i, (ast, prst, _src) in enumerate(presets):
        x = (GAP + (i % COLS) * (CELL_W + GAP) - 6) * scale
        y = (GAP + (i // COLS) * (CELL_H + GAP) - 6) * scale
        box = tuple(int(v) for v in (x, y, x + (CELL_W + 12) * scale, y + (CELL_H + 12) * scale))
        a = got.crop(box).convert('L').point(lambda v: 255 if v > 240 else 0)
        b = ref.crop(box).convert('L').point(lambda v: 255 if v > 240 else 0)
        diff = ImageChops.difference(a, b)
        frac = sum(1 for v in diff.getdata() if v) / (diff.width * diff.height)
        if frac > 0.04:
            bad.append((ast, prst, round(frac, 3)))
    side = Image.new('RGB', (got.width * 2, got.height), 'white')
    side.paste(got, (0, 0))
    side.paste(ref, (got.width, 0))
    keep = os.environ.get('VECSTAMP_KEEP')
    if keep:
        side.save(os.path.join(HERE, 'presets-compare.png'))
        print('wrote tests/presets-compare.png (left: VSDX in LibreOffice, right: reference)')
    shutil.rmtree(tmp, ignore_errors=True)
    print(f'{len(presets)} presets compared, {len(bad)} differ by more than 4 % of their cell')
    for b in bad:
        print('  ', b)
    if bad:
        sys.exit(1)
    print('OK')


if __name__ == '__main__':
    main()
