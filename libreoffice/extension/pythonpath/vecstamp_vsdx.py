"""Visio .vsdx writer for VecStamp (Shi Yin).

Two kinds of drawings are written:

* native  - every exported shape becomes a real Visio shape (geometry, fill, line,
            arrowheads, rotation, editable rich text, groups); pictures and other
            objects that have no Visio equivalent are embedded as EMF pictures;
* picture - the whole drawing embedded as one EMF picture (exact appearance).

Shapes are described with plain dicts (see ``shape`` below); the PowerPoint add-in
(src/VecStamp.bas, section "Native VSDX") writes the same XML.

Units: inches, Visio coordinates (y grows upwards, origin at the bottom-left of
the page or of the parent group).

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
The package parts live in vsdx_template/ (copied from src/vsdx by the build).
"""
import datetime
import math
import os
import re
import struct
import zipfile

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vsdx_template')
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
KAPPA = 0.5522847498307936          # cubic Bezier approximation of a quarter circle

# arrowhead styles (shared names) -> Visio BeginArrow / EndArrow index
ARROWS = {'none': 0, 'triangle': 4, 'open': 1, 'stealth': 5, 'diamond': 22, 'oval': 10, 'square': 11}
# dash styles -> Visio LinePattern
PATTERNS = {'solid': 1, 'dash': 2, 'dot': 3, 'dashdot': 4, 'dashdotdot': 5, 'longdash': 9,
            'longdashdot': 7, 'rounddot': 10}
_CJK = re.compile('[⺀-鿿가-힯豈-﫿＀-￯]')
_BAD = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')


def num(v):
    """Number as text for XML: '.' decimal separator, at most 6 decimals."""
    if abs(v) < 5e-7:
        return '0'
    s = ('%.6f' % v).rstrip('0').rstrip('.')
    return '0' if s in ('-0', '') else s


def esc_attr(s):
    return (_BAD.sub('', s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace("'", '&apos;').replace('"', '&quot;'))


def esc_text(s):
    return _BAD.sub('', s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def hexcolor(rgb):
    """(r, g, b) -> '#RRGGBB'."""
    return '#%02X%02X%02X' % tuple(max(0, min(255, int(round(c)))) for c in rgb)


def cell(name, value, unit=None, formula=None):
    v = value if isinstance(value, str) else num(value)
    s = "<Cell N='%s' V='%s'" % (name, esc_attr(v))
    if unit:
        s += " U='%s'" % unit
    if formula:
        s += " F='%s'" % esc_attr(formula)
    return s + '/>'


# ------------------------------------------------------------------ geometry
class Path:
    """One sub-path in local shape coordinates (inches, y up)."""

    def __init__(self, closed=True, nofill=False, noline=False):
        self.closed, self.nofill, self.noline = closed, nofill, noline
        self.segs = []          # ('M', x, y) | ('L', x, y) | ('C', x1, y1, x2, y2, x, y) | ('E', cx, cy, rx, ry)

    def move(self, x, y):
        self.segs.append(('M', x, y))
        return self

    def line(self, x, y):
        self.segs.append(('L', x, y))
        return self

    def cubic(self, x1, y1, x2, y2, x, y):
        self.segs.append(('C', x1, y1, x2, y2, x, y))
        return self

    def ellipse(self, cx, cy, rx, ry):
        self.segs.append(('E', cx, cy, rx, ry))
        return self


def rect_path(w, h, r=0.0, rx=None, ry=None):
    """Rectangle (optionally with rounded corners rx / ry) as a closed path."""
    rx = r if rx is None else rx
    ry = r if ry is None else ry
    rx, ry = max(0.0, min(rx, w / 2)), max(0.0, min(ry, h / 2))
    p = Path()
    if rx <= 0 or ry <= 0:
        return p.move(0, 0).line(w, 0).line(w, h).line(0, h).line(0, 0)
    kx, ky = rx * KAPPA, ry * KAPPA
    p.move(rx, 0).line(w - rx, 0).cubic(w - rx + kx, 0, w, ry - ky, w, ry)
    p.line(w, h - ry).cubic(w, h - ry + ky, w - rx + kx, h, w - rx, h)
    p.line(rx, h).cubic(rx - kx, h, 0, h - ry + ky, 0, h - ry)
    p.line(0, ry).cubic(0, ry - ky, rx - kx, 0, rx, 0)
    return p


def geometry_xml(paths, w, h):
    out = []
    for ix, p in enumerate(paths):
        if not p.segs:
            continue
        rows = []
        first = None
        last = None
        for k, s in enumerate(p.segs, 1):
            t = s[0]
            if t == 'E':
                cx, cy, rx, ry = s[1:]
                rows.append("<Row T='Ellipse' IX='%d'>%s%s%s%s%s%s</Row>" % (
                    k, cell('X', cx), cell('Y', cy), cell('A', cx + rx), cell('B', cy),
                    cell('C', cx), cell('D', cy + ry)))
                continue
            if t == 'M':
                rows.append("<Row T='MoveTo' IX='%d'>%s%s</Row>" % (k, cell('X', s[1]), cell('Y', s[2])))
                first = last = (s[1], s[2])
            elif t == 'L':
                rows.append("<Row T='LineTo' IX='%d'>%s%s</Row>" % (k, cell('X', s[1]), cell('Y', s[2])))
                last = (s[1], s[2])
            elif t == 'C':
                x1, y1, x2, y2, x, y = s[1:]
                nurbs = 'NURBS(1,3,1,1,%s,%s,0,1,%s,%s,0,1)' % (num(x1), num(y1), num(x2), num(y2))
                rows.append("<Row T='NURBSTo' IX='%d'>%s%s%s%s%s%s%s</Row>" % (
                    k, cell('X', x), cell('Y', y), cell('A', 1), cell('B', 1), cell('C', 0), cell('D', 1),
                    cell('E', nurbs, formula=nurbs)))
                last = (x, y)
        if p.closed and first and last and (abs(first[0] - last[0]) > 1e-6 or abs(first[1] - last[1]) > 1e-6):
            rows.append("<Row T='LineTo' IX='%d'>%s%s</Row>" % (len(p.segs) + 1, cell('X', first[0]),
                                                                cell('Y', first[1])))
        nofill = p.nofill or not p.closed
        out.append("<Section N='Geometry' IX='%d'>%s%s%s%s</Section>" % (
            ix, cell('NoFill', 1 if nofill else 0), cell('NoLine', 1 if p.noline else 0),
            cell('NoShow', 0), ''.join(rows)))
    return ''.join(out)


# ------------------------------------------------------------------ text
def text_xml(text, w, h):
    """Character / Paragraph sections, text block cells and the <Text> element.

    One Character row per run and one Paragraph row per paragraph, in text order."""
    char_rows, para_rows, body = [], [], []
    paras = text.get('paras', [])
    for pi, para in enumerate(paras):
        spline = para.get('spline', -1.2)
        para_rows.append("<Row IX='%d'>%s%s%s%s%s</Row>" % (
            pi, cell('HorzAlign', para.get('align', 1)), cell('SpLine', spline, 'PT' if spline > 0 else None),
            cell('SpBefore', para.get('before', 0) / 72.0, 'PT'), cell('SpAfter', para.get('after', 0) / 72.0, 'PT'),
            cell('Bullet', 0)))
        body.append("<pp IX='%d'/>" % pi)
        for run in para.get('runs') or []:
            font = run.get('font') or 'Calibri'
            style = int(bool(run.get('bold'))) | 2 * int(bool(run.get('italic'))) | 4 * int(bool(run.get('underline')))
            ix = len(char_rows)
            char_rows.append("<Row IX='%d'>%s%s%s%s%s%s%s</Row>" % (
                ix, cell('Font', font), cell('AsianFont', run.get('asian') or font),
                cell('Color', hexcolor(run.get('color', (0, 0, 0)))), cell('Style', style),
                cell('Size', run.get('size', 18) / 72.0, 'PT'), cell('Pos', int(run.get('pos', 0))),
                cell('LangID', 'zh-CN' if _CJK.search(run.get('text', '')) else 'en-US')))
            body.append("<cp IX='%d'/>" % ix)
            body.append(esc_text(run.get('text', '').replace('\r', '\n').replace('\x0b', '\u2028')))
        if pi < len(paras) - 1:
            body.append('\n')
    l, r, t, b = text.get('margins', (7.2, 7.2, 3.6, 3.6))
    cells = [cell('LeftMargin', l / 72.0, 'PT'), cell('RightMargin', r / 72.0, 'PT'),
             cell('TopMargin', t / 72.0, 'PT'), cell('BottomMargin', b / 72.0, 'PT'),
             cell('VerticalAlign', text.get('valign', 1))]
    if text.get('vertical'):
        cells.append(cell('TextDirection', 1))
    angle = text.get('angle', 0.0)
    box = text.get('box')                     # (x, y, w, h) of the text area inside the shape
    bx, by, tw, th = box if box else (0.0, 0.0, w, h)
    if abs(math.sin(angle)) > 0.7:
        tw, th = th, tw
    if text.get('nowrap'):
        tw = max(tw * 2, tw + 1.0)
    if box or angle or text.get('nowrap'):
        align = text['paras'][0].get('align', 1) if text.get('paras') else 1
        cx = bx + (box[2] if box else w) / 2
        if not text.get('nowrap') or align == 1:
            px = cx
        elif align == 2:
            px = bx + (box[2] if box else w) - tw / 2
        else:
            px = bx + tw / 2
        cells += [cell('TxtPinX', px), cell('TxtPinY', by + (box[3] if box else h) / 2), cell('TxtWidth', tw),
                  cell('TxtHeight', th), cell('TxtLocPinX', tw / 2), cell('TxtLocPinY', th / 2),
                  cell('TxtAngle', angle)]
    sections = "<Section N='Character'>%s</Section><Section N='Paragraph'>%s</Section>" % (
        ''.join(char_rows), ''.join(para_rows))
    return ''.join(cells), sections, '<Text>%s</Text>' % ''.join(body)


# ------------------------------------------------------------------ shapes
class Writer:
    """Turns shape dicts into page XML. Media (EMF pictures) are collected in self.media."""

    def __init__(self):
        self.next_id = 1
        self.media = []

    def _xform(self, sh):
        w, h = max(sh['w'], 1e-4), max(sh['h'], 1e-4)
        return ''.join([cell('PinX', sh['x'] + w / 2), cell('PinY', sh['y'] + h / 2), cell('Width', w),
                        cell('Height', h), cell('LocPinX', w / 2, formula='Width*0.5'),
                        cell('LocPinY', h / 2, formula='Height*0.5'), cell('Angle', sh.get('angle', 0.0)),
                        cell('FlipX', 0), cell('FlipY', 0)])

    @staticmethod
    def _style(sh):
        out = []
        fill = sh.get('fill')
        if fill:
            out += [cell('FillForegnd', hexcolor(fill['color'])), cell('FillPattern', 1),
                    cell('FillForegndTrans', fill.get('trans', 0))]
            stops = fill.get('stops')
            if stops and len(stops) > 1:
                out += [cell('FillGradientEnabled', 1), cell('FillGradientDir', 0),
                        cell('FillGradientAngle', fill.get('angle', 0.0)), cell('RotateGradientWithShape', 1)]
        else:
            out.append(cell('FillPattern', 0))
        line = sh.get('line')
        if line:
            out += [cell('LineWeight', max(line.get('weight', 0.75), 0.1) / 72.0, 'PT'),
                    cell('LineColor', hexcolor(line['color'])), cell('LinePattern', line.get('pattern', 1)),
                    cell('LineColorTrans', line.get('trans', 0))]
            for end in ('Begin', 'End'):
                a = line.get(end.lower() + '_arrow', 0)
                if a:
                    out += [cell(end + 'Arrow', a), cell(end + 'ArrowSize', line.get(end.lower() + '_size', 2))]
        else:
            out.append(cell('LinePattern', 0))
        return ''.join(out)

    @staticmethod
    def _gradient(sh):
        fill = sh.get('fill') or {}
        stops = fill.get('stops')
        if not stops or len(stops) < 2:
            return ''
        rows = ''.join("<Row IX='%d'>%s%s%s</Row>" % (
            i, cell('GradientStopColor', hexcolor(c)), cell('GradientStopColorTrans', t),
            cell('GradientStopPosition', pos)) for i, (pos, c, t) in enumerate(stops))
        return "<Section N='FillGradient'>%s</Section>" % rows

    def shape_xml(self, sh):
        sid = self.next_id
        self.next_id += 1
        name = esc_attr(sh.get('name') or 'Shape %d' % sid)
        kind = sh.get('kind', 'shape')
        head = "<Shape ID='%d' NameU='%s' Name='%s' Type='%s' LineStyle='0' FillStyle='0' TextStyle='0'>" % (
            sid, name, name, {'group': 'Group', 'foreign': 'Foreign'}.get(kind, 'Shape'))
        w, h = max(sh['w'], 1e-4), max(sh['h'], 1e-4)
        if kind == 'group':
            kids = ''.join(self.shape_xml(c) for c in sh.get('children', []))
            return head + self._xform(sh) + '<Shapes>%s</Shapes></Shape>' % kids
        if kind == 'foreign':
            self.media.append(sh['image'])
            rid = 'rId%d' % len(self.media)
            return (head + self._xform(sh) + cell('ResizeMode', 0) +
                    cell('ImgOffsetX', 0, formula='ImgWidth*0') + cell('ImgOffsetY', 0, formula='ImgHeight*0') +
                    cell('ImgWidth', w, formula='Width*1') + cell('ImgHeight', h, formula='Height*1') +
                    cell('LinePattern', 0) + cell('FillPattern', 0) +
                    geometry_xml([rect_path(w, h)], w, h).replace("<Cell N='NoLine' V='0'/>",
                                                                   "<Cell N='NoLine' V='1'/>") +
                    "<ForeignData ForeignType='EnhMetaFile' ObjectWidth='%s' ObjectHeight='%s'>"
                    "<Rel r:id='%s'/></ForeignData></Shape>" % (num(w), num(h), rid))
        body = [head, self._xform(sh), self._style(sh)]
        text = sh.get('text')
        tcells = tsections = telement = ''
        if text and any(r.get('text') for p in text.get('paras', []) for r in p.get('runs', [])):
            tcells, tsections, telement = text_xml(text, w, h)
        body += [tcells, self._gradient(sh), tsections, geometry_xml(sh.get('geom', []), w, h), telement, '</Shape>']
        return ''.join(body)

    def page_xml(self, shapes):
        return ''.join(self.shape_xml(s) for s in shapes)

    def rels_xml(self):
        return ''.join('<Relationship Id="rId%d" Type="%s" Target="../media/image%d.emf"/>' % (i, REL_IMAGE, i)
                       for i in range(1, len(self.media) + 1))


def _template_entries():
    entries = []
    with open(os.path.join(TEMPLATE_DIR, 'map.txt'), encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if '=' in line:
                entries.append(tuple(line.split('=', 1)))
    return entries


def save(out_path, page_w, page_h, shapes, version='1.2.1', title='VecStamp drawing'):
    """Writes a one-page .vsdx with the given shapes."""
    writer = Writer()
    shapes_xml = writer.page_xml(shapes)
    values = {
        'PW': num(page_w), 'PH': num(page_h), 'W': num(page_w), 'H': num(page_h),
        'CX': num(page_w / 2), 'CY': num(page_h / 2), 'LX': num(page_w / 2), 'LY': num(page_h / 2),
        'VERSION': esc_text(version), 'TITLE': esc_text(title),
        'CREATED': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        'SHAPES': shapes_xml, 'RELS': writer.rels_xml(),
    }
    tmp = out_path + '.part'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for src, dst in _template_entries():
            text = open(os.path.join(TEMPLATE_DIR, src), encoding='utf-8').read()
            for k in ('SHAPES', 'RELS'):          # filled last: they may contain braces of their own
                text = text.replace('{' + k + '}', '\0' + k + '\0')
            for k, v in values.items():
                if k not in ('SHAPES', 'RELS'):
                    text = text.replace('{' + k + '}', v)
            text = text.replace('\0SHAPES\0', values['SHAPES']).replace('\0RELS\0', values['RELS'])
            z.writestr(dst, text.encode('utf-8'))
        for i, data in enumerate(writer.media, 1):
            z.writestr('visio/media/image%d.emf' % i, data)
    os.replace(tmp, out_path)
    return out_path


def emf_frame_inches(emf_bytes):
    """Size of an EMF picture from its header frame (0.01 mm units)."""
    if len(emf_bytes) < 40 or struct.unpack('<I', emf_bytes[:4])[0] != 1:
        raise ValueError('not an EMF file')
    left, top, right, bottom = struct.unpack('<4i', emf_bytes[24:40])
    return max(right - left, 1) / 2540.0, max(bottom - top, 1) / 2540.0


def build_vsdx(emf_path, out_path, version='1.2.1', title='VecStamp drawing', size_in=None):
    """Picture mode: the whole drawing as one embedded EMF picture."""
    data = open(emf_path, 'rb').read()
    w, h = size_in or emf_frame_inches(data)
    shape = {'kind': 'foreign', 'name': 'VecStamp', 'x': 0.0, 'y': 0.0, 'w': w, 'h': h, 'image': data}
    return save(out_path, w, h, [shape], version, title)
