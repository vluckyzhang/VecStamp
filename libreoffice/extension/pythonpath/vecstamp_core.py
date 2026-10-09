"""VecStamp (Shi Yin) for LibreOffice Impress / Draw - export engine.

Exports the current selection (or whole slides) as EMF, EMZ, SVG, WMF, PNG,
PDF or VSDX, with optional background (white, beige, grid, custom colour),
transparency and margin. No user interface here: see vecstamp_ui.py.

Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.
"""
import gzip
import json
import os
import re
import shutil
import sys
import tempfile
import time

import uno
from com.sun.star.awt import Point, Size
from com.sun.star.beans import PropertyValue

import vecstamp_shapes
import vecstamp_vsdx

VERSION = '1.2.1'
AUTHOR = 'vluckyzhang'
EMAIL = 'vluckyzhang@gmail.com'
REPO = 'https://github.com/vluckyzhang/VecStamp'

# key: (extension, label)
FORMATS = [
    ('emf', 'EMF 增强型图元文件'),
    ('emz', 'EMZ 压缩的增强型图元文件'),
    ('svg', 'SVG 可缩放矢量图形'),
    ('wmf', 'WMF Windows 图元文件'),
    ('png', 'PNG 高分辨率位图'),
    ('pdf', 'PDF 矢量文档'),
    ('vsdx', 'VSDX Visio 绘图'),
]
FMT_INDEX = {ext: i for i, (ext, _) in enumerate(FORMATS)}
MEDIA = {'emf': 'image/x-emf', 'wmf': 'image/x-wmf', 'svg': 'image/svg+xml', 'png': 'image/png'}
BACKGROUNDS = ['透明', '白色', '米色', '网格', '自定义颜色']
BG_NONE, BG_WHITE, BG_BEIGE, BG_GRID, BG_CUSTOM = range(5)
DPI_PRESETS = [150, 300, 600]
DPI_MAX = 3000
PNG_WARN_PIXELS = 120_000_000
PT = 2540.0 / 72.0          # 1 pt in 1/100 mm

DEFAULTS = {
    'format': 0, 'mode': 0, 'naming': 0, 'dpi_index': 1, 'dpi_custom': 1200,
    'bg_type': 0, 'bg_color': [255, 248, 235], 'bg_alpha': 0, 'grid_step': 10.0,
    'margin': 0.0, 'open_folder': False, 'folder': '', 'last_folder': '', 'vsdx_native': True,
}


class Cancelled(Exception):
    pass


# --------------------------------------------------------------------- settings
def settings_path():
    if sys.platform.startswith('win'):
        base = os.environ.get('APPDATA') or os.path.expanduser('~')
    elif sys.platform == 'darwin':
        base = os.path.expanduser('~/Library/Application Support')
    else:
        base = os.environ.get('XDG_CONFIG_HOME') or os.path.expanduser('~/.config')
    return os.path.join(base, 'VecStamp', 'libreoffice-settings.json')


class Settings(dict):
    def __init__(self, path=None):
        super().__init__(DEFAULTS)
        self.path = path or settings_path()
        try:
            with open(self.path, encoding='utf-8') as f:
                data = json.load(f)
            for k, v in data.items():
                if k in DEFAULTS:
                    self[k] = v
        except (OSError, ValueError):
            pass
        self.normalize()

    def normalize(self):
        def clamp(key, lo, hi, cast):
            try:
                v = cast(self[key])
            except (TypeError, ValueError):
                v = DEFAULTS[key]
            self[key] = min(max(v, lo), hi)
        clamp('format', 0, len(FORMATS) - 1, int)
        clamp('mode', 0, 1, int)
        clamp('naming', 0, 1, int)
        clamp('dpi_index', 0, 3, int)
        clamp('dpi_custom', 1, DPI_MAX, int)
        clamp('bg_type', 0, 4, int)
        clamp('bg_alpha', 0, 100, int)
        clamp('grid_step', 2.0, 200.0, float)
        clamp('margin', 0.0, 500.0, float)
        self['vsdx_native'] = bool(self.get('vsdx_native', True))
        c = self.get('bg_color')
        if not (isinstance(c, (list, tuple)) and len(c) == 3 and all(isinstance(x, int) and 0 <= x <= 255 for x in c)):
            self['bg_color'] = list(DEFAULTS['bg_color'])

    def save(self):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump({k: self[k] for k in DEFAULTS}, f, ensure_ascii=False, indent=2)

    def reset(self):
        self.clear()
        self.update(DEFAULTS)
        self['bg_color'] = list(DEFAULTS['bg_color'])

    @property
    def dpi(self):
        i = self['dpi_index']
        return self['dpi_custom'] if i == 3 else DPI_PRESETS[i]

    @property
    def ext(self):
        return FORMATS[self['format']][0]


def parse_color(text):
    """'255,248,235', '255 248 235', 'rgb(255,248,235)', '#FFF8EB', 'FFF8EB', '#FFF' -> [r, g, b] or None."""
    t = text.strip().lower().replace('rgb', '').replace('(', '').replace(')', '').replace('，', ',')
    h = t[1:] if t.startswith('#') else t
    if re.fullmatch(r'[0-9a-f]{6}|[0-9a-f]{3}', h):
        if len(h) == 3:
            h = ''.join(c * 2 for c in h)
        return [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    parts = [p for p in re.split(r'[,;\s]+', t) if p]
    if len(parts) == 3 and all(re.fullmatch(r'\d{1,3}', p) for p in parts):
        vals = [int(p) for p in parts]
        if all(0 <= v <= 255 for v in vals):
            return vals
    return None


def color_text(c):
    return '%d,%d,%d' % tuple(c)


def rgb_int(c):
    return (c[0] << 16) | (c[1] << 8) | c[2]


def bg_color(settings):
    t = settings['bg_type']
    if t == BG_BEIGE:
        return [245, 240, 225]
    if t == BG_CUSTOM:
        return list(settings['bg_color'])
    return [255, 255, 255]


# --------------------------------------------------------------------- helpers
def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def props(*pairs):
    return tuple(pv(n, v) for n, v in pairs)


def to_url(path):
    return uno.systemPathToFileUrl(os.path.abspath(path))


_TEMP_DIR = None


def temp_dir():
    """Per-user scratch folder (shared /tmp folders of other users are never reused)."""
    global _TEMP_DIR
    if _TEMP_DIR and os.path.isdir(_TEMP_DIR):
        return _TEMP_DIR
    try:
        import getpass
        user = re.sub(r'[^A-Za-z0-9_.-]', '_', getpass.getuser())
    except Exception:
        user = 'user'
    d = os.path.join(tempfile.gettempdir(), 'VecStamp-' + user)
    try:
        os.makedirs(d, mode=0o700, exist_ok=True)
        if not os.access(d, os.W_OK):
            raise PermissionError(d)
    except OSError:
        d = tempfile.mkdtemp(prefix='VecStamp-')
    _TEMP_DIR = d
    return d


def temp_file(ext):
    fd, p = tempfile.mkstemp(prefix='vs_', suffix='.' + ext, dir=temp_dir())
    os.close(fd)
    return p


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\t\r\n]', '_', s or '').strip().rstrip('.')
    return (s[:80] or 'shape')


def unique_path(folder, base, ext):
    p = os.path.join(folder, '%s.%s' % (base, ext))
    k = 2
    while os.path.exists(p):
        p = os.path.join(folder, '%s_%d.%s' % (base, k, ext))
        k += 1
    return p


def doc_base_name(doc):
    url = doc.getURL()
    if url:
        name = os.path.splitext(os.path.basename(uno.fileUrlToSystemPath(url)))[0]
    else:
        name = doc.getTitle() or 'VecStamp'
        name = os.path.splitext(name)[0]
    return safe_name(name)


def doc_folder(doc):
    url = doc.getURL()
    if url.startswith('file:'):
        d = os.path.dirname(uno.fileUrlToSystemPath(url))
        if os.path.isdir(d):
            return d
    return ''


def desktop_folder():
    for d in (os.path.join(os.path.expanduser('~'), 'Desktop'), os.path.join(os.path.expanduser('~'), '桌面')):
        if os.path.isdir(d):
            return d
    return os.path.expanduser('~')


def default_folder(doc, settings):
    for d in (settings['folder'], doc_folder(doc)):
        if d and os.path.isdir(d):
            return d
    return desktop_folder()


def initial_folder(doc, settings):
    d = settings['last_folder']
    return d if d and os.path.isdir(d) else default_folder(doc, settings)


# --------------------------------------------------------------------- shapes
class Target:
    """What to export: a list of shapes on one page."""

    def __init__(self, page, shapes, page_index=None):
        self.page = page
        self.shapes = list(shapes)
        self.page_index = page_index


def iter_shapes(container):
    return [container.getByIndex(i) for i in range(container.getCount())]


def is_empty_placeholder(shape):
    try:
        info = shape.getPropertySetInfo()
        if info.hasPropertyByName('IsEmptyPresentationObject'):
            return bool(shape.getPropertyValue('IsEmptyPresentationObject'))
    except Exception:
        pass
    return False


def page_shapes(page):
    return [s for s in iter_shapes(page) if not is_empty_placeholder(s)]


def selection_target(doc):
    """Target for the current selection, or None when nothing is selected."""
    ctrl = doc.getCurrentController()
    sel = ctrl.getSelection()
    if sel is None:
        return None
    if hasattr(sel, 'getCount'):
        shapes = iter_shapes(sel)
    elif hasattr(sel, 'supportsService') and sel.supportsService('com.sun.star.drawing.Shape'):
        shapes = [sel]
    else:
        return None
    if not shapes:
        return None
    page = ctrl.getCurrentPage()
    return Target(page, shapes, page_index_of(doc, page))


def current_page_target(doc):
    page = doc.getCurrentController().getCurrentPage()
    shapes = page_shapes(page)
    return Target(page, shapes, page_index_of(doc, page)) if shapes else None


def page_index_of(doc, page):
    pages = doc.getDrawPages()
    for i in range(pages.getCount()):
        if pages.getByIndex(i) == page:
            return i + 1
    return None


def bound_rect(shapes):
    xs0, ys0, xs1, ys1 = [], [], [], []
    for s in shapes:
        r = s.getPropertyValue('BoundRect')
        xs0.append(r.X)
        ys0.append(r.Y)
        xs1.append(r.X + r.Width)
        ys1.append(r.Y + r.Height)
    return min(xs0), min(ys0), max(xs1), max(ys1)


def make_collection(ctx, shapes):
    coll = ctx.ServiceManager.createInstanceWithContext('com.sun.star.drawing.ShapeCollection', ctx)
    for s in shapes:
        coll.add(s)
    return coll


class Composition:
    """Temporarily adds background / grid / margin shapes under the target on its page."""

    def __init__(self, ctx, doc, target, settings):
        self.ctx, self.doc, self.target, self.settings = ctx, doc, target, settings
        self.added = []
        self.undo = None
        self.modified = doc.isModified()

    def __enter__(self):
        s = self.settings
        bg, margin = s['bg_type'], s['margin'] * PT
        x0, y0, x1, y1 = bound_rect(self.target.shapes)
        x0, y0, x1, y1 = x0 - margin, y0 - margin, x1 + margin, y1 + margin
        self.rect = (x0, y0, x1, y1)
        if bg == BG_NONE and margin <= 0:
            return self
        try:
            self.undo = self.doc.getUndoManager()
            self.undo.lock()
        except Exception:
            self.undo = None
        page = self.target.page
        z = min(sh.getPropertyValue('ZOrder') for sh in self.target.shapes)
        rect = self.doc.createInstance('com.sun.star.drawing.RectangleShape')
        page.add(rect)
        rect.setPosition(Point(int(x0), int(y0)))
        rect.setSize(Size(int(x1 - x0), int(y1 - y0)))
        rect.setPropertyValue('LineStyle', uno.Enum('com.sun.star.drawing.LineStyle', 'NONE'))
        rect.setPropertyValue('FillStyle', uno.Enum('com.sun.star.drawing.FillStyle', 'SOLID'))
        rect.setPropertyValue('FillColor', rgb_int(bg_color(s)))
        rect.setPropertyValue('FillTransparence', 100 if bg == BG_NONE else int(s['bg_alpha']))
        self.added.append(rect)
        if bg == BG_GRID:
            self._grid(page, x0, y0, x1, y1)
        for i, shp in enumerate(self.added):
            shp.setPropertyValue('ZOrder', z + i)
        return self

    def _grid(self, page, x0, y0, x1, y1):
        step = self.settings['grid_step'] * PT
        while ((x1 - x0) + (y1 - y0)) / step > 600:
            step *= 2
        lines = []
        i, x = 1, x0 + step
        while x < x1 - 1:
            lines.append(((x, y0), (x, y1), i % 5 == 0))
            i += 1
            x = x0 + i * step
        i, y = 1, y0 + step
        while y < y1 - 1:
            lines.append(((x0, y), (x1, y), i % 5 == 0))
            i += 1
            y = y0 + i * step
        for (ax, ay), (bx, by), major in lines:
            ln = self.doc.createInstance('com.sun.star.drawing.LineShape')
            page.add(ln)
            poly = uno.Any('[][]com.sun.star.awt.Point', ((Point(int(ax), int(ay)), Point(int(bx), int(by))),))
            uno.invoke(ln, 'setPropertyValue', ('PolyPolygon', poly))
            ln.setPropertyValue('LineColor', 0xB0B8C4 if major else 0xDEE2E8)
            ln.setPropertyValue('LineWidth', int((0.75 if major else 0.5) * PT))
            self.added.append(ln)

    def shapes(self):
        return self.added + self.target.shapes

    def __exit__(self, *exc):
        for shp in self.added:
            try:
                self.target.page.remove(shp)
            except Exception:
                pass
        if self.undo is not None:
            try:
                self.undo.unlock()
            except Exception:
                pass
        try:
            self.doc.setModified(self.modified)
        except Exception:
            pass
        return False


# --------------------------------------------------------------------- export
def graphic_export(ctx, shapes, path, media, filter_data=()):
    flt = ctx.ServiceManager.createInstanceWithContext('com.sun.star.drawing.GraphicExportFilter', ctx)
    flt.setSourceDocument(make_collection(ctx, shapes))
    args = [pv('URL', to_url(path)), pv('MediaType', media)]
    if filter_data:
        args.append(pv('FilterData', uno.Any('[]com.sun.star.beans.PropertyValue', tuple(filter_data))))
    flt.filter(tuple(args))
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise RuntimeError('LibreOffice did not write %s' % path)


def png_pixels(rect, dpi):
    x0, y0, x1, y1 = rect
    return max(1, round((x1 - x0) / 2540.0 * dpi)), max(1, round((y1 - y0) / 2540.0 * dpi))


def fix_svg(path):
    """Ensures the root element is <svg> in the SVG namespace. Returns False if not SVG."""
    data = open(path, 'rb').read()
    if data[:2] == b'\xff\xfe':
        return b'<\x00s\x00v\x00g\x00' in data
    head = data[:65536]
    m = re.search(rb'<(?![?!])([A-Za-z_][\w:.-]*)', re.sub(rb'<!--.*?-->', lambda x: b' ' * len(x.group()), head, flags=re.S))
    if not m:
        return False
    name = m.group(1).lower()
    if name != b'svg':
        return name.endswith(b':svg')
    end = data.find(b'>', m.start())
    tag = data[m.start():end]
    ins = b''
    if b'http://www.w3.org/2000/svg' not in tag:
        ins += b' xmlns="http://www.w3.org/2000/svg"'
    if b'xlink:' in data and b'xmlns:xlink' not in tag:
        ins += b' xmlns:xlink="http://www.w3.org/1999/xlink"'
    if ins:
        pos = m.start() + 4
        open(path, 'wb').write(data[:pos] + ins + data[pos:])
    return True


def export_target(ctx, doc, target, ext, out_path, settings, confirm_big=None):
    """Exports one target to out_path. confirm_big(width_px, height_px) -> bool for huge PNGs.

    LibreOffice writes to an ASCII-named temporary file first; Python then moves it to
    out_path, so destinations with non-ASCII names work regardless of the process locale.
    """
    work = temp_file(ext)
    try:
        _export_to(ctx, doc, target, ext, work, settings, confirm_big)
        if os.path.exists(out_path):
            os.remove(out_path)
        shutil.move(work, out_path)
        if os.name == 'posix':                  # mkstemp files are 0600: use normal permissions
            mask = os.umask(0)
            os.umask(mask)
            os.chmod(out_path, 0o666 & ~mask)
    finally:
        _remove(work)
    return out_path


def _export_to(ctx, doc, target, ext, out_path, settings, confirm_big):
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    with Composition(ctx, doc, target, settings) as comp:
        shapes = comp.shapes()
        if ext in ('emf', 'wmf', 'svg'):
            graphic_export(ctx, shapes, out_path, MEDIA[ext])
            if ext == 'svg' and not fix_svg(out_path):
                raise RuntimeError('SVG export produced an invalid file')
        elif ext == 'png':
            w, h = png_pixels(comp.rect, settings.dpi)
            if w * h > PNG_WARN_PIXELS and confirm_big is not None and not confirm_big(w, h):
                raise Cancelled()
            graphic_export(ctx, shapes, out_path, MEDIA['png'], [pv('PixelWidth', w), pv('PixelHeight', h)])
        elif ext == 'emz':
            t = temp_file('emf')
            try:
                graphic_export(ctx, shapes, t, MEDIA['emf'])
                with open(t, 'rb') as src, gzip.open(out_path, 'wb', compresslevel=9) as dst:
                    shutil.copyfileobj(src, dst)
            finally:
                _remove(t)
        elif ext == 'pdf':
            t = temp_file('svg')
            try:
                graphic_export(ctx, shapes, t, MEDIA['svg'])
                _svg_to_pdf(ctx, desktop, t, comp.rect, out_path)
            finally:
                _remove(t)
        elif ext == 'vsdx':
            if settings['vsdx_native'] and export_vsdx_native(ctx, doc, shapes, comp.rect, out_path):
                return
            t = temp_file('emf')
            try:
                graphic_export(ctx, shapes, t, MEDIA['emf'])
                x0, y0, x1, y1 = comp.rect
                vecstamp_vsdx.build_vsdx(t, out_path, VERSION, doc_base_name(doc),
                                         size_in=((x1 - x0) / 2540.0, (y1 - y0) / 2540.0))
            finally:
                _remove(t)
        else:
            raise ValueError('unknown format: %s' % ext)


def export_vsdx_native(ctx, doc, shapes, rect, out_path):
    """Visio drawing with native, editable shapes. Returns False to fall back to picture mode."""
    try:
        dicts, pw, ph, col = vecstamp_shapes.collect(ctx, shapes, rect,
                                                     vecstamp_shapes.emf_exporter(ctx, graphic_export))
        if not dicts:
            return False
        vecstamp_vsdx.save(out_path, pw, ph, dicts, VERSION, doc_base_name(doc))
        return True
    except Exception:
        return False


def _svg_to_pdf(ctx, desktop, svg_path, rect, out_path):
    """Places the SVG on a Draw page of exactly the drawing's size and saves it as PDF."""
    gp = ctx.ServiceManager.createInstanceWithContext('com.sun.star.graphic.GraphicProvider', ctx)
    graphic = gp.queryGraphic(props(('URL', to_url(svg_path))))
    x0, y0, x1, y1 = rect
    w, h = max(int(x1 - x0), 10), max(int(y1 - y0), 10)
    d = desktop.loadComponentFromURL('private:factory/sdraw', '_blank', 0, props(('Hidden', True)))
    try:
        page = d.getDrawPages().getByIndex(0)
        page.setPropertyValue('BorderLeft', 0)
        page.setPropertyValue('BorderRight', 0)
        page.setPropertyValue('BorderTop', 0)
        page.setPropertyValue('BorderBottom', 0)
        page.setPropertyValue('Width', w)
        page.setPropertyValue('Height', h)
        g = d.createInstance('com.sun.star.drawing.GraphicObjectShape')
        page.add(g)
        g.setPropertyValue('Graphic', graphic)
        g.setPosition(Point(0, 0))
        g.setSize(Size(w, h))
        d.storeToURL(to_url(out_path), props(('FilterName', 'draw_pdf_Export')))
    finally:
        d.close(True)


def _remove(p):
    try:
        os.remove(p)
    except OSError:
        pass


def export_pages(ctx, doc, folder, ext, settings, page_numbers=None, confirm_big=None, progress=None):
    """One file per page. Returns (files, skipped_pages). progress(i, n) is called before each page."""
    pages = doc.getDrawPages()
    base = doc_base_name(doc)
    files, skipped = [], 0
    for i in range(pages.getCount()):
        if page_numbers and (i + 1) not in page_numbers:
            continue
        if progress is not None:
            progress(i + 1, pages.getCount())
        page = pages.getByIndex(i)
        shapes = page_shapes(page)
        if not shapes:
            skipped += 1
            continue
        path = unique_path(folder, '%s_第%d页' % (base, i + 1), ext)
        export_target(ctx, doc, Target(page, shapes, i + 1), ext, path, settings, confirm_big)
        files.append(path)
    return files, skipped


def picture_copy(ctx, doc, target):
    """Adds an EMF picture of the target next to it and selects it."""
    t = temp_file('emf')
    try:
        graphic_export(ctx, target.shapes, t, MEDIA['emf'])
        gp = ctx.ServiceManager.createInstanceWithContext('com.sun.star.graphic.GraphicProvider', ctx)
        graphic = gp.queryGraphic(props(('URL', to_url(t))))
    finally:
        _remove(t)
    x0, y0, x1, y1 = bound_rect(target.shapes)
    g = doc.createInstance('com.sun.star.drawing.GraphicObjectShape')
    target.page.add(g)
    g.setPropertyValue('Graphic', graphic)
    g.setPosition(Point(int(x0 + 500), int(y0 + 500)))
    g.setSize(Size(int(x1 - x0), int(y1 - y0)))
    g.setPropertyValue('Name', 'VecStamp EMF ' + time.strftime('%H%M%S'))
    try:
        doc.getCurrentController().select(g)
    except Exception:
        pass
    return g
