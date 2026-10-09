"""Generate the promotional images used by README.md (assets/promo/*.png).

    pip install cairosvg pillow
    python build/make_promo.py

All artwork is drawn here as SVG and rendered with cairosvg. Fonts: Noto Sans
CJK SC (Regular / Medium / Bold / Black) and DejaVu Sans - on Debian / Ubuntu
`sudo apt install fonts-noto-cjk fonts-dejavu`.

assets/screenshots/*.png are real captures of LibreOffice Impress on Linux with
the extension installed (see docs/DEVELOPMENT.md); this script only frames and
arranges them. Images that show the PowerPoint ribbon are labelled as mock-ups.
"""
import base64
import io
import os
import subprocess
import sys

import cairosvg
from PIL import Image, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import make_icons as mi  # noqa: E402

PROMO = os.path.join(ROOT, 'assets', 'promo')
SHOTS = os.path.join(ROOT, 'assets', 'screenshots')
ICONS = os.path.join(ROOT, 'src', 'icons')

SEAL, SEAL_DARK, INK, PAPER, BEIGE = mi.SEAL, mi.SEAL_DARK, mi.INK, mi.PAPER, mi.BEIGE
FMT = mi.FMT_COLORS
MUTED = '#6B7280'
SOFT = '#EADFD5'
GREEN = '#2E8B57'

F = 'Noto Sans CJK SC'            # regular
FM = 'Noto Sans CJK SC Medium'
FB = 'Noto Sans CJK SC Black'
FBOLD = 'Noto Sans CJK SC:bold'    # rendered as family F + font-weight bold
LATIN = 'DejaVu Sans'
MONO = 'DejaVu Sans Mono'


def version():
    import re
    src = open(os.path.join(ROOT, 'libreoffice', 'extension', 'pythonpath', 'vecstamp_core.py'), encoding='utf-8').read()
    return re.search(r"^VERSION = '([^']+)'", src, re.M).group(1)


# ------------------------------------------------------------------ primitives
def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


_fonts = {}


def _font(family, size):
    key = (family, size)
    if key not in _fonts:
        try:
            spec = subprocess.run(['fc-match', '-f', '%{file}|%{index}', family], capture_output=True,
                                  text=True, check=True).stdout
            path, idx = spec.split('|')
            _fonts[key] = ImageFont.truetype(path, size, index=int(idx or 0))
        except Exception:
            _fonts[key] = None
    return _fonts[key]


def tw(s, size, family=F):
    """Width of a text run in px (measured with the real font when available)."""
    f = _font(family, int(round(size)))
    if f is not None:
        return f.getlength(s)
    return sum(size * (1.0 if ord(c) > 0x2E80 else 0.3 if c == ' ' else 0.6) for c in s)


def T(x, y, s, size=20, fill=INK, family=F, anchor='start', opacity=None, weight=None):
    op = f' opacity="{opacity}"' if opacity is not None else ''
    if family == FBOLD:
        family, weight = F, 'bold'
    wt = f' font-weight="{weight}"' if weight else ''
    return (f'<text x="{x:.1f}" y="{y:.1f}" font-family="{family}" font-size="{size}"{wt} fill="{fill}" '
            f'text-anchor="{anchor}"{op}>{esc(s)}</text>')


def icon(body, x, y, size, rotate=0):
    rot = f' rotate({rotate},128,128)' if rotate else ''
    return f'<g transform="translate({x:.1f},{y:.1f}) scale({size / 256:.5f}){rot}">{body}</g>'


def shadow(x, y, w, h, r, depth=9, alpha=0.028, dy=5):
    out = []
    for i in range(depth, 0, -1):
        g = i * 0.9
        out.append(f'<rect x="{x - g:.1f}" y="{y - g + dy:.1f}" width="{w + 2 * g:.1f}" height="{h + 2 * g:.1f}" '
                   f'rx="{r + g:.1f}" fill="#4A3426" opacity="{alpha}"/>')
    return ''.join(out)


def card(x, y, w, h, r=18, fill='#FFFFFF', stroke=None, sh=True):
    st = f' stroke="{stroke}" stroke-width="1.5"' if stroke else ''
    return (shadow(x, y, w, h, r) if sh else '') + \
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{fill}"{st}/>'


def pill(x, y, label, fill, fg='#FFFFFF', h=34, size=16, family=FBOLD, pad=15, stroke=None):
    w = tw(label, size, family) + 2 * pad
    st = f' stroke="{stroke}" stroke-width="1.6"' if stroke else ''
    return (f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h}" rx="{h / 2}" fill="{fill}"{st}/>'
            + T(x + w / 2, y + h / 2 + size * 0.36, label, size, fg, family, 'middle')), w


def data_uri(path_or_image):
    if isinstance(path_or_image, Image.Image):
        buf = io.BytesIO()
        path_or_image.save(buf, 'PNG')
        raw = buf.getvalue()
    else:
        raw = open(path_or_image, 'rb').read()
    return 'data:image/png;base64,' + base64.b64encode(raw).decode('ascii')


def image(src, x, y, w, h, clip_r=0, cid=None):
    img = (f'<image x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" '
           f'xlink:href="{data_uri(src)}" preserveAspectRatio="none"/>')
    if clip_r and cid:
        return (f'<clipPath id="{cid}"><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{clip_r}"/></clipPath>'
                f'<g clip-path="url(#{cid})">{img}</g>')
    return img


def ribbon_icon(name, x, y, size):
    return image(os.path.join(ICONS, name + '.png'), x, y, size, size)


def caret(cx, cy, s=5, color=MUTED):
    """A small drop-down triangle (the CJK font has no U+25BE)."""
    return f'<path d="M{cx - s},{cy - s * 0.5} L{cx + s},{cy - s * 0.5} L{cx},{cy + s * 0.6} Z" fill="{color}"/>'


def cursor(x, y, s=1.0):
    return (f'<path transform="translate({x},{y}) scale({s})" d="M0,0 L0,26 L7,20 L12,31 L17,29 L12,18 L21,18 Z" '
            f'fill="#FFFFFF" stroke="{INK}" stroke-width="2" stroke-linejoin="round"/>')


def check(x, y, color=GREEN, r=13):
    return (f'<circle cx="{x}" cy="{y}" r="{r}" fill="{color}"/>'
            f'<path d="M{x - r * 0.45:.1f},{y + r * 0.02:.1f} L{x - r * 0.1:.1f},{y + r * 0.38:.1f} '
            f'L{x + r * 0.5:.1f},{y - r * 0.36:.1f}" fill="none" stroke="#FFFFFF" stroke-width="{r * 0.26:.1f}" '
            f'stroke-linecap="round" stroke-linejoin="round"/>')


def cross(x, y, color=SEAL, r=13):
    d = r * 0.38
    return (f'<circle cx="{x}" cy="{y}" r="{r}" fill="{color}"/>'
            f'<path d="M{x - d},{y - d} L{x + d},{y + d} M{x + d},{y - d} L{x - d},{y + d}" stroke="#FFFFFF" '
            f'stroke-width="{r * 0.26:.1f}" stroke-linecap="round"/>')


def seal_stamp(cx, cy, size, rotate=-8):
    """A square cinnabar seal reading 矢印 (top to bottom, like a name seal)."""
    h = size / 2
    return (f'<g transform="translate({cx},{cy}) rotate({rotate})" opacity="0.94">'
            f'<rect x="{-h}" y="{-h}" width="{size}" height="{size}" rx="{size * 0.1}" fill="{SEAL}"/>'
            f'<rect x="{-h + size * 0.07}" y="{-h + size * 0.07}" width="{size * 0.86}" height="{size * 0.86}" '
            f'rx="{size * 0.06}" fill="none" stroke="{PAPER}" stroke-width="{size * 0.035}"/>'
            + T(0, -size * 0.04, '矢', size * 0.38, PAPER, FB, 'middle')
            + T(0, size * 0.36, '印', size * 0.38, PAPER, FB, 'middle') + '</g>')


def mini_diagram(x, y, w, selected=True, title='协同框架'):
    """The sample slide used throughout: four boxes around a centre ellipse (280 x 158 units)."""
    s = w / 280
    b = [f'<g transform="translate({x},{y}) scale({s:.5f})">',
         '<rect x="0" y="0" width="280" height="158" rx="6" fill="#FFFFFF" stroke="#D9D2CB" stroke-width="1.2"/>',
         T(140, 22, title, 11, INK, FM, 'middle')]
    boxes = [(22, 36, FMT['svg'], '降碳'), (198, 36, GREEN, '减污'), (22, 100, FMT['vsdx'], '扩绿'),
             (198, 100, FMT['png'], '增长')]
    for bx, by, col, label in boxes:
        right = bx < 140
        x1 = bx + 60 if right else bx
        x2 = 100 if right else 180
        mid = 92 if right else 188
        b.append(f'<path d="M{x1},{by + 12} H{mid} V79 H{x2 - (3 if right else -3)}" fill="none" '
                 f'stroke="#8A8F98" stroke-width="1.3"/>')
    b.append('<path d="M97,75.5 L103,79 L97,82.5 Z" fill="#8A8F98"/><path d="M183,75.5 L177,79 L183,82.5 Z" '
             'fill="#8A8F98"/>')
    for bx, by, col, label in boxes:
        b.append(f'<rect x="{bx}" y="{by}" width="60" height="24" rx="5" fill="{col}"/>'
                 + T(bx + 30, by + 16, label, 10.5, '#FFFFFF', FM, 'middle'))
    b.append(f'<ellipse cx="140" cy="79" rx="38" ry="20" fill="{SEAL}"/>' + T(140, 83, '协同增效', 10.5, '#FFFFFF',
                                                                              FM, 'middle'))
    b.append(T(140, 146, '政策 · 技术 · 市场 · 数字', 8, MUTED, F, 'middle'))
    if selected:
        b.append('<rect x="16" y="30" width="248" height="102" fill="none" stroke="#3B82F6" stroke-width="1" '
                 'stroke-dasharray="3,2"/>')
        for hx in (16, 140, 264):
            for hy in (30, 81, 132):
                if (hx, hy) != (140, 81):
                    b.append(f'<rect x="{hx - 3}" y="{hy - 3}" width="6" height="6" fill="#FFFFFF" '
                             f'stroke="#3B82F6" stroke-width="1.1"/>')
    b.append('</g>')
    return ''.join(b)


def top_band(x, y, w, r, color, cid, hb=7):
    """A coloured strip along the top edge of a rounded card."""
    return (f'<clipPath id="{cid}"><rect x="{x}" y="{y}" width="{w}" height="{r * 3}" rx="{r}"/></clipPath>'
            f'<rect x="{x}" y="{y}" width="{w}" height="{hb}" fill="{color}" clip-path="url(#{cid})"/>')


def fit(s, size, width, family=F, where=''):
    if tw(s, size, family) > width:
        print(f'warning: text too wide ({tw(s, size, family):.0f} > {width:.0f}px) {where}: {s}')
    return s


def svg_doc(w, h, body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
            f'width="{w}" height="{h}" viewBox="0 0 {w} {h}">{body}</svg>')


# When set, images are drawn on this flat colour instead of the rounded paper card and saved
# as <name>-web.png (used by the website, which frames the images itself).
WEB_BG = None


def render(name, w, h, body, scale=1.5):
    os.makedirs(PROMO, exist_ok=True)
    if WEB_BG:
        name += '-web'
    out = os.path.join(PROMO, name + '.png')
    cairosvg.svg2png(bytestring=svg_doc(w, h, body).encode('utf-8'), write_to=out,
                     output_width=int(w * scale), output_height=int(h * scale))
    Image.open(out).save(out, optimize=True)
    print('wrote', os.path.relpath(out, ROOT))


def background(w, h, r=28, grid=False):
    if WEB_BG:
        return f'<rect width="{w}" height="{h}" fill="{WEB_BG}"/>'
    b = f'<rect width="{w}" height="{h}" rx="{r}" fill="{PAPER}"/>'
    if grid:
        b += ('<defs><pattern id="gmin" width="24" height="24" patternUnits="userSpaceOnUse">'
              f'<path d="M24,0 H0 V24" fill="none" stroke="{SOFT}" stroke-width="1"/></pattern>'
              '<pattern id="gmaj" width="120" height="120" patternUnits="userSpaceOnUse">'
              '<path d="M120,0 H0 V120" fill="none" stroke="#E0D2C6" stroke-width="1.4"/></pattern>'
              '<linearGradient id="fade" x1="0" y1="0" x2="1" y2="0">'
              f'<stop offset="0.38" stop-color="{PAPER}" stop-opacity="1"/>'
              f'<stop offset="0.75" stop-color="{PAPER}" stop-opacity="0"/></linearGradient>'
              f'<clipPath id="bgclip"><rect width="{w}" height="{h}" rx="{r}"/></clipPath></defs>'
              f'<g clip-path="url(#bgclip)"><rect width="{w}" height="{h}" fill="url(#gmin)"/>'
              f'<rect width="{w}" height="{h}" fill="url(#gmaj)"/>'
              f'<rect width="{w}" height="{h}" fill="url(#fade)"/></g>')
    return b


# ------------------------------------------------------------------ images
def hero():
    w, h = 1280, 640
    b = [background(w, h, 32, grid=True)]
    b.append(icon(mi.logo_body(), 72, 70, 118))
    b.append(T(212, 158, '矢印', 84, INK, FB))
    b.append(T(212 + tw('矢印', 84, FB) + 16, 158, 'VecStamp', 50, SEAL, LATIN, weight='bold'))
    b.append(T(76, 262, '选中即印，一键导出矢量图', 42, INK, FBOLD))
    b.append(T(78, 306, 'PowerPoint 加载项 · LibreOffice 扩展', 19, MUTED, F))
    b.append(T(78 + tw('PowerPoint 加载项 · LibreOffice 扩展', 19) + 20, 306, 'Windows · macOS · Linux', 19, SEAL, FM))
    x = 76
    for ext in ('emf', 'emz', 'svg', 'wmf', 'pdf', 'vsdx', 'png'):
        p, pw = pill(x, 340, ext.upper(), FMT[ext], h=38, size=17, family=LATIN)
        b.append(p.replace(f'font-family="{LATIN}"', f'font-family="{LATIN}" font-weight="bold"'))
        x += pw + 10
    feats = ['透明 · 白色 · 米色 · 网格 · 自定义背景（色盘 / RGB / 透明度）',
             'PNG 150 – 3000 DPI 任意分辨率，图形之外不留白',
             'Visio 原生形状 VSDX，文字可编辑 · SVG 支持 PowerPoint 2013+']
    for i, f in enumerate(feats):
        y = 425 + i * 40
        b.append(check(88, y - 7, SEAL, 10) + T(110, y, f, 19, INK, F))
    x = 76
    for label, fill, fg, stroke in (('v' + version(), SEAL, '#FFFFFF', None), ('MIT 开源', PAPER, INK, INK),
                                    ('永久免费 · 开源 · 无广告', PAPER, INK, INK)):
        p, pw = pill(x, 560, label, fill, fg, h=36, size=16, stroke=stroke)
        b.append(p)
        x += pw + 12

    # illustration: slide -> seal -> vector files
    b.append(shadow(744, 92, 400, 226, 9, alpha=0.035))
    b.append(mini_diagram(744, 92, 400, title='物流业降碳减污扩绿增长协同框架'))
    b.append(cursor(948, 226, 1.1))
    b.append(seal_stamp(1156, 312, 84, -9))
    for i, (ext, dx, dy, rot) in enumerate((('emf', 752, 382, -9), ('svg', 880, 360, 0), ('pdf', 1008, 382, 9))):
        b.append(f'<g opacity="0.16"><ellipse cx="{dx + 75}" cy="{dy + 160}" rx="54" ry="9" fill="{INK}"/></g>')
        b.append(icon(mi.doc_body(FMT[ext], ext.upper()), dx, dy, 150, rot))
    render('hero', w, h, ''.join(b))


def workflow():
    w, h = 1280, 430
    b = [background(w, h)]
    cards = [(46, '1', '选中图形', '单个、多个或组合形状均可，也可整页导出'),
             (466, '2', '点击「矢印」', '功能区一键导出，或从下拉菜单选格式'),
             (886, '3', '得到矢量文件', '插入 Word / LaTeX / Visio，放大不失真')]
    for x, num, title, sub in cards:
        b.append(card(x, 36, 348, 356, 20))
        b.append(f'<circle cx="{x + 40}" cy="{36 + 290}" r="20" fill="{SEAL}"/>'
                 + T(x + 40, 36 + 297, num, 21, '#FFFFFF', LATIN, 'middle', weight='bold'))
        b.append(T(x + 72, 36 + 299, title, 25, INK, FBOLD))
        b.append(T(x + 24, 36 + 336, sub, 15.5, MUTED, F))
    for ax in (424, 844):
        b.append(f'<path d="M{ax},214 l20,0 m-9,-10 l10,10 l-10,10" fill="none" stroke="{SEAL}" stroke-width="4" '
                 'stroke-linecap="round" stroke-linejoin="round"/>')
    # 1: selection
    b.append(mini_diagram(70, 70, 300))
    b.append(cursor(250, 170, 1.0))
    # 2: ribbon fragment
    x0, y0 = 486, 66
    b.append(f'<rect x="{x0}" y="{y0}" width="308" height="200" rx="10" fill="#F7F5F3" stroke="#E2DCD6"/>')
    tabs = [('开始', False), ('插入', False), ('设计', False), ('矢印', True)]
    tx = x0 + 16
    for t, active in tabs:
        if active:
            b.append(f'<rect x="{tx - 8}" y="{y0 + 10}" width="{tw(t, 15) + 16}" height="30" rx="5" fill="#FFFFFF"/>'
                     f'<rect x="{tx - 2}" y="{y0 + 36}" width="{tw(t, 15) + 4}" height="3" rx="1.5" fill="{SEAL}"/>')
        b.append(T(tx, y0 + 30, t, 15, SEAL if active else MUTED, FBOLD if active else F))
        tx += tw(t, 15) + 30
    b.append(f'<rect x="{x0 + 10}" y="{y0 + 44}" width="288" height="146" rx="6" fill="#FFFFFF"/>')
    b.append(f'<rect x="{x0 + 22}" y="{y0 + 56}" width="92" height="118" rx="8" fill="#FBE9E7" '
             f'stroke="{SEAL}" stroke-width="1.5"/>')
    b.append(ribbon_icon('export64', x0 + 36, y0 + 66, 64))
    b.append(T(x0 + 68, y0 + 150, '导出 SVG', 14, INK, FM, 'middle') + caret(x0 + 68, y0 + 162, 4.5))
    for i, ext in enumerate(('emf', 'pdf', 'png')):
        yy = y0 + 64 + i * 38
        b.append(ribbon_icon(ext + '32', x0 + 134, yy, 24) + T(x0 + 166, yy + 18, ext.upper(), 15, INK, LATIN))
    for i, ext in enumerate(('emz', 'vsdx', 'svg')):
        yy = y0 + 64 + i * 38
        b.append(ribbon_icon(ext + '32', x0 + 214, yy, 24) + T(x0 + 246, yy + 18, ext.upper(), 15, INK, LATIN))
    b.append(cursor(x0 + 102, y0 + 96, 1.0))
    # 3: files
    for ext, dx, dy, rot in (('emf', 906, 92, -8), ('svg', 1010, 76, 0), ('pdf', 1114, 92, 8)):
        b.append(icon(mi.doc_body(FMT[ext], ext.upper()), dx, dy, 118, rot))
    b.append(T(1060, 246, '∞', 40, SEAL, LATIN, 'middle', weight='bold'))
    b.append(T(1060, 272, '无限放大依然锐利', 15, MUTED, F, 'middle'))
    render('workflow', w, h, ''.join(b))


FORMAT_CARDS = [
    ('emf', '增强型图元文件', ['Word / WPS / PPT 插图首选', '放大不失真，可取消组合编辑'], ['矢量']),
    ('emz', '压缩的 EMF', ['与 EMF 同样清晰，体积更小', '便于邮件、网盘传输'], ['矢量']),
    ('svg', '可缩放矢量图形', ['网页、Inkscape、Illustrator', 'PowerPoint 2013+ 内置转换'], ['矢量', 'v1.2.1 增强']),
    ('wmf', 'Windows 图元文件', ['兼容只认 WMF 的老旧软件', '不支持透明和渐变'], ['矢量']),
    ('pdf', '矢量文档', ['页面尺寸与图形完全一致', '期刊投稿、LaTeX 插图、打印'], ['文档']),
    ('vsdx', 'Visio 绘图', ['Visio 原生形状，文字可编辑', '无需安装 Visio 即可生成'], ['文档', 'v1.2.1 原生']),
    ('png', '高分辨率位图', ['150 / 300 / 600 DPI 一键切换', '自定义分辨率，最高 3000 DPI'], ['位图']),
]


def formats():
    w, h = 1280, 648
    b = [background(w, h)]
    cw, ch, gap = 280, 272, 24
    x0 = (w - 4 * cw - 3 * gap) / 2
    for i in range(8):
        x = x0 + (i % 4) * (cw + gap)
        y = 36 + (i // 4) * (ch + gap)
        if i == 7:
            b.append(card(x, y, cw, ch, 20, INK))
            b.append(icon(mi.pages_body(), x + 22, y + 22, 76))
            b.append(T(x + 116, y + 64, '还有更多', 28, '#FFFFFF', FB))
            b.append(T(x + 117, y + 92, '批量与工具', 16, '#C9CDD3', F))
            for k, line in enumerate(['逐页导出整份演示文稿', '一键生成 EMF 图片副本', '自动命名 · 默认文件夹',
                                      '边距 · 合并或逐个导出']):
                b.append(check(x + 32, y + 145 + k * 32, SEAL, 8) + T(x + 50, y + 151 + k * 32, line, 16, '#FFFFFF', F))
            continue
        ext, sub, lines, tags = FORMAT_CARDS[i]
        col = FMT[ext]
        b.append(card(x, y, cw, ch, 20))
        b.append(top_band(x, y, cw, 20, col, f'band{i}'))
        b.append(icon(mi.doc_body(col, ext.upper()), x + 18, y + 22, 84))
        b.append(T(x + 112, y + 66, ext.upper(), 32, INK, LATIN, weight='bold'))
        b.append(T(x + 113, y + 94, sub, 16, MUTED, F))
        b.append(f'<line x1="{x + 22}" y1="{y + 124}" x2="{x + cw - 22}" y2="{y + 124}" stroke="#EFE8E2" stroke-width="1.5"/>')
        for k, line in enumerate(lines):
            b.append(f'<circle cx="{x + 28}" cy="{y + 155 + k * 32}" r="3.5" fill="{col}"/>'
                     + T(x + 40, y + 161 + k * 32, fit(line, 15.5, cw - 58, where=ext), 15.5, INK, F))
        tx = x + 22
        for t in tags:
            new = t.startswith('v1.2')
            p, pw = pill(tx, y + ch - 50, t, SEAL if new else '#F3EEE9', '#FFFFFF' if new else INK,
                         h=28, size=13.5, pad=12)
            b.append(p)
            tx += pw + 8
    render('formats', w, h, ''.join(b))


def bg_tile(x, y, s, kind, cid):
    b = [f'<clipPath id="{cid}"><rect x="{x}" y="{y}" width="{s}" height="{s}" rx="16"/></clipPath>',
         f'<g clip-path="url(#{cid})">']
    if kind in ('none', 'custom'):
        b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" fill="url(#checker)"/>')
    if kind == 'white':
        b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" fill="#FFFFFF"/>')
    elif kind == 'beige':
        b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" fill="rgb(245,240,225)"/>')
    elif kind == 'custom':
        b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" fill="rgb(196,226,255)" opacity="0.7"/>')
    elif kind == 'grid':
        b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" fill="#FFFFFF"/>')
        step = s / 20
        minor, major = [], []
        for k in range(21):
            p = k * step
            (major if k % 5 == 0 else minor).append(f'M{x + p:.1f},{y} V{y + s} M{x},{y + p:.1f} H{x + s}')
        b.append(f'<path d="{" ".join(minor)}" stroke="rgb(222,226,232)" stroke-width="0.8"/>')
        b.append(f'<path d="{" ".join(major)}" stroke="rgb(176,184,196)" stroke-width="1.2"/>')
    # the sample figure
    fx, fy = x + s * 0.12, y + s * 0.215
    u = s / 224
    b.append(f'<g transform="translate({fx:.1f},{fy:.1f}) scale({u:.4f})">'
             f'<rect x="0" y="10" width="74" height="52" rx="9" fill="{FMT["svg"]}"/>'
             + T(37, 43, '输入', 15, '#FFFFFF', FM, 'middle') +
             f'<path d="M80,36 H104" stroke="{INK}" stroke-width="3"/><path d="M102,29 L114,36 L102,43 Z" fill="{INK}"/>'
             f'<ellipse cx="146" cy="36" rx="32" ry="26" fill="{SEAL}"/>' + T(146, 41, '处理', 15, '#FFFFFF', FM, 'middle') +
             f'<rect x="0" y="86" width="178" height="30" rx="7" fill="{GREEN}"/>'
             + T(89, 106, 'Figure 1  示例图', 13, '#FFFFFF', F, 'middle') + '</g>')
    b.append('</g>')
    b.append(f'<rect x="{x}" y="{y}" width="{s}" height="{s}" rx="16" fill="none" stroke="#E2D8CF" stroke-width="1.5"/>')
    if kind == 'custom':
        b.append(f'<circle cx="{x + s - 30}" cy="{y + 30}" r="20" fill="#FFFFFF"/>'
                 + icon(mi.swatch_body('custom'), x + s - 48, y + 12, 36))
    return ''.join(b)


def backgrounds():
    w, h = 1280, 380
    b = [background(w, h),
         '<defs><pattern id="checker" width="16" height="16" patternUnits="userSpaceOnUse">'
         '<rect width="16" height="16" fill="#FFFFFF"/><rect width="8" height="8" fill="#E6E8EC"/>'
         '<rect x="8" y="8" width="8" height="8" fill="#E6E8EC"/></pattern></defs>']
    items = [('none', '透明', 'PNG 保留 Alpha 通道'), ('white', '白色', 'RGB 255, 255, 255'),
             ('beige', '米色', 'RGB 245, 240, 225'), ('grid', '网格', '细线 + 每 5 格粗线，间距可调'),
             ('custom', '自定义', '色盘 · RGB · 透明度')]
    s, gap = 220, 25
    x0 = (w - 5 * s - 4 * gap) / 2
    for i, (kind, title, sub) in enumerate(items):
        x = x0 + i * (s + gap)
        b.append(shadow(x, 34, s, s, 16, alpha=0.022))
        b.append(bg_tile(x, 34, s, kind, f'tile{i}'))
        b.append(T(x + s / 2, 300, title, 24, INK, FBOLD, 'middle'))
        b.append(T(x + s / 2, 330, sub, 15, MUTED, F, 'middle'))
    render('backgrounds', w, h, ''.join(b))


def vector_vs_bitmap():
    w, h = 1280, 540
    b = [background(w, h)]
    sample = ('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="64" viewBox="0 0 96 64">'
              '<rect width="96" height="64" fill="#FFFFFF"/>'
              + T(4, 46, '矢', 42, INK, FB)
              + mi.arrow_curve((50, 54), (54, 20), (74, 54), (84, 16), 4.5, SEAL, head=0.9)
              + f'<circle cx="64" cy="22" r="6" fill="none" stroke="{FMT["svg"]}" stroke-width="2"/></svg>')
    zw, zh = 576, 384
    small = Image.open(io.BytesIO(cairosvg.svg2png(bytestring=sample.encode(), output_width=96, output_height=64)))
    pixelated = small.convert('RGB').resize((zw, zh), Image.NEAREST)
    sharp = Image.open(io.BytesIO(cairosvg.svg2png(bytestring=sample.encode(), output_width=zw * 2,
                                                   output_height=zh * 2))).convert('RGB')
    for k, (x, img, title, ok) in enumerate(((44, pixelated, '截图 / 粘贴为图片', False),
                                             (660, sharp, '矢印导出  EMF · SVG · PDF', True))):
        b.append((check if ok else cross)(x + 18, 50, r=15) + T(x + 44, 58, title, 24, INK, FBOLD))
        b.append(card(x - 8, 82, zw + 16, zh + 16, 18))
        b.append(image(img, x, 90, zw, zh, 12, f'zoom{k}'))
    b.append(T(w / 2, 516, '示意：同一图形放大 6 倍后的边缘（左：96 DPI 截图；右：矢量重新渲染）', 15.5, MUTED, F, 'middle'))
    render('vector-vs-bitmap', w, h, ''.join(b), scale=1)


def platform_glyph(kind, x, y, s=64):
    u = s / 64
    g = f'<g transform="translate({x},{y}) scale({u:.4f})">'
    if kind == 'win':
        g += (f'<rect x="4" y="6" width="56" height="38" rx="5" fill="#FFFFFF" stroke="{INK}" stroke-width="4"/>'
              f'<path d="M24,44 L22,56 H42 L40,44" fill="none" stroke="{INK}" stroke-width="4" stroke-linejoin="round"/>'
              f'<line x1="16" y1="58" x2="48" y2="58" stroke="{INK}" stroke-width="4" stroke-linecap="round"/>'
              f'<rect x="12" y="14" width="40" height="22" rx="2" fill="{FMT["svg"]}" opacity="0.25"/>')
    elif kind == 'mac':
        g += (f'<rect x="10" y="10" width="44" height="32" rx="4" fill="#FFFFFF" stroke="{INK}" stroke-width="4"/>'
              f'<path d="M2,48 H62 L58,54 H6 Z" fill="{INK}"/>'
              f'<rect x="16" y="16" width="32" height="20" rx="2" fill="{SEAL}" opacity="0.22"/>')
    else:
        g += (f'<rect x="4" y="8" width="56" height="48" rx="6" fill="{INK}"/>'
              '<path d="M14,24 L24,32 L14,40" fill="none" stroke="#7CE38B" stroke-width="4" stroke-linecap="round" '
              'stroke-linejoin="round"/><line x1="28" y1="42" x2="44" y2="42" stroke="#FFFFFF" stroke-width="4" '
              'stroke-linecap="round"/>')
    return g + '</g>'


def platforms():
    w, h = 1280, 372
    b = [background(w, h)]
    items = [('win', 'Windows', 'PowerPoint 2013 及以上', '2016 / 2019 / 2021 / 2024 / Microsoft 365', '安装或卸载-Windows.bat',
              '双击运行，菜单中自选安装或卸载', FMT['svg']),
             ('mac', 'macOS', 'PowerPoint for Mac 16.82 及以上', '辅助脚本处理沙盒内的保存与打包', '安装或卸载-macOS.command',
              '双击运行，菜单中自选安装或卸载', SEAL),
             ('linux', 'Linux', 'LibreOffice Impress / Draw 6.4+', '矢印 LibreOffice 扩展（.oxt）', '安装或卸载-Linux.sh',
              '也可在任意系统的 LibreOffice 中双击 .oxt', GREEN)]
    cw, gap = 384, 24
    x0 = (w - 3 * cw - 2 * gap) / 2
    for i, (kind, name, host, sub, inst, note, col) in enumerate(items):
        x = x0 + i * (cw + gap)
        b.append(card(x, 34, cw, 300, 20))
        b.append(top_band(x, 34, cw, 20, col, f'pband{i}'))
        b.append(platform_glyph(kind, x + 26, 64, 60))
        b.append(T(x + 104, 106, name, 32, INK, LATIN, weight='bold'))
        b.append(T(x + 28, 170, host, 19, INK, FM))
        b.append(T(x + 28, 200, sub, 15.5, MUTED, F))
        b.append(f'<rect x="{x + 24}" y="226" width="{cw - 48}" height="42" rx="10" fill="#F6F1EC"/>')
        b.append(T(x + 42, 253, '▶', 13, col, LATIN) + T(x + 64, 254, inst, 16.5, INK, FM))
        b.append(T(x + 28, 302, note, 14.5, MUTED, F))
    render('platforms', w, h, ''.join(b))


# ---- ribbon mock-up ----
class Ribbon:
    def __init__(self, x, y, h):
        self.x, self.y, self.h = x, y, h
        self.b = []

    def small(self, x, y, img, label):
        self.b.append(ribbon_icon(img, x, y + 2, 20) + T(x + 27, y + 18, label, 14.5, INK, F))
        return 27 + tw(label, 14.5)

    def big(self, x, img, lines, w=74, highlight=False):
        y = self.y + 8
        if highlight:
            self.b.append(f'<rect x="{x}" y="{y}" width="{w}" height="{self.h - 46}" rx="5" fill="#FBE9E7" '
                          f'stroke="{SEAL}" stroke-width="1.2"/>')
        self.b.append(ribbon_icon(img, x + (w - 40) / 2, y + 8, 40))
        for k, line in enumerate(lines):
            if line == 'CARET':
                self.b.append(caret(x + w / 2, y + 68 + k * 18, 4.5))
            else:
                self.b.append(T(x + w / 2, y + 72 + k * 18, line, 14.5, INK, F, 'middle'))
        return w

    def field(self, x, y, label, value, fw, drop=False, swatch=None):
        self.b.append(T(x, y + 18, label, 14.5, INK, F))
        fx = x + tw(label, 14.5) + 8
        self.b.append(f'<rect x="{fx:.1f}" y="{y + 1}" width="{fw}" height="24" rx="3" fill="#FFFFFF" stroke="#C8C6C4"/>')
        vx = fx + 6
        if swatch:
            self.b.append(ribbon_icon(swatch, vx, y + 4, 18))
            vx += 23
        self.b.append(T(vx, y + 18, value, 14, INK, F))
        if drop:
            self.b.append(caret(fx + fw - 12, y + 13, 4.5))
        return fx - x + fw

    def checkbox(self, x, y, label, on):
        self.b.append(f'<rect x="{x}" y="{y + 4}" width="16" height="16" rx="3" fill="{SEAL if on else "#FFFFFF"}" '
                      f'stroke="{SEAL if on else "#8A8886"}" stroke-width="1.3"/>')
        if on:
            self.b.append(f'<path d="M{x + 3.5},{y + 12} L{x + 7},{y + 15.5} L{x + 13},{y + 8}" fill="none" '
                          'stroke="#FFFFFF" stroke-width="2" stroke-linecap="round"/>')
        self.b.append(T(x + 23, y + 18, label, 14.5, INK, F))
        return 23 + tw(label, 14.5)

    def group_end(self, x0, x1, label):
        self.b.append(T((x0 + x1) / 2, self.y + self.h - 12, label, 13.5, MUTED, F, 'middle'))
        self.b.append(f'<line x1="{x1 + 8}" y1="{self.y + 8}" x2="{x1 + 8}" y2="{self.y + self.h - 8}" '
                      'stroke="#E1DFDD" stroke-width="1.2"/>')
        return x1 + 18


OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']
OFFICE_STANDARD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0']
RECENT_DEMO = ['FFF8EB', 'F5F0E1', '2F6FB0', 'C8372D']


def theme_column(hexrgb):
    """Base colour plus the five tint / shade rows of PowerPoint's theme colour gallery."""
    import colorsys
    r, g, b = (int(hexrgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if l >= 0.999:
        steps = [('d', .05), ('d', .15), ('d', .25), ('d', .35), ('d', .5)]
    elif l <= 0.001:
        steps = [('l', .5), ('l', .35), ('l', .25), ('l', .15), ('l', .05)]
    elif l > 0.8:
        steps = [('d', .1), ('d', .25), ('d', .5), ('d', .75), ('d', .9)]
    elif l < 0.2:
        steps = [('l', .9), ('l', .75), ('l', .5), ('l', .25), ('l', .1)]
    else:
        steps = [('l', .8), ('l', .6), ('l', .4), ('d', .25), ('d', .5)]
    out = ['#' + hexrgb]
    for kind, k in steps:
        ll = l + (1 - l) * k if kind == 'l' else l * (1 - k)
        rr, gg, bb = colorsys.hls_to_rgb(h, ll, s)
        out.append('#%02X%02X%02X' % (round(rr * 255), round(gg * 255), round(bb * 255)))
    return out


def color_gallery(x, y):
    """The drop-down of the 背景颜色 gallery (same layout as PowerPoint's colour picker)."""
    sw, gap = 18, 5
    w = 10 * sw + 9 * gap + 28
    hgt = 368
    b = [shadow(x, y, w, hgt, 6, depth=10, alpha=0.035, dy=6),
         f'<rect x="{x}" y="{y}" width="{w}" height="{hgt}" rx="6" fill="#FFFFFF" stroke="#C8C6C4"/>']
    cx = x + 14
    b.append(f'<rect x="{x + 1}" y="{y + 1}" width="{w - 2}" height="26" rx="5" fill="#F3F2F1"/>'
             + T(cx, y + 19, '主题颜色', 13.5, '#323130', FBOLD))
    yy = y + 38
    cols = [theme_column(c) for c in OFFICE_THEME]
    for i, col in enumerate(cols):
        sx = cx + i * (sw + gap)
        b.append(f'<rect x="{sx}" y="{yy}" width="{sw}" height="{sw}" fill="{col[0]}" stroke="#C8C6C4" stroke-width="0.8"/>')
        for k in range(5):
            b.append(f'<rect x="{sx}" y="{yy + sw + 6 + k * sw}" width="{sw}" height="{sw}" fill="{col[k + 1]}" '
                     'stroke="#E1DFDD" stroke-width="0.6"/>')
    yy += sw * 6 + 18
    b.append(f'<rect x="{x + 1}" y="{yy}" width="{w - 2}" height="26" fill="#F3F2F1"/>'
             + T(cx, yy + 18, '标准色', 13.5, '#323130', FBOLD))
    yy += 34
    for i, c in enumerate(OFFICE_STANDARD):
        b.append(f'<rect x="{cx + i * (sw + gap)}" y="{yy}" width="{sw}" height="{sw}" fill="#{c}" stroke="#C8C6C4" '
                 'stroke-width="0.6"/>')
    yy += sw + 10
    b.append(f'<rect x="{x + 1}" y="{yy}" width="{w - 2}" height="26" fill="#F3F2F1"/>'
             + T(cx, yy + 18, '最近使用的颜色', 13.5, '#323130', FBOLD))
    yy += 34
    for i, c in enumerate(RECENT_DEMO):
        sx = cx + i * (sw + gap)
        b.append(f'<rect x="{sx}" y="{yy}" width="{sw}" height="{sw}" fill="#{c}" stroke="#C8C6C4" stroke-width="0.6"/>')
        if i == 0:
            b.append(f'<rect x="{sx - 2.5}" y="{yy - 2.5}" width="{sw + 5}" height="{sw + 5}" fill="none" '
                     f'stroke="{SEAL}" stroke-width="1.6"/>')
    yy += sw + 12
    b.append(f'<line x1="{x + 8}" y1="{yy}" x2="{x + w - 8}" y2="{yy}" stroke="#E1DFDD"/>')
    for k, (img, label) in enumerate((('palette32', '其他颜色(M)…'), ('eyedropper32', '取色器(E)'))):
        iy = yy + 8 + k * 30
        b.append(ribbon_icon(img, cx, iy + 2, 20) + T(cx + 30, iy + 18, label, 14, INK, F))
    return ''.join(b), w


def ribbon():
    rx, ry = 28, 28
    r = Ribbon(rx + 14, ry + 50, 160)
    x = r.x
    y1, y2, y3 = r.y + 10, r.y + 44, r.y + 78
    # group: export
    g0 = x
    x += r.big(x, 'export64', ['导出 EMF', 'CARET'], 78, highlight=True) + 10
    for col in (('emf', 'emz', 'svg'), ('pdf', 'vsdx', 'png')):
        cw = max(r.small(x, yy, ext + '32', ext.upper()) for ext, yy in zip(col, (y1, y2, y3)))
        x += cw + 14
    x = r.group_end(g0, x - 6, '导出选中图形')
    # group: tools
    g0 = x
    x += r.big(x, 'pages64', ['逐页导出'], 70) + 10
    cw = max(r.small(x, y1, 'picture32', 'EMF 图片副本'), r.small(x, y2, 'folder32', '打开导出文件夹'))
    x = r.group_end(g0, x + cw, '批量与工具')
    # group: background (dropDown, RGB box, colour gallery | alpha, grid, margin)
    g0 = x
    gal_x = x
    fw = 120
    c1 = max(r.field(x, y1, '背景', '自定义颜色', fw, drop=True, swatch='bgcustom32'),
             r.field(x, y2, 'RGB', '255,248,235', fw + tw('背景', 14.5) - tw('RGB', 14.5)))
    r.b.append(f'<rect x="{x - 4}" y="{y3 - 1}" width="{tw("背景颜色", 14.5) + 56}" height="28" rx="3" '
               f'fill="#F3F2F1" stroke="#C8C6C4"/>')
    r.b.append(f'<rect x="{x + 1}" y="{y3 + 5}" width="16" height="16" fill="#FFF8EB" stroke="#8A8886"/>'
               + T(x + 25, y3 + 18, '背景颜色', 14.5, INK, F) + caret(x + 37 + tw('背景颜色', 14.5), y3 + 13, 4.5))
    x += c1 + 22
    lw = tw('透明度%', 14.5)
    c2 = max(r.field(x, y1, '透明度%', '0', 46), r.field(x, y2, '网格(磅)', '10', 46 + lw - tw('网格(磅)', 14.5)),
             r.field(x, y3, '边距(磅)', '6', 46 + lw - tw('边距(磅)', 14.5)))
    x = r.group_end(g0, x + c2, '背景')
    # group: settings
    g0 = x
    c1 = max(r.field(x, y1, '默认格式', 'EMF', 150, drop=True, swatch='emf32'),
             r.field(x, y2, '多个图形', '合并为一个文件', 150, drop=True),
             r.field(x, y3, '保存方式', '每次选择保存位置', 150, drop=True))
    x += c1 + 22
    lw = tw('PNG 分辨率', 14.5)
    c2 = max(r.field(x, y1, 'PNG 分辨率', '自定义', 92, drop=True),
             r.field(x, y2, '自定义 DPI', '1200', 92 + lw - tw('自定义 DPI', 14.5)),
             r.checkbox(x, y3, 'VSDX 原生形状', True))
    x += c2 + 22
    c3 = max(r.checkbox(x, y1, '完成后打开文件夹', False), r.small(x, y2, 'setfolder32', '默认文件夹…'),
             r.small(x, y3, 'reset32', '恢复默认设置'))
    x += c3
    r.b.append(f'<path d="M{x - 4},{r.y + r.h - 22} h8 v-8 M{x - 4},{r.y + r.h - 14} l8,-8" fill="none" '
               f'stroke="{MUTED}" stroke-width="1.2"/>')
    x = r.group_end(g0, x + 6, '导出设置')
    # group: about
    g0 = x
    x += r.big(x, 'logo64', ['关于矢印'], 74) + 8
    x += r.big(x, 'donate64', ['赞赏作者', 'CARET'], 74) + 10
    cw = max(r.small(x, y1, 'help32', '使用说明'), r.small(x, y2, 'mail32', '反馈'), r.small(x, y3, 'home32', '项目主页'))
    x += cw
    r.b.append(T((g0 + x) / 2, r.y + r.h - 12, '关于', 13.5, MUTED, F, 'middle'))

    # frame sized to the content
    rw = x + 28 - rx
    w = int(rx * 2 + rw)
    pop_y = y3 + 30
    popup, pop_w = color_gallery(gal_x - 4, pop_y)
    h = int(pop_y + 368 + 30)
    b = [background(w, h), card(rx, ry, rw, 214, 12, '#FFFFFF'),
         f'<path d="M{rx},{ry + 12} a12,12 0 0 1 12,-12 H{rx + rw - 12} a12,12 0 0 1 12,12 V{ry + 46} H{rx} Z" '
         'fill="#F3F2F1"/>']
    tx = rx + 26
    for t in ('文件', '开始', '插入', '绘图', '设计', '切换', '动画', '幻灯片放映', '审阅', '视图', '矢印'):
        active = t == '矢印'
        if active:
            b.append(f'<rect x="{tx - 12}" y="{ry + 6}" width="{tw(t, 16) + 24}" height="40" rx="4" fill="#FFFFFF"/>'
                     f'<rect x="{tx - 4}" y="{ry + 40}" width="{tw(t, 16) + 8}" height="3" rx="1.5" fill="{SEAL}"/>')
        b.append(T(tx, ry + 32, t, 16, SEAL if active else '#484644', FBOLD if active else F))
        tx += tw(t, 16) + 34
    p, pw = pill(rx + rw - 200, ry + 9, '示意图 · 非截图', SEAL, h=28, size=14)
    b.append(p)
    b += r.b
    b.append(popup)
    # notes beside the open gallery
    nx = gal_x + pop_w + 40
    ny = ry + 214 + 54
    notes = [('背景颜色', '与 PowerPoint 相同的取色面板：主题颜色、标准色、最近使用的颜色'),
             ('其他颜色 / 取色器', '打开 PowerPoint 自带的「颜色」对话框，或从幻灯片上直接吸取颜色'),
             ('VSDX 原生形状', '导出的 Visio 绘图由原生形状组成，文字可直接编辑'),
             ('赞赏作者', '矢印永久免费开源；觉得好用可以扫码请作者喝杯咖啡')]
    for k, (head, line) in enumerate(notes):
        yy = ny + k * 66
        b.append(f'<circle cx="{nx + 6}" cy="{yy - 6}" r="5" fill="{SEAL}"/>' + T(nx + 22, yy, head, 17, INK, FBOLD))
        b.append(T(nx + 22, yy + 26, fit(line, 15, w - nx - 60, where='ribbon note'), 15, MUTED, F))
    render('ribbon', w, h, ''.join(b), scale=1.25)


def window(x, y, w, h, title, img, cid, bar=30):
    b = [shadow(x, y, w, h + bar, 10, depth=12, alpha=0.03, dy=8),
         f'<clipPath id="{cid}"><rect x="{x}" y="{y}" width="{w}" height="{h + bar}" rx="10"/></clipPath>',
         f'<g clip-path="url(#{cid})">',
         f'<rect x="{x}" y="{y}" width="{w}" height="{bar}" fill="#E9E5E1"/>',
         image(img, x, y + bar, w, h), '</g>',
         f'<rect x="{x}" y="{y}" width="{w}" height="{h + bar}" rx="10" fill="none" stroke="#CFC7BF" stroke-width="1"/>']
    for k, c in enumerate(('#C9C3BD', '#C9C3BD', '#C9C3BD')):
        b.append(f'<circle cx="{x + w - 22 - k * 20}" cy="{y + bar / 2}" r="5.5" fill="{c}"/>')
    b.append(T(x + 16, y + bar / 2 + 5, title, 13.5, '#4B4F56', F))
    return ''.join(b)


def linux_showcase():
    w, h = 1280, 830
    b = [background(w, h)]
    menu = Image.open(os.path.join(SHOTS, 'libreoffice-menu.png'))
    dlg = Image.open(os.path.join(SHOTS, 'libreoffice-settings.png'))
    picker = Image.open(os.path.join(SHOTS, 'libreoffice-colorpicker.png'))
    mw = 880
    b.append(window(36, 34, mw, mw * menu.height / menu.width, '无标题 1 - LibreOffice Impress', menu, 'win1'))
    pw = 262
    b.append(window(982, 34, pw, pw * picker.height / picker.width, '选色', picker, 'win3'))
    dw = 458
    b.append(window(786, 330, dw, dw * dlg.height / dlg.width, '矢印 VecStamp - 导出设置', dlg, 'win2'))
    p, _ = pill(36, h - 58, 'Linux 实机截图', GREEN, h=32, size=15)
    b.append(p)
    b.append(T(190, h - 36, 'Ubuntu 24.04 · LibreOffice 24.2：「矢印」菜单、导出设置、取色器', 16, MUTED, F))
    render('linux', w, h, ''.join(b), scale=1.25)


def social():
    """1280 x 640 social preview (GitHub: Settings -> Social preview) - same artwork as the hero."""
    src = os.path.join(PROMO, 'hero.png')
    Image.open(src).resize((1280, 640), Image.LANCZOS).save(os.path.join(PROMO, 'social-preview.png'), optimize=True)
    print('wrote assets/promo/social-preview.png')


def main():
    global WEB_BG
    hero()
    workflow()
    formats()
    backgrounds()
    vector_vs_bitmap()
    platforms()
    ribbon()
    if os.path.exists(os.path.join(SHOTS, 'libreoffice-menu.png')):
        linux_showcase()
    social()
    # flat-background copies for the website
    WEB_BG = '#EEF1F5'
    ribbon()
    if os.path.exists(os.path.join(SHOTS, 'libreoffice-menu.png')):
        linux_showcase()
    WEB_BG = None


if __name__ == '__main__':
    main()
