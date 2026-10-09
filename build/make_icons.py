"""Generate the VecStamp logo, README banner and ribbon icons.

All artwork is drawn here as SVG (pure geometry, no fonts except the
format labels / banner text) and rendered to PNG with cairosvg.

    pip install cairosvg
    python build/make_icons.py
"""
import math
import os

import cairosvg

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, 'assets')
ICONS = os.path.join(ROOT, 'src', 'icons')

SEAL = '#C8372D'      # cinnabar seal red
SEAL_DARK = '#9E2A22'
INK = '#2B2F36'
PAPER = '#FFF7EF'
FMT_COLORS = {
    'emf': SEAL,
    'emz': '#2E8B57',
    'svg': '#2F6FB0',
    'wmf': '#7A4FB0',
    'png': '#D98324',
    'pdf': '#B0306A',
    'vsdx': '#138D90',
}
BEIGE = '#F5F0E1'
LO_ICONS = os.path.join(ROOT, 'libreoffice', 'extension', 'icons')


def svg(body, size=256):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
            f'viewBox="0 0 256 256">{body}</svg>')


def arrow_curve(p0, c1, c2, p3, width, color, head=1.0):
    """A cubic Bezier stroke ending in a filled arrowhead (the 'vector' in VecStamp)."""
    dx, dy = p3[0] - c2[0], p3[1] - c2[1]
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    hl, hw = width * 2.3 * head, width * 1.55 * head
    base = (p3[0] - ux * hl * 0.55, p3[1] - uy * hl * 0.55)       # curve stops inside the head
    tip = (p3[0] + ux * hl * 0.45, p3[1] + uy * hl * 0.45)
    back = (p3[0] - ux * hl * 0.55, p3[1] - uy * hl * 0.55)
    left = (back[0] - uy * hw, back[1] + ux * hw)
    right = (back[0] + uy * hw, back[1] - ux * hw)
    return (f'<path d="M{p0[0]},{p0[1]} C{c1[0]},{c1[1]} {c2[0]},{c2[1]} {base[0]:.1f},{base[1]:.1f}" '
            f'fill="none" stroke="{color}" stroke-width="{width}" stroke-linecap="round"/>'
            f'<path d="M{tip[0]:.1f},{tip[1]:.1f} L{left[0]:.1f},{left[1]:.1f} L{right[0]:.1f},{right[1]:.1f} Z" '
            f'fill="{color}" stroke="{color}" stroke-width="{width * 0.35:.1f}" stroke-linejoin="round"/>')


def logo_body(detail=True):
    p0, c1, c2, p3 = (70, 184), (92, 104), (156, 186), (184, 86)
    b = (f'<rect x="14" y="14" width="228" height="228" rx="46" fill="{SEAL}"/>'
         f'<rect x="30" y="30" width="196" height="196" rx="34" fill="none" stroke="{PAPER}" '
         f'stroke-width="7" opacity="0.92"/>')
    if detail:
        # pen-tool handle from the start anchor to its control point
        b += (f'<line x1="{p0[0]}" y1="{p0[1]}" x2="{c1[0]}" y2="{c1[1]}" stroke="{PAPER}" '
              f'stroke-width="5" opacity="0.85"/>'
              f'<circle cx="{c1[0]}" cy="{c1[1]}" r="9" fill="{SEAL}" stroke="{PAPER}" stroke-width="5"/>')
    b += arrow_curve(p0, c1, c2, p3, 15 if detail else 22, PAPER)
    s = 24 if detail else 30
    b += (f'<rect x="{p0[0] - s / 2}" y="{p0[1] - s / 2}" width="{s}" height="{s}" rx="3" '
          f'fill="{PAPER}" stroke="{SEAL_DARK}" stroke-width="4"/>')
    return b


def doc_body(color, label=None, glyph='curve'):
    """A page with a folded corner, a coloured band and a small vector glyph."""
    b = (f'<path d="M58,18 H160 L206,64 V226 a12,12 0 0 1 -12,12 H58 a12,12 0 0 1 -12,-12 V30 '
         f'a12,12 0 0 1 12,-12 Z" fill="#FFFFFF" stroke="{INK}" stroke-width="10" stroke-linejoin="round"/>'
         f'<path d="M160,18 V52 a12,12 0 0 0 12,12 H206" fill="none" stroke="{INK}" stroke-width="10" '
         f'stroke-linejoin="round"/>')
    if label:
        b += (f'<rect x="20" y="138" width="216" height="86" rx="16" fill="{color}"/>'
              f'<text x="128" y="203" text-anchor="middle" font-family="DejaVu Sans, Arial, sans-serif" '
              f'font-weight="bold" font-size="66" fill="#FFFFFF">{label}</text>')
        b += arrow_curve((78, 118), (92, 70), (130, 112), (150, 64), 12, color, head=0.9)
    else:
        b += f'<rect x="46" y="170" width="160" height="56" fill="{color}"/>'
        b += (f'<path d="M46,170 H206 V226 a12,12 0 0 1 -12,12 H58 a12,12 0 0 1 -12,-12 Z" fill="{color}"/>')
        b += arrow_curve((80, 146), (94, 82), (138, 136), (162, 76), 16, color, head=0.9)
    return b


def export_body():
    b = doc_body(SEAL)
    # a small seal "stamp" badge in the corner
    b += (f'<circle cx="196" cy="196" r="50" fill="{SEAL}" stroke="#FFFFFF" stroke-width="10"/>'
          f'<path d="M172,206 L196,176 L220,206" fill="none" stroke="#FFFFFF" stroke-width="14" '
          f'stroke-linecap="round" stroke-linejoin="round"/>'
          f'<line x1="196" y1="178" x2="196" y2="222" stroke="#FFFFFF" stroke-width="14" stroke-linecap="round"/>')
    return b


def pages_body():
    b = ''
    for i, (x, y) in enumerate([(78, 22), (56, 44), (34, 66)]):
        fill = '#FFFFFF' if i == 2 else '#F2E9E4'
        b += (f'<rect x="{x}" y="{y}" width="150" height="168" rx="14" fill="{fill}" stroke="{INK}" '
              f'stroke-width="10"/>')
    b += arrow_curve((62, 196), (78, 122), (120, 182), (146, 112), 14, SEAL, head=0.9)
    b += (f'<circle cx="200" cy="200" r="44" fill="{SEAL}" stroke="#FFFFFF" stroke-width="9"/>'
          f'<text x="200" y="219" text-anchor="middle" font-family="DejaVu Sans, Arial, sans-serif" '
          f'font-weight="bold" font-size="52" fill="#FFFFFF">N</text>')
    return b


def picture_body():
    return (f'<rect x="22" y="42" width="212" height="172" rx="18" fill="#FFFFFF" stroke="{INK}" stroke-width="12"/>'
            + arrow_curve((60, 178), (80, 96), (140, 168), (176, 92), 16, SEAL, head=0.9)
            + f'<rect x="48" y="166" width="24" height="24" rx="3" fill="#FFFFFF" stroke="{SEAL}" stroke-width="6"/>')


def folder_body(badge=None):
    b = (f'<path d="M22,62 a14,14 0 0 1 14,-14 H98 l20,22 H220 a14,14 0 0 1 14,14 V200 a14,14 0 0 1 -14,14 '
         f'H36 a14,14 0 0 1 -14,-14 Z" fill="#F6C76B" stroke="{INK}" stroke-width="11" stroke-linejoin="round"/>'
         f'<path d="M22,104 H234" stroke="{INK}" stroke-width="9"/>')
    if badge == 'pin':
        b += (f'<circle cx="190" cy="186" r="44" fill="{SEAL}" stroke="#FFFFFF" stroke-width="9"/>'
              f'<path d="M190,160 v52 M164,186 h52" stroke="#FFFFFF" stroke-width="13" stroke-linecap="round"/>')
    elif badge == 'open':
        b += (f'<path d="M150,160 h56 m-22,-22 l22,22 l-22,22" fill="none" stroke="{SEAL}" stroke-width="14" '
              f'stroke-linecap="round" stroke-linejoin="round"/>')
    return b


def help_body():
    return (f'<circle cx="128" cy="128" r="104" fill="#FFFFFF" stroke="{INK}" stroke-width="13"/>'
            f'<path d="M94,100 a34,34 0 1 1 50,30 c-12,7 -16,14 -16,28" fill="none" stroke="{SEAL}" '
            f'stroke-width="20" stroke-linecap="round"/>'
            f'<circle cx="128" cy="190" r="12" fill="{SEAL}"/>')


def mail_body():
    return (f'<rect x="20" y="52" width="216" height="152" rx="18" fill="#FFFFFF" stroke="{INK}" stroke-width="12"/>'
            f'<path d="M30,66 L128,140 L226,66" fill="none" stroke="{SEAL}" stroke-width="14" '
            f'stroke-linecap="round" stroke-linejoin="round"/>')


def reset_body():
    return (f'<path d="M200,128 a72,72 0 1 1 -24,-54" fill="none" stroke="{INK}" stroke-width="20" '
            f'stroke-linecap="round"/>'
            f'<path d="M206,40 L192,96 L138,78 Z" fill="{SEAL}" stroke="{SEAL}" stroke-width="8" stroke-linejoin="round"/>')


def settings_body():
    teeth = ''.join(
        f'<rect x="116" y="18" width="24" height="44" rx="6" fill="{INK}" '
        f'transform="rotate({a} 128 128)"/>' for a in range(0, 360, 45))
    return (teeth + f'<circle cx="128" cy="128" r="76" fill="{INK}"/>'
            f'<circle cx="128" cy="128" r="34" fill="#FFFFFF"/>')


def swatch_body(kind):
    """Background swatches for the ribbon drop-down."""
    frame = f'<rect x="20" y="20" width="216" height="216" rx="30" fill="none" stroke="{INK}" stroke-width="14"/>'
    clip = '<clipPath id="c"><rect x="20" y="20" width="216" height="216" rx="30"/></clipPath>'
    if kind == 'none':
        cells = ''.join(f'<rect x="{20 + i * 54}" y="{20 + j * 54}" width="54" height="54" fill="#C9CED6"/>'
                        for i in range(4) for j in range(4) if (i + j) % 2 == 0)
        inner = f'<rect x="20" y="20" width="216" height="216" fill="#FFFFFF"/>{cells}'
    elif kind == 'white':
        inner = '<rect x="20" y="20" width="216" height="216" fill="#FFFFFF"/>'
    elif kind == 'beige':
        inner = f'<rect x="20" y="20" width="216" height="216" fill="{BEIGE}"/>'
    elif kind == 'grid':
        lines = ''.join(f'<line x1="{20 + k * 36}" y1="20" x2="{20 + k * 36}" y2="236" stroke="#9AA5B4" stroke-width="7"/>'
                        f'<line x1="20" y1="{20 + k * 36}" x2="236" y2="{20 + k * 36}" stroke="#9AA5B4" stroke-width="7"/>'
                        for k in range(1, 6))
        inner = f'<rect x="20" y="20" width="216" height="216" fill="#FFFFFF"/>{lines}'
    else:  # custom: colour wheel
        cols = ['#E53935', '#FB8C00', '#FDD835', '#43A047', '#1E88E5', '#8E24AA']
        segs = ''
        for k, c in enumerate(cols):
            a0, a1 = math.radians(k * 60 - 90), math.radians((k + 1) * 60 - 90)
            x0, y0 = 128 + 150 * math.cos(a0), 128 + 150 * math.sin(a0)
            x1, y1 = 128 + 150 * math.cos(a1), 128 + 150 * math.sin(a1)
            segs += f'<path d="M128,128 L{x0:.1f},{y0:.1f} A150,150 0 0 1 {x1:.1f},{y1:.1f} Z" fill="{c}"/>'
        inner = segs + '<circle cx="128" cy="128" r="40" fill="#FFFFFF"/>'
    return f'<defs>{clip}</defs><g clip-path="url(#c)">{inner}</g>{frame}'


def palette_body():
    dots = [(88, 84, '#E53935'), (142, 64, '#FDD835'), (190, 96, '#43A047'), (196, 152, '#1E88E5')]
    b = (f'<path d="M128,18 C58,18 14,66 14,128 C14,190 62,238 120,238 C150,238 152,214 140,198 '
         f'C128,182 136,160 160,160 L190,160 C222,160 242,138 242,108 C242,56 192,18 128,18 Z" '
         f'fill="#FFF3DD" stroke="{INK}" stroke-width="12" stroke-linejoin="round"/>')
    b += ''.join(f'<circle cx="{x}" cy="{y}" r="22" fill="{c}"/>' for x, y, c in dots)
    b += f'<circle cx="74" cy="168" r="24" fill="#FFFFFF" stroke="{INK}" stroke-width="10"/>'
    return b


def eyedropper_body():
    """A pipette picking up colour (generic glyph)."""
    return (f'<path d="M168,40 a30,30 0 0 1 42,42 l-26,26 l10,10 l-16,16 l-60,-60 l16,-16 l10,10 Z" '
            f'fill="{INK}" stroke="{INK}" stroke-width="6" stroke-linejoin="round"/>'
            f'<path d="M116,104 L46,174 C38,182 36,196 40,206 L28,226 L36,230 L50,216 C60,220 74,218 82,210 L152,140" '
            f'fill="#FFFFFF" stroke="{INK}" stroke-width="12" stroke-linejoin="round"/>'
            f'<path d="M52,184 L94,184 L68,210 C62,214 54,214 50,208 C46,202 46,192 52,184 Z" fill="{SEAL}"/>')


def donate_body():
    """A coffee cup with a heart - 'buy the author a coffee'."""
    return (f'<path d="M40,104 H184 V170 a54,54 0 0 1 -54,54 H94 a54,54 0 0 1 -54,-54 Z" fill="#FFFFFF" '
            f'stroke="{INK}" stroke-width="12" stroke-linejoin="round"/>'
            f'<path d="M184,124 h16 a26,26 0 0 1 0,52 h-18" fill="none" stroke="{INK}" stroke-width="12"/>'
            f'<path d="M112,196 C70,170 64,148 78,136 C90,126 106,132 112,144 C118,132 134,126 146,136 '
            f'C160,148 154,170 112,196 Z" fill="{SEAL}"/>'
            f'<path d="M86,30 c-14,18 14,30 0,52 M118,24 c-14,20 14,32 0,56 M150,30 c-14,18 14,30 0,52" '
            f'fill="none" stroke="#C9CED6" stroke-width="10" stroke-linecap="round"/>')


def home_body():
    return (f'<path d="M40,124 L128,42 L216,124" fill="none" stroke="{INK}" stroke-width="16" '
            f'stroke-linecap="round" stroke-linejoin="round"/>'
            f'<path d="M66,112 V214 H190 V112" fill="#FFFFFF" stroke="{INK}" stroke-width="12" stroke-linejoin="round"/>'
            f'<rect x="108" y="150" width="40" height="64" rx="4" fill="{SEAL}"/>')


DONATE = os.path.join(ASSETS, 'donate')
DONATE_CARDS = {
    # name: (source file, crop box of QR code + name in the original image, header colour, header text)
    'alipay': ('alipay-original.jpg', (196, 656, 1062, 1640), '#1677FF', '支付宝扫码赞赏'),
    'wechat': ('wechat-original.png', (331, 372, 939, 1082), '#07C160', '微信扫码赞赏'),
}


def donate_card(name, w, h):
    """QR card (PNG bytes): coloured header + the QR code cropped from the user's image."""
    import base64
    import io
    from PIL import Image
    src, box, color, title = DONATE_CARDS[name]
    im = Image.open(os.path.join(DONATE, src)).convert('RGB').crop(box)
    head = round(h * 0.16)
    avail_w, avail_h = w - 8, h - head - 8
    k = min(avail_w / im.width, avail_h / im.height)
    qw, qh = round(im.width * k), round(im.height * k)
    im = im.resize((qw * 2, qh * 2), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, 'PNG')
    data = base64.b64encode(buf.getvalue()).decode('ascii')
    body = (f'<rect width="{w}" height="{h}" rx="{h * 0.04:.1f}" fill="#FFFFFF" stroke="#D9D9D9" stroke-width="1.5"/>'
            f'<path d="M0,{h * 0.04:.1f} a{h * 0.04:.1f},{h * 0.04:.1f} 0 0 1 {h * 0.04:.1f},-{h * 0.04:.1f} '
            f'H{w - h * 0.04:.1f} a{h * 0.04:.1f},{h * 0.04:.1f} 0 0 1 {h * 0.04:.1f},{h * 0.04:.1f} V{head} H0 Z" fill="{color}"/>'
            f'<text x="{w / 2}" y="{head * 0.68:.1f}" text-anchor="middle" font-family="Noto Sans CJK SC" '
            f'font-weight="bold" font-size="{head * 0.5:.1f}" fill="#FFFFFF">{title}</text>'
            f'<image x="{(w - qw) / 2:.1f}" y="{head + 4 + (avail_h - qh) / 2:.1f}" width="{qw}" height="{qh}" '
            f'xlink:href="data:image/png;base64,{data}"/>')
    svg_text = (f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
                f'width="{w}" height="{h}" viewBox="0 0 {w} {h}">{body}</svg>')
    return cairosvg.svg2png(bytestring=svg_text.encode('utf-8'), output_width=w, output_height=h)


def write_donate():
    """Ribbon gallery cards (240 x 302), LibreOffice About dialog cards and README / website images."""
    from PIL import Image
    for name in DONATE_CARDS:
        open(os.path.join(ICONS, 'qr' + name + '.png'), 'wb').write(donate_card(name, 240, 302))
        open(os.path.join(LO_ICONS, 'donate_' + name + '.png'), 'wb').write(donate_card(name, 180, 227))
        src = Image.open(os.path.join(DONATE, DONATE_CARDS[name][0])).convert('RGB')
        k = 420 / src.width
        src.resize((420, round(src.height * k)), Image.LANCZOS).save(os.path.join(DONATE, name + '.png'), optimize=True)


def write(name, body, sizes):
    os.makedirs(ICONS, exist_ok=True)
    for s in sizes:
        out = os.path.join(ICONS, f'{name}{s}.png')
        cairosvg.svg2png(bytestring=svg(body, s).encode(), write_to=out, output_width=s, output_height=s)


def write_lo_icons():
    """16 / 26 px toolbar images for the LibreOffice extension."""
    os.makedirs(LO_ICONS, exist_ok=True)
    items = {'logo': logo_body(detail=False), 'export': export_body(), 'pages': pages_body(),
             'picture': picture_body(), 'settings': settings_body(), 'help': help_body(),
             'folder': folder_body('open'), 'mail': mail_body(), 'donate': donate_body(), 'home': home_body()}
    for fmt, color in FMT_COLORS.items():
        items[fmt] = doc_body(color)
    for name, body in items.items():
        for s in (16, 26):
            cairosvg.svg2png(bytestring=svg(body, s).encode(), write_to=os.path.join(LO_ICONS, f'{name}_{s}.png'),
                             output_width=s, output_height=s)
    cairosvg.svg2png(bytestring=svg(logo_body(), 64).encode(), write_to=os.path.join(LO_ICONS, 'logo_64.png'),
                     output_width=64, output_height=64)


def banner():
    w, h = 1280, 400
    logo = logo_body()
    body = (f'<rect width="{w}" height="{h}" rx="36" fill="{PAPER}"/>'
            f'<g transform="translate(92,72) scale(1.0)">{logo}</g>'
            f'<text x="400" y="200" font-family="Noto Sans CJK SC, sans-serif" font-weight="900" font-size="128" '
            f'fill="{INK}">矢印</text>'
            f'<text x="680" y="200" font-family="DejaVu Sans, Arial, sans-serif" font-weight="bold" font-size="92" '
            f'fill="{SEAL}">VecStamp</text>'
            f'<text x="404" y="282" font-family="Noto Sans CJK SC, sans-serif" font-size="40" fill="{INK}" '
            f'opacity="0.78">选中即印 · 幻灯片图形一键导出矢量图</text>')
    s = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
         f'{body}</svg>')
    with open(os.path.join(ASSETS, 'banner.svg'), 'w', encoding='utf-8') as f:
        f.write(s)
    cairosvg.svg2png(bytestring=s.encode(), write_to=os.path.join(ASSETS, 'banner.png'))


def main():
    os.makedirs(ASSETS, exist_ok=True)
    with open(os.path.join(ASSETS, 'logo.svg'), 'w', encoding='utf-8') as f:
        f.write(svg(logo_body()))
    for s in (16, 32, 64, 128, 256, 512):
        body = logo_body(detail=s >= 64)
        cairosvg.svg2png(bytestring=svg(body, s).encode(),
                         write_to=os.path.join(ASSETS, f'logo-{s}.png'), output_width=s, output_height=s)
    write('logo', logo_body(), [64])
    write('logo', logo_body(detail=False), [32])
    write('export', export_body(), [64])
    for fmt, color in FMT_COLORS.items():
        write(fmt, doc_body(color, fmt.upper()), [64])
        write(fmt, doc_body(color), [32])
    write('pages', pages_body(), [64])
    write('picture', picture_body(), [32, 64])
    write('folder', folder_body('open'), [32])
    write('setfolder', folder_body('pin'), [32])
    write('help', help_body(), [32])
    write('mail', mail_body(), [32])
    write('reset', reset_body(), [32])
    write('settings', settings_body(), [32])
    for kind in ('none', 'white', 'beige', 'grid', 'custom'):
        write('bg' + kind, swatch_body(kind), [32])
    write('palette', palette_body(), [32])
    write('eyedropper', eyedropper_body(), [32])
    write('donate', donate_body(), [64])
    write('home', home_body(), [32])
    write_lo_icons()
    write_donate()
    banner()
    print('icons written to', ICONS)


if __name__ == '__main__':
    main()
