"""Prepare the promotional website in website/ (static files, deploy anywhere).

    pip install pillow cairosvg
    python build/make_website.py

Copies and resizes the images the page uses into website/assets and keeps the
version number in website/config.js and index.html in sync with the add-in.
Run it after make_promo.py so the screenshots are current.
"""
import os
import re
import sys

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SITE = os.path.join(ROOT, 'website')
OUT = os.path.join(SITE, 'assets')
sys.path.insert(0, HERE)


def version():
    src = open(os.path.join(ROOT, 'libreoffice', 'extension', 'pythonpath', 'vecstamp_core.py'), encoding='utf-8').read()
    return re.search(r"^VERSION = '([^']+)'", src, re.M).group(1)


def resized(src, width):
    im = Image.open(os.path.join(ROOT, src))
    if im.width > width:
        im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)
    return im


def save_pair(im, name):
    """PNG fallback plus a smaller WebP for browsers that support it."""
    png = os.path.join(OUT, name + '.png')
    im.convert('RGBA' if im.mode in ('RGBA', 'LA', 'P') else 'RGB').save(png, optimize=True)
    im.save(os.path.join(OUT, name + '.webp'), 'WEBP', quality=88, method=6)
    print(f'wrote website/assets/{name}.png / .webp  {im.width}x{im.height}')


def icons():
    import shutil
    shutil.copyfile(os.path.join(ROOT, 'assets', 'logo.svg'), os.path.join(OUT, 'logo.svg'))
    try:
        import cairosvg
        svg = open(os.path.join(ROOT, 'assets', 'logo.svg'), 'rb').read()
        cairosvg.svg2png(bytestring=svg, write_to=os.path.join(OUT, 'apple-touch-icon.png'), output_width=180, output_height=180)
        big = os.path.join(OUT, '_logo256.png')
        cairosvg.svg2png(bytestring=svg, write_to=big, output_width=256, output_height=256)
        Image.open(big).save(os.path.join(SITE, 'favicon.ico'), sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
        os.remove(big)
    except ImportError:
        Image.open(os.path.join(ROOT, 'assets', 'logo-256.png')).resize((180, 180), Image.LANCZOS).save(
            os.path.join(OUT, 'apple-touch-icon.png'))
        Image.open(os.path.join(ROOT, 'assets', 'logo-256.png')).save(
            os.path.join(SITE, 'favicon.ico'), sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
    print('wrote website/favicon.ico, assets/logo.svg, assets/apple-touch-icon.png')


def qr_codes():
    from make_icons import DONATE_CARDS, DONATE
    for name, (src, box, _color, _title) in DONATE_CARDS.items():
        im = Image.open(os.path.join(DONATE, src)).convert('RGB').crop(box)
        im = im.resize((440, round(im.height * 440 / im.width)), Image.LANCZOS)
        im.save(os.path.join(OUT, f'qr-{name}.png'), optimize=True)
        print(f'wrote website/assets/qr-{name}.png  {im.width}x{im.height}')


def sync_version(v):
    cfg = os.path.join(SITE, 'config.js')
    s = open(cfg, encoding='utf-8').read()
    s = re.sub(r"(version:\s*')[^']*(')", lambda m: m.group(1) + v + m.group(2), s, count=1)
    open(cfg, 'w', encoding='utf-8', newline='\n').write(s)
    page = os.path.join(SITE, 'index.html')
    s = open(page, encoding='utf-8').read()
    s = re.sub(r'(<span data-v>)[^<]*(</span>)', lambda m: m.group(1) + v + m.group(2), s)
    open(page, 'w', encoding='utf-8', newline='\n').write(s)
    print('website version:', v)


def main():
    os.makedirs(OUT, exist_ok=True)
    icons()
    save_pair(resized('assets/promo/ribbon-web.png', 1600), 'ribbon')
    save_pair(resized('assets/promo/linux-web.png', 1400), 'linux')
    resized('assets/promo/social-preview.png', 1200).convert('RGB').save(os.path.join(OUT, 'og.png'), optimize=True)
    print('wrote website/assets/og.png')
    qr_codes()
    sync_version(version())


if __name__ == '__main__':
    main()
