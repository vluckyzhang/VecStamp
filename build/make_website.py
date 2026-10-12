"""Prepare the promotional website in website/ (static files, deploy anywhere).

    pip install pillow cairosvg
    python build/make_website.py            # images, version and the four language versions
    python build/make_website.py --pages    # only the language versions and SEO files (no image work)

- copies and resizes the images the page uses into website/assets and keeps the version number
  in website/config.js and index.html in sync with the add-in (run it after make_promo.py);
- website/index.html is the Simplified Chinese page and the source of the other languages:
  website/zh-TW/, website/en/ and website/ja/ are generated from it with the translations in
  website/i18n.json (every Chinese text node and attribute must have an entry, otherwise the
  build stops and lists what is missing);
- every page gets canonical / hreflang links, Open Graph tags and structured data
  (SoftwareApplication and FAQPage), and website/sitemap.xml and robots.txt are written,
  all based on siteUrl in website/config.js.
"""
import datetime
import html
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SITE = os.path.join(ROOT, 'website')
OUT = os.path.join(SITE, 'assets')
sys.path.insert(0, HERE)

LANGS = ('zh-CN', 'zh-TW', 'en', 'ja')
DIRS = {'zh-CN': '', 'zh-TW': 'zh-TW/', 'en': 'en/', 'ja': 'ja/'}
NAMES = {'zh-CN': '简体中文', 'zh-TW': '繁體中文', 'en': 'English', 'ja': '日本語'}
OG_LOCALE = {'zh-CN': 'zh_CN', 'zh-TW': 'zh_TW', 'en': 'en_US', 'ja': 'ja_JP'}
LANG_LABEL = {'zh-CN': '语言 / Language', 'zh-TW': '語言 / Language', 'en': 'Language', 'ja': '言語 / Language'}
SITE_NAME = {'zh-CN': '矢印 VecStamp', 'zh-TW': '矢印 VecStamp', 'en': 'VecStamp', 'ja': '矢印 VecStamp'}
# CJK fonts in the order that shows the right glyph shapes for each language
FONTS = {
    'zh-CN': '-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "HarmonyOS Sans SC", "MiSans", '
             '"Microsoft YaHei UI", "Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei", sans-serif',
    'zh-TW': '-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang TC", "Microsoft JhengHei UI", "Microsoft JhengHei", '
             '"Noto Sans CJK TC", "Source Han Sans TC", "PingFang SC", "Microsoft YaHei UI", sans-serif',
    'en': '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", '
          '"Microsoft YaHei UI", "Noto Sans CJK SC", sans-serif',
    'ja': '-apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic UI", '
          '"Yu Gothic", "Meiryo", "Noto Sans CJK JP", "Source Han Sans JP", sans-serif',
}
CJK = re.compile(r'[　-〿㐀-鿿＀-￯]')
REPO = 'https://github.com/vluckyzhang/VecStamp'


# ------------------------------------------------------------------ images and version
def version():
    src = open(os.path.join(ROOT, 'libreoffice', 'extension', 'pythonpath', 'vecstamp_core.py'), encoding='utf-8').read()
    return re.search(r"^VERSION = '([^']+)'", src, re.M).group(1)


def resized(src, width):
    from PIL import Image
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
    from PIL import Image
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
    from PIL import Image
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


# ------------------------------------------------------------------ languages
def site_url():
    cfg = open(os.path.join(SITE, 'config.js'), encoding='utf-8').read()
    m = re.search(r"siteUrl:\s*'([^']+)'", cfg)
    url = m.group(1) if m else 'http://vluckyzhang.cn/shiyin/'
    return url if url.endswith('/') else url + '/'


def load_table():
    with open(os.path.join(SITE, 'i18n.json'), encoding='utf-8') as f:
        return json.load(f)


class Translator:
    def __init__(self, table, lang):
        self.table, self.lang, self.missing = table, lang, []

    def __call__(self, text):
        if self.lang == 'zh-CN' or not CJK.search(text):
            return text
        entry = self.table.get(text)
        if entry is None or self.lang not in entry:
            if text not in self.missing:
                self.missing.append(text)
            return text
        return entry[self.lang]


ATTRS = re.compile(r'\b(alt|aria-label|title|content|placeholder|data-badge)="([^"]*)"')


def localize_markup(chunk, tr):
    def text_node(m):
        raw = m.group(1)
        core = raw.strip()
        if not core or not CJK.search(core):
            return m.group(0)
        lead, trail = raw[:len(raw) - len(raw.lstrip())], raw[len(raw.rstrip()):]
        return '>' + lead + html.escape(tr(html.unescape(core)), quote=False) + trail + '<'

    def attr(m):
        value = html.unescape(m.group(2))
        return '%s="%s"' % (m.group(1), html.escape(tr(value), quote=True)) if CJK.search(value) else m.group(0)
    chunk = re.sub(r'>([^<>]+)<', text_node, chunk)
    return ATTRS.sub(attr, chunk)


def localize(page, tr):
    """Translates text nodes and attributes; <style>, <script> and comments are kept, except the
    JSON block with the texts the script uses (id="i18n")."""
    out, pos = [], 0
    for m in re.finditer(r'<style\b.*?</style>|<script\b.*?</script>|<!--.*?-->', page, re.S):
        out.append(localize_markup(page[pos:m.start()], tr))
        block = m.group(0)
        if block.startswith('<script type="application/json" id="i18n">'):
            body = block[len('<script type="application/json" id="i18n">'):-len('</script>')]
            data = {k: tr(v) for k, v in json.loads(body).items()}
            block = '<script type="application/json" id="i18n">' + json.dumps(data, ensure_ascii=False) + '</script>'
        out.append(block)
        pos = m.end()
    out.append(localize_markup(page[pos:], tr))
    return ''.join(out)


def strip_tags(s):
    return re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', '', s))).strip()


def lang_menu(lang):
    up = '../' if DIRS[lang] else ''
    items = []
    for code in LANGS:
        cur = ' aria-current="true"' if code == lang else ''
        href = (up + DIRS[code]) or './'
        items.append(f'          <li><a href="{href}" hreflang="{code}" lang="{code}"{cur}>{NAMES[code]}</a></li>')
    return ('<!-- LANG:BEGIN (written by build/make_website.py) -->\n'
            '      <details class="lang" id="lang">\n'
            f'        <summary aria-label="{LANG_LABEL[lang]}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            'stroke-width="1.7" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c2.6 '
            '2.6 3.9 5.6 3.9 9s-1.3 6.4-3.9 9c-2.6-2.6-3.9-5.6-3.9-9S9.4 5.6 12 3Z"/></svg>'
            f'<span>{NAMES[lang]}</span></summary>\n'
            '        <ul>\n' + '\n'.join(items) + '\n        </ul>\n'
            '      </details>\n'
            '      <!-- LANG:END -->')


def og_image(lang):
    name = 'og.png' if lang == 'zh-CN' else f'og-{lang}.png'
    return name if os.path.exists(os.path.join(OUT, name)) else 'og.png'


def seo_block(page, lang, base, ver):
    url = base + DIRS[lang]
    og = base + 'assets/' + og_image(lang)
    desc = html.unescape(re.search(r'<meta name="description" content="([^"]*)"', page).group(1))
    title = html.unescape(re.search(r'<title>([^<]*)</title>', page).group(1))
    features = [strip_tags(x) for x in re.findall(r'<dt>(.*?)</dt>', page, re.S)]
    faq = []
    for q, a in re.findall(r'<details>\s*<summary>(.*?)</summary>\s*<div>(.*?)</div>\s*</details>', page, re.S):
        faq.append({'@type': 'Question', 'name': strip_tags(q),
                    'acceptedAnswer': {'@type': 'Answer', 'text': strip_tags(a)}})
    app = {
        '@context': 'https://schema.org', '@type': 'SoftwareApplication',
        'name': SITE_NAME[lang], 'alternateName': ['VecStamp', '矢印', 'Shi Yin'],
        'description': desc, 'url': url, 'inLanguage': lang,
        'applicationCategory': 'MultimediaApplication', 'applicationSubCategory': 'PowerPoint add-in / LibreOffice extension',
        'operatingSystem': 'Windows, macOS, Linux', 'softwareVersion': ver,
        'softwareRequirements': 'Microsoft PowerPoint 2013+ (Windows), PowerPoint for Mac 16.82+, LibreOffice 6.4+',
        'downloadUrl': REPO + '/releases/latest', 'image': og, 'screenshot': base + 'assets/ribbon.png',
        'license': 'https://opensource.org/licenses/MIT', 'isAccessibleForFree': True,
        'offers': {'@type': 'Offer', 'price': '0', 'priceCurrency': 'USD'},
        'author': {'@type': 'Person', 'name': 'vluckyzhang', 'email': 'vluckyzhang@gmail.com', 'url': 'https://github.com/vluckyzhang'},
        'sameAs': [REPO], 'featureList': features,
    }
    lines = ['<!-- SEO:BEGIN (canonical, hreflang, Open Graph, structured data: written by build/make_website.py) -->',
             f'<link rel="canonical" href="{url}">']
    for code in LANGS:
        lines.append(f'<link rel="alternate" hreflang="{code}" href="{base}{DIRS[code]}">')
    lines += [f'<link rel="alternate" hreflang="x-default" href="{base}">',
              f'<meta property="og:url" content="{url}">',
              f'<meta property="og:site_name" content="{SITE_NAME[lang]}">',
              f'<meta property="og:image" content="{og}">',
              '<meta property="og:image:width" content="1200">', '<meta property="og:image:height" content="600">',
              f'<meta property="og:locale" content="{OG_LOCALE[lang]}">']
    lines += [f'<meta property="og:locale:alternate" content="{OG_LOCALE[c]}">' for c in LANGS if c != lang]
    lines += ['<meta name="twitter:card" content="summary_large_image">',
              f'<meta name="twitter:title" content="{html.escape(title)}">',
              f'<meta name="twitter:description" content="{html.escape(desc)}">',
              f'<meta name="twitter:image" content="{og}">',
              '<script type="application/ld+json">' + json.dumps(app, ensure_ascii=False) + '</script>']
    if faq:
        lines.append('<script type="application/ld+json">' + json.dumps(
            {'@context': 'https://schema.org', '@type': 'FAQPage', 'inLanguage': lang, 'mainEntity': faq},
            ensure_ascii=False) + '</script>')
    lines.append('<!-- SEO:END -->')
    return '\n'.join(lines)


def build_page(src, lang, table, base, ver):
    tr = Translator(table, lang)
    # the language menu and the SEO block are written per language, not translated
    page, n1 = re.subn(r'<!-- LANG:BEGIN.*?<!-- LANG:END -->', '<!--LANG-->', src, count=1, flags=re.S)
    page, n2 = re.subn(r'<!-- SEO:BEGIN.*?<!-- SEO:END -->', '<!--SEO-->', page, count=1, flags=re.S)
    if not (n1 and n2):
        raise SystemExit('website/index.html: LANG / SEO markers missing')
    if lang != 'zh-CN':
        page = localize(page, tr)
    root_attr = ' data-root="../"' if DIRS[lang] else ''
    page = page.replace('<html lang="zh-CN">', f'<html lang="{lang}"{root_attr}>', 1)
    page = re.sub(r'--sans: [^;]+;', lambda m: '--sans: ' + FONTS[lang] + ';', page, count=1)
    page = page.replace('<!--LANG-->', lang_menu(lang), 1)
    if DIRS[lang]:                                     # pages in a sub-folder: shared files are one level up
        page = re.sub(r'((?:src|href|srcset)=")(assets/|favicon\.ico|config\.js)', r'\1../\2', page)
        page = page.replace('url("assets/', 'url("../assets/')
    page = page.replace('<!--SEO-->', seo_block(page, lang, base, ver), 1)
    return page, tr.missing


def sitemap(base):
    today = datetime.date.today().isoformat()
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:xhtml="http://www.w3.org/1999/xhtml">']
    for lang in LANGS:
        out += ['  <url>', f'    <loc>{base}{DIRS[lang]}</loc>', f'    <lastmod>{today}</lastmod>',
                '    <changefreq>monthly</changefreq>', f'    <priority>{"1.0" if lang == "zh-CN" else "0.9"}</priority>']
        out += [f'    <xhtml:link rel="alternate" hreflang="{c}" href="{base}{DIRS[c]}"/>' for c in LANGS]
        out += [f'    <xhtml:link rel="alternate" hreflang="x-default" href="{base}"/>', '  </url>']
    out.append('</urlset>')
    return '\n'.join(out) + '\n'


def pages():
    base, ver = site_url(), version()
    table = load_table()
    src_path = os.path.join(SITE, 'index.html')
    src = open(src_path, encoding='utf-8').read()
    missing_all = {}
    for lang in LANGS:
        page, missing = build_page(src, lang, table, base, ver)
        if missing:
            missing_all[lang] = missing
            continue
        path = os.path.join(SITE, DIRS[lang], 'index.html')
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(page)
        print(f'wrote website/{DIRS[lang]}index.html')
    if missing_all:
        for lang, items in missing_all.items():
            print(f'{lang}: {len(items)} text(s) missing in website/i18n.json:')
            for t in items:
                print('   ', t)
        raise SystemExit(1)
    with open(os.path.join(SITE, 'sitemap.xml'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(sitemap(base))
    with open(os.path.join(SITE, 'robots.txt'), 'w', encoding='utf-8', newline='\n') as f:
        f.write(f'User-agent: *\nAllow: /\n\nSitemap: {base}sitemap.xml\n')
    print(f'wrote website/sitemap.xml and robots.txt ({base})')


def main():
    if '--pages' not in sys.argv[1:]:
        os.makedirs(OUT, exist_ok=True)
        icons()
        save_pair(resized('assets/promo/ribbon-web.png', 1600), 'ribbon')
        save_pair(resized('assets/promo/linux-web.png', 1400), 'linux')
        resized('assets/promo/social-preview.png', 1200).convert('RGB').save(os.path.join(OUT, 'og.png'), optimize=True)
        for lang in LANGS[1:]:                         # share images in the other languages
            src = f'assets/promo/social-preview.{lang}.png'
            if os.path.exists(os.path.join(ROOT, src)):
                resized(src, 1200).convert('RGB').save(os.path.join(OUT, f'og-{lang}.png'), optimize=True)
        print('wrote website/assets/og.png (+ og-<lang>.png)')
        qr_codes()
        sync_version(version())
    pages()


if __name__ == '__main__':
    main()
