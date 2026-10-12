"""Checks the interface languages (zh-CN, zh-TW, en, ja) without PowerPoint or LibreOffice.

- both string tables are complete (all four languages) and use the same {0}, {1}, ... in every language;
- every key the VBA module, the ribbon and the LibreOffice extension ask for exists;
- the LoadStrings code that build/build_bas.py writes into dist/VecStamp.bas yields exactly the texts of
  src/i18n/strings.json (the generated VBA is interpreted here, U("hex") decoded as at run time).

    python3 tests/test_i18n.py          # run python3 build/build.py (or build_bas.py) first
"""
import json
import os
import re
import sys
import xml.dom.minidom

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, 'build'))
import i18n_tool  # noqa: E402

LANGS = ('zh-CN', 'zh-TW', 'en', 'ja')
# ribbon controls whose supertip is built in VBA (VS_GetSupertip) instead of "rb.<id>.sup"
DYNAMIC_SUPERTIP = {'vsBtnExport', 'vsBtnFolder', 'vsBtnOpenFolder', 'vsGalColor', 'vsEbRgb', 'vsBgCustomMac',
                    'vsMnuBg', 'vsMnuLang'}


def load(path):
    with open(os.path.join(ROOT, path), encoding='utf-8') as f:
        return f.read()


def vba_keys(bas):
    keys = set(re.findall(r'Tx\("([\w.:]+)"', bas))
    for a, b in re.findall(r'Tx\(IIf\(.*?, "([\w.]+)", "([\w.]+)"\)', bas):
        keys |= {a, b}
    keys = {k for k in keys if not k.endswith('.')}
    keys |= {'fmt.' + e for e in ('emf', 'emz', 'svg', 'wmf', 'png', 'pdf', 'vsdx')}
    keys |= {'xps.open', 'xps.write', 'xps.done'}
    return keys


def ribbon_problems(ui, table):
    bad = []
    doc = xml.dom.minidom.parseString(ui.encode('utf-8'))
    for el in doc.getElementsByTagName('*'):
        cid = el.getAttribute('id')
        if not cid:
            continue
        alias = 'vsMnuBg' if cid.startswith('vsBgSw') else cid
        fmt_item = cid.startswith('vsMnu') and cid[5:].isupper()
        if el.getAttribute('getLabel') and cid != 'vsMnuLang' and not fmt_item and 'rb.' + cid not in table:
            bad.append('label of ' + cid)
        if el.getAttribute('getTitle') and 'rb.' + cid not in table:
            bad.append('title of ' + cid)
        if el.getAttribute('getDescription') and 'rb.' + cid + '.desc' not in table:
            bad.append('description of ' + cid)
        if el.getAttribute('getScreentip') and not ('rb.' + alias + '.tip' in table or 'rb.' + alias in table
                                                   or el.getAttribute('getLabel')):
            bad.append('screentip of ' + cid)
        if el.getAttribute('getSupertip') and alias not in DYNAMIC_SUPERTIP and 'rb.' + alias + '.sup' not in table:
            bad.append('supertip of ' + cid)
        if el.getAttribute('label') and re.search(r'[㐀-鿿]', el.getAttribute('label')) \
                and not cid.startswith('vsLang'):
            bad.append('untranslated static label of ' + cid)
    return bad


def decode_u(hexs):
    return ''.join(chr(int(hexs[i:i + 4], 16)) for i in range(0, len(hexs), 4))


def vba_literal(expr):
    expr = expr.strip()
    if expr == 'vbCrLf':
        return '\r\n'
    m = re.fullmatch(r'U\("([0-9A-F]*)"\)', expr)
    if m:
        return decode_u(m.group(1))
    m = re.fullmatch(r'"((?:[^"]|"")*)"', expr)
    if m:
        return m.group(1).replace('""', '"')
    raise ValueError('unexpected VBA expression: ' + expr)


def interpret_strings(dist):
    """Runs the generated LoadStrings subs: {lang: {key: text}}."""
    subs = dict(re.findall(r'^Private Sub (LoadStrings_\w+)\(\)\r?\n(.*?)^End Sub', dist, re.M | re.S))
    disp = re.search(r'^Private Sub LoadStrings\(ByVal lang As String\)(.*?)^End Sub', dist, re.M | re.S).group(1)
    order, cur = {}, None
    for line in disp.splitlines():
        line = line.strip()
        m = re.fullmatch(r'Case "([\w-]+)"', line)
        if m:
            cur = m.group(1)
        elif line == 'Case Else':
            cur = 'zh-CN'
        elif line.startswith('LoadStrings_'):
            order.setdefault(cur, []).append(line)
    out = {}
    for lang, names in order.items():
        got = {}
        for name in names:
            s = ''
            for line in subs[name].splitlines():
                line = line.strip()
                if not line or line.startswith('Dim '):
                    continue
                if line == 's = ""':
                    s = ''
                elif line.startswith('s = s & '):
                    s += vba_literal(line[8:])
                elif line.startswith('gStr.Add '):
                    value, key = line[9:].rsplit(', ', 1)
                    key = vba_literal(key)
                    if key in got:
                        raise AssertionError('duplicate key ' + key)
                    got[key] = s if value == 's' else vba_literal(value)
                else:
                    raise AssertionError('unexpected line in %s: %s' % (name, line))
        out[lang] = got
    return out


def main():
    failures = 0
    if i18n_tool.check():
        failures += 1
    vba = json.loads(load('src/i18n/strings.json'))
    lo = json.loads(load('libreoffice/extension/pythonpath/vecstamp_strings.json'))
    bas = load('src/VecStamp.bas')
    missing = sorted(vba_keys(bas) - set(vba))
    if missing:
        print('VBA keys missing in src/i18n/strings.json:', missing)
        failures += 1
    bad = ribbon_problems(load('src/customUI14.xml'), vba)
    if bad:
        print('ribbon texts missing:', bad)
        failures += 1
    used = set()
    pydir = os.path.join(ROOT, 'libreoffice', 'extension', 'pythonpath')
    for name in os.listdir(pydir):
        if name.endswith('.py'):
            src = open(os.path.join(pydir, name), encoding='utf-8').read()
            used |= set(re.findall(r"T\(\s*'([\w.:]+)'", src)) | set(re.findall(r"if \w+ else '([\w.:]+)'", src))
            used |= set(re.findall(r"'(bg\.\w+)'", src))
    used = {k for k in used if not k.endswith('.')}
    missing = sorted(used - set(lo))
    if missing:
        print('LibreOffice keys missing in vecstamp_strings.json:', missing)
        failures += 1
    # non-ASCII literals left in the VBA module must be names that are never translated
    allowed = {'简体中文', '繁體中文', '日本語', '…', '……'}
    left = [lit for lit in re.findall(r'"([^"\n]*)"', bas) if re.search(r'[^\x00-\x7f]', lit) and lit not in allowed]
    if left:
        print('untranslated VBA literals:', left[:10])
        failures += 1
    dist_path = os.path.join(ROOT, 'dist', 'VecStamp.bas')
    if os.path.exists(dist_path):
        got = interpret_strings(open(dist_path, encoding='ascii').read())
        for lang in LANGS:
            want = {k: v[lang].replace('\n', '\r\n') for k, v in vba.items()}
            if got.get(lang) != want:
                diff = [k for k in want if got.get(lang, {}).get(k) != want[k]]
                print(f'generated LoadStrings differ for {lang}: {diff[:5]}')
                failures += 1
        print(f'generated VBA strings: {", ".join(f"{l}={len(got.get(l, {}))}" for l in LANGS)}')
    else:
        print('SKIP generated VBA check: dist/VecStamp.bas not built')
    print(f'PowerPoint strings: {len(vba)}, LibreOffice strings: {len(lo)}, languages: {", ".join(LANGS)}')
    print('OK' if not failures else f'{failures} problem(s)')
    sys.exit(1 if failures else 0)


if __name__ == '__main__':
    main()
