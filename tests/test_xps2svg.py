"""Checks the XPS -> SVG converter used for PowerPoint 2013 - 2021 (src/helpers/xps2svg.ps1).

    python3 tests/test_xps2svg.py            # needs PowerShell 7 (pwsh) or Windows PowerShell
    PWSH=/path/to/pwsh python3 tests/test_xps2svg.py
    VECSTAMP_KEEP=1 python3 tests/test_xps2svg.py   # keep the SVG in tests/xps/out.svg

tests/xps/sample.xps (built by make_xps_fixture.py) contains what PowerPoint's XPS export
uses: canvases with transforms and clips, solid / gradient / image brushes, static resources,
dashed strokes and Glyphs with obfuscated embedded fonts (Latin and CJK subsets).
tests/xps/ref-1.png is a reference rendering of the same page made with xpstopng.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPT = os.path.join(ROOT, 'src', 'helpers', 'xps2svg.ps1')
SAMPLE = os.path.join(HERE, 'xps', 'sample.xps')
SVG = '{http://www.w3.org/2000/svg}'


def find_pwsh():
    for c in (os.environ.get('PWSH'), shutil.which('pwsh'), shutil.which('powershell')):
        if c and os.path.exists(c):
            return c
    return None


def fail(msg):
    print('FAIL:', msg)
    sys.exit(1)


def main():
    pwsh = find_pwsh()
    if not pwsh:
        print('SKIP: PowerShell not found (set PWSH=/path/to/pwsh)')
        return
    tmp = tempfile.mkdtemp(prefix='vs-xps-')
    out = os.path.join(tmp, 'out.svg')
    # The add-in runs the script exactly like this: read as text and invoke as a script block.
    cmd = ("& ([scriptblock]::Create([IO.File]::ReadAllText('%s',[Text.Encoding]::UTF8))) "
           "-Xps '%s' -Out '%s' -Frame '60,40,520,300' -Slide '960,540'") % (SCRIPT, SAMPLE, out)
    r = subprocess.run([pwsh, '-NoProfile', '-NonInteractive', '-Command', cmd], capture_output=True, text=True,
                       timeout=180)
    if r.returncode != 0 or not os.path.exists(out):
        print(r.stdout, r.stderr)
        fail('converter exit code %d' % r.returncode)
    raw = open(out, 'rb').read()
    if raw.startswith(b'\xef\xbb\xbf'):
        fail('SVG must be written without a BOM')
    root = ET.fromstring(raw)
    if root.tag != SVG + 'svg':
        fail('root element is %s' % root.tag)
    tags = {}
    for el in root.iter():
        tags[el.tag.replace(SVG, '')] = tags.get(el.tag.replace(SVG, ''), 0) + 1
    style = ''.join(el.text or '' for el in root.iter(SVG + 'style'))
    checks = {
        'paths': tags.get('path', 0) >= 3,
        'text': tags.get('text', 0) >= 1,
        'gradient': tags.get('linearGradient', 0) + tags.get('radialGradient', 0) >= 1,
        'image': tags.get('image', 0) + tags.get('pattern', 0) >= 1,
        'clip': tags.get('clipPath', 0) >= 1,
        'embedded fonts': style.count('@font-face') >= 2,
        'viewBox': bool(root.get('viewBox')),
    }
    for name, ok in checks.items():
        print(('ok   ' if ok else 'FAIL ') + name)
    print('elements:', dict(sorted(tags.items())))
    print('size: %d bytes, viewBox %s' % (len(raw), root.get('viewBox')))
    if os.environ.get('VECSTAMP_KEEP'):
        shutil.copyfile(out, os.path.join(HERE, 'xps', 'out.svg'))
        print('kept tests/xps/out.svg')
    shutil.rmtree(tmp, ignore_errors=True)
    if not all(checks.values()):
        fail('missing SVG features')
    print('OK')


if __name__ == '__main__':
    main()
