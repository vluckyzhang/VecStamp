"""End-to-end test of the LibreOffice extension (Linux / any LibreOffice with Python).

    python3 build/build.py               # builds dist/VecStamp-LibreOffice.oxt
    python3 tests/test_libreoffice.py    # run as a normal user, not root

What it does
1. creates a throw-away LibreOffice profile in a temp folder (your own profile is untouched)
2. installs dist/VecStamp-LibreOffice.oxt into it with unopkg
3. starts LibreOffice headless, draws a few shapes in a new Impress document and selects them
4. calls the extension's non-interactive "selftest" for every background type, which exports
   the selection in all 7 formats, then checks every file (magic bytes, SVG root, VSDX parts)
5. shuts LibreOffice down and deletes the temp profile (set VECSTAMP_KEEP=1 to keep the files)

Needs: LibreOffice 6.4+, its Python-UNO bridge (Debian/Ubuntu: python3-uno) and a python3
that can "import uno" (the system python3 on most distributions).
"""
import gzip
import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OXT = os.path.join(ROOT, 'dist', 'VecStamp-LibreOffice.oxt')
BACKGROUNDS = {'transparent': 0, 'white': 1, 'beige': 2, 'grid': 3, 'custom': 4}


def find_program(name):
    for cand in (shutil.which(name), f'/usr/lib/libreoffice/program/{name}',
                 f'/usr/lib64/libreoffice/program/{name}', f'/opt/libreoffice/program/{name}',
                 f'/Applications/LibreOffice.app/Contents/MacOS/{name}'):
        if cand and os.path.exists(cand):
            return cand
    sys.exit(f'{name} not found - install LibreOffice first')


def free_port():
    with socket.socket() as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


# ------------------------------------------------------------------ file checks
def check_emf(data):
    assert data[40:44] == b' EMF', 'no EMF signature'


def check_file(ext, path, native_vsdx=True):
    data = open(path, 'rb').read()
    assert len(data) > 100, f'{ext}: file too small ({len(data)} bytes)'
    if ext == 'emf':
        check_emf(data)
    elif ext == 'emz':
        check_emf(gzip.decompress(data))
    elif ext == 'wmf':
        assert data[:4] in (b'\xd7\xcd\xc6\x9a', b'\x01\x00\x09\x00', b'\x02\x00\x09\x00'), 'no WMF header'
    elif ext == 'png':
        assert data[:8] == b'\x89PNG\r\n\x1a\n', 'no PNG signature'
    elif ext == 'pdf':
        assert data[:5] == b'%PDF-', 'no PDF header'
    elif ext == 'svg':
        text = data.decode('utf-8', 'replace')
        head = text[:2000]
        assert '<svg' in head and 'xmlns="http://www.w3.org/2000/svg"' in head, 'root is not an SVG element'
    elif ext == 'vsdx':
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            assert names[0] == '[Content_Types].xml', '[Content_Types].xml must be first'
            for part in ('visio/document.xml', 'visio/pages/page1.xml', '_rels/.rels'):
                assert part in names, f'missing {part}'
            page = z.read('visio/pages/page1.xml').decode('utf-8')
            native = page.count("Type='Shape'")
            media = [n for n in names if n.startswith('visio/media/')]
            if native_vsdx:
                assert native >= 2 and "<Section N='Geometry'" in page, 'no native Visio shapes'
                assert '<Text>' in page, 'text was not kept as Visio text'
            else:
                assert media, 'no embedded picture'
                check_emf(z.read(media[0]))
    return len(data)


# ------------------------------------------------------------------ the test
def run(tmp):
    import uno
    from com.sun.star.awt import Point, Size
    from com.sun.star.beans import PropertyValue

    def pv(name, value):
        p = PropertyValue()
        p.Name, p.Value = name, value
        return p

    profile = 'file://' + os.path.join(tmp, 'profile')
    env_arg = '-env:UserInstallation=' + profile
    print('installing extension into a temp profile ...')
    subprocess.run([find_program('unopkg'), 'add', '-f', env_arg, OXT], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    port = free_port()
    office = subprocess.Popen([find_program('soffice'), env_arg, '--headless', '--invisible', '--norestore',
                               f'--accept=socket,host=127.0.0.1,port={port};urp;'],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    desktop = None
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError('LibreOffice did not start')
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        doc = desktop.loadComponentFromURL('private:factory/simpress', '_blank', 0, (pv('Hidden', True),))
        page = doc.DrawPages.getByIndex(0)
        shapes = []
        for kind, x, y, w, h in (('Rectangle', 2000, 2000, 5000, 3000), ('Ellipse', 8000, 2500, 3000, 3000)):
            s = doc.createInstance(f'com.sun.star.drawing.{kind}Shape')
            page.add(s)
            s.Position, s.Size = Point(x, y), Size(w, h)
            shapes.append(s)
        shapes[0].String = '矢印 VecStamp'
        shapes[1].String = '协同'
        coll = smgr.createInstanceWithContext('com.sun.star.drawing.ShapeCollection', ctx)
        for s in shapes:
            coll.add(s)
        doc.CurrentController.select(coll)

        job = smgr.createInstanceWithContext('com.vluckyzhang.vecstamp.Exporter', ctx)
        assert job is not None, 'extension component is not registered'
        failures = 0
        for name, bg in BACKGROUNDS.items():
            out = os.path.join(tmp, 'out-' + name)
            os.makedirs(out)
            native = name != 'white'                    # one run in picture mode
            with open(os.path.join(out, 'settings.json'), 'w', encoding='utf-8') as f:
                json.dump({'bg_type': bg, 'bg_alpha': 30 if bg == 4 else 0, 'bg_color': [200, 230, 255],
                           'margin': 6, 'dpi_index': 3, 'dpi_custom': 450, 'vsdx_native': native}, f)
            job.trigger('selftest:' + out)
            report = json.load(open(os.path.join(out, 'selftest.json'), encoding='utf-8'))
            sizes = []
            for ext in report['files']:
                try:
                    sizes.append(f'{ext}={check_file(ext, os.path.join(out, "selftest." + ext), native) // 1024 + 1}K')
                except AssertionError as exc:
                    failures += 1
                    sizes.append(f'{ext}=FAIL({exc})')
            print(f'  {name:<12}', ' '.join(sizes))
        doc.close(True)
        cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
        menu = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess',
                                              (pv('nodepath', '/org.openoffice.Office.Addons/AddonUI/OfficeMenuBar'),))
        items = menu.getByName('com.vluckyzhang.vecstamp.menu').getByName('Submenu').getElementNames()
        print(f'  menu entries: {len(items)}   extension version: {report["version"]}   '
              f'about dialog controls: {len(report["about"])}')
        return failures
    finally:
        try:                                  # also on failure: never leave an office running
            desktop.terminate()
        except Exception:
            pass
        try:
            office.wait(15)
        except subprocess.TimeoutExpired:
            office.kill()


def main():
    if hasattr(os, 'geteuid') and os.geteuid() == 0:
        sys.exit('please run as a normal user: unopkg refuses to run as root')
    if not os.path.exists(OXT):
        sys.exit('dist/VecStamp-LibreOffice.oxt not found - run python3 build/build.py first')
    tmp = tempfile.mkdtemp(prefix='vecstamp-test-')
    try:
        failures = run(tmp)
    finally:
        if os.environ.get('VECSTAMP_KEEP'):
            print('kept output in', tmp)
        else:
            shutil.rmtree(tmp, ignore_errors=True)
    print('OK' if not failures else f'{failures} check(s) failed')
    sys.exit(1 if failures else 0)


if __name__ == '__main__':
    main()
