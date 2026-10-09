"""One-stop build for VecStamp (run from anywhere):

    pip install cairosvg python-pptx pillow
    python build/build.py            # add --promo to also redraw the README / website images

Steps
1. make_icons.py         - logo, banner, ribbon and toolbar icons (skipped if cairosvg is missing)
2. build_bas.py          - src/VecStamp.bas -> dist/VecStamp.bas (ASCII-only, Chinese as U("hex"),
                           VSDX template parts embedded)
3. check_vba.py          - static checks for the Windows and macOS compile paths (source and build)
4. check_ribbon.py       - ribbon callbacks, ids and icons
5. make_template.py      - dist/VecStamp-template.pptm with the ribbon and icons
6. make_lo_extension.py  - dist/VecStamp-LibreOffice.oxt (Linux / any LibreOffice)
7. make_promo.py         - README images in assets/promo (only with --promo; needs Noto Sans CJK)
8. make_website.py       - copies images and the version number into website/ (only with --promo)

The final step, turning the template + module into VecStamp.ppam, needs
PowerPoint itself: run the Windows installer and choose [4], or follow the
manual steps in README.md (works on Windows and macOS).
"""
import importlib.util
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def run(*args):
    print('>', ' '.join(args))
    subprocess.check_call([sys.executable, *args], cwd=ROOT)


def main():
    if importlib.util.find_spec('cairosvg'):
        run(os.path.join(HERE, 'make_icons.py'))
    else:
        print('cairosvg not installed: keeping the existing icons')
    run(os.path.join(HERE, 'build_bas.py'), os.path.join('src', 'VecStamp.bas'), os.path.join('dist', 'VecStamp.bas'))
    run(os.path.join(HERE, 'check_vba.py'), os.path.join('src', 'VecStamp.bas'))
    run(os.path.join(HERE, 'check_vba.py'), os.path.join('dist', 'VecStamp.bas'))
    run(os.path.join(HERE, 'check_ribbon.py'))
    run(os.path.join(HERE, 'make_template.py'))
    run(os.path.join(HERE, 'make_lo_extension.py'))
    if '--promo' in sys.argv[1:]:
        run(os.path.join(HERE, 'make_promo.py'))
        run(os.path.join(HERE, 'make_website.py'))
    print('\nDone. Next: build VecStamp.ppam with PowerPoint (see docs/DEVELOPMENT.md).')


if __name__ == '__main__':
    main()
