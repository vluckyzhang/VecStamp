"""Cross-checks src/customUI14.xml against the VBA module and the icons.

    python build/check_ribbon.py

- every callback named in the ribbon exists as a Public Sub in src/VecStamp.bas
- control ids are unique, every image="..." has a PNG in src/icons
- (optional) validates against Microsoft's customUI14.xsd when you pass its path:
      python build/check_ribbon.py path/to/customui14.xsd   (needs: pip install xmlschema)
"""
import os
import re
import sys
import xml.dom.minidom

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    ui_path = os.path.join(ROOT, 'src', 'customUI14.xml')
    ui = open(ui_path, encoding='utf-8').read()
    bas = open(os.path.join(ROOT, 'src', 'VecStamp.bas'), encoding='utf-8').read()
    xml.dom.minidom.parseString(ui.encode('utf-8'))
    errors = []
    for cb in sorted(set(re.findall(r'(?:on\w+|get\w+)="(VS_\w+)"', ui))):
        if not re.search(r'^Public Sub ' + cb + r'\(', bas, re.M):
            errors.append(f'callback {cb} is not a Public Sub in src/VecStamp.bas')
    ids = re.findall(r'\bid="([^"]+)"', ui)
    errors += [f'duplicate id {i}' for i in sorted(set(ids)) if ids.count(i) > 1]
    for img in sorted(set(re.findall(r'\simage="(\w+)"', ui))):
        if not os.path.exists(os.path.join(ROOT, 'src', 'icons', img + '.png')):
            errors.append(f'missing icon src/icons/{img}.png')
    if len(sys.argv) > 1:
        import xmlschema
        schema = xmlschema.XMLSchema(sys.argv[1], validation='lax')
        errors += [f'schema: {e.reason}' for e in schema.iter_errors(ui_path)]
    for e in errors:
        print('  ', e)
    print('ribbon:', 'OK' if not errors else f'{len(errors)} problem(s)')
    sys.exit(1 if errors else 0)


if __name__ == '__main__':
    main()
