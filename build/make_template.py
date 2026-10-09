"""Build dist/VecStamp-template.pptm: an empty macro-enabled presentation that
already carries the VecStamp ribbon (src/customUI14.xml) and its icons.

PowerPoint (Windows or Mac) opens this template, the VBA module is imported,
and the file is saved as VecStamp.ppam. The ribbon parts travel along.

    pip install python-pptx
    python build/make_template.py
"""
import io
import os
import re
import zipfile

from pptx import Presentation
from pptx.util import Inches

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'src')
OUT = os.path.join(ROOT, 'dist', 'VecStamp-template.pptm')

MAIN_PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
MAIN_PPTM = 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
REL_UI14 = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'


def main():
    ui = open(os.path.join(SRC, 'customUI14.xml'), 'rb').read()
    image_ids = sorted(set(re.findall(rb'\simage="([A-Za-z0-9_]+)"', ui)))
    images = {}
    for raw in image_ids:
        iid = raw.decode()
        path = os.path.join(SRC, 'icons', iid + '.png')
        if not os.path.exists(path):
            raise SystemExit(f'missing icon for image="{iid}": {path}')
        images[iid] = open(path, 'rb').read()

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prs.core_properties.title = 'VecStamp add-in template'
    prs.core_properties.author = 'vluckyzhang'
    buf = io.BytesIO()
    prs.save(buf)

    src = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == '[Content_Types].xml':
                text = data.decode('utf-8').replace(MAIN_PPTX, MAIN_PPTM)
                if 'Extension="png"' not in text:
                    text = text.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1)
                data = text.encode('utf-8')
            elif info.filename == '_rels/.rels':
                text = data.decode('utf-8')
                text = text.replace('</Relationships>',
                                    f'<Relationship Id="rIdVecStampUI" Type="{REL_UI14}" '
                                    'Target="customUI/customUI14.xml"/></Relationships>')
                data = text.encode('utf-8')
            z.writestr(info.filename, data)
        z.writestr('customUI/customUI14.xml', ui)
        rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for iid, data in images.items():
            z.writestr(f'customUI/images/{iid}.png', data)
            rels.append(f'<Relationship Id="{iid}" Type="{REL_IMAGE}" Target="images/{iid}.png"/>')
        rels.append('</Relationships>')
        z.writestr('customUI/_rels/customUI14.xml.rels', ''.join(rels))
    print(f'wrote {OUT} ({len(images)} icons)')


if __name__ == '__main__':
    main()
