"""Builds tests/xps/sample.xps: a FixedPage that exercises what PowerPoint's XPS output uses
(canvases with transforms and clips, solid / gradient / image brushes, static resources,
PathGeometry elements, dashed strokes and Glyphs with obfuscated embedded fonts).

    python3 tests/make_xps_fixture.py        # needs fontTools and Pillow
"""
import io
import os
import uuid
import zipfile

from fontTools import subset
from fontTools.ttLib import TTFont
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'xps', 'sample.xps')
LATIN = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
CJK = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'


def subset_font(path, text, number=0):
    opts = subset.Options()
    opts.name_IDs = ['*']
    opts.name_languages = ['*']
    opts.notdef_outline = True
    font = TTFont(path, fontNumber=number)
    s = subset.Subsetter(opts)
    s.populate(text=text)
    s.subset(font)
    buf = io.BytesIO()
    font.save(buf)
    return buf.getvalue(), font


def obfuscate(data, guid):
    key = bytes.fromhex(guid.replace('-', ''))[::-1]
    b = bytearray(data)
    for i in range(32):
        b[i] ^= key[i % 16]
    return bytes(b)


def advances(font, text):
    cmap = font.getBestCmap()
    hmtx = font['hmtx']
    upm = font['head'].unitsPerEm
    return ';'.join(',%.2f' % (hmtx[cmap[ord(c)]][0] * 100.0 / upm) for c in text)


def main():
    cjk_text = '矢印协同框架'
    lat_text = 'VecStamp XPS'
    lat, lat_font = subset_font(LATIN, lat_text + 'Hello')
    cjk, cjk_font = subset_font(CJK, cjk_text, 2)
    g1, g2 = str(uuid.uuid4()).upper(), str(uuid.uuid4()).upper()
    img = Image.new('RGB', (64, 32), '#FFF8EB')
    d = ImageDraw.Draw(img)
    d.rectangle([4, 4, 28, 28], fill='#C8372D')
    d.ellipse([36, 4, 60, 28], fill='#2F6FB0')
    png = io.BytesIO()
    img.save(png, 'PNG', dpi=(192, 192))
    page = f'''<FixedPage Width="960" Height="540" xmlns="http://schemas.microsoft.com/xps/2005/06" xmlns:x="http://schemas.microsoft.com/xps/2005/06/resourcedictionary-key" xml:lang="zh-CN">
 <FixedPage.Resources><ResourceDictionary>
  <SolidColorBrush x:Key="b0" Color="#FF138D90"/>
  <PathGeometry x:Key="g0" Figures="M 700,300 L 820,300 760,400 Z"/>
 </ResourceDictionary></FixedPage.Resources>
 <Path Data="M 0,0 L 960,0 960,540 0,540 Z" Fill="#FFFFFFFF"/>
 <Canvas RenderTransform="1.3333333,0,0,1.3333333,0,0">
  <Path Data="F1 M 60,60 L 240,60 240,150 60,150 Z" Fill="#FF2F6FB0" Stroke="#FF1A3D66" StrokeThickness="1.5"/>
  <Glyphs OriginX="80" OriginY="112" FontRenderingEmSize="20" FontUri="/Documents/1/Resources/Fonts/{g2}.odttf" UnicodeString="{cjk_text[:2]}" Indices="{advances(cjk_font, cjk_text[:2])}" Fill="#FFFFFFFF"/>
  <Path Data="M 300,80 C 360,20 420,180 480,90" Stroke="#FF2E8B57" StrokeThickness="4" StrokeDashArray="2 1" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
  <Path Data="M 520,105 A 60,40 0 1 1 640,105 A 60,40 0 1 1 520,105 Z" Fill="#80C8372D"/>
 </Canvas>
 <Path Data="M 80,260 L 380,260 380,360 80,360 Z">
  <Path.Fill><LinearGradientBrush MappingMode="Absolute" StartPoint="80,0" EndPoint="380,0" SpreadMethod="Pad">
   <LinearGradientBrush.GradientStops><GradientStop Color="#FF2F6FB0" Offset="0"/><GradientStop Color="#FFFFF8EB" Offset="1"/></LinearGradientBrush.GradientStops>
  </LinearGradientBrush></Path.Fill>
 </Path>
 <Path Data="M 420,310 A 70,50 0 1 1 560,310 A 70,50 0 1 1 420,310 Z">
  <Path.Fill><RadialGradientBrush MappingMode="Absolute" Center="490,310" GradientOrigin="470,290" RadiusX="70" RadiusY="50">
   <RadialGradientBrush.GradientStops><GradientStop Color="#FFFFFFFF" Offset="0"/><GradientStop Color="#FFD98324" Offset="1"/></RadialGradientBrush.GradientStops>
  </RadialGradientBrush></Path.Fill>
 </Path>
 <Path Data="{{StaticResource g0}}" Fill="{{StaticResource b0}}"/>
 <Canvas Clip="M 600,430 L 900,430 900,500 600,500 Z" Opacity="0.9">
  <Path Data="M 600,430 L 728,430 728,494 600,494 Z">
   <Path.Fill><ImageBrush ImageSource="/Documents/1/Resources/Images/1.png" Viewbox="0,0,32,16" ViewboxUnits="Absolute" Viewport="600,430,128,64" ViewportUnits="Absolute" TileMode="None"/></Path.Fill>
  </Path>
  <Path Fill="#FF7A4FB0"><Path.Data><PathGeometry FillRule="NonZero"><PathFigure StartPoint="760,440" IsClosed="true"><PolyLineSegment Points="880,440 880,490"/><ArcSegment Point="760,490" Size="60,25" RotationAngle="0" IsLargeArc="false" SweepDirection="Clockwise"/></PathFigure></PathGeometry></Path.Data></Path>
 </Canvas>
 <Glyphs OriginX="80" OriginY="460" FontRenderingEmSize="40" FontUri="/Documents/1/Resources/Fonts/{g1}.odttf" UnicodeString="{lat_text}" Fill="#FF2B2F36"/>
 <Glyphs OriginX="80" OriginY="510" FontRenderingEmSize="28" FontUri="/Documents/1/Resources/Fonts/{g2}.odttf" UnicodeString="{cjk_text}" Indices="{advances(cjk_font, cjk_text)}" Fill="#FFC8372D" StyleSimulations="BoldSimulation"/>
</FixedPage>'''
    ct = '''<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="fdseq" ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/><Default Extension="fdoc" ContentType="application/vnd.ms-package.xps-fixeddocument+xml"/><Default Extension="fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/><Default Extension="odttf" ContentType="application/vnd.ms-package.obfuscated-opentype"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'''
    rels = '''<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" Target="/FixedDocumentSequence.fdseq" Id="R0"/></Relationships>'''
    fdseq = '<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06"><DocumentReference Source="Documents/1/FixedDocument.fdoc"/></FixedDocumentSequence>'
    fdoc = '<FixedDocument xmlns="http://schemas.microsoft.com/xps/2005/06"><PageContent Source="Pages/1.fpage"/></FixedDocument>'
    prels = f'''<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.microsoft.com/xps/2005/06/required-resource" Target="../Resources/Fonts/{g1}.odttf" Id="R1"/><Relationship Type="http://schemas.microsoft.com/xps/2005/06/required-resource" Target="../Resources/Fonts/{g2}.odttf" Id="R2"/><Relationship Type="http://schemas.microsoft.com/xps/2005/06/required-resource" Target="../Resources/Images/1.png" Id="R3"/></Relationships>'''
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', rels)
        z.writestr('FixedDocumentSequence.fdseq', fdseq)
        z.writestr('Documents/1/FixedDocument.fdoc', fdoc)
        z.writestr('Documents/1/Pages/1.fpage', page)
        z.writestr('Documents/1/Pages/_rels/1.fpage.rels', prels)
        z.writestr(f'Documents/1/Resources/Fonts/{g1}.odttf', obfuscate(lat, g1))
        z.writestr(f'Documents/1/Resources/Fonts/{g2}.odttf', obfuscate(cjk, g2))
        z.writestr('Documents/1/Resources/Images/1.png', png.getvalue())
    print('wrote', OUT)


if __name__ == '__main__':
    main()
