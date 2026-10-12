"""Compiles the DrawingML preset shape definitions into a compact form for the VBA geometry engine.

src/presets/presetShapeDefinitions.xml.gz is the ECMA-376 presetShapeDefinitions.xml (see
THIRD_PARTY_NOTICES.md). Every PowerPoint AutoShapeType maps to one preset; for each preset we
emit one ASCII string that the VBA engine (VxPresetGeo in src/VecStamp.bas) evaluates at run time
with the shape's own size and adjustment values, so arcs stay arcs in the exported VSDX.

Compact format (sections separated by "|"):
  0  adjust defaults, comma separated; "!" marks an angle adjustment (PowerPoint reports degrees)
  1  guides, ";" separated: <op>,<arg>,<arg>,<arg>  (op codes in OPS below)
  2  paths, "~" separated: <fill f|n>,<stroke 1|0>,<path w>,<path h>:<cmd>;<cmd>...
       M,x,y  L,x,y  A,wR,hR,stAng,swAng  C,x1,y1,x2,y2,x,y  Q,x1,y1,x,y  Z
  3  text rectangle l,t,r,b
Arguments: $n = guide n (0-based), @n = adjust value n (1-based), integer literal, or a builtin
name (w, h, l, t, r, b, hc, vc, ss, wd2, hd4, ssd8, cd4, k34 = 3cd4, k38, k58, k78, ...).

    python build/make_presets.py          # prints statistics; used by build_bas.py
"""
import gzip
import os
import re
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'src', 'presets', 'presetShapeDefinitions.xml.gz')
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
OPS = {'*/': 'm', '+-': 'a', '+/': 'd', '?:': 'q', 'abs': 'b', 'at2': 't', 'cat2': 'c', 'cos': 'o', 'max': 'x',
       'min': 'n', 'mod': 'h', 'pin': 'p', 'sat2': 's', 'sin': 'i', 'sqrt': 'r', 'tan': 'g', 'val': 'v'}
BUILTINS = {'w', 'h', 'l', 't', 'r', 'b', 'hc', 'vc', 'ss', 'ls', 'wd2', 'wd3', 'wd4', 'wd5', 'wd6', 'wd8', 'wd10',
            'wd12', 'wd32', 'hd2', 'hd3', 'hd4', 'hd5', 'hd6', 'hd8', 'hd10', 'hd12', 'hd32', 'ssd2', 'ssd4', 'ssd6',
            'ssd8', 'ssd16', 'ssd32', 'cd2', 'cd4', 'cd8', '3cd4', '3cd8', '5cd8', '7cd8'}

# builtins starting with a digit get letter-first names, so a leading digit always means a number
ALIASES = {'3cd4': 'k34', '3cd8': 'k38', '5cd8': 'k58', '7cd8': 'k78'}

# MsoAutoShapeType value -> preset name (from python-pptx's MSO_AUTO_SHAPE_TYPE, which follows the
# Office object model). Kept literal so the build does not depend on python-pptx for this.
AUTOSHAPES = {
    1: 'rect', 2: 'parallelogram', 3: 'trapezoid', 4: 'diamond', 5: 'roundRect', 6: 'octagon', 7: 'triangle',
    8: 'rtTriangle', 9: 'ellipse', 10: 'hexagon', 11: 'plus', 12: 'pentagon', 13: 'can', 14: 'cube',
    15: 'bevel', 16: 'foldedCorner', 17: 'smileyFace', 18: 'donut', 19: 'noSmoking', 20: 'blockArc',
    21: 'heart', 22: 'lightningBolt', 23: 'sun', 24: 'moon', 25: 'arc', 26: 'bracketPair', 27: 'bracePair',
    28: 'plaque', 29: 'leftBracket', 30: 'rightBracket', 31: 'leftBrace', 32: 'rightBrace', 33: 'rightArrow',
    34: 'leftArrow', 35: 'upArrow', 36: 'downArrow', 37: 'leftRightArrow', 38: 'upDownArrow', 39: 'quadArrow',
    40: 'leftRightUpArrow', 41: 'bentArrow', 42: 'uturnArrow', 43: 'leftUpArrow', 44: 'bentUpArrow',
    45: 'curvedRightArrow', 46: 'curvedLeftArrow', 47: 'curvedUpArrow', 48: 'curvedDownArrow',
    49: 'stripedRightArrow', 50: 'notchedRightArrow', 51: 'homePlate', 52: 'chevron', 53: 'rightArrowCallout',
    54: 'leftArrowCallout', 55: 'upArrowCallout', 56: 'downArrowCallout', 57: 'leftRightArrowCallout',
    58: 'upDownArrowCallout', 59: 'quadArrowCallout', 60: 'circularArrow', 61: 'flowChartProcess',
    62: 'flowChartAlternateProcess', 63: 'flowChartDecision', 64: 'flowChartInputOutput',
    65: 'flowChartPredefinedProcess', 66: 'flowChartInternalStorage', 67: 'flowChartDocument',
    68: 'flowChartMultidocument', 69: 'flowChartTerminator', 70: 'flowChartPreparation',
    71: 'flowChartManualInput', 72: 'flowChartManualOperation', 73: 'flowChartConnector',
    74: 'flowChartOffpageConnector', 75: 'flowChartPunchedCard', 76: 'flowChartPunchedTape',
    77: 'flowChartSummingJunction', 78: 'flowChartOr', 79: 'flowChartCollate', 80: 'flowChartSort',
    81: 'flowChartExtract', 82: 'flowChartMerge', 83: 'flowChartOnlineStorage', 84: 'flowChartDelay',
    85: 'flowChartMagneticTape', 86: 'flowChartMagneticDisk', 87: 'flowChartMagneticDrum',
    88: 'flowChartDisplay', 89: 'irregularSeal1', 90: 'irregularSeal2', 91: 'star4', 92: 'star5', 93: 'star8',
    94: 'star16', 95: 'star24', 96: 'star32', 97: 'ribbon2', 98: 'ribbon', 99: 'ellipseRibbon2',
    100: 'ellipseRibbon', 101: 'verticalScroll', 102: 'horizontalScroll', 103: 'wave', 104: 'doubleWave',
    105: 'wedgeRectCallout', 106: 'wedgeRoundRectCallout', 107: 'wedgeEllipseCallout', 108: 'cloudCallout',
    109: 'borderCallout1', 110: 'borderCallout2', 111: 'borderCallout3', 112: 'borderCallout3',
    113: 'accentCallout1', 114: 'accentCallout2', 115: 'accentCallout3', 116: 'accentCallout3', 117: 'callout1',
    118: 'callout2', 119: 'callout3', 120: 'callout3', 121: 'accentBorderCallout1', 122: 'accentBorderCallout2',
    123: 'accentBorderCallout3', 124: 'accentBorderCallout3', 125: 'actionButtonBlank', 126: 'actionButtonHome',
    127: 'actionButtonHelp', 128: 'actionButtonInformation', 129: 'actionButtonBackPrevious',
    130: 'actionButtonForwardNext', 131: 'actionButtonBeginning', 132: 'actionButtonEnd',
    133: 'actionButtonReturn', 134: 'actionButtonDocument', 135: 'actionButtonSound', 136: 'actionButtonMovie',
    137: 'wedgeRoundRectCallout', 139: 'flowChartOfflineStorage', 140: 'leftRightRibbon', 141: 'diagStripe',
    142: 'pie', 143: 'nonIsoscelesTrapezoid', 144: 'decagon', 145: 'heptagon', 146: 'dodecagon', 147: 'star6',
    148: 'star7', 149: 'star10', 150: 'star12', 151: 'round1Rect', 152: 'round2SameRect', 153: 'round2DiagRect',
    154: 'snipRoundRect', 155: 'snip1Rect', 156: 'snip2SameRect', 157: 'snip2DiagRect', 158: 'frame',
    159: 'halfFrame', 160: 'teardrop', 161: 'chord', 162: 'corner', 163: 'mathPlus', 164: 'mathMinus',
    165: 'mathMultiply', 166: 'mathDivide', 167: 'mathEqual', 168: 'mathNotEqual', 169: 'cornerTabs',
    170: 'squareTabs', 171: 'plaqueTabs', 172: 'gear6', 173: 'gear9', 174: 'funnel', 175: 'pieWedge',
    176: 'leftCircularArrow', 177: 'leftRightCircularArrow', 178: 'swooshArrow', 179: 'cloud', 180: 'chartX',
    181: 'chartStar', 182: 'chartPlus', 183: 'lineInv',
}


def _fill_autoshapes_from_pptx():
    """Overrides the literal table with python-pptx's when available (it is authoritative)."""
    try:
        from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE as M
    except ImportError:
        return dict(AUTOSHAPES)
    table = {}
    for m in M:
        xv = getattr(m, 'xml_value', None)
        if xv and m.value > 0:
            table.setdefault(m.value, xv)
    return table


def load():
    with gzip.open(SRC, 'rb') as f:
        return ET.parse(f).getroot()


def angle_adjusts(shape, adj_names):
    """Adjust names that end up as angles (arcTo angles or the angle argument of sin / cos / tan)."""
    deps = {}
    gd = shape.find(A + 'gdLst')
    for g in (gd if gd is not None else []):
        deps[g.get('name')] = g.get('fmla').split()
    angle_names = set()
    for _name, f in deps.items():
        if f[0] in ('sin', 'cos', 'tan'):
            angle_names.add(f[2])
    for c in shape.find(A + 'pathLst').iter():
        if c.tag == A + 'arcTo':
            angle_names.update([c.get('stAng'), c.get('swAng')])
    if shape.tag in ('circularArrow', 'leftCircularArrow', 'leftRightCircularArrow'):
        angle_names.update(['adj2', 'adj3', 'adj4'])      # ECMA-376: these three are angles
    # walk back through val / pin / +- guides
    result, todo, seen = set(), list(angle_names), set()
    while todo:
        n = todo.pop()
        if n in seen:
            continue
        seen.add(n)
        if n in adj_names:
            result.add(n)
            continue
        f = deps.get(n)
        if f and f[0] in ('val', 'pin', '+-', '?:', 'min', 'max', 'abs'):
            todo.extend(a for a in f[1:] if not re.match(r'^-?\d+$', a))
    return result


def compile_shape(shape):
    av = shape.find(A + 'avLst')
    gd = shape.find(A + 'gdLst')
    adj = [(g.get('name'), g.get('fmla').split()[1]) for g in (av if av is not None else [])]
    adj_index = {n: i + 1 for i, (n, _) in enumerate(adj)}
    angles = angle_adjusts(shape, set(adj_index))
    guides = [(g.get('name'), g.get('fmla').split()) for g in (gd if gd is not None else [])]
    guide_index = {}

    def ref(tok):
        if re.match(r'^-?\d+$', tok):
            return tok
        if tok in guide_index:
            return '$%d' % guide_index[tok]
        if tok in adj_index:
            return '@%d' % adj_index[tok]
        if tok in BUILTINS:
            return ALIASES.get(tok, tok)
        raise ValueError(f'{shape.tag}: unknown reference {tok}')

    gparts = []
    for i, (name, f) in enumerate(guides):
        gparts.append(','.join([OPS[f[0]]] + [ref(a) for a in f[1:]]))
        guide_index[name] = i       # a guide may refer to earlier guides only
    paths = []
    for p in shape.find(A + 'pathLst'):
        fill = 'n' if p.get('fill') == 'none' else 'f'
        stroke = '0' if p.get('stroke') == 'false' else '1'
        cmds = []
        for c in p:
            tag = c.tag.replace(A, '')
            pts = [ref(v) for pt in c.findall(A + 'pt') for v in (pt.get('x'), pt.get('y'))]
            if tag == 'moveTo':
                cmds.append('M,' + ','.join(pts))
            elif tag == 'lnTo':
                cmds.append('L,' + ','.join(pts))
            elif tag == 'cubicBezTo':
                cmds.append('C,' + ','.join(pts))
            elif tag == 'quadBezTo':
                cmds.append('Q,' + ','.join(pts))
            elif tag == 'arcTo':
                cmds.append('A,' + ','.join(ref(c.get(k)) for k in ('wR', 'hR', 'stAng', 'swAng')))
            elif tag == 'close':
                cmds.append('Z')
        paths.append(f"{fill},{stroke},{p.get('w', '0')},{p.get('h', '0')}:" + ';'.join(cmds))
    rect = shape.find(A + 'rect')
    rtxt = ','.join(ref(rect.get(k)) for k in ('l', 't', 'r', 'b')) if rect is not None else 'l,t,r,b'
    adefs = ','.join(('!' if n in angles else '') + v for n, v in adj)
    out = '|'.join([adefs, ';'.join(gparts), '~'.join(paths), rtxt])
    assert '"' not in out and '|' not in adefs
    return out


def compile_all():
    root = load()
    by_name = {c.tag: c for c in root}
    table = _fill_autoshapes_from_pptx()
    compiled = {}
    for ast, prst in sorted(table.items()):
        if prst in by_name:
            compiled[ast] = (prst, compile_shape(by_name[prst]))
    return compiled


def vba_block(max_chars=30000, width=900):
    """VBA: PresetSrc1..N(ast) and PresetSource(ast) returning the compact definition."""
    compiled = compile_all()
    by_prst = {}
    for ast, (prst, src) in compiled.items():
        by_prst.setdefault((prst, src), []).append(ast)
    funcs, cur, size = [], [], 0
    for (prst, src), asts in sorted(by_prst.items(), key=lambda kv: min(kv[1])):
        lines = ['        Case ' + ', '.join(str(a) for a in sorted(asts)) + "      ' " + prst]
        for i in range(0, len(src), width):
            chunk = src[i:i + width]
            lines.append(f'            s = s & "{chunk}"' if i else f'            s = "{chunk}"')
        n = sum(len(x) for x in lines)
        if cur and size + n > max_chars:
            funcs.append(cur)
            cur, size = [], 0
        cur += lines
        size += n
    if cur:
        funcs.append(cur)
    code = []
    for k, body in enumerate(funcs, 1):
        code += [f'Private Function PresetSrc{k}(ByVal ast As Long) As String', '    Dim s As String',
                 '    Select Case ast'] + body + ['    End Select', f'    PresetSrc{k} = s', 'End Function', '']
    code += ['Private Function PresetSource(ByVal ast As Long) As String', '    Dim s As String']
    for k in range(1, len(funcs) + 1):
        code.append(f'    If Len(s) = 0 Then s = PresetSrc{k}(ast)')
    code += ['    PresetSource = s', 'End Function', '']
    return code


if __name__ == '__main__':
    c = compile_all()
    total = sum(len(s) for _, s in {v: 0 for v in c.values()})
    print(len(c), 'autoshape types,', len({p for p, _ in c.values()}), 'presets,', total, 'chars')
    print('angle adjustments:', sorted({(p, s.split('|')[0]) for p, s in c.values() if '!' in s.split('|')[0]}))
    print('VBA lines:', len(vba_block()))
