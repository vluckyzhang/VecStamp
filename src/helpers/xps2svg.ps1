# VecStamp (Shi Yin) - XPS page -> SVG converter.
#
# Used by the PowerPoint add-in on PowerPoint versions that cannot export SVG themselves
# (Windows, PowerPoint 2013 - 2021 / LTSC): the add-in prints the drawing to XPS with
# PowerPoint ("ExportAsFixedFormat") and this script turns the first page into SVG,
# cropped to the drawing. Paths, gradients, images, clips and transforms are converted
# 1:1; text stays text (font family, size, exact glyph positions) and the subset fonts
# embedded in the XPS are embedded in the SVG as well.
#
# Windows PowerShell 5.1 and PowerShell 7 compatible; no modules needed.
# The add-in embeds this file (build/build_bas.py), writes it to its temp folder and runs:
#   & ([scriptblock]::Create((Get-Content -Raw xps2svg.ps1))) -Xps in.xps -Out out.svg -Frame "l,t,w,h" -Slide "w,h"
# Frame and Slide are in points (slide coordinates). Exit code 0 = success.
#
# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>. MIT License.

param(
    [Parameter(Mandatory = $true)][string]$Xps,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Frame = '',
    [string]$Slide = '',
    [switch]$NoEmbedFonts
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$Inv = [Globalization.CultureInfo]::InvariantCulture
$KeyNs = 'http://schemas.microsoft.com/xps/2005/06/resourcedictionary-key'

# ------------------------------------------------------------------ helpers
function N([double]$v) {
    if ([Math]::Abs($v) -lt 0.0005) { return '0' }
    return $v.ToString('0.###', $Inv)
}

function Num([string]$s) {
    return [double]::Parse($s.Trim(), [Globalization.NumberStyles]::Float, $Inv)
}

function Esc([string]$s) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        $c = [int]$ch
        if ($c -eq 38) { [void]$sb.Append('&amp;') }
        elseif ($c -eq 60) { [void]$sb.Append('&lt;') }
        elseif ($c -eq 62) { [void]$sb.Append('&gt;') }
        elseif ($c -eq 34) { [void]$sb.Append('&quot;') }
        elseif ($c -lt 32 -and $c -ne 9 -and $c -ne 10 -and $c -ne 13) { }
        elseif ($c -eq 0xFFFE -or $c -eq 0xFFFF) { }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Numbers([string]$s) {
    $list = New-Object System.Collections.Generic.List[double]
    foreach ($m in [regex]::Matches($s, '[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')) { $list.Add((Num $m.Value)) }
    return , $list.ToArray()
}

# 2D affine matrices as [a, b, c, d, e, f] (x' = a x + c y + e, y' = b x + d y + f), XPS / SVG order
function MatMul($m, $n) {
    # result = apply m first, then n
    return @(($m[0] * $n[0] + $m[1] * $n[2]), ($m[0] * $n[1] + $m[1] * $n[3]),
        ($m[2] * $n[0] + $m[3] * $n[2]), ($m[2] * $n[1] + $m[3] * $n[3]),
        ($m[4] * $n[0] + $m[5] * $n[2] + $n[4]), ($m[4] * $n[1] + $m[5] * $n[3] + $n[5]))
}

function MatText($m) {
    return 'matrix(' + (($m | ForEach-Object { N $_ }) -join ' ') + ')'
}

function ParseMatrix([string]$s) {
    $v = Numbers $s
    if ($v.Count -ne 6) { return $null }
    return @($v[0], $v[1], $v[2], $v[3], $v[4], $v[5])
}

# ------------------------------------------------------------------ package
$Zip = [IO.Compression.ZipFile]::OpenRead($Xps)
$Parts = @{}
foreach ($e in $Zip.Entries) {
    $name = '/' + [Uri]::UnescapeDataString($e.FullName.Replace('\', '/'))
    $name = [regex]::Replace($name, '/\[\d+\]\.(last\.)?piece$', '')
    $Parts[$name.ToLowerInvariant()] = $e
}

function ReadBytes([string]$part) {
    $e = $Parts[$part.ToLowerInvariant()]
    if ($null -eq $e) { return $null }
    $s = $e.Open()
    $ms = New-Object IO.MemoryStream
    $s.CopyTo($ms)
    $s.Dispose()
    return , $ms.ToArray()
}

function ReadXml([string]$part) {
    $b = ReadBytes $part
    if ($null -eq $b) { return $null }
    $x = New-Object System.Xml.XmlDocument
    $x.PreserveWhitespace = $true
    $x.Load((New-Object IO.MemoryStream (, $b)))
    return $x
}

function Resolve([string]$base, [string]$uri) {
    $u = [Uri]::UnescapeDataString($uri.Trim())
    if ($u.StartsWith('/')) { return $u }
    $dir = $base.Substring(0, $base.LastIndexOf('/') + 1)
    $stack = New-Object System.Collections.ArrayList
    foreach ($seg in ($dir + $u).Split('/')) {
        if ($seg -eq '' -or $seg -eq '.') { continue }
        if ($seg -eq '..') { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) } ; continue }
        [void]$stack.Add($seg)
    }
    return '/' + ($stack -join '/')
}

function FirstPage {
    $rels = ReadXml '/_rels/.rels'
    if ($rels) {
        foreach ($r in $rels.DocumentElement.ChildNodes) {
            if ($r.LocalName -eq 'Relationship' -and $r.GetAttribute('Type') -match 'fixedrepresentation$') {
                $seqPart = Resolve '/' $r.GetAttribute('Target')
                $seq = ReadXml $seqPart
                if (-not $seq) { break }
                foreach ($d in $seq.DocumentElement.ChildNodes) {
                    if ($d.LocalName -ne 'DocumentReference') { continue }
                    $docPart = Resolve $seqPart $d.GetAttribute('Source')
                    $doc = ReadXml $docPart
                    if (-not $doc) { continue }
                    foreach ($p in $doc.DocumentElement.ChildNodes) {
                        if ($p.LocalName -eq 'PageContent') { return (Resolve $docPart $p.GetAttribute('Source')) }
                    }
                }
            }
        }
    }
    $names = @($Parts.Keys | Where-Object { $_ -like '*.fpage' } | Sort-Object)
    if ($names.Count) { return $names[0] }
    throw 'no FixedPage in XPS'
}

# ------------------------------------------------------------------ output state
$Defs = New-Object System.Text.StringBuilder
$Body = New-Object System.Text.StringBuilder
$script:IdN = 0
$Fonts = @{}          # font part -> info hashtable
$Resources = @{}      # x:Key -> XmlElement
$script:Box = $null   # content bounding box in page units [x0, y0, x1, y1]
$script:PageW = 0.0
$script:PageH = 0.0

function NewId([string]$prefix) {
    $script:IdN++
    return $prefix + $script:IdN
}

function AddBox($m, [double[]]$pts, [double]$pad) {
    # pts: x0, y0, x1, y1, ... in local coordinates; m: matrix to page coordinates
    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {
        $x = $m[0] * $pts[$i] + $m[2] * $pts[$i + 1] + $m[4]
        $y = $m[1] * $pts[$i] + $m[3] * $pts[$i + 1] + $m[5]
        if ($null -eq $script:Box) { $script:Box = @(($x - $pad), ($y - $pad), ($x + $pad), ($y + $pad)) }
        else {
            if ($x - $pad -lt $script:Box[0]) { $script:Box[0] = $x - $pad }
            if ($y - $pad -lt $script:Box[1]) { $script:Box[1] = $y - $pad }
            if ($x + $pad -gt $script:Box[2]) { $script:Box[2] = $x + $pad }
            if ($y + $pad -gt $script:Box[3]) { $script:Box[3] = $y + $pad }
        }
    }
}

# ------------------------------------------------------------------ geometry
function FigureBounds([string]$d) {
    # absolute points of an abbreviated-syntax path (end and control points) for bounding boxes
    $pts = New-Object System.Collections.Generic.List[double]
    $toks = [regex]::Matches($d, '[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?')
    $cmd = 'M'; $x = 0.0; $y = 0.0; $sx = 0.0; $sy = 0.0
    $nums = New-Object System.Collections.Generic.List[double]
    $flush = {
        param($c, $vals)
        $rel = ($c -cmatch '[a-z]')
        $C = $c.ToUpperInvariant()
        $k = 0
        $arity = @{ 'M' = 2; 'L' = 2; 'H' = 1; 'V' = 1; 'C' = 6; 'Q' = 4; 'S' = 4; 'T' = 2; 'A' = 7; 'Z' = 0 }[$C]
        if ($null -eq $arity -or $arity -eq 0) { return }
        while ($k + $arity -le $vals.Count) {
            switch ($C) {
                'H' { $script:px = $(if ($rel) { $script:px + $vals[$k] } else { $vals[$k] }) }
                'V' { $script:py = $(if ($rel) { $script:py + $vals[$k] } else { $vals[$k] }) }
                'A' {
                    $nx = $vals[$k + 5]; $ny = $vals[$k + 6]
                    if ($rel) { $nx += $script:px; $ny += $script:py }
                    $r = [Math]::Max([Math]::Abs($vals[$k]), [Math]::Abs($vals[$k + 1]))
                    $pts.Add($script:px - $r); $pts.Add($script:py - $r); $pts.Add($script:px + $r); $pts.Add($script:py + $r)
                    $script:px = $nx; $script:py = $ny
                }
                default {
                    for ($j = 0; $j -lt $arity; $j += 2) {
                        $nx = $vals[$k + $j]; $ny = $vals[$k + $j + 1]
                        if ($rel) { $nx += $script:px; $ny += $script:py }
                        $pts.Add($nx); $pts.Add($ny)
                        if ($j + 2 -ge $arity) { $script:px = $nx; $script:py = $ny }
                    }
                }
            }
            $pts.Add($script:px); $pts.Add($script:py)
            $k += $arity
            if ($C -eq 'M') { $C = 'L' }
        }
    }
    $script:px = 0.0; $script:py = 0.0
    foreach ($t in $toks) {
        $v = $t.Value
        if ($v -cmatch '^[A-Za-z]$') {
            if ($v -eq 'F' -or $v -eq 'f') { $cmd = 'F'; $nums.Clear(); continue }
            if ($cmd -ne 'F') { & $flush $cmd $nums }
            $nums.Clear()
            $cmd = $v
        }
        else { $nums.Add((Num $v)) }
    }
    if ($cmd -ne 'F') { & $flush $cmd $nums }
    return , $pts.ToArray()
}

function GeometryFromElement($g) {
    # PathGeometry element -> @{ d = svg path; rule = 'nonzero'|'evenodd' }
    $rule = 'evenodd'
    if ($g.GetAttribute('FillRule') -eq 'NonZero') { $rule = 'nonzero' }
    $sb = New-Object System.Text.StringBuilder
    $fig = $g.GetAttribute('Figures')
    if ($fig) { [void]$sb.Append(($fig -replace '^\s*[Ff][01]\s*', '')) }
    foreach ($f in $g.ChildNodes) {
        if ($f.LocalName -ne 'PathFigure') { continue }
        [void]$sb.Append(' M ' + $f.GetAttribute('StartPoint'))
        foreach ($s in $f.ChildNodes) {
            switch ($s.LocalName) {
                'PolyLineSegment' { [void]$sb.Append(' L ' + $s.GetAttribute('Points')) }
                'PolyBezierSegment' { [void]$sb.Append(' C ' + $s.GetAttribute('Points')) }
                'PolyQuadraticBezierSegment' { [void]$sb.Append(' Q ' + $s.GetAttribute('Points')) }
                'ArcSegment' {
                    $sz = $s.GetAttribute('Size'); $rot = $s.GetAttribute('RotationAngle'); if (-not $rot) { $rot = '0' }
                    $large = $(if ($s.GetAttribute('IsLargeArc') -eq 'true') { '1' } else { '0' })
                    $sweep = $(if ($s.GetAttribute('SweepDirection') -eq 'Clockwise') { '1' } else { '0' })
                    [void]$sb.Append(' A ' + $sz + ' ' + $rot + ' ' + $large + ' ' + $sweep + ' ' + $s.GetAttribute('Point'))
                }
            }
        }
        if ($f.GetAttribute('IsClosed') -eq 'true') { [void]$sb.Append(' Z') }
    }
    return @{ d = $sb.ToString().Trim(); rule = $rule }
}

function Geometry([string]$attr, $node, [string]$propName) {
    # Data / Clip attribute (abbreviated syntax or StaticResource) or <X.Data> child element
    if ($attr) {
        if ($attr -match '^\{StaticResource\s+(.+)\}$') {
            $r = $Resources[$matches[1].Trim()]
            if ($r) { return (GeometryFromElement $r) }
            return $null
        }
        $rule = 'evenodd'
        if ($attr -match '^\s*[Ff]1') { $rule = 'nonzero' }
        return @{ d = ($attr -replace '^\s*[Ff][01]\s*', ''); rule = $rule }
    }
    foreach ($c in $node.ChildNodes) {
        if ($c.LocalName -eq $propName) {
            foreach ($g in $c.ChildNodes) { if ($g.LocalName -eq 'PathGeometry') { return (GeometryFromElement $g) } }
        }
    }
    return $null
}

# ------------------------------------------------------------------ colours and brushes
function Color([string]$s) {
    # -> @{ rgb = '#rrggbb'; a = 0..1 }
    $s = $s.Trim()
    if ($s -match '^#([0-9A-Fa-f]{8})$') {
        $h = $matches[1]
        return @{ rgb = '#' + $h.Substring(2, 6).ToLowerInvariant(); a = [Convert]::ToInt32($h.Substring(0, 2), 16) / 255.0 }
    }
    if ($s -match '^#([0-9A-Fa-f]{6})$') { return @{ rgb = '#' + $matches[1].ToLowerInvariant(); a = 1.0 } }
    if ($s -match '^sc#') {
        $v = Numbers $s.Substring(3)
        if ($v.Count -eq 3) { $v = @(1.0) + $v }
        $c = @()
        for ($i = 1; $i -le 3; $i++) {
            $x = [Math]::Max(0.0, [Math]::Min(1.0, $v[$i]))
            $x = $(if ($x -le 0.0031308) { 12.92 * $x } else { 1.055 * [Math]::Pow($x, 1 / 2.4) - 0.055 })
            $c += [int][Math]::Round($x * 255)
        }
        return @{ rgb = ('#{0:x2}{1:x2}{2:x2}' -f $c[0], $c[1], $c[2]); a = [Math]::Max(0.0, [Math]::Min(1.0, $v[0])) }
    }
    return @{ rgb = '#000000'; a = 1.0 }
}

function Stops($brush) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $brush.ChildNodes) {
        if ($c.LocalName -notlike '*.GradientStops') { continue }
        foreach ($s in $c.ChildNodes) {
            if ($s.LocalName -ne 'GradientStop') { continue }
            $col = Color $s.GetAttribute('Color')
            [void]$sb.Append('<stop offset="' + (N (Num $s.GetAttribute('Offset'))) + '" stop-color="' + $col.rgb + '"')
            if ($col.a -lt 1) { [void]$sb.Append(' stop-opacity="' + (N $col.a) + '"') }
            [void]$sb.Append('/>')
        }
    }
    return $sb.ToString()
}

function BrushTransform($brush) {
    $t = $brush.GetAttribute('Transform')
    if ($t -match '^\{StaticResource') { $t = '' }
    if ($t) { return (ParseMatrix $t) }
    foreach ($c in $brush.ChildNodes) {
        if ($c.LocalName -like '*.Transform') {
            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }
        }
    }
    return $null
}

function ImageInfo([byte[]]$b) {
    # -> @{ mime; w; h; dpiX; dpiY } (pixel size and resolution)
    $info = @{ mime = 'image/png'; w = 0; h = 0; dpiX = 96.0; dpiY = 96.0 }
    if ($b.Length -gt 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50) {
        $info.w = [int](U32 $b 16)
        $info.h = [int](U32 $b 20)
        $i = 8
        while ($i + 12 -le $b.Length) {
            $len = [int](U32 $b $i)
            $typ = [Text.Encoding]::ASCII.GetString($b, $i + 4, 4)
            if ($typ -eq 'pHYs' -and $b[$i + 16] -eq 1) {
                $px = U32 $b ($i + 8)
                $py = U32 $b ($i + 12)
                if ($px -gt 0) { $info.dpiX = $px * 0.0254 }
                if ($py -gt 0) { $info.dpiY = $py * 0.0254 }
            }
            if ($typ -eq 'IDAT' -or $typ -eq 'IEND' -or $len -lt 0) { break }
            $i += 12 + $len
        }
    }
    elseif ($b.Length -gt 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8) {
        $info.mime = 'image/jpeg'
        $i = 2
        while ($i + 9 -lt $b.Length) {
            if ($b[$i] -ne 0xFF) { $i++; continue }
            $mk = $b[$i + 1]
            $len = ([int]$b[$i + 2] -shl 8) -bor $b[$i + 3]
            if ($mk -eq 0xE0 -and $b[$i + 4] -eq 0x4A -and $b[$i + 11] -eq 1) {
                $dx = ([int]$b[$i + 12] -shl 8) -bor $b[$i + 13]; $dy = ([int]$b[$i + 14] -shl 8) -bor $b[$i + 15]
                if ($dx -gt 0) { $info.dpiX = $dx; $info.dpiY = $dy }
            }
            if ($mk -ge 0xC0 -and $mk -le 0xCF -and $mk -ne 0xC4 -and $mk -ne 0xC8 -and $mk -ne 0xCC) {
                $info.h = ([int]$b[$i + 5] -shl 8) -bor $b[$i + 6]
                $info.w = ([int]$b[$i + 7] -shl 8) -bor $b[$i + 8]
                break
            }
            $i += 2 + $len
        }
    }
    elseif ($b.Length -gt 4 -and (($b[0] -eq 0x49 -and $b[1] -eq 0x49) -or ($b[0] -eq 0x4D -and $b[1] -eq 0x4D))) {
        $info.mime = 'image/tiff'
    }
    return $info
}

function Rect4([string]$s) {
    $v = Numbers $s
    if ($v.Count -lt 4) { return @(0.0, 0.0, 1.0, 1.0) }
    return @($v[0], $v[1], $v[2], $v[3])
}

function Brush($brush, [string]$base) {
    # -> @{ paint = 'none' | '#rrggbb' | 'url(#id)'; a = opacity factor }
    if ($null -eq $brush) { return @{ paint = 'none'; a = 1.0 } }
    $op = 1.0
    if ($brush.GetAttribute('Opacity')) { $op = Num $brush.GetAttribute('Opacity') }
    switch ($brush.LocalName) {
        'SolidColorBrush' {
            $c = Color $brush.GetAttribute('Color')
            return @{ paint = $c.rgb; a = $c.a * $op }
        }
        'LinearGradientBrush' {
            $id = NewId 'g'
            $s = Numbers $brush.GetAttribute('StartPoint'); $e = Numbers $brush.GetAttribute('EndPoint')
            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })
            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }
            [void]$Defs.Append('<linearGradient id="' + $id + '" gradientUnits="' + $units + '" spreadMethod="' + $spread + '" x1="' + (N $s[0]) + '" y1="' + (N $s[1]) + '" x2="' + (N $e[0]) + '" y2="' + (N $e[1]) + '"')
            $m = BrushTransform $brush
            if ($m) { [void]$Defs.Append(' gradientTransform="' + (MatText $m) + '"') }
            [void]$Defs.Append('>' + (Stops $brush) + '</linearGradient>')
            return @{ paint = 'url(#' + $id + ')'; a = $op }
        }
        'RadialGradientBrush' {
            $id = NewId 'g'
            $c = Numbers $brush.GetAttribute('Center'); $o = Numbers $brush.GetAttribute('GradientOrigin')
            $rx = Num $brush.GetAttribute('RadiusX'); $ry = Num $brush.GetAttribute('RadiusY')
            if ($rx -le 0) { $rx = 0.001 }
            if ($ry -le 0) { $ry = 0.001 }
            $units = $(if ($brush.GetAttribute('MappingMode') -eq 'RelativeToBoundingBox') { 'objectBoundingBox' } else { 'userSpaceOnUse' })
            $spread = $brush.GetAttribute('SpreadMethod').ToLowerInvariant(); if (-not $spread) { $spread = 'pad' }
            $k = $ry / $rx
            $scale = @(1.0, 0.0, 0.0, $k, 0.0, ($c[1] - $k * $c[1]))
            $m = BrushTransform $brush
            if ($m) { $scale = MatMul $scale $m }
            $fy = $c[1] + ($o[1] - $c[1]) / $k
            [void]$Defs.Append('<radialGradient id="' + $id + '" gradientUnits="' + $units + '" spreadMethod="' + $spread + '" cx="' + (N $c[0]) + '" cy="' + (N $c[1]) + '" r="' + (N $rx) + '" fx="' + (N $o[0]) + '" fy="' + (N $fy) + '" gradientTransform="' + (MatText $scale) + '">' + (Stops $brush) + '</radialGradient>')
            return @{ paint = 'url(#' + $id + ')'; a = $op }
        }
        'ImageBrush' {
            $src = $brush.GetAttribute('ImageSource')
            if ($src -match '^\{') { return @{ paint = 'none'; a = 1.0 } }
            $part = Resolve $base $src
            $bytes = ReadBytes $part
            if ($null -eq $bytes) { return @{ paint = 'none'; a = 1.0 } }
            $info = ImageInfo $bytes
            $vb = Rect4 $brush.GetAttribute('Viewbox'); $vp = Rect4 $brush.GetAttribute('Viewport')
            $nw = $(if ($info.w -gt 0) { $info.w * 96.0 / $info.dpiX } else { $vb[2] })
            $nh = $(if ($info.h -gt 0) { $info.h * 96.0 / $info.dpiY } else { $vb[3] })
            if ($brush.GetAttribute('ViewboxUnits') -eq 'RelativeToBoundingBox') {
                $vb = @(($vb[0] * $nw), ($vb[1] * $nh), ($vb[2] * $nw), ($vb[3] * $nh))
            }
            if ($vb[2] -le 0) { $vb[2] = 1 }
            if ($vb[3] -le 0) { $vb[3] = 1 }
            $sx = $vp[2] / $vb[2]; $sy = $vp[3] / $vb[3]
            $id = NewId 'p'
            [void]$Defs.Append('<pattern id="' + $id + '" patternUnits="userSpaceOnUse" x="' + (N $vp[0]) + '" y="' + (N $vp[1]) + '" width="' + (N $vp[2]) + '" height="' + (N $vp[3]) + '"')
            $m = BrushTransform $brush
            if ($m) { [void]$Defs.Append(' patternTransform="' + (MatText $m) + '"') }
            [void]$Defs.Append('><image x="' + (N (-$vb[0] * $sx)) + '" y="' + (N (-$vb[1] * $sy)) + '" width="' + (N ($nw * $sx)) + '" height="' + (N ($nh * $sy)) + '" preserveAspectRatio="none" xlink:href="data:' + $info.mime + ';base64,' + [Convert]::ToBase64String($bytes) + '"/></pattern>')
            return @{ paint = 'url(#' + $id + ')'; a = $op }
        }
    }
    return @{ paint = 'none'; a = 1.0 }
}

function BrushOf($node, [string]$attr, [string]$prop, [string]$base) {
    $v = $node.GetAttribute($attr)
    if ($v) {
        if ($v -match '^\{StaticResource\s+(.+)\}$') { return (Brush $Resources[$matches[1].Trim()] $base) }
        $c = Color $v
        return @{ paint = $c.rgb; a = $c.a }
    }
    foreach ($c in $node.ChildNodes) {
        if ($c.LocalName -eq $prop) {
            foreach ($b in $c.ChildNodes) { if ($b.NodeType -eq 'Element') { return (Brush $b $base) } }
        }
    }
    return $null
}

# ------------------------------------------------------------------ fonts
function U16([byte[]]$b, [int]$o) { return ([int]$b[$o] -shl 8) -bor $b[$o + 1] }
function U32([byte[]]$b, [int]$o) { return ([int64]$b[$o] -shl 24) -bor ([int64]$b[$o + 1] -shl 16) -bor ([int64]$b[$o + 2] -shl 8) -bor [int64]$b[$o + 3] }
function S16([byte[]]$b, [int]$o) { $v = U16 $b $o; if ($v -ge 32768) { $v -= 65536 }; return $v }

function FontTables([byte[]]$b) {
    $t = @{}
    if ($b.Length -lt 12) { return $t }
    $off = 0
    if ([Text.Encoding]::ASCII.GetString($b, 0, 4) -eq 'ttcf') { $off = [int](U32 $b 12) }
    $n = U16 $b ($off + 4)
    for ($i = 0; $i -lt $n; $i++) {
        $r = $off + 12 + 16 * $i
        if ($r + 16 -gt $b.Length) { break }
        $t[[Text.Encoding]::ASCII.GetString($b, $r, 4)] = @([int](U32 $b ($r + 8)), [int](U32 $b ($r + 12)))
    }
    return $t
}

function IsFont([byte[]]$b) {
    if ($b.Length -lt 4) { return $false }
    $tag = [Text.Encoding]::ASCII.GetString($b, 0, 4)
    return ($tag -eq 'OTTO' -or $tag -eq 'true' -or $tag -eq 'ttcf' -or ($b[0] -eq 0 -and $b[1] -eq 1 -and $b[2] -eq 0 -and $b[3] -eq 0))
}

function Deobfuscate([byte[]]$b, [string]$part) {
    if (IsFont $b) { return $b }
    $m = [regex]::Match($part, '([0-9A-Fa-f]{8})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{12})')
    if (-not $m.Success) { return $b }
    $hex = ($m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value + $m.Groups[4].Value + $m.Groups[5].Value)
    $fwd = New-Object byte[] 16
    for ($i = 0; $i -lt 16; $i++) { $fwd[$i] = [Convert]::ToByte($hex.Substring(2 * $i, 2), 16) }
    $rev = New-Object byte[] 16
    for ($i = 0; $i -lt 16; $i++) { $rev[$i] = $fwd[15 - $i] }
    foreach ($key in @($rev, $fwd)) {
        $c = [byte[]]$b.Clone()
        for ($i = 0; $i -lt 32 -and $i -lt $c.Length; $i++) { $c[$i] = $c[$i] -bxor $key[$i % 16] }
        if (IsFont $c) { return $c }
    }
    return $b
}

function FontInfo([string]$part) {
    if ($Fonts.ContainsKey($part)) { return $Fonts[$part] }
    $info = @{ family = 'sans-serif'; local = ''; weight = 400; italic = $false; upm = 1000; adv = $null; cmap = $null; rev = $null; id = ''; data = $null }
    $Fonts[$part] = $info
    $b = ReadBytes $part
    if ($null -eq $b) { return $info }
    $b = Deobfuscate $b $part
    if (-not (IsFont $b)) { return $info }
    $info.data = $b
    $t = FontTables $b
    try {
        if ($t.ContainsKey('name')) {
            $o = $t['name'][0]
            $count = U16 $b ($o + 2); $str = $o + (U16 $b ($o + 4))
            $best = @{}
            for ($i = 0; $i -lt $count; $i++) {
                $r = $o + 6 + 12 * $i
                $plat = U16 $b $r; $lang = U16 $b ($r + 4); $nid = U16 $b ($r + 6)
                $len = U16 $b ($r + 8); $off = U16 $b ($r + 10)
                if ($plat -ne 3 -or ($nid -ne 1 -and $nid -ne 16)) { continue }
                $s = [Text.Encoding]::BigEndianUnicode.GetString($b, $str + $off, $len)
                $key = $(if ($lang -eq 0x409) { 'en' } elseif ($lang -eq 0x804 -or $lang -eq 0x404) { 'zh' } else { 'x' }) + $nid
                $best[$key] = $s
            }
            foreach ($k in @('en16', 'en1', 'x16', 'x1')) { if ($best[$k]) { $info.family = $best[$k]; break } }
            foreach ($k in @('zh16', 'zh1')) { if ($best[$k] -and $best[$k] -ne $info.family) { $info.local = $best[$k]; break } }
        }
        if ($t.ContainsKey('OS/2')) {
            $o = $t['OS/2'][0]
            $info.weight = U16 $b ($o + 4)
            $info.italic = ((U16 $b ($o + 62)) -band 1) -eq 1
        }
        if ($t.ContainsKey('head')) { $info.upm = U16 $b ($t['head'][0] + 18) }
        if ($t.ContainsKey('hhea') -and $t.ContainsKey('hmtx')) {
            $nh = U16 $b ($t['hhea'][0] + 34)
            $o = $t['hmtx'][0]
            $adv = New-Object int[] ([Math]::Max($nh, 1))
            for ($i = 0; $i -lt $nh; $i++) { $adv[$i] = U16 $b ($o + 4 * $i) }
            $info.adv = $adv
        }
        if ($t.ContainsKey('cmap')) {
            $o = $t['cmap'][0]
            $n = U16 $b ($o + 2)
            $sub = -1
            for ($i = 0; $i -lt $n; $i++) {
                $r = $o + 4 + 8 * $i
                $plat = U16 $b $r; $eid = U16 $b ($r + 2); $so = $o + [int](U32 $b ($r + 4))
                $fmt = U16 $b $so
                if ($plat -eq 3 -and ($eid -eq 1 -or $eid -eq 10 -or $eid -eq 0) -and ($fmt -eq 4 -or $fmt -eq 12)) { $sub = $so; if ($fmt -eq 12) { break } }
            }
            if ($sub -ge 0) {
                $map = @{}
                $fmt = U16 $b $sub
                if ($fmt -eq 4) {
                    $segX2 = U16 $b ($sub + 6)
                    $ends = $sub + 14; $starts = $ends + $segX2 + 2; $deltas = $starts + $segX2; $ranges = $deltas + $segX2
                    for ($s = 0; $s -lt $segX2 / 2; $s++) {
                        $e = U16 $b ($ends + 2 * $s); $st = U16 $b ($starts + 2 * $s)
                        $dl = U16 $b ($deltas + 2 * $s); $ro = U16 $b ($ranges + 2 * $s)
                        if ($st -eq 0xFFFF) { continue }
                        for ($c = $st; $c -le $e; $c++) {
                            if ($ro -eq 0) { $g = ($c + $dl) % 65536 }
                            else {
                                $ga = $ranges + 2 * $s + $ro + 2 * ($c - $st)
                                $g = U16 $b $ga
                                if ($g -ne 0) { $g = ($g + $dl) % 65536 }
                            }
                            if ($g -ne 0) { $map[$c] = $g }
                        }
                    }
                }
                else {
                    $ng = [int](U32 $b ($sub + 12))
                    for ($s = 0; $s -lt $ng; $s++) {
                        $r = $sub + 16 + 12 * $s
                        $st = [int](U32 $b $r); $e = [int](U32 $b ($r + 4)); $g0 = [int](U32 $b ($r + 8))
                        if ($e - $st -gt 70000) { continue }
                        for ($c = $st; $c -le $e; $c++) { $map[$c] = $g0 + ($c - $st) }
                    }
                }
                $info.cmap = $map
            }
        }
    }
    catch { }
    $info.id = NewId 'f'
    if (-not $NoEmbedFonts -and [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'ttcf') {
        [void]$Defs.Append('<style>@font-face{font-family:''' + $info.id + ''';src:url(data:font/ttf;base64,' + [Convert]::ToBase64String($b) + ');}</style>')
    }
    return $info
}

function GlyphAdvance($info, [int]$gid) {
    if ($null -eq $info.adv -or $info.adv.Count -eq 0) { return 0.0 }
    if ($gid -ge $info.adv.Count) { $gid = $info.adv.Count - 1 }
    if ($gid -lt 0) { return 0.0 }
    return $info.adv[$gid] / [double]$info.upm
}

function GidOf($info, [int]$cp) {
    if ($null -eq $info.cmap) { return -1 }
    $g = $info.cmap[$cp]
    if ($null -eq $g) { return -1 }
    return [int]$g
}

function CharFromGid($info, [int]$gid) {
    if ($null -eq $info.cmap) { return $null }
    if ($null -eq $info.rev) {
        $info.rev = @{}
        foreach ($k in $info.cmap.Keys) { if (-not $info.rev.ContainsKey($info.cmap[$k])) { $info.rev[$info.cmap[$k]] = $k } }
    }
    $cp = $info.rev[$gid]
    if ($null -eq $cp) { return $null }
    return [char]::ConvertFromUtf32([int]$cp)
}

# ------------------------------------------------------------------ elements
function TransformOf($node, [string]$prop) {
    $t = $node.GetAttribute('RenderTransform')
    if ($t -match '^\{StaticResource\s+(.+)\}$') {
        $r = $Resources[$matches[1].Trim()]
        if ($r) { return (ParseMatrix $r.GetAttribute('Matrix')) }
        return $null
    }
    if ($t) { return (ParseMatrix $t) }
    foreach ($c in $node.ChildNodes) {
        if ($c.LocalName -eq $prop) {
            foreach ($m in $c.ChildNodes) { if ($m.LocalName -eq 'MatrixTransform') { return (ParseMatrix $m.GetAttribute('Matrix')) } }
        }
    }
    return $null
}

function CommonAttrs($node, [string]$kind, $ctm, [ref]$mOut, [string]$base) {
    # transform, clip and opacity of an element -> SVG attribute string; mOut receives the new CTM
    $s = ''
    $m = TransformOf $node ($kind + '.RenderTransform')
    $cur = $ctm
    if ($m) { $s += ' transform="' + (MatText $m) + '"'; $cur = MatMul $m $ctm }
    $mOut.Value = $cur
    $clip = Geometry $node.GetAttribute('Clip') $node ($kind + '.Clip')
    if ($clip -and $clip.d) {
        $id = NewId 'c'
        [void]$Defs.Append('<clipPath id="' + $id + '"><path d="' + $clip.d + '" clip-rule="' + $clip.rule + '"/></clipPath>')
        $s += ' clip-path="url(#' + $id + ')"'
    }
    $op = 1.0
    if ($node.GetAttribute('Opacity')) { $op = Num $node.GetAttribute('Opacity') }
    $mask = BrushOf $node 'OpacityMask' ($kind + '.OpacityMask') $base
    if ($mask -and $mask.paint -match '^#') { $op *= $mask.a }
    if ($op -lt 0.999) { $s += ' opacity="' + (N $op) + '"' }
    return $s
}

function LoadResources($node, [string]$base) {
    foreach ($c in $node.ChildNodes) {
        if ($c.LocalName -notlike '*.Resources') { continue }
        foreach ($d in $c.ChildNodes) {
            if ($d.LocalName -ne 'ResourceDictionary') { continue }
            $src = $d.GetAttribute('Source')
            if ($src) {
                $part = Resolve $base $src
                $x = ReadXml $part
                if ($x) { foreach ($r in $x.DocumentElement.ChildNodes) { if ($r.NodeType -eq 'Element') { $k = $r.GetAttribute('Key', $KeyNs); if ($k) { $Resources[$k] = $r } } } }
            }
            foreach ($r in $d.ChildNodes) {
                if ($r.NodeType -ne 'Element') { continue }
                $k = $r.GetAttribute('Key', $KeyNs)
                if ($k) { $Resources[$k] = $r }
            }
        }
    }
}

function IsPageBackground($pts, $ctm) {
    if ($null -eq $pts -or $pts.Count -lt 4) { return $false }
    $x0 = [double]::MaxValue; $y0 = [double]::MaxValue; $x1 = -[double]::MaxValue; $y1 = -[double]::MaxValue
    for ($i = 0; $i + 1 -lt $pts.Count; $i += 2) {
        $x = $ctm[0] * $pts[$i] + $ctm[2] * $pts[$i + 1] + $ctm[4]
        $y = $ctm[1] * $pts[$i] + $ctm[3] * $pts[$i + 1] + $ctm[5]
        $x0 = [Math]::Min($x0, $x); $y0 = [Math]::Min($y0, $y); $x1 = [Math]::Max($x1, $x); $y1 = [Math]::Max($y1, $y)
    }
    return ($x0 -le 0.5 -and $y0 -le 0.5 -and $x1 -ge $script:PageW - 0.5 -and $y1 -ge $script:PageH - 0.5)
}

function DoPath($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {
    $g = Geometry $node.GetAttribute('Data') $node 'Path.Data'
    if ($null -eq $g -or -not $g.d) { return }
    $pts = FigureBounds $g.d
    $m = $null
    $attrs = CommonAttrs $node 'Path' $ctm ([ref]$m) $base
    if (IsPageBackground $pts $m) { return }
    $fill = BrushOf $node 'Fill' 'Path.Fill' $base
    $stroke = BrushOf $node 'Stroke' 'Path.Stroke' $base
    [void]$sb.Append('<path d="' + $g.d + '"' + $attrs)
    if ($null -eq $fill -or $fill.paint -eq 'none') { [void]$sb.Append(' fill="none"') }
    else {
        [void]$sb.Append(' fill="' + $fill.paint + '"')
        if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity="' + (N $fill.a) + '"') }
        if ($g.rule -eq 'evenodd') { [void]$sb.Append(' fill-rule="evenodd"') }
    }
    $pad = 0.0
    if ($stroke -and $stroke.paint -ne 'none') {
        $w = 1.0
        if ($node.GetAttribute('StrokeThickness')) { $w = Num $node.GetAttribute('StrokeThickness') }
        $pad = $w / 2
        [void]$sb.Append(' stroke="' + $stroke.paint + '" stroke-width="' + (N $w) + '"')
        if ($stroke.a -lt 0.999) { [void]$sb.Append(' stroke-opacity="' + (N $stroke.a) + '"') }
        $join = $node.GetAttribute('StrokeLineJoin')
        if ($join -eq 'Round') { [void]$sb.Append(' stroke-linejoin="round"') }
        elseif ($join -eq 'Bevel') { [void]$sb.Append(' stroke-linejoin="bevel"') }
        else {
            $ml = 10.0
            if ($node.GetAttribute('StrokeMiterLimit')) { $ml = Num $node.GetAttribute('StrokeMiterLimit') }
            [void]$sb.Append(' stroke-miterlimit="' + (N ([Math]::Max($ml, 1.0))) + '"')
        }
        $cap = $node.GetAttribute('StrokeStartLineCap')
        if ($cap -eq 'Round') { [void]$sb.Append(' stroke-linecap="round"') }
        elseif ($cap -eq 'Square') { [void]$sb.Append(' stroke-linecap="square"') }
        $dash = $node.GetAttribute('StrokeDashArray')
        if ($dash) {
            $dv = Numbers $dash
            if ($dv.Count -gt 0) {
                [void]$sb.Append(' stroke-dasharray="' + (($dv | ForEach-Object { N ([Math]::Max($_ * $w, 0.01)) }) -join ' ') + '"')
                if ($node.GetAttribute('StrokeDashOffset')) { [void]$sb.Append(' stroke-dashoffset="' + (N ((Num $node.GetAttribute('StrokeDashOffset')) * $w)) + '"') }
                if ($node.GetAttribute('StrokeDashCap') -eq 'Round') { [void]$sb.Append(' stroke-linecap="round"') }
            }
        }
    }
    [void]$sb.Append('/>')
    if (($fill -and $fill.paint -ne 'none') -or ($stroke -and $stroke.paint -ne 'none')) { AddBox $m $pts $pad }
}

function DoGlyphs($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {
    $fontPart = Resolve $base $node.GetAttribute('FontUri')
    $info = FontInfo $fontPart
    $em = Num $node.GetAttribute('FontRenderingEmSize')
    $ox = Num $node.GetAttribute('OriginX'); $oy = Num $node.GetAttribute('OriginY')
    $text = $node.GetAttribute('UnicodeString')
    if ($text.StartsWith('{}')) { $text = $text.Substring(2) }
    $idx = $node.GetAttribute('Indices')
    $entries = @()
    if ($idx) { $entries = $idx.Split(';') }
    $hasClusters = $idx -match '\('
    $rtl = $false
    if ($node.GetAttribute('BidiLevel')) { $rtl = ((Num $node.GetAttribute('BidiLevel')) % 2) -eq 1 }
    # code points of the text
    $cps = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $text.Length; $i++) {
        if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }
        else { $cps.Add([int]$text[$i]) }
    }
    if ($cps.Count -eq 0 -and $entries.Count -gt 0) {
        # glyph indices only: recover the characters from the font's cmap
        $chars = ''
        foreach ($e in $entries) {
            $gi = ($e.Split(',')[0] -replace '\(.*\)', '').Trim()
            if ($gi -match '^\d+$') { $c = CharFromGid $info ([int]$gi); if ($c) { $chars += $c } }
        }
        if (-not $chars) { return }
        $text = $chars
        for ($i = 0; $i -lt $text.Length; $i++) {
            if ([char]::IsHighSurrogate($text[$i]) -and $i + 1 -lt $text.Length) { $cps.Add([char]::ConvertToUtf32($text[$i], $text[$i + 1])); $i++ }
            else { $cps.Add([int]$text[$i]) }
        }
    }
    if ($cps.Count -eq 0) { return }
    # x position of every character from the advance widths
    $xs = New-Object System.Collections.Generic.List[string]
    $x = $ox
    $dir = $(if ($rtl) { -1.0 } else { 1.0 })
    $total = 0.0
    for ($i = 0; $i -lt $cps.Count; $i++) {
        $gid = -1; $adv = $null; $uoff = 0.0
        if (-not $hasClusters -and $i -lt $entries.Count) {
            $f = $entries[$i].Split(',')
            if ($f[0].Trim() -match '^\d+$') { $gid = [int]$f[0].Trim() }
            if ($f.Count -gt 1 -and $f[1].Trim()) { $adv = (Num $f[1]) / 100.0 }
            if ($f.Count -gt 2 -and $f[2].Trim()) { $uoff = (Num $f[2]) / 100.0 * $em }
        }
        if ($gid -lt 0) { $gid = GidOf $info $cps[$i] }
        if ($null -eq $adv) { $adv = GlyphAdvance $info $gid }
        $xs.Add((N ($x + $dir * $uoff)))
        $x += $dir * $adv * $em
        $total += $adv * $em
    }
    $fill = BrushOf $node 'Fill' 'Glyphs.Fill' $base
    if ($null -eq $fill) { $fill = @{ paint = '#000000'; a = 1.0 } }
    $m = $null
    $attrs = CommonAttrs $node 'Glyphs' $ctm ([ref]$m) $base
    $fam = "'" + $info.family.Replace("'", '') + "'"
    if ($info.local) { $fam += ", '" + $info.local.Replace("'", '') + "'" }
    if ($info.data -and -not $NoEmbedFonts) { $fam = $info.id + ', ' + $fam }
    $weight = $info.weight
    $sim = $node.GetAttribute('StyleSimulations')
    if ($sim -match 'Bold') { $weight = 700 }
    [void]$sb.Append('<text xml:space="preserve"' + $attrs + ' x="' + ($xs -join ' ') + '" y="' + (N $oy) + '" font-family="' + (Esc $fam) + '" font-size="' + (N $em) + '"')
    if ($weight -ge 600) { [void]$sb.Append(' font-weight="' + $weight + '"') }
    if ($info.italic -or $sim -match 'Italic') { [void]$sb.Append(' font-style="italic"') }
    if ($fill.paint -ne '#000000') { [void]$sb.Append(' fill="' + $fill.paint + '"') }
    if ($fill.a -lt 0.999) { [void]$sb.Append(' fill-opacity="' + (N $fill.a) + '"') }
    [void]$sb.Append('>' + (Esc $text) + '</text>')
    $x0 = [Math]::Min($ox, $ox + $dir * $total); $x1 = [Math]::Max($ox, $ox + $dir * $total)
    AddBox $m @($x0, ($oy - 0.95 * $em), $x1, ($oy + 0.3 * $em)) 0.0
}

function DoChildren($node, $ctm, [string]$base, [System.Text.StringBuilder]$sb) {
    foreach ($c in $node.ChildNodes) {
        if ($c.NodeType -ne 'Element') { continue }
        switch ($c.LocalName) {
            'Path' { DoPath $c $ctm $base $sb }
            'Glyphs' { DoGlyphs $c $ctm $base $sb }
            'Canvas' {
                LoadResources $c $base
                $m = $null
                $attrs = CommonAttrs $c 'Canvas' $ctm ([ref]$m) $base
                $inner = New-Object System.Text.StringBuilder
                DoChildren $c $m $base $inner
                if ($inner.Length -gt 0) { [void]$sb.Append('<g' + $attrs + '>' + $inner.ToString() + '</g>') }
            }
        }
    }
}

# ------------------------------------------------------------------ main
try {
    $page = FirstPage
    $doc = ReadXml $page
    if (-not $doc) { throw "cannot read $page" }
    $root = $doc.DocumentElement
    $script:PageW = Num $root.GetAttribute('Width')
    $script:PageH = Num $root.GetAttribute('Height')
    LoadResources $root $page
    DoChildren $root @(1.0, 0.0, 0.0, 1.0, 0.0, 0.0) $page $Body
    if ($Body.Length -eq 0) { throw 'empty page' }

    # crop: the drawing's frame (from PowerPoint) united with what was actually drawn
    $crop = $null
    if ($Frame -and $Slide) {
        $f = Numbers $Frame; $s = Numbers $Slide
        if ($f.Count -eq 4 -and $s.Count -eq 2 -and $s[0] -gt 0) {
            $k = $script:PageW / $s[0]
            $crop = @(($f[0] * $k), ($f[1] * $k), (($f[0] + $f[2]) * $k), (($f[1] + $f[3]) * $k))
        }
    }
    if ($script:Box) {
        $b = @([Math]::Max($script:Box[0], 0.0), [Math]::Max($script:Box[1], 0.0),
            [Math]::Min($script:Box[2], $script:PageW), [Math]::Min($script:Box[3], $script:PageH))
        if ($null -eq $crop) { $crop = $b }
        elseif ($b[2] -gt $b[0] -and $b[3] -gt $b[1]) {
            $crop = @([Math]::Min($crop[0], $b[0]), [Math]::Min($crop[1], $b[1]), [Math]::Max($crop[2], $b[2]), [Math]::Max($crop[3], $b[3]))
        }
    }
    if ($null -eq $crop) { $crop = @(0.0, 0.0, $script:PageW, $script:PageH) }
    $cw = [Math]::Max($crop[2] - $crop[0], 0.1); $ch = [Math]::Max($crop[3] - $crop[1], 0.1)

    $svg = New-Object System.Text.StringBuilder
    [void]$svg.Append('<?xml version="1.0" encoding="UTF-8"?>' + "`n")
    [void]$svg.Append('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" width="' + (N ($cw * 0.75)) + 'pt" height="' + (N ($ch * 0.75)) + 'pt" viewBox="' + (N $crop[0]) + ' ' + (N $crop[1]) + ' ' + (N $cw) + ' ' + (N $ch) + '">' + "`n")
    [void]$svg.Append('<!-- Created by VecStamp (https://github.com/vluckyzhang/VecStamp) from PowerPoint XPS output -->' + "`n")
    if ($Defs.Length -gt 0) { [void]$svg.Append('<defs>' + $Defs.ToString() + '</defs>' + "`n") }
    [void]$svg.Append($Body.ToString() + "`n</svg>`n")
    $enc = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Out, $svg.ToString(), $enc)
    $Zip.Dispose()
    exit 0
}
catch {
    try { $Zip.Dispose() } catch { }
    [Console]::Error.WriteLine('xps2svg: ' + $_.Exception.Message)
    exit 1
}
