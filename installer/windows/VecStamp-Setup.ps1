# 矢印 VecStamp - Windows 安装 / 卸载程序
# Copyright (c) 2026 vluckyzhang <vluckyzhang@gmail.com>
# Released under the MIT License.
#
# 由根目录的「安装或卸载-Windows.bat」调用，也可以直接运行：
#   powershell -ExecutionPolicy Bypass -File installer\windows\VecStamp-Setup.ps1 [-Action Install|Uninstall|Trust|Build]
#
# 本程序只修改当前 Windows 用户（HKCU）的设置：
#   · 安装：把 VecStamp.ppam 复制到 %APPDATA%\Microsoft\AddIns，并登记为 PowerPoint 自动加载的加载项
#   · 从源码生成 .ppam 时（需要你确认）：临时开启「信任对 VBA 工程对象模型的访问」，完成后立即恢复
#   · 受信任位置：仅在你选择菜单 [3] 时才添加
#   · 卸载：删除以上所有内容

param(
    [ValidateSet('Menu', 'Install', 'Uninstall', 'Trust', 'Build')]
    [string]$Action = 'Menu'
)

$ErrorActionPreference = 'Stop'
$Version    = '1.2.1'
$AddinName  = 'VecStamp'
$LegacyName = 'VectorExport'                       # v1.0「PPT 矢量导出加载项」
$Root       = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$DistPpam   = Join-Path $Root 'dist\VecStamp.ppam'
$BasFile    = Join-Path $Root 'dist\VecStamp.bas'
$Template   = Join-Path $Root 'dist\VecStamp-template.pptm'
$AddinDir   = Join-Path $env:APPDATA 'Microsoft\AddIns'
$Target     = Join-Path $AddinDir "$AddinName.ppam"
$OfficeKey  = 'HKCU:\Software\Microsoft\Office'
$SettingsKeys = @('HKCU:\Software\VecStamp', 'HKCU:\Software\PPTVectorExport')
$UiRelType  = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'

# ------------------------------------------------------------------ helpers
function Say([string]$Msg = '', [string]$Color = 'Gray') { Write-Host $Msg -ForegroundColor $Color }

function Ask-YesNo([string]$Question, [bool]$Default) {
    $hint = if ($Default) { '[Y/n]' } else { '[y/N]' }
    $ans = Read-Host "$Question $hint"
    if ([string]::IsNullOrWhiteSpace($ans)) { return $Default }
    return $ans.Trim().ToLower().StartsWith('y')
}

function Show-Banner {
    Clear-Host
    Say ''
    Say '   ════════════════════════════════════════════' 'DarkRed'
    Say '     矢印 VecStamp  ·  PowerPoint 矢量导出加载项' 'Red'
    Say "     版本 $Version  ·  Windows 安装器" 'Gray'
    Say '     Copyright (c) 2026 vluckyzhang  ·  MIT License' 'DarkGray'
    Say '   ════════════════════════════════════════════' 'DarkRed'
    Say ''
}

function Assert-PowerPointClosed {
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        throw '检测到 PowerPoint 正在运行。请保存并关闭所有 PowerPoint 窗口后再试。'
    }
}

function Get-OfficeVersions {
    $list = @()
    if (Test-Path $OfficeKey) {
        $list = @(Get-ChildItem $OfficeKey | Where-Object { $_.PSChildName -match '^\d+\.0$' -and [double]$_.PSChildName -ge 14 } |
                  ForEach-Object { $_.PSChildName })
    }
    if ($list -notcontains '16.0') { $list += '16.0' }
    return $list
}

function Get-PowerPointVersion {
    $installed = @(Get-OfficeVersions | Where-Object { Test-Path "$OfficeKey\$_\PowerPoint" } | Sort-Object { [double]$_ })
    if ($installed.Count -gt 0) { return $installed[-1] }
    return '16.0'
}

function Release-Com($Obj) {
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Obj) } catch { }
    }
}

function Wait-PowerPointExit {
    for ($i = 0; $i -lt 40; $i++) {
        if (-not (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 500
    }
}

function Load-Packaging {
    try { Add-Type -AssemblyName WindowsBase } catch { Add-Type -AssemblyName System.IO.Packaging }
}

function Copy-PartStream($From, $To) {
    $in = $From.GetStream([System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    $out = $To.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    try { $in.CopyTo($out) } finally { $in.Close(); $out.Close() }
}

# Makes sure the .ppam carries the ribbon; copies it from the template if PowerPoint dropped it.
function Ensure-Ribbon([string]$Ppam) {
    Load-Packaging
    $uiUri = New-Object System.Uri('/customUI/customUI14.xml', [System.UriKind]::Relative)
    $dst = [System.IO.Packaging.Package]::Open($Ppam, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        if ($dst.PartExists($uiUri)) { return }
        Say '   正在写入功能区界面…'
        $src = [System.IO.Packaging.Package]::Open($Template, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        try {
            $srcUi = $src.GetPart($uiUri)
            $newUi = $dst.CreatePart($uiUri, 'application/xml', [System.IO.Packaging.CompressionOption]::Normal)
            Copy-PartStream $srcUi $newUi
            foreach ($rel in @($srcUi.GetRelationships())) {
                $imgUri = [System.IO.Packaging.PackUriHelper]::ResolvePartUri($uiUri, $rel.TargetUri)
                if (-not $dst.PartExists($imgUri)) {
                    $img = $dst.CreatePart($imgUri, 'image/png', [System.IO.Packaging.CompressionOption]::NotCompressed)
                    Copy-PartStream ($src.GetPart($imgUri)) $img
                }
                [void]$newUi.CreateRelationship($rel.TargetUri, [System.IO.Packaging.TargetMode]::Internal,
                                                $rel.RelationshipType, $rel.Id)
            }
            foreach ($r in @($dst.GetRelationshipsByType($UiRelType))) { $dst.DeleteRelationship($r.Id) }
            $relTarget = New-Object System.Uri('customUI/customUI14.xml', [System.UriKind]::Relative)
            [void]$dst.CreateRelationship($relTarget, [System.IO.Packaging.TargetMode]::Internal, $UiRelType, 'rIdVecStampUI')
        }
        finally { $src.Close() }
    }
    finally { $dst.Close() }
}

function Show-ManualBuild {
    Say ''
    Say '手动生成 VecStamp.ppam（约 1 分钟）：' 'Cyan'
    Say "  1. 用 PowerPoint 打开 $Template"
    Say '  2. 按 Alt+F11 打开 VBA 编辑器，「文件 → 导入文件」选择 dist\VecStamp.bas'
    Say '  3. 点「调试 → 编译 VBAProject」，没有提示即代码正常；关闭 VBA 编辑器'
    Say '  4. 「文件 → 另存为」，类型选「PowerPoint 加载项 (*.ppam)」，保存为 dist\VecStamp.ppam'
    Say '  5. 关闭 PowerPoint，重新运行安装器，选择 [1] 安装'
}

# Builds dist\VecStamp.ppam from source with the local PowerPoint (asks for consent first).
function Build-Ppam {
    foreach ($f in @($BasFile, $Template)) {
        if (-not (Test-Path -LiteralPath $f)) { throw "缺少文件 $f ，请完整解压项目后再试。" }
    }
    Assert-PowerPointClosed
    Say '需要用本机 PowerPoint 从源码生成加载项文件 VecStamp.ppam。' 'Cyan'
    Say '这一步会临时开启 PowerPoint 的「信任对 VBA 工程对象模型的访问」，' 'Yellow'
    Say '以便把 dist\VecStamp.bas 导入模板；生成完成后（无论成功与否）立即恢复为原来的设置。' 'Yellow'
    if (-not (Ask-YesNo '是否继续？' $true)) {
        Show-ManualBuild
        return $false
    }

    $backup = @{}
    $versions = Get-OfficeVersions
    foreach ($v in $versions) {
        $k = "$OfficeKey\$v\PowerPoint\Security"
        if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
        $backup[$v] = (Get-ItemProperty -Path $k -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        New-ItemProperty -Path $k -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
    }

    $app = $null; $pres = $null; $comps = $null; $comp = $null
    $tmpOut = Join-Path $env:TEMP "VecStamp-build-$PID.ppam"
    try {
        Say '   正在启动 PowerPoint 并导入代码（约 10 秒）…'
        $app = New-Object -ComObject PowerPoint.Application
        $pres = $app.Presentations.Open($Template, -1, -1, 0)    # ReadOnly, Untitled, no window
        try {
            $comps = $pres.VBProject.VBComponents
            $comp = $comps.Import($BasFile)
        }
        catch {
            throw ("无法向 PowerPoint 写入 VBA 代码（$($_.Exception.Message)）。" +
                   '可能是电脑策略禁止了「信任对 VBA 工程对象模型的访问」，请改用手动生成（见使用说明）。')
        }
        if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force }
        try { $pres.SaveAs($tmpOut, 30) }                         # 30 = ppSaveAsOpenXMLAddin
        catch {
            throw ('PowerPoint 未能保存加载项，通常是 VBA 代码编译出错。请按手动生成的步骤操作，' +
                   '在第 3 步「调试 → 编译」时把提示和黄色高亮的代码行截图反馈给作者。')
        }
        $pres.Close()
    }
    finally {
        Release-Com $comp; Release-Com $comps; Release-Com $pres
        if ($null -ne $app) {
            try { $app.Quit() } catch { }
            Release-Com $app
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        foreach ($v in @($backup.Keys)) {
            $k = "$OfficeKey\$v\PowerPoint\Security"
            if ($null -eq $backup[$v]) { Remove-ItemProperty -Path $k -Name AccessVBOM -ErrorAction SilentlyContinue }
            else { New-ItemProperty -Path $k -Name AccessVBOM -Value $backup[$v] -PropertyType DWord -Force | Out-Null }
        }
        Say '   已恢复「信任对 VBA 工程对象模型的访问」的原设置。' 'DarkGray'
    }
    Wait-PowerPointExit
    if (-not (Test-Path -LiteralPath $tmpOut)) { throw '没有生成加载项文件。' }
    Ensure-Ribbon $tmpOut
    Copy-Item -LiteralPath $tmpOut -Destination $DistPpam -Force
    Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue
    Say "   已生成：$DistPpam" 'Green'
    return $true
}

function Remove-Registration([string[]]$Names) {
    if (-not (Test-Path $OfficeKey)) { return }
    foreach ($v in Get-ChildItem $OfficeKey | Where-Object { $_.PSChildName -match '^\d+\.0$' }) {
        foreach ($n in $Names) {
            foreach ($sub in @("PowerPoint\AddIns\$n", "PowerPoint\Security\Trusted Locations\$n")) {
                $k = Join-Path $v.PSPath $sub
                if (Test-Path -LiteralPath $k) { Remove-Item -LiteralPath $k -Recurse -Force }
            }
        }
    }
}

# ------------------------------------------------------------------ actions
function Do-Install {
    Assert-PowerPointClosed
    if (-not (Test-Path -LiteralPath $DistPpam)) {
        Say '未找到预先生成的 dist\VecStamp.ppam。'
        if (-not (Build-Ppam)) { return }
    }
    else {
        Say "使用已生成的加载项：$DistPpam"
        Ensure-Ribbon $DistPpam
    }

    $legacy = Join-Path $AddinDir "$LegacyName.ppam"
    if (Test-Path -LiteralPath $legacy) {
        Say '检测到旧版「PPT 矢量导出加载项 (VectorExport)」，其功能已全部包含在矢印中，将一并移除。' 'Yellow'
        Remove-Item -LiteralPath $legacy -Force
    }
    Remove-Registration @($LegacyName)

    New-Item -ItemType Directory -Path $AddinDir -Force | Out-Null
    Copy-Item -LiteralPath $DistPpam -Destination $Target -Force
    $ver = Get-PowerPointVersion
    $k = "$OfficeKey\$ver\PowerPoint\AddIns\$AddinName"
    New-Item -Path $k -Force | Out-Null
    New-ItemProperty -Path $k -Name 'Path' -Value $Target -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $k -Name 'AutoLoad' -Value -1 -PropertyType DWord -Force | Out-Null

    Say ''
    Say '安装完成！' 'Green'
    Say "  加载项：$Target"
    Say "  已登记为 PowerPoint（Office $ver）启动时自动加载。"
    Say '  打开 PowerPoint，在「视图」右侧找到「矢印」选项卡即可使用。'
    Say '  如果提示宏已被禁用，请重新运行本程序并选择 [3]。' 'DarkGray'
}

function Do-Uninstall {
    Assert-PowerPointClosed
    if (-not (Ask-YesNo '确定要卸载矢印 VecStamp 吗？' $false)) { return }
    Remove-Registration @($AddinName, $LegacyName)
    foreach ($f in @($Target, (Join-Path $AddinDir "$LegacyName.ppam"))) {
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force; Say "  已删除 $f" }
    }
    if (Ask-YesNo '同时删除导出设置（默认格式、文件夹等）吗？' $true) {
        foreach ($k in $SettingsKeys) { if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force } }
    }
    Say ''
    Say '卸载完成。' 'Green'
}

function Do-Trust {
    $ver = Get-PowerPointVersion
    Say '把加载项文件夹加入 PowerPoint 的「受信任位置」：' 'Cyan'
    Say "  $AddinDir\"
    Say '只有在打开 PowerPoint 后看到加载项的宏被禁用时才需要这样做。'
    Say '该文件夹默认只存放 Office 加载项；卸载时会自动移除这项设置。' 'DarkGray'
    if (-not (Ask-YesNo '是否添加？' $false)) { return }
    $tl = "$OfficeKey\$ver\PowerPoint\Security\Trusted Locations\$AddinName"
    New-Item -Path $tl -Force | Out-Null
    New-ItemProperty -Path $tl -Name 'Path' -Value ($AddinDir + '\') -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $tl -Name 'Description' -Value 'VecStamp PowerPoint add-in' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $tl -Name 'AllowSubfolders' -Value 0 -PropertyType DWord -Force | Out-Null
    Say '已添加。重新打开 PowerPoint 后生效。' 'Green'
}

function Do-Build {
    if (Build-Ppam) {
        Say '可以把 dist\VecStamp.ppam 发布到 Release，供 macOS 用户和其他电脑直接安装。' 'DarkGray'
    }
}

function Run-Action([string]$Name) {
    try {
        switch ($Name) {
            'Install'   { Do-Install }
            'Uninstall' { Do-Uninstall }
            'Trust'     { Do-Trust }
            'Build'     { Do-Build }
        }
        return $true
    }
    catch {
        Say ''
        Say "出错了：$($_.Exception.Message)" 'Red'
        return $false
    }
}

# ------------------------------------------------------------------ entry
if ($Action -ne 'Menu') {
    Show-Banner
    if (Run-Action $Action) { exit 0 } else { exit 1 }
}

while ($true) {
    Show-Banner
    if (Test-Path -LiteralPath $Target) { Say "   当前状态：已安装（$Target）" 'Green' }
    else { Say '   当前状态：未安装' 'Yellow' }
    Say ''
    Say '   [1] 安装 / 更新'
    Say '   [2] 卸载'
    Say '   [3] 修复：宏被禁用时，把加载项文件夹设为受信任位置'
    Say '   [4] 仅从源码生成 dist\VecStamp.ppam（开发者 / 给 Mac 用）'
    Say '   [0] 退出'
    Say ''
    $choice = Read-Host '   请输入数字后回车'
    $name = switch ($choice.Trim()) { '1' { 'Install' } '2' { 'Uninstall' } '3' { 'Trust' } '4' { 'Build' } '0' { 'Exit' } default { '' } }
    if ($name -eq 'Exit') { break }
    if ($name -eq '') { continue }
    Say ''
    [void](Run-Action $name)
    Say ''
    [void](Read-Host '按回车键返回菜单')
}
