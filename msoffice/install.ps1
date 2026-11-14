<#
.SYNOPSIS
    Instalira (ili uklanja) Translit za Microsoft Word, Excel i PowerPoint.

.DESCRIPTION
    Pravi dodatke na ovom računaru, iz izvornog koda u folderu src:
      Word        Translit.dotm  ->  Word STARTUP folder (učitava se sam)
      Excel       Translit.xlam  ->  %APPDATA%\Microsoft\AddIns (uključuje se kao dodatak)
      PowerPoint  Translit.ppam  ->  %APPDATA%\Microsoft\AddIns (uključuje se kao dodatak)
    i u svaki ubacuje karticu "Translit" (ribbon\customUI14.xml).

    Da bi uvezao VBA kod, skript privremeno uključuje opciju
    "Trust access to the VBA project object model" i posle je vraća kako je bila.

    Word, Excel i PowerPoint moraju biti zatvoreni.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall
#>
param(
    [switch]$Uninstall,
    [ValidateSet('Word', 'Excel', 'PowerPoint')]
    [string[]]$Apps = @('Word', 'Excel', 'PowerPoint')
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir = Join-Path $here 'src'
$ribbonDir = Join-Path $here 'ribbon'
$addinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'

$modules = @{
    Word       = @('TranslitCore.bas', 'TranslitNative.bas', 'TranslitSettings.bas', 'TranslitWord.bas')
    Excel      = @('TranslitCore.bas', 'TranslitNative.bas', 'TranslitSettings.bas', 'TranslitShapes.bas', 'TranslitExcel.bas')
    PowerPoint = @('TranslitCore.bas', 'TranslitNative.bas', 'TranslitSettings.bas', 'TranslitShapes.bas', 'TranslitPowerPoint.bas')
}
$progIds = @{ Word = 'Word.Application'; Excel = 'Excel.Application'; PowerPoint = 'PowerPoint.Application' }
$processes = @{ Word = 'WINWORD'; Excel = 'EXCEL'; PowerPoint = 'POWERPNT' }

function Say([string]$text, [string]$color = 'Gray') { Write-Host $text -ForegroundColor $color }

function Release($obj) {
    if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }
}

function New-OfficeApp([string]$app) {
    try { return New-Object -ComObject $progIds[$app] }
    catch { return $null }
}

function Get-OfficeVersions([string]$app) {
    Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d$' -and (Test-Path (Join-Path $_.PSPath $app)) } |
        ForEach-Object { $_.PSChildName }
}

# ---------------------------------------------------------------- pristup VBA projektu

$vbomBackup = @()

function Enable-VbaAccess([string]$app, [string]$version) {
    $policy = "HKCU:\Software\Policies\Microsoft\Office\$version\$app\Security"
    $p = Get-ItemProperty $policy -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($p -and $p.AccessVBOM -eq 0) {
        throw "Pristup VBA projektu je zabranjen politikom (Group Policy) za $app. Obratite se administratoru."
    }
    $key = "HKCU:\Software\Microsoft\Office\$version\$app\Security"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    $old = Get-ItemProperty $key -Name AccessVBOM -ErrorAction SilentlyContinue
    $script:vbomBackup += [pscustomobject]@{ Key = $key; Had = [bool]$old; Value = $(if ($old) { $old.AccessVBOM } else { $null }) }
    Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
}

# Pokreće aplikaciju, uključi pristup VBA projektu za njenu verziju i pokrene je ponovo
# (Office to podešavanje čita pri pokretanju).
function Start-AppForVba([string]$app) {
    $o = New-OfficeApp $app
    if (-not $o) { return $null }
    $version = $o.Version
    Enable-VbaAccess $app $version
    try { $o.Quit() } catch {}
    Release $o
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 800
    return New-OfficeApp $app
}

function Restore-VbaAccess {
    foreach ($b in $script:vbomBackup) {
        if ($b.Had) { Set-ItemProperty -Path $b.Key -Name AccessVBOM -Value $b.Value -Type DWord }
        else { Remove-ItemProperty -Path $b.Key -Name AccessVBOM -ErrorAction SilentlyContinue }
    }
    $script:vbomBackup = @()
}

# ---------------------------------------------------------------- pravljenje dodatka

function Import-Modules($vbProject, [string]$app, [string]$tempDir) {
    foreach ($m in $modules[$app]) {
        # VBA uvozi samo ASCII sa CRLF krajevima redova
        $text = [IO.File]::ReadAllText((Join-Path $srcDir $m), [Text.Encoding]::ASCII)
        $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
        $tmp = Join-Path $tempDir $m
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        [void]$vbProject.VBComponents.Import($tmp)
    }
}

function Get-VbProject($container, [string]$app) {
    try {
        $p = $container.VBProject
        [void]$p.VBComponents.Count
        return $p
    }
    catch {
        throw "$app ne dozvoljava pristup VBA projektu. U $app-u uključite: File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', pa ponovo pokrenite instalaciju (posle je možete isključiti)."
    }
}

# Ubacuje karticu (customUI14.xml + ikonice) u Office Open XML paket
function Add-Ribbon([string]$path) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $utf8 = New-Object Text.UTF8Encoding($false)
    $zip = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $put = {
            param([string]$name, [byte[]]$bytes)
            $old = $zip.GetEntry($name)
            if ($old) { $old.Delete() }
            $s = $zip.CreateEntry($name).Open()
            $s.Write($bytes, 0, $bytes.Length)
            $s.Close()
        }
        $read = {
            param([string]$name)
            $r = New-Object IO.StreamReader($zip.GetEntry($name).Open(), $utf8)
            $t = $r.ReadToEnd()
            $r.Close()
            $t
        }
        & $put 'customUI/customUI14.xml' ([IO.File]::ReadAllBytes((Join-Path $ribbonDir 'customUI14.xml')))
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        foreach ($img in Get-ChildItem (Join-Path $ribbonDir 'images') -Filter *.png) {
            & $put "customUI/images/$($img.Name)" ([IO.File]::ReadAllBytes($img.FullName))
            $rels += '<Relationship Id="' + $img.BaseName + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/' + $img.Name + '"/>'
        }
        $rels += '</Relationships>'
        & $put 'customUI/_rels/customUI14.xml.rels' $utf8.GetBytes($rels)

        $root = & $read '_rels/.rels'
        if ($root -notmatch 'customUI14\.xml') {
            $root = $root -replace '</Relationships>', '<Relationship Id="rIdTranslitUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
            & $put '_rels/.rels' $utf8.GetBytes($root)
        }
        $types = & $read '[Content_Types].xml'
        if ($types -notmatch 'Extension="png"') {
            $types = $types -replace '<Types([^>]*)>', '<Types$1><Default Extension="png" ContentType="image/png"/>'
            & $put '[Content_Types].xml' $utf8.GetBytes($types)
        }
    }
    finally { $zip.Dispose() }
}

function Build-Word([string]$tempDir) {
    $word = Start-AppForVba 'Word'
    if (-not $word) { Say '  Word nije instaliran - preskačem.' 'Yellow'; return }
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()
        Import-Modules (Get-VbProject $doc 'Word') 'Word' $tempDir
        $out = Join-Path $tempDir 'Translit.dotm'
        $doc.SaveAs2($out, 15)   # wdFormatXMLTemplateMacroEnabled
        $doc.Close(0)
        Release $doc
        $startup = $word.StartupPath
        if (-not $startup) { $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP' }
    }
    finally { try { $word.Quit(0) } catch {} ; Release $word }
    Add-Ribbon $out
    if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
    Copy-Item $out (Join-Path $startup 'Translit.dotm') -Force
    Say "  Word: $(Join-Path $startup 'Translit.dotm')" 'Green'
}

function Build-Excel([string]$tempDir) {
    $xl = Start-AppForVba 'Excel'
    if (-not $xl) { Say '  Excel nije instaliran - preskačem.' 'Yellow'; return }
    $dest = Join-Path $addinsDir 'Translit.xlam'
    try {
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        Import-Modules (Get-VbProject $wb 'Excel') 'Excel' $tempDir
        $out = Join-Path $tempDir 'Translit.xlam'
        $wb.SaveAs($out, 55)   # xlOpenXMLAddIn
        $wb.Close($false)
        Release $wb
    }
    finally { try { $xl.Quit() } catch {} ; Release $xl }
    Add-Ribbon $out
    if (-not (Test-Path $addinsDir)) { New-Item -ItemType Directory -Path $addinsDir -Force | Out-Null }
    Copy-Item $out $dest -Force

    # uključi dodatak (Excel to upisuje u registar pri zatvaranju)
    $xl = New-OfficeApp 'Excel'
    try {
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $ai = $xl.AddIns.Add($dest, $false)
        $ai.Installed = $true
        $wb.Close($false)
        Release $ai; Release $wb
        $version = $xl.Version
    }
    finally { try { $xl.Quit() } catch {} ; Release $xl }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Register-ExcelAddin $version $dest
    Say "  Excel: $dest" 'Green'
}

# Ako Excel nije zapamtio dodatak, upiši ga sam (Options\OPEN, OPEN1, ...)
function Register-ExcelAddin([string]$version, [string]$dest) {
    $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    $props = Get-ItemProperty $key
    $open = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' }
    if ($open | Where-Object { $_.Value -like '*Translit.xlam*' }) { return }
    $i = 0
    while ($props.PSObject.Properties.Name -contains $(if ($i -eq 0) { 'OPEN' } else { "OPEN$i" })) { $i++ }
    $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
    Set-ItemProperty -Path $key -Name $name -Value ('"' + $dest + '"')
}

function Build-PowerPoint([string]$tempDir) {
    $pp = Start-AppForVba 'PowerPoint'
    if (-not $pp) { Say '  PowerPoint nije instaliran - preskačem.' 'Yellow'; return }
    $dest = Join-Path $addinsDir 'Translit.ppam'
    try {
        $version = $pp.Version
        $pres = $pp.Presentations.Add(0)   # bez prozora
        Import-Modules (Get-VbProject $pres 'PowerPoint') 'PowerPoint' $tempDir
        $out = Join-Path $tempDir 'Translit.ppam'
        $pres.SaveAs($out, 30)   # ppSaveAsOpenXMLAddin
        $pres.Close()
        Release $pres
    }
    finally { try { $pp.Quit() } catch {} ; Release $pp }
    Add-Ribbon $out
    if (-not (Test-Path $addinsDir)) { New-Item -ItemType Directory -Path $addinsDir -Force | Out-Null }
    Copy-Item $out $dest -Force
    $key = "HKCU:\Software\Microsoft\Office\$version\PowerPoint\AddIns\Translit"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    Set-ItemProperty -Path $key -Name Path -Value $dest
    Set-ItemProperty -Path $key -Name AutoLoad -Value 1 -Type DWord
    Say "  PowerPoint: $dest" 'Green'
}

# ---------------------------------------------------------------- uklanjanje

function Remove-Translit {
    $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\Translit.dotm'
    if ('Word' -in $Apps -and (Test-Path $startup)) { Remove-Item $startup -Force; Say "  obrisan $startup" }
    if ('Excel' -in $Apps) {
        foreach ($v in Get-OfficeVersions 'Excel') {
            $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
            if (Test-Path $key) {
                (Get-ItemProperty $key).PSObject.Properties |
                    Where-Object { $_.Name -match '^OPEN\d*$' -and $_.Value -like '*Translit.xlam*' } |
                    ForEach-Object { Remove-ItemProperty -Path $key -Name $_.Name; Say "  uklonjen Excel $v $($_.Name)" }
            }
        }
        $f = Join-Path $addinsDir 'Translit.xlam'
        if (Test-Path $f) { Remove-Item $f -Force; Say "  obrisan $f" }
    }
    if ('PowerPoint' -in $Apps) {
        foreach ($v in Get-OfficeVersions 'PowerPoint') {
            $key = "HKCU:\Software\Microsoft\Office\$v\PowerPoint\AddIns\Translit"
            if (Test-Path $key) { Remove-Item $key -Recurse -Force; Say "  uklonjen PowerPoint $v dodatak" }
        }
        $f = Join-Path $addinsDir 'Translit.ppam'
        if (Test-Path $f) { Remove-Item $f -Force; Say "  obrisan $f" }
    }
    Say 'Opcije (HKCU\Software\VB and VBA Program Settings\Translit) i spisak izuzetaka (%APPDATA%\Translit) su ostali.'
}

# ---------------------------------------------------------------- glavni deo

$running = foreach ($a in $Apps) { Get-Process $processes[$a] -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object { $a } }
if ($running) {
    Say "Zatvorite pa pokrenite ponovo: $($running -join ', ')." 'Red'
    exit 1
}

if ($Uninstall) {
    Say 'Uklanjam Translit...' 'Cyan'
    Remove-Translit
    Say 'Gotovo.' 'Green'
    exit 0
}

$tempDir = Join-Path $env:TEMP ('translit_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$failed = $false
Say 'Instaliram Translit...' 'Cyan'
try {
    foreach ($a in $Apps) {
        try {
            switch ($a) {
                'Word' { Build-Word $tempDir }
                'Excel' { Build-Excel $tempDir }
                'PowerPoint' { Build-PowerPoint $tempDir }
            }
        }
        catch {
            $failed = $true
            Say "  ${a}: $($_.Exception.Message)" 'Red'
        }
    }
}
finally {
    Restore-VbaAccess
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failed) { Say 'Instalacija nije potpuno uspela (vidi poruke iznad).' 'Red'; exit 1 }
Say 'Gotovo. Otvorite Word, Excel ili PowerPoint - kartica "Translit" je na traci.' 'Green'
