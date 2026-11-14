# Pravi ikonice za LibreOffice (libreoffice/icons) i MS Office (msoffice/ribbon/images).
#   pwsh tools/make_icons.ps1
Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root "libreoffice\icons"
New-Item -ItemType Directory -Force $outDir | Out-Null

$blue = [System.Drawing.Color]::FromArgb(255, 31, 95, 191)
$green = [System.Drawing.Color]::FromArgb(255, 34, 128, 62)
$orange = [System.Drawing.Color]::FromArgb(255, 204, 102, 0)
$gray = [System.Drawing.Color]::FromArgb(255, 110, 110, 110)

# naziv, mali znak (izvor), veliki znak (cilj), boja
$icons = @(
    @("cyr", "L", "Ћ", $blue),
    @("lat", "Л", "Ć", $green),
    @("yulat", "}", "Ć", $orange),
    @("yucyr", "}", "Ћ", $orange)
)

function Draw-Icon($path, $size, $small, $big, $color) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = "AntiAlias"
    $g.TextRenderingHint = "AntiAliasGridFit"
    $g.Clear([System.Drawing.Color]::Transparent)
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = "Center"; $fmt.LineAlignment = "Center"
    $fmt.FormatFlags = "NoClip"

    $bigFont = New-Object System.Drawing.Font("Segoe UI", ($size * 0.78), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush $color
    $rect = New-Object System.Drawing.RectangleF ($size * 0.18), ($size * 0.12), ($size * 0.86), ($size * 0.94)
    $g.DrawString($big, $bigFont, $brush, $rect, $fmt)

    if ($small) {
        $smallFont = New-Object System.Drawing.Font("Segoe UI", ($size * 0.42), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $gbrush = New-Object System.Drawing.SolidBrush $gray
        $srect = New-Object System.Drawing.RectangleF (-$size * 0.04), (-$size * 0.06), ($size * 0.48), ($size * 0.5)
        $g.DrawString($small, $smallFont, $gbrush, $srect, $fmt)
    }
    $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

foreach ($i in $icons) {
    foreach ($s in 16, 26, 32) {
        Draw-Icon (Join-Path $outDir ("{0}_{1}.png" -f $i[0], $s)) $s $i[1] $i[2] $i[3]
    }
}
# podešavanja: zupčanik
foreach ($s in 16, 26, 32) {
    Draw-Icon (Join-Path $outDir ("dialog_{0}.png" -f $s)) $s "" ([string][char]0x2699) $gray
}
# ikonica ekstenzije
Draw-Icon (Join-Path $outDir "translit_64.png") 64 "L" "Ћ" $blue
Write-Output "Ikonice su u $outDir"

# MS Office kartica koristi iste ikonice (32 px)
$msDir = Join-Path $root "msoffice\ribbon\images"
New-Item -ItemType Directory -Force $msDir | Out-Null
foreach ($n in "cyr", "lat", "yulat", "yucyr") { Copy-Item (Join-Path $outDir "${n}_32.png") (Join-Path $msDir "$n.png") -Force }
Copy-Item (Join-Path $outDir "dialog_32.png") (Join-Path $msDir "exceptions.png") -Force
Write-Output "Ikonice za MS Office su u $msDir"
