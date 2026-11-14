# Provera msoffice/install.ps1 bez Office-a:
#  - skript se parsira bez sintaksnih grešaka
#  - Add-Ribbon ispravno ubacuje karticu i ikonice u Office Open XML paket (i drugi put ništa ne duplira)
#
#   powershell -ExecutionPolicy Bypass -File tests/install_test.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$script = Join-Path $root 'msoffice\install.ps1'
$fail = 0

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$tokens, [ref]$errors)
if ($errors.Count) { $errors | ForEach-Object { Write-Host "GREŠKA sintaksa: $($_.Message) (red $($_.Extent.StartLineNumber))" }; $fail++ }
else { Write-Host 'ok     install.ps1 se parsira' }

# Add-Ribbon iz skripta, bez pokretanja instalacije
$fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Add-Ribbon' }, $true) | Select-Object -First 1
$ribbonDir = Join-Path $root 'msoffice\ribbon'
. ([scriptblock]::Create($fn.Extent.Text))

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$tmp = Join-Path $env:TEMP ('translit_ribbon_test_' + [guid]::NewGuid().ToString('N') + '.docx')
$zip = [IO.Compression.ZipFile]::Open($tmp, 'Create')
function Add-Text($name, $text) { $s = New-Object IO.StreamWriter($zip.CreateEntry($name).Open()); $s.Write($text); $s.Close() }
Add-Text '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
Add-Text '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
Add-Text 'word/document.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Proba</w:t></w:r></w:p></w:body></w:document>'
$zip.Dispose()

Add-Ribbon $tmp
Add-Ribbon $tmp   # drugi put ne sme ništa da duplira

$zip = [IO.Compression.ZipFile]::OpenRead($tmp)
function Read-Entry($name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $r = New-Object IO.StreamReader($e.Open()); $t = $r.ReadToEnd(); $r.Close(); $t }
$names = $zip.Entries | ForEach-Object FullName
$checks = @(
    @('customUI14.xml u paketu', ($names -contains 'customUI/customUI14.xml')),
    @('ikonice u paketu', (@($names | Where-Object { $_ -like 'customUI/images/*.png' }).Count -eq 5)),
    @('veza ka kartici u _rels/.rels (jednom)', ([regex]::Matches((Read-Entry '_rels/.rels'), 'ui/extensibility').Count -eq 1)),
    @('png u [Content_Types].xml (jednom)', ([regex]::Matches((Read-Entry '[Content_Types].xml'), 'Extension="png"').Count -eq 1)),
    @('veze ka ikonicama', ([regex]::Matches((Read-Entry 'customUI/_rels/customUI14.xml.rels'), 'Target="images/').Count -eq 5)),
    @('nema duplih unosa', (($names | Group-Object | Where-Object Count -gt 1).Count -eq 0))
)
foreach ($c in $checks) { if ($c[1]) { Write-Host "ok     $($c[0])" } else { Write-Host "GREŠKA $($c[0])"; $fail++ } }

# svaka ikonica iz customUI14.xml mora imati vezu
$ui = [xml](Read-Entry 'customUI/customUI14.xml')
$rels = [xml](Read-Entry 'customUI/_rels/customUI14.xml.rels')
$ids = $rels.Relationships.Relationship | ForEach-Object Id
$images = Select-Xml -Xml $ui -XPath '//*[@image]' | ForEach-Object { $_.Node.image }
$missing = $images | Where-Object { $ids -notcontains $_ }
if ($missing) { Write-Host "GREŠKA ikonice bez veze: $missing"; $fail++ } else { Write-Host "ok     sve ikonice kartice imaju vezu ($($images.Count))" }
$zip.Dispose()
Remove-Item $tmp

if ($fail) { Write-Host "`nNEUSPELO: $fail"; exit 1 } else { Write-Host "`nSVE U REDU" }
