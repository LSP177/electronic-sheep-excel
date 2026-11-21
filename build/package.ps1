param(
    [string]$OutDir = (Join-Path (Split-Path $PSScriptRoot -Parent) ('dist/1.0-oss-' + (Get-Date -Format 'yyyyMMdd-HHmmss')))
)
$ErrorActionPreference = 'Stop'
$product = Split-Path $PSScriptRoot -Parent
$dist = [IO.Path]::GetFullPath($OutDir)
if (Test-Path -LiteralPath $dist) { throw "Output directory already exists: $dist" }

$coreBuild = Get-Content (Join-Path $product 'build/output/Core/build-evidence.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ((Get-FileHash (Join-Path $product 'build/output/Core/DataHelper_PocEngine.xlam')).Hash -ne $coreBuild.CoreSHA256) {
    throw 'Core binary changed since build'
}

$stageRoot = Join-Path $env:TEMP ('es_excel_pkg_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Force $dist, $stageRoot | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression
$csc = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'

$hostBuild = Get-Content (Join-Path $product 'build/output/TaskPane/build-evidence.json') -Raw | ConvertFrom-Json
foreach ($entry in $hostBuild.Sources) {
    if ((Get-FileHash (Join-Path $product $entry.Path)).Hash -ne $entry.SHA256) { throw ('Host build is stale: ' + $entry.Path) }
}
foreach ($entry in $hostBuild.Outputs) {
    if ((Get-FileHash (Join-Path $product ('build/output/TaskPane/' + $entry.Path))).Hash -ne $entry.SHA256) { throw ('Host output mismatch: ' + $entry.Path) }
}

foreach ($edition in @('legacy', 'taskpane')) {
    $stage = Join-Path $stageRoot $edition
    New-Item -ItemType Directory -Force (Join-Path $stage 'Core') | Out-Null
    Copy-Item (Join-Path $product 'build/output/Core/DataHelper_PocEngine.xlam') (Join-Path $stage 'Core')
    Copy-Item (Join-Path $product 'version.json') (Join-Path $stage 'product.json')
    Copy-Item (Join-Path $PSScriptRoot 'RunningWorkbook.cs') $stage
    Copy-Item -LiteralPath (Join-Path $product 'installer/电子羊表格.ico') (Join-Path $stage 'app.ico')
    $script = [IO.File]::ReadAllText((Join-Path $product 'installer/install.ps1'))
    [IO.File]::WriteAllText((Join-Path $stage 'install.ps1'), $script, [Text.UTF8Encoding]::new($true))

    $template = Join-Path $stage 'Start.xlsx'
    $zip = [IO.Compression.ZipFile]::Open($template, [IO.Compression.ZipArchiveMode]::Create)
    $xml = @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="开始使用" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
    }
    try {
        foreach ($name in $xml.Keys) {
            $entry = $zip.CreateEntry($name)
            $writer = New-Object IO.StreamWriter($entry.Open(), [Text.UTF8Encoding]::new($false))
            try { $writer.Write($xml[$name]) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }

    if ($edition -eq 'taskpane') {
        $files = @('DataHelperTaskPanePoc.dll', 'DataHelperTaskPanePoc.dll.manifest', 'DataHelperTaskPanePoc.vsto', 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'WebView2Loader.dll', 'index.html', 'bridge-client.js', 'poc-signing.cer')
        New-Item -ItemType Directory -Force (Join-Path $stage 'TaskPane') | Out-Null
        foreach ($name in $files) {
            $source = Join-Path $product ('build/output/TaskPane/' + $name)
            if (-not (Test-Path $source -PathType Leaf)) { throw ('Required runtime asset missing: ' + $name) }
            Copy-Item $source (Join-Path $stage 'TaskPane')
        }
    }

    $hashes = @(Get-ChildItem $stage -Recurse -File | ForEach-Object { @{ Path = $_.FullName.Substring($stage.Length + 1); SHA256 = (Get-FileHash $_.FullName).Hash } })
    $hashes | ConvertTo-Json | Set-Content (Join-Path $stage 'payload-hashes.json') -Encoding UTF8
    $payload = Join-Path $stageRoot ($edition + '.zip')
    [IO.Compression.ZipFile]::CreateFromDirectory($stage, $payload)
    $hash = (Get-FileHash $payload).Hash
    $info = Join-Path $stageRoot ($edition + '.cs')
    [IO.File]::WriteAllText($info, ('internal static class BuildInfo { internal const string Edition="' + $edition + '"; internal const string Hash="' + $hash + '"; }'))

    $name = if ($edition -eq 'legacy') { '电子羊表格_1.0_兼容版_安装.exe' } else { '电子羊表格_1.0_现代版_安装.exe' }
    $out = Join-Path $dist $name
    & $csc /nologo /target:winexe /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll ('/resource:' + $payload + ',payload.zip') ('/win32icon:' + (Join-Path $product 'installer/电子羊表格.ico')) ('/out:' + $out) (Join-Path $product 'installer/Setup.cs') $info
    if ($LASTEXITCODE -ne 0) { throw 'Bootstrap compilation failed' }
}
Write-Output "Packaged successfully into: $dist"
