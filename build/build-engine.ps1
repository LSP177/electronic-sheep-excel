param(
    [string]$BaselineDir = (Join-Path (Split-Path $PSScriptRoot -Parent) 'assets/baseline'),
    [string]$PythonExe = ''
)
$ErrorActionPreference = 'Stop'
$product = Split-Path $PSScriptRoot -Parent
if (-not (Test-Path -LiteralPath $BaselineDir)) {
    throw "Baseline directory not found: $BaselineDir. Pass -BaselineDir with DataHelper_PocEngine.xlam and TaskPaneHost.xlsx seed files."
}
$work = Join-Path $env:TEMP ('es_excel_engine_' + [Guid]::NewGuid().ToString('N'))
$out = Join-Path $product 'build/output/Core'
$importFolder = Join-Path $env:TEMP ('DH1Import_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work, $out, $importFolder | Out-Null
. (Join-Path $PSScriptRoot 'prepare-import.ps1')
Add-Type -Path (Join-Path $PSScriptRoot 'RunningWorkbook.cs')
Add-Type 'public static class DH1Native { [System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hwnd,out uint pid); }'

$engine = Join-Path $work 'DataHelper_PocEngine.xlam'
$hostFile = Join-Path $work 'BuildHost.xlsx'
Copy-Item (Join-Path $BaselineDir 'DataHelper_PocEngine.xlam') $engine
Copy-Item (Join-Path $BaselineDir 'TaskPaneHost.xlsx') $hostFile

$key = 'HKCU:/Software/Microsoft/Office/16.0/Excel/Security'
$prior = Get-ItemProperty $key -ErrorAction SilentlyContinue
$had = $prior -and $null -ne $prior.PSObject.Properties['AccessVBOM']
$old = if ($had) { $prior.AccessVBOM } else { $null }
$mutex = New-Object Threading.Mutex($false, 'Local\ExcelDataHelper_VBABuild')
$owned = $false; $changed = $false; $excel = $null; $process = $null

try {
    try { $owned = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $owned = $true }
    if (-not $owned) { throw 'Another VBA build is active' }
    if ($had -and $old -ne 0) { throw 'AccessVBOM already active; do not adopt another build setting' }
    if (-not (Test-Path $key)) { New-Item $key -Force | Out-Null }
    New-ItemProperty $key AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null; $changed = $true

    $excelPath = 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'
    if (-not (Test-Path $excelPath)) { $excelPath = 'C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE' }
    $process = Start-Process $excelPath -ArgumentList @('/x', '/automation', ('"' + $hostFile + '"')) -WindowStyle Hidden -PassThru
    $deadline = (Get-Date).AddSeconds(45)
    do {
        Start-Sleep -Milliseconds 300; $hostBook = [RunningWriteWorkbook]::Find($hostFile)
        if ($hostBook) { try { $candidate = $hostBook.Application; if ($candidate -and $candidate.Workbooks -and $candidate.Hwnd) { $excel = $candidate } } catch {} }
    } while (-not $excel -and (Get-Date) -lt $deadline)
    if (-not $excel) { throw 'Owned Excel did not become ready' }

    $actual = 0
    [void][DH1Native]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$actual)
    if ($actual -ne $process.Id) { $excel = $null; throw 'Excel PID mismatch' }
    $excel.DisplayAlerts = $false; $excel.EnableEvents = $false

    $loaded = $null; try { $loaded = $excel.Workbooks.Item('DataHelper_PocEngine.xlam') } catch {}
    if ($loaded) {
        if ([string]::Equals($loaded.FullName, $engine, [StringComparison]::OrdinalIgnoreCase)) { throw 'Build engine already open' }
        $loaded.Close($false); $loaded = $null
    }

    $book = $excel.Workbooks.Open($engine, 0, $false)
    if (-not $book -or -not [string]::Equals($book.FullName, $engine, [StringComparison]::OrdinalIgnoreCase)) { throw 'Engine seed did not open in build Excel' }

    $files = @(Get-ChildItem (Join-Path $product 'core/business') -File) + @(Get-ChildItem (Join-Path $product 'core/runtime') -File) + @(Get-Item (Join-Path $product 'frontends/legacy/DHLegacyRibbon.bas'))
    foreach ($file in $files) {
        $name = $file.BaseName; $component = $null
        try { $component = $book.VBProject.VBComponents.Item($name) } catch {}
        if ($component) { $book.VBProject.VBComponents.Remove($component) }
        $import = Join-Path $importFolder $file.Name
        [IO.File]::WriteAllText($import, (ConvertTo-AsciiVba ([IO.File]::ReadAllText($file.FullName))), [Text.Encoding]::ASCII)
        $added = $book.VBProject.VBComponents.Import($import)
        if ($added.Name -ne $name) { throw ('Unexpected module import rename: ' + $name) }
    }

    $book.Save()
    $health = $excel.Run("'DataHelper_PocEngine.xlam'!DH1_Health")
    if ($health -ne 'DataHelper/1.0.0/shared-engine') { throw 'Shared engine compile/health failed' }
    $modules = @(); foreach ($component in $book.VBProject.VBComponents) { $modules += $component.Name }
    $book.Close($false)

    Add-Type -AssemblyName System.IO.Compression.FileSystem, System.IO.Compression
    $zip = [IO.Compression.ZipFile]::Open($engine, [IO.Compression.ZipArchiveMode]::Update)
    try {
        if ($zip.GetEntry('customUI/customUI.xml')) { throw 'Office 2007 customUI is forbidden' }
        $entry = $zip.GetEntry('customUI/customUI14.xml'); if (-not $entry) { throw 'Seed lacks existing Ribbon relationship' }
        $entry.Delete(); $entry = $zip.CreateEntry('customUI/customUI14.xml')
        $writer = New-Object IO.StreamWriter($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write([IO.File]::ReadAllText((Join-Path $product 'frontends/legacy/customUI14.xml'))) } finally { $writer.Dispose() }

        $metadata = $zip.GetEntry('docProps/core.xml')
        if ($metadata) {
            $reader = New-Object IO.StreamReader($metadata.Open()); try { $properties = New-Object Xml.XmlDocument; $properties.LoadXml($reader.ReadToEnd()) } finally { $reader.Dispose() }
            $namespace = New-Object Xml.XmlNamespaceManager($properties.NameTable); $namespace.AddNamespace('dc', 'http://purl.org/dc/elements/1.1/')
            $title = $properties.SelectSingleNode('//dc:title', $namespace)
            if (-not $title) { $title = $properties.CreateElement('dc', 'title', 'http://purl.org/dc/elements/1.1/'); [void]$properties.DocumentElement.AppendChild($title) }
            $title.InnerText = '电子羊表格'; $metadata.Delete(); $metadata = $zip.CreateEntry('docProps/core.xml')
            $writer = New-Object IO.StreamWriter($metadata.Open(), [Text.UTF8Encoding]::new($false)); try { $writer.Write($properties.OuterXml) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose() }

    $pyCmd = $PythonExe
    if ([string]::IsNullOrWhiteSpace($pyCmd)) {
        $found = Get-Command python -ErrorAction SilentlyContinue
        if ($found) { $pyCmd = $found.Source }
    }
    if (-not [string]::IsNullOrWhiteSpace($pyCmd)) {
        & $pyCmd (Join-Path $PSScriptRoot 'sanitize-core.py') $engine
        if ($LASTEXITCODE -ne 0) { throw 'Core metadata sanitization failed' }
    }

    Copy-Item $engine $out -Force
    [ordered]@{
        Health = $health
        Modules = $modules
        CoreSHA256 = (Get-FileHash $engine).Hash
        Time = (Get-Date -Format o)
    } | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $out 'build-evidence.json') -Encoding UTF8
} finally {
    if ($excel) { try { $excel.DisplayAlerts = $false; $excel.Quit() } catch {} }
    $book = $null; $hostBook = $null; $excel = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($process -and -not $process.HasExited -and -not $process.WaitForExit(5000)) {
        $info = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $process.Id)
        if ($info -and $info.CommandLine -and $info.CommandLine.Contains($hostFile)) { Stop-Process -Id $process.Id -Force }
    }
    if ($changed) { if ($had) { Set-ItemProperty $key AccessVBOM $old } else { Remove-ItemProperty $key AccessVBOM } }
    if ($owned) { $mutex.ReleaseMutex() }; $mutex.Dispose()
}
