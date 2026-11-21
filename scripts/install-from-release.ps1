param(
    [ValidateSet('auto', 'modern', 'legacy')][string]$Edition = 'auto',
    [string]$Tag = 'latest',
    [switch]$Launch
)
$ErrorActionPreference = 'Stop'

$repo = 'LSP177/electronic-sheep-excel'
$baseUrl = if ($Tag -eq 'latest') {
    "https://github.com/$repo/releases/latest/download"
} else {
    "https://github.com/$repo/releases/download/$Tag"
}

function Get-ExcelBitness {
    $views = if ([Environment]::Is64BitOperatingSystem) {
        @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    } else {
        @([Microsoft.Win32.RegistryView]::Registry32)
    }
    foreach ($view in $views) {
        $base = $null; $key = $null
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, $view)
            $key = $base.OpenSubKey('CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32', $false)
            if (-not $key) { continue }
            $server = [string]$key.GetValue('')
            $m = [regex]::Match($server, '^\s*(?:"(?<exe>[^"]+\.exe)"|(?<exe>.+?\.exe))(?:\s|$)', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $m.Success) { continue }
            $exe = [Environment]::ExpandEnvironmentVariables($m.Groups['exe'].Value.Trim())
            if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { continue }
            $fs = [IO.File]::Open($exe, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            $br = New-Object IO.BinaryReader($fs)
            try {
                $fs.Position = 60
                $pe = $br.ReadInt32()
                $fs.Position = $pe + 4
                $machine = $br.ReadUInt16()
                return @{ Exe = $exe; Machine = $machine; Is64Bit = ($machine -eq 0x8664) }
            } finally {
                $br.Dispose(); $fs.Dispose()
            }
        } catch {} finally {
            if ($key) { $key.Dispose() }
            if ($base) { $base.Dispose() }
        }
    }
    throw '未检测到 Microsoft Excel（仅支持 Windows 桌面版 Microsoft Excel，暂不支持 WPS）。'
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw '检测到正在运行的 Excel 进程，请先保存工作簿并关闭所有 Excel 窗口后再执行安装。'
}

$excelInfo = Get-ExcelBitness
Write-Host "[预检] 检测到 Microsoft Excel: $($excelInfo.Exe) (64位=$($excelInfo.Is64Bit))"

$selected = $Edition
if ($selected -eq 'auto') {
    $selected = if ($excelInfo.Is64Bit) { 'modern' } else { 'legacy' }
} elseif ($selected -eq 'modern' -and -not $excelInfo.Is64Bit) {
    throw '当前检测到 32 位 Microsoft Excel，现代版仅支持 64 位，请改用 -Edition legacy。'
}
Write-Host "[方案] 自动选择安装版本: $selected"

$workDir = Join-Path $env:TEMP ('ElectronicSheepInstall_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $workDir | Out-Null

try {
    $shaUrl = "$baseUrl/SHA256SUMS.txt"
    $shaPath = Join-Path $workDir 'SHA256SUMS.txt'
    Write-Host "[下载] 获取校验清单: $shaUrl"
    Invoke-WebRequest -Uri $shaUrl -OutFile $shaPath -UseBasicParsing

    $assetName = if ($selected -eq 'modern') {
        'ElectronicSheepExcel-v1.0.0-oss-Modern-Setup.zip'
    } else {
        'ElectronicSheepExcel-v1.0.0-oss-Legacy-Setup.exe'
    }

    $expectedHash = $null
    foreach ($rawLine in (Get-Content $shaPath -Encoding UTF8)) {
        $line = $rawLine.Trim([char]0xFEFF, ' ', "`t", "`r", "`n")
        if ($line -match '^([A-Fa-f0-9]{64})\s+(.+)$' -and $matches[2].Trim() -eq $assetName) {
            $expectedHash = $matches[1].ToUpperInvariant()
            break
        }
    }
    if (-not $expectedHash) { throw "在 SHA256SUMS.txt 中未找到 $assetName 的校验值。" }

    $assetUrl = "$baseUrl/$assetName"
    $assetPath = Join-Path $workDir $assetName
    Write-Host "[下载] 正在下载安装包: $assetUrl"
    Invoke-WebRequest -Uri $assetUrl -OutFile $assetPath -UseBasicParsing

    $actualHash = (Get-FileHash -LiteralPath $assetPath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "安装包 SHA-256 校验失败！预期 $expectedHash，实际 $actualHash"
    }
    Write-Host "[校验] SHA-256 校验通过: $actualHash"

    $setupExe = $assetPath
    if ($assetName.EndsWith('.zip')) {
        $extractDir = Join-Path $workDir 'extracted'
        Expand-Archive -Path $assetPath -DestinationPath $extractDir -Force
        $exeFound = Get-ChildItem -Path $extractDir -Filter '*.exe' | Select-Object -First 1
        if (-not $exeFound) { throw '压缩包内未找到安装可执行文件。' }
        $setupExe = $exeFound.FullName
    }

    $setupArgs = if ($Launch) { @() } else { @('--no-launch') }
    Write-Host "[安装] 正在运行安装程序: $($setupExe)"
    $proc = Start-Process -FilePath $setupExe -ArgumentList $setupArgs -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        $logDir = Join-Path $env:LOCALAPPDATA 'DataHelper\SetupLogs'
        $latestLog = Get-ChildItem -Path $logDir -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $detail = if ($latestLog) { Get-Content $latestLog.FullName -Raw } else { "请查看日志目录: $logDir" }
        throw "安装程序返回错误码 $($proc.ExitCode)。日志详情:`n$detail"
    }

    Write-Host "[完成] 电子羊表格 ($selected) 安装成功！"
    if ($selected -eq 'modern') {
        Write-Host "打开方式: 启动 Microsoft Excel -> 功能区「电子羊表格」->「工作台」->「打开/关闭电子羊表格」（或双击桌面快捷方式「电子羊表格 现代版」）。"
    } else {
        Write-Host "打开方式: 启动 Microsoft Excel -> 功能区「电子羊表格」（或双击桌面快捷方式「电子羊表格 兼容版」）。"
    }
} finally {
    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
