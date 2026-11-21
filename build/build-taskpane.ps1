param(
    [Parameter(Mandatory = $true)][ValidatePattern('^[A-Fa-f0-9]{40}$')][string]$CertificateThumbprint,
    [string]$AssetsDir = (Join-Path (Split-Path $PSScriptRoot -Parent) 'assets/taskpane_deps')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$out = Join-Path $root 'build/output/TaskPane'
New-Item -ItemType Directory -Force -Path $out | Out-Null

if (-not (Test-Path -LiteralPath $AssetsDir)) {
    throw "Build dependencies directory not found: $AssetsDir. Pass -AssetsDir containing Office/VSTO ReferenceAssemblies, OfficeTools, MSBuild, and WebView2 assemblies."
}

$assets = Join-Path $AssetsDir 'build_assets'
$refs = Join-Path $assets 'ReferenceAssemblies'
$tasks = Join-Path $assets 'OfficeTools'
$web = Join-Path $AssetsDir 'dist/DeveloperPoC/app'
$msbuild = Join-Path $assets 'MSBuild/Current/Bin/MSBuild.exe'
$mage = Join-Path $assets 'NETFX48/mage.exe'

$interop = Join-Path $refs 'Microsoft.Office.Interop.Excel.dll'
$office = Join-Path $refs 'office.dll'
$referenceNames = @(
    'Microsoft.Office.Tools.dll',
    'Microsoft.Office.Tools.Common.dll',
    'Microsoft.Office.Tools.Excel.dll',
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
    'Microsoft.VisualStudio.Tools.Applications.Runtime.dll',
    'Microsoft.Office.Tools.v4.0.Framework.dll'
)
$compilerRefs = @('/r:System.dll', '/r:System.Core.dll', '/r:System.Windows.Forms.dll', '/r:System.Drawing.dll', '/r:System.Web.Extensions.dll', '/r:Microsoft.CSharp.dll', ('/r:' + $interop), ('/r:' + $office))
foreach ($name in $referenceNames) { $compilerRefs += ('/r:' + (Join-Path $refs $name)) }

foreach ($name in @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll')) {
    $srcDll = Join-Path $web ('lib/net462/' + $name)
    if (-not (Test-Path $srcDll)) { $srcDll = Join-Path $web $name }
    Copy-Item $srcDll $out -Force
    $compilerRefs += ('/r:' + (Join-Path $out $name))
}
$loaderDll = Join-Path $web 'runtimes/win-x64/native/WebView2Loader.dll'
if (-not (Test-Path $loaderDll)) { $loaderDll = Join-Path $web 'WebView2Loader.dll' }
Copy-Item $loaderDll $out -Force
Copy-Item (Join-Path $refs 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll') $out -Force
Copy-Item (Join-Path $root 'frontends/taskpane/index.html') $out -Force
Copy-Item (Join-Path $root 'frontends/taskpane/bridge-client.js') $out -Force

$csc = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe' }
if (-not (Test-Path $csc)) { throw '.NET Framework C# compiler is missing' }

& $csc /nologo /target:library /platform:anycpu /codepage:65001 ('/out:' + (Join-Path $out 'DataHelperTaskPanePoc.dll')) $compilerRefs (Join-Path $root 'frontends/taskpane/ThisAddIn.cs') (Join-Path $root 'frontends/taskpane/PocControl.cs') (Join-Path $root 'frontends/taskpane/BridgeProtocol.cs')
if ($LASTEXITCODE -ne 0) { throw 'VSTO assembly compilation failed' }

$proj = Join-Path $root 'frontends/taskpane/Manifest.proj'
& $msbuild $proj /nologo /t:AppManifest ('/p:OfficeTasks=' + $tasks) ('/p:Out=' + $out) /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'Official VSTO application manifest generation failed' }

$certificate = Get-Item ('Cert:/CurrentUser/My/' + $CertificateThumbprint) -ErrorAction Stop
if (-not $certificate.HasPrivateKey) { throw 'Signing certificate must have a private key in CurrentUser/My' }
if (-not @($certificate.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq '1.3.6.1.5.5.7.3.3' }).Count) { throw 'Code-signing EKU (1.3.6.1.5.5.7.3.3) missing' }
Export-Certificate -Cert $certificate -FilePath (Join-Path $out 'poc-signing.cer') | Out-Null

& $mage -Sign (Join-Path $out 'DataHelperTaskPanePoc.dll.manifest') -CertHash $certificate.Thumbprint -Algorithm sha256RSA
if ($LASTEXITCODE -ne 0) { throw 'Application manifest signing failed' }
& $msbuild $proj /nologo /t:Deployment ('/p:OfficeTasks=' + $tasks) ('/p:Out=' + $out) /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'Deployment manifest generation failed' }
& $mage -Sign (Join-Path $out 'DataHelperTaskPanePoc.vsto') -CertHash $certificate.Thumbprint -Algorithm sha256RSA
if ($LASTEXITCODE -ne 0) { throw 'Deployment manifest signing failed' }

$sourceFiles = Get-ChildItem (Join-Path $root 'frontends/taskpane') -File
@{
    Time = (Get-Date -Format o)
    Sources = @($sourceFiles | ForEach-Object { @{ Path = $_.FullName.Substring($root.Length + 1); SHA256 = (Get-FileHash $_.FullName).Hash } })
    Outputs = @(Get-ChildItem $out -File | Where-Object Name -ne 'build-evidence.json' | ForEach-Object { @{ Path = $_.Name; SHA256 = (Get-FileHash $_.FullName).Hash } })
} | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $out 'build-evidence.json') -Encoding UTF8
