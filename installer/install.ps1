param([ValidateSet('legacy','taskpane')][string]$Edition,
 [ValidateSet('Install','Launch','Uninstall')][string]$Mode='Install',[switch]$NoLaunch)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$homePath=Join-Path $env:LOCALAPPDATA 'DataHelper/1.0.0'
$settings='HKCU:/Software/DataHelper/1.0'
$registration='HKCU:/Software/Microsoft/Office/Excel/Addins/DataHelper.TaskPane'
$options='HKCU:/Software/Microsoft/Office/16.0/Excel/Options'
$engine=Join-Path $homePath 'Core/DataHelper_PocEngine.xlam'
$manifest=Join-Path $homePath 'TaskPane/DataHelperTaskPanePoc.vsto'
function EnsureKey([string]$path){if(-not(Test-Path $path)){New-Item $path -Force|Out-Null}}
function GetExcelServerCommand([Microsoft.Win32.RegistryView]$view){
 $base=$null;$key=$null
 try{
  $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot,$view)
  $key=$base.OpenSubKey('CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32',$false)
  if($key){return [string]$key.GetValue('')}
 }catch{return $null}finally{if($key){$key.Dispose()};if($base){$base.Dispose()}}
}
function ExcelPath {
 $views=if([Environment]::Is64BitOperatingSystem){@([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)}else{@([Microsoft.Win32.RegistryView]::Registry32)}
 foreach($view in $views){
  $server=GetExcelServerCommand $view
  if([string]::IsNullOrWhiteSpace($server)){continue}
  $commandMatch=[regex]::Match($server,'^\s*(?:"(?<exe>[^"]+\.exe)"|(?<exe>.+?\.exe))(?:\s|$)',[Text.RegularExpressions.RegexOptions]::IgnoreCase)
  if(-not $commandMatch.Success){continue}
  $path=[Environment]::ExpandEnvironmentVariables($commandMatch.Groups['exe'].Value.Trim())
  if(-not [IO.Path]::IsPathRooted($path) -or [IO.Path]::GetFileName($path) -ine 'EXCEL.EXE'){continue}
  if(Test-Path -LiteralPath $path -PathType Leaf){return $path}
 }
 throw '未检测到 Microsoft Excel，本软件不支持 WPS。请安装 Windows Microsoft Excel；若已安装，请修复其安装登记后重试。'
}
function ReadExcelMachine([string]$path){
 $stream=$null;$reader=$null
 try{
  $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
  $reader=New-Object IO.BinaryReader($stream)
  if($stream.Length -lt 64 -or $reader.ReadUInt16() -ne 0x5A4D){throw 'Invalid DOS header'}
  $stream.Position=60;$pe=$reader.ReadInt32()
  if($pe -lt 64 -or $pe -gt ($stream.Length-6)){throw 'Invalid PE offset'}
  $stream.Position=$pe
  if($reader.ReadUInt32() -ne 0x4550){throw 'Invalid PE signature'}
  return $reader.ReadUInt16()
 }catch{throw '无法识别 Microsoft Excel 程序位数；请修复 Office 安装后重试。'}
 finally{if($reader){$reader.Dispose()}elseif($stream){$stream.Dispose()}}
}
function RequireModernExcel {
 $machine=ReadExcelMachine (ExcelPath)
 if($machine -eq 0x14c){throw '现代版需要 64 位 Office；当前检测到 32 位 Office，请改用电子羊表格兼容版。'}
 if($machine -ne 0x8664){throw '现代版需要本包支持的 64 位 Office（x64）；当前架构不受支持，请改用兼容版或联系维护者。'}
}
function PayloadHash([string]$path){
 $stream=$null;$sha=$null
 try{
  $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
  $sha=[Security.Cryptography.SHA256]::Create()
  return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')
 }catch{
  $cause=$_.Exception;$locked=$false
  while($cause){$code=$cause.HResult -band 65535;if($code -eq 32 -or $code -eq 33){$locked=$true};$cause=$cause.InnerException}
  if($locked){throw ('Payload file is in use: '+$path+'; hash was not compared')}
  throw ('Payload read failed: '+$path+'; hash was not compared; '+$_.Exception.Message)
 }finally{if($sha){$sha.Dispose()};if($stream){$stream.Dispose()}}
}
function ExportRegistryKeyNative([string]$keyPath,[string]$file){
 # R13 fix: export a .reg-compatible snapshot with the native registry API instead
 # of shelling out to reg.exe (blocked by policy on managed machines).
 $parts=$keyPath.Split('\');$rootName=$parts[0]
 $sub=($parts[1..($parts.Length-1)] -join '\')
 $root=$null
 if($rootName -eq 'HKCU'){$root=[Microsoft.Win32.Registry]::CurrentUser}
 elseif($rootName -eq 'HKLM'){$root=[Microsoft.Win32.Registry]::LocalMachine}
 else{throw ('Unsupported registry root: '+$rootName)}
 function ConvertTo-RegHex([object]$bytes){
  if($null -eq $bytes){return ''}
  $arr=@([byte[]]$bytes)
  if($arr.Count -eq 0){return ''}
  return (($arr | ForEach-Object { $_.ToString('x2') }) -join ',')
 }
 function Add-Escaped([string]$text){ return $text.Replace('\','\\').Replace('"','\"') }
 function Write-KeyTo($key,[string]$display,$lines){
  $lines.Add('['+$display+']')
  foreach($valueName in $key.GetValueNames()){
   $kind=[string]$key.GetValueKind($valueName)
   $raw=$key.GetValue($valueName,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
   $label=if($valueName -eq ''){'@'}else{'"'+(Add-Escaped $valueName)+'"'}
   switch($kind){
    'String'{ $lines.Add($label+'="'+(Add-Escaped ([string]$raw))+'"') }
    'ExpandString'{ $bytes=[Text.Encoding]::Unicode.GetBytes(([string]$raw)+[char]0); $lines.Add($label+'=hex(2):'+(ConvertTo-RegHex $bytes)) }
    'DWord'{ $lines.Add($label+'=dword:'+('{0:x8}' -f ([uint32]([int64]$raw)))) }
    'QWord'{ $bytes=[BitConverter]::GetBytes([int64]$raw); $lines.Add($label+'=hex(b):'+(ConvertTo-RegHex $bytes)) }
    'MultiString'{ $bytes=[Text.Encoding]::Unicode.GetBytes((([string[]]$raw) -join [char]0)+[char]0+[char]0); $lines.Add($label+'=hex(7):'+(ConvertTo-RegHex $bytes)) }
    'Binary'{ $lines.Add($label+'=hex:'+(ConvertTo-RegHex ([byte[]]$raw))) }
    default{ $lines.Add($label+'=hex:'+(ConvertTo-RegHex ([byte[]]$raw))) }
   }
  }
  foreach($childName in $key.GetSubKeyNames()){
   $child=$key.OpenSubKey($childName,$false)
   try{ if($null -ne $child){ $lines.Add(''); Write-KeyTo $child ($display+'\'+$childName) $lines } } finally{ if($child){$child.Dispose()} }
  }
 }
 $lines=[System.Collections.Generic.List[string]]::new()
 $lines.Add('Windows Registry Editor Version 5.00')
 $lines.Add('')
 $key=$root.OpenSubKey($sub,$false)
 if($null -eq $key){throw ('Registry key not found: '+$keyPath)}
 try{ Write-KeyTo $key ('HKEY_CURRENT_USER\'+$sub) $lines } finally{ $key.Dispose() }
 [IO.File]::WriteAllText($file,($lines -join "`r`n"),[Text.Encoding]::Unicode)
}
function SaveRegistryBefore {
 $folder=Join-Path $env:LOCALAPPDATA ('DataHelper/SetupLogs/registry_'+(Get-Date -Format 'yyyyMMdd_HHmmssfff'))
 New-Item -ItemType Directory -Force $folder|Out-Null
 # Snapshot is an audit aid for restore; a failure is recorded next to the
 # snapshot but must not block install/launch (R13).
 foreach($entry in @(@('Options','HKCU\Software\Microsoft\Office\16.0\Excel\Options'),@('Product','HKCU\Software\DataHelper\1.0'),@('Host','HKCU\Software\Microsoft\Office\Excel\Addins\DataHelper.TaskPane'))){
  $test='Registry::HKEY_CURRENT_USER'+$entry[1].Substring(4)
  if(Test-Path $test){
   try{ ExportRegistryKeyNative $entry[1] (Join-Path $folder ($entry[0]+'.reg')) }
   catch{ (([string]$_.Exception.Message)+"`r`n")|Set-Content (Join-Path $folder ($entry[0]+'.snapshot-warning.txt')) -Encoding UTF8 }
  }
 }
}
function RequireExcelClosed {
 # Do not block or terminate another agent/user's unrelated Excel process.
 # Exclusive access is the actual prerequisite for changing this installation.
 # R13 fix: probe all runtime files we overwrite, not only the engine xlam:
 # an Excel that merely hosts the task pane keeps WebView2 DLLs locked even
 # when the engine is not loaded.
 $targets=@($engine,(Join-Path $homePath 'TaskPane/WebView2Loader.dll'),(Join-Path $homePath 'TaskPane/DataHelperTaskPanePoc.dll'))
 foreach($target in $targets){
  if(Test-Path -LiteralPath $target){
   $handle=$null
   try{$handle=[IO.File]::Open($target,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)}
   catch{throw ('请先保存并关闭正在使用电子羊表格的 Excel，再安装或切换（文件被占用：'+(Split-Path $target -Leaf)+'）。安装器不会终止 Excel。')}
   finally{if($handle){$handle.Dispose()}}
  }
 }
}
function MigrateKnownLegacy {
 $saved=Join-Path $homePath 'previous-open.json'
 if(Test-Path $saved){return}
 $previous=@()
 foreach($property in (Get-ItemProperty $options).PSObject.Properties){
  if($property.Name -notmatch '^OPEN\d*$'){continue}
  $value=[string]$property.Value
  if($value -eq '"DataHelper_v1.0-rc1.xlam"'){
   $old=Join-Path $env:APPDATA 'Microsoft/AddIns/DataHelper_v1.0-rc1.xlam'
   if((Test-Path $old) -and (PayloadHash $old) -eq '93968156F84838D9C1BE18B69C47C89F2EE7A9045F60DC935C74651C81BF663F'){
    $previous+=@{Name=$property.Name;Value=$value;File=$old;SHA256=(PayloadHash $old)}
    Remove-ItemProperty $options $property.Name
   }
  }
 }
 ConvertTo-Json -InputObject @($previous) -Depth 4|Set-Content $saved -Encoding UTF8
}
function CheckPayload {
 $items=Get-Content (Join-Path $PSScriptRoot 'payload-hashes.json') -Raw|ConvertFrom-Json
 foreach($item in $items){
  $path=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot $item.Path))
  if(-not $path.StartsWith($PSScriptRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe payload path'}
  if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw ('Payload read failed: missing '+$item.Path+'; hash was not compared')}
  if((PayloadHash $path) -ne $item.SHA256){throw ('Payload hash mismatch: '+$item.Path)}
 }
}
function RemoveOwnOpen {
 if(Test-Path $options){foreach($property in (Get-ItemProperty $options).PSObject.Properties){
  if($property.Name -match '^OPEN\d*$' -and [string]$property.Value -eq ('/R "'+$engine+'"')){Remove-ItemProperty $options $property.Name}
 }}
}
function RemoveOwnModern {
 if(Test-Path $registration){
  if((Get-ItemProperty $registration).Manifest -ne (([Uri]$manifest).AbsoluteUri+'|vstolocal')){throw '已有同名加载项注册不属于本安装器，未覆盖。'}
  Remove-Item -LiteralPath $registration
 }
 $trust=Join-Path $homePath 'internal-trust.json'
 if(Test-Path $trust){
  $entry=Get-Content $trust -Raw|ConvertFrom-Json
  if($entry.Key -match '^HKCU:/Software/Microsoft/VSTO/Security/Inclusion/[a-f0-9-]+$' -and (Test-Path $entry.Key)){
   $value=Get-ItemProperty $entry.Key
   if($value.Url -eq $entry.Url -and $value.PublicKey -eq $entry.PublicKey){Remove-Item -LiteralPath $entry.Key}
  }
 }
}
function SetFrontend {
 EnsureKey $settings
 New-ItemProperty $settings InstallRoot -Value $homePath -PropertyType String -Force|Out-Null
 New-ItemProperty $settings Frontend -Value $Edition -PropertyType String -Force|Out-Null
 if($Edition -eq 'legacy'){RemoveOwnModern}
 else{
  if(-not(Test-Path $manifest)){throw '现代版文件尚未安装。'}
  $certificate=New-Object Security.Cryptography.X509Certificates.X509Certificate2((Join-Path $homePath 'TaskPane/poc-signing.cer'))
  if($certificate.NotAfter -lt (Get-Date)){throw '安装证书已过期，请获取更新的安装包；请勿调低安全设置。'}
  $url=([Uri]$manifest).AbsoluteUri
  $trust=Join-Path $homePath 'internal-trust.json'
  if(Test-Path $trust){RemoveOwnModern}
  $key='HKCU:/Software/Microsoft/VSTO/Security/Inclusion/'+[Guid]::NewGuid().ToString()
  $publicKey=$certificate.PublicKey.Key.ToXmlString($false)
  New-Item $key -Force|Out-Null
  New-ItemProperty $key Url -Value $url -PropertyType String|Out-Null
  New-ItemProperty $key PublicKey -Value $publicKey -PropertyType String|Out-Null
  @{Key=$key;Url=$url;PublicKey=$publicKey;Scope='Internal URL only; no Root/TrustedPublisher import'}|ConvertTo-Json|Set-Content $trust -Encoding UTF8
  EnsureKey $registration
  New-ItemProperty $registration Manifest -Value ($url+'|vstolocal') -PropertyType String -Force|Out-Null
  New-ItemProperty $registration FriendlyName -Value '电子羊表格' -PropertyType String -Force|Out-Null
  New-ItemProperty $registration Description -Value '电子羊表格现代工作台' -PropertyType String -Force|Out-Null
  New-ItemProperty $registration LoadBehavior -Value 3 -PropertyType DWord -Force|Out-Null
 }
}
function LaunchExcel {
 $exe=ExcelPath
 if(-not(Test-Path -LiteralPath $exe)){throw '未找到 Microsoft Excel。'}
 $launch=Join-Path $homePath ('Runtime/launch_'+[Guid]::NewGuid().ToString('N'))
 New-Item -ItemType Directory -Force $launch|Out-Null
 $book=Join-Path $launch 'DataHelper_Start.xlsx'
 Copy-Item (Join-Path $homePath 'Start.xlsx') $book
 $process=Start-Process -FilePath $exe -ArgumentList @('/x',('"'+$book+'"')) -PassThru
 @{Pid=$process.Id;Workbook=$book;Edition=$Edition;Time=(Get-Date -Format o)}|ConvertTo-Json|Set-Content (Join-Path $homePath 'last-launch.json') -Encoding UTF8
 Add-Type -Path (Join-Path $homePath 'RunningWorkbook.cs')
 if(-not ('DHSetupNative' -as [type])){Add-Type 'public static class DHSetupNative { [System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hwnd,out uint pid); }'}
 $deadline=(Get-Date).AddSeconds(40);$ownedBook=$null;$excel=$null
 do{
  $ownedBook=[RunningWriteWorkbook]::Find($book)
  if($null -ne $ownedBook){
   try{
    $candidateApp=$ownedBook.Application
    if($null -ne $candidateApp -and $null -ne $candidateApp.Workbooks -and $candidateApp.Hwnd){$excel=$candidateApp}
   }catch{}
  }
  if($null -eq $excel){Start-Sleep -Milliseconds 300}
 }while($null -eq $excel -and (Get-Date) -lt $deadline)
 if($null -eq $excel){throw 'Excel 已启动但未完成加载，请查看 Excel 提示；安装器没有修改宏安全策略。'}
 [uint32]$actualPid=0;[void][DHSetupNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$actualPid)
 if($actualPid -ne $process.Id){throw '启动工作簿对应的 Excel 进程与本次启动不一致，未继续加载核心。'}
 # OPEN is the normal startup registration. If Excel skipped startup add-ins,
 # load this installation by full path, never a bare filename resolved in Documents.
 $loadedEngine=$null
 try{$loadedEngine=$excel.Workbooks.Item('DataHelper_PocEngine.xlam')}catch{}
 if($loadedEngine){
  if(-not [string]::Equals($loadedEngine.FullName,$engine,[StringComparison]::OrdinalIgnoreCase)){throw '此 Excel 已加载另一份同名电子羊表格核心；请关闭后从本版快捷方式重新启动。'}
 }else{
  $loadedEngine=$excel.Workbooks.Open($engine,0,$true)
  if(-not $loadedEngine -or -not [string]::Equals($loadedEngine.FullName,$engine,[StringComparison]::OrdinalIgnoreCase)){throw '未能从安装目录加载电子羊表格核心；请查看 Excel 的宏策略提示。'}
 }
 $health=$excel.Run("'DataHelper_PocEngine.xlam'!DH1_Health")
 if($health -ne 'DataHelper/1.0.0/shared-engine'){throw '加载项健康检查未通过，请查看 Excel 的宏策略提示。'}
 $hostState=$null
 if($Edition -eq 'taskpane'){
  $addin=$excel.COMAddIns.Item('DataHelper.TaskPane')
  if(-not $addin.Connect -or -not $addin.Object){throw '现代宿主未加载，请查看安装日志和 Excel 加载项状态。'}
  do{Start-Sleep -Milliseconds 250;$hostState=$addin.Object.StateJson()|ConvertFrom-Json}while(-not @($hostState.panes|Where-Object ready).Count -and (Get-Date) -lt $deadline)
  if(-not @($hostState.panes|Where-Object ready).Count){throw '任务窗格未就绪，请查看本地宿主日志。'}
 }
 @{Health=$health;Edition=$Edition;HostState=$hostState;Pid=$process.Id;Workbook=$book;Time=(Get-Date -Format o)}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $homePath 'startup-health.json') -Encoding UTF8
}
function CleanRuntime {
 # N-1: same rule as the host: only <pid>_<guid> folders whose Excel is gone, and our own
 # launch copies. WebProfile (receipts), backups and audit data stay.
 $removed=0;$kept=0
 if(ProgramPathProblem 'install.ps1'){Write-Warning 'Runtime cleanup skipped: installation path is not safe';return}
 $hostRuntime=Join-Path $env:LOCALAPPDATA 'DataHelper/Runtime'
 if((Test-Path -LiteralPath $hostRuntime) -and ((Get-Item -LiteralPath $hostRuntime -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0){foreach($dir in Get-ChildItem -LiteralPath $hostRuntime -Directory){
  if(($dir.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){$kept++;continue}
  if($dir.Name -notmatch '^(\d+)_[0-9a-f]{32}$'){continue}
  $alive=$false
  try{$p=Get-Process -Id ([int]$Matches[1]) -ErrorAction Stop;$alive=($p.ProcessName -eq 'EXCEL') -and ($p.StartTime -le $dir.CreationTime.AddSeconds(5))}catch [Microsoft.PowerShell.Commands.ProcessCommandException]{$alive=$false}catch{$alive=$true}
  if($alive){$kept++;continue}
  Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
  if(Test-Path -LiteralPath $dir.FullName){$kept++}else{$removed++}
 }}
 $launchRoot=Join-Path $homePath 'Runtime'
 if((Test-Path -LiteralPath $launchRoot) -and ((Get-Item -LiteralPath $launchRoot -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0){foreach($dir in Get-ChildItem -LiteralPath $launchRoot -Directory -Filter 'launch_*'){
  if(($dir.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){$kept++;continue}
  Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
  if(Test-Path -LiteralPath $dir.FullName){$kept++}else{$removed++}
 }}
 'Runtime cleanup: removed='+$removed+' kept='+$kept
}
function ProgramFiles {
 @('Core\DataHelper_PocEngine.xlam','install.ps1','Start.xlsx','product.json','RunningWorkbook.cs','app.ico',
 'TaskPane\DataHelperTaskPanePoc.dll','TaskPane\DataHelperTaskPanePoc.dll.manifest','TaskPane\DataHelperTaskPanePoc.vsto',
 'TaskPane\Microsoft.Office.Tools.Common.v4.0.Utilities.dll','TaskPane\Microsoft.Web.WebView2.Core.dll',
 'TaskPane\Microsoft.Web.WebView2.WinForms.dll','TaskPane\WebView2Loader.dll','TaskPane\index.html','TaskPane\bridge-client.js','TaskPane\poc-signing.cer')
}
function ProgramPathProblem([string]$relative){
 $relative=$relative.Replace('/','\')
 $allowed=@(ProgramFiles)+@('installed-files.json','last-install.json','last-launch.json','startup-health.json','internal-trust.json')
 if($relative -notin $allowed){return 'invalid-record'}
 $root=[IO.Path]::GetFullPath($homePath).TrimEnd('\');$path=[IO.Path]::GetFullPath((Join-Path $root $relative))
 if(-not $path.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){return 'invalid-record'}
 # Never follow a junction/symlink, including an installation ancestor.
 $cursor=$path
 while($cursor){
  try{if(([IO.File]::GetAttributes($cursor) -band [IO.FileAttributes]::ReparsePoint) -ne 0){return 'reparse-point'}}
  catch [IO.FileNotFoundException]{}catch [IO.DirectoryNotFoundException]{}catch{return 'path-unreadable'}
  $parent=[IO.Path]::GetDirectoryName($cursor);if($parent -eq $cursor){break};$cursor=$parent
 }
 return $null
}
function SaveInstalledFiles {
 $inventory=Join-Path $homePath 'installed-files.json';$files=@{}
 $problem=ProgramPathProblem 'installed-files.json';if($problem){throw ('Installation inventory path rejected: '+$problem)}
 if(Test-Path -LiteralPath $inventory){
  $prior=Get-Content -LiteralPath $inventory -Raw -Encoding UTF8|ConvertFrom-Json
  if($prior.Schema -ne 'DataHelper.InstallFiles.v1'){throw 'Unrecognized installation inventory; no ownership record overwritten'}
  foreach($entry in $prior.Files){$path=([string]$entry.Path).Replace('/','\');if($path -in @(ProgramFiles) -and [string]$entry.SHA256 -match '^[0-9a-fA-F]{64}$'){$files[$path]=@{Path=$path;SHA256=[string]$entry.SHA256}}}
 }
 foreach($entry in (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'payload-hashes.json') -Raw -Encoding UTF8|ConvertFrom-Json)){
  $path=([string]$entry.Path).Replace('/','\')
  if($path -notin @(ProgramFiles) -or [string]$entry.SHA256 -notmatch '^[0-9a-fA-F]{64}$'){throw 'Unexpected installation inventory entry'}
  $files[$path]=@{Path=$path;SHA256=[string]$entry.SHA256}
 }
 @{Schema='DataHelper.InstallFiles.v1';Files=@($files.Values)}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $inventory -Encoding UTF8
}
function RemoveInstalledProgramFiles {
 $removed=[Collections.Generic.List[string]]::new();$kept=[Collections.Generic.List[object]]::new()
 $inventory=Join-Path $homePath 'installed-files.json';$problem=ProgramPathProblem 'installed-files.json'
 if(-not $problem -and -not(Test-Path -LiteralPath $inventory)){$problem='inventory-missing'}
 if($problem){return @{Removed=@();Preserved=@(@{Path='installed-files.json';Reason=$problem});ProgramRootRemoved=$false}}
 try{$record=Get-Content -LiteralPath $inventory -Raw -Encoding UTF8|ConvertFrom-Json;if($record.Schema -ne 'DataHelper.InstallFiles.v1' -or -not $record.Files){throw 'Invalid inventory'}}
 catch{return @{Removed=@();Preserved=@(@{Path='installed-files.json';Reason='inventory-invalid'});ProgramRootRemoved=$false}}
 foreach($entry in $record.Files){
  $relative=([string]$entry.Path).Replace('/','\');$problem=ProgramPathProblem $relative
  if(-not $problem -and [string]$entry.SHA256 -notmatch '^[0-9a-fA-F]{64}$'){$problem='invalid-record'}
  if($problem){$kept.Add(@{Path=$relative;Reason=$problem});continue}
  $path=Join-Path $homePath $relative
  if(-not(Test-Path -LiteralPath $path -PathType Leaf)){continue}
  try{
   if((PayloadHash $path) -ne $entry.SHA256){$kept.Add(@{Path=$relative;Reason='hash-changed'});continue}
   Remove-Item -LiteralPath $path -Force -ErrorAction Stop;$removed.Add($relative)
  }catch{$kept.Add(@{Path=$relative;Reason='delete-failed';Error=$_.Exception.Message})}
 }
 $metadata=@('last-install.json','last-launch.json','startup-health.json','internal-trust.json')
 if($kept.Count -eq 0){$metadata+='installed-files.json'}
 foreach($relative in $metadata){
  $problem=ProgramPathProblem $relative;$path=Join-Path $homePath $relative
  if($problem){$kept.Add(@{Path=$relative;Reason=$problem});continue}
  if(Test-Path -LiteralPath $path -PathType Leaf){try{Remove-Item -LiteralPath $path -Force -ErrorAction Stop;$removed.Add($relative)}catch{$kept.Add(@{Path=$relative;Reason='delete-failed';Error=$_.Exception.Message})}}
 }
 # Only prune known empty folders; no recursive deletion of the installation tree.
 foreach($relative in @('Core','TaskPane','Runtime','')){
  $path=if($relative){Join-Path $homePath $relative}else{$homePath}
  try{if([IO.Directory]::Exists($path) -and ([IO.File]::GetAttributes($path) -band [IO.FileAttributes]::ReparsePoint) -eq 0 -and [IO.Directory]::GetFileSystemEntries($path).Count -eq 0){[IO.Directory]::Delete($path,$false)}}
  catch{$kept.Add(@{Path=$relative;Reason='directory-retained';Error=$_.Exception.Message})}
 }
 if(Test-Path -LiteralPath $homePath){foreach($entry in Get-ChildItem -LiteralPath $homePath -Force){$kept.Add(@{Path=$entry.Name;Reason='unmanaged-or-retained'})}}
 return @{Removed=@($removed.ToArray());Preserved=@($kept.ToArray());ProgramRootRemoved=(-not(Test-Path -LiteralPath $homePath))}
}
if($Mode -ne 'Uninstall'){
 try{if($Edition -eq 'taskpane'){RequireModernExcel}else{[void](ExcelPath)}}catch{
  # Known preflight failure: concise UI, full setup log; no registration/payload writes.
  [Console]::Error.WriteLine('DH_SETUP_EXPECTED_ERROR:'+($_.Exception.Message -replace '[\r\n]+',' '))
  exit 2
 }
}
RequireExcelClosed
SaveRegistryBefore
if($Mode -eq 'Uninstall'){
 CleanRuntime
 $removedEdition=$Edition
 $property='Installed_'+$Edition
 if(Test-Path $settings){Remove-ItemProperty $settings $property -ErrorAction SilentlyContinue}
 $other=if($Edition -eq 'legacy'){'taskpane'}else{'legacy'}
 $remaining=Get-ItemProperty $settings -ErrorAction SilentlyContinue
 if($remaining -and $remaining.PSObject.Properties['Installed_'+$other] -and $remaining.('Installed_'+$other)){
  $Edition=$other;SetFrontend
 }else{
  RemoveOwnModern;RemoveOwnOpen
  $saved=Join-Path $homePath 'previous-open.json'
  if(Test-Path $saved){foreach($entry in (Get-Content $saved -Raw|ConvertFrom-Json)){
   if((Test-Path $entry.File) -and (PayloadHash $entry.File) -eq $entry.SHA256){
    $index=0;$name='OPEN';while($null -ne (Get-ItemProperty $options).PSObject.Properties[$name]){$index++;$name='OPEN'+$index}
    New-ItemProperty $options $name -Value $entry.Value -PropertyType String|Out-Null
   }
  };Remove-Item -LiteralPath $saved}
  if(Test-Path $settings){Remove-Item -LiteralPath $settings}
 }
 $desktop=[Environment]::GetFolderPath('DesktopDirectory')
 $name=if($property -eq 'Installed_legacy'){'Legacy'}else{'TaskPane'}
 $editionName=if($name -eq 'Legacy'){'兼容版'}else{'现代版'}
 foreach($file in @((Join-Path $desktop ('DataHelper 1.0 '+$name+'.lnk')),(Join-Path $desktop ('Uninstall DataHelper '+$name+'.lnk')),
  (Join-Path $desktop ('电子羊表格 '+$editionName+'.lnk')),(Join-Path $desktop ('卸载电子羊表格 '+$editionName+'.lnk')))){
  if(Test-Path $file){Remove-Item -LiteralPath $file}
 }
 $remainingFrontend=if(Test-Path $settings){(Get-ItemProperty $settings).Frontend}else{$null}
 # Audit survives outside the program directory; user data and another frontend stay.
 $logRoot=Join-Path $env:LOCALAPPDATA 'DataHelper/SetupLogs';New-Item -ItemType Directory -Force $logRoot|Out-Null
 $logFile=Join-Path $logRoot ('uninstall_'+(Get-Date -Format 'yyyyMMdd_HHmmssfff')+'_'+[Guid]::NewGuid().ToString('N')+'.json')
 $audit=@{Version='1.0.0';Edition=$removedEdition;Mode='Uninstall';RemainingFrontend=$remainingFrontend;Time=(Get-Date -Format o);PublicRelease=$true;Cleanup=@{Status='pending'};UserData='Preserved: ExcelDataHelper backups/details/logs and DataHelper WebProfile'}
 $audit|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $logFile -Encoding UTF8
 if($remainingFrontend){
  $audit.Cleanup=@{Status='shared-frontend-retained';Removed=@();Preserved=@();ProgramRootRemoved=$false}
  $audit|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $homePath 'last-install.json') -Encoding UTF8
 }else{$audit.Cleanup=RemoveInstalledProgramFiles}
 $audit|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $logFile -Encoding UTF8
 if(@($audit.Cleanup.Preserved).Count){Write-Warning ('Program files retained; see '+$logFile)}
 'UNINSTALLED (user backup/audit preserved); cleanup log: '+$logFile;exit 0
}
if($Mode -eq 'Install'){
 CheckPayload
 if($Edition -eq 'taskpane'){
  $net=(Get-ItemProperty 'HKLM:/Software/Microsoft/NET Framework Setup/NDP/v4/Full' -ErrorAction SilentlyContinue).Release
  if($net -lt 528040){throw '现代版需要 .NET Framework 4.8；请先安装微软运行库。'}
  $vsto=$false
  foreach($path in @('HKLM:/SOFTWARE/Microsoft/VSTO Runtime Setup','HKLM:/SOFTWARE/WOW6432Node/Microsoft/VSTO Runtime Setup')){
   foreach($version in @('v4','v4R','v4M')){
    $entry=Get-ItemProperty ($path+'/'+$version) -ErrorAction SilentlyContinue
    if($entry -and ($entry.Install -eq 1 -or $entry.InstallerPath)){$vsto=$true}
   }
  }
  if(-not $vsto){throw '现代版需要 Microsoft VSTO Runtime；兼容版不需要此组件。'}
  $wv=Get-ChildItem 'HKLM:/SOFTWARE/WOW6432Node/Microsoft/EdgeUpdate/Clients','HKCU:/Software/Microsoft/EdgeUpdate/Clients' -ErrorAction SilentlyContinue|Get-ItemProperty|Where-Object { $_.name -like '*WebView2*' -and $_.pv -and $_.pv -ne '0.0.0.0' }
  if(-not $wv){throw '现代版需要 Microsoft Edge WebView2 Runtime；请先安装微软运行库。'}
 }
 if(Test-Path $registration){if((Get-ItemProperty $registration).Manifest -ne (([Uri]$manifest).AbsoluteUri+'|vstolocal')){throw '同名加载项注册冲突，未改动。'}}
 New-Item -ItemType Directory -Force $homePath|Out-Null
 foreach($item in (Get-Content (Join-Path $PSScriptRoot 'payload-hashes.json') -Raw|ConvertFrom-Json)){
  $destination=Join-Path $homePath $item.Path
  New-Item -ItemType Directory -Force (Split-Path $destination -Parent)|Out-Null
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot $item.Path) -Destination $destination -Force
 }
 SaveInstalledFiles
 EnsureKey $options
 MigrateKnownLegacy
 RemoveOwnOpen
 $index=0;$name='OPEN'
 while($null -ne (Get-ItemProperty $options).PSObject.Properties[$name]){$index++;$name='OPEN'+$index}
 New-ItemProperty $options $name -Value ('/R "'+$engine+'"') -PropertyType String|Out-Null
 EnsureKey $settings
 New-ItemProperty $settings ('Installed_'+$Edition) -Value 1 -PropertyType DWord -Force|Out-Null
 # Startup/Uninstall shortcuts invoke this verified local script; no manual Add-ins dialog.
 Add-Type -Path (Join-Path $homePath 'RunningWorkbook.cs')
 $desktop=[Environment]::GetFolderPath('DesktopDirectory')
 $name=if($Edition -eq 'legacy'){'Legacy'}else{'TaskPane'}
 $editionName=if($Edition -eq 'legacy'){'兼容版'}else{'现代版'}
 foreach($operation in @('Launch','Uninstall')){
  $label=if($operation -eq 'Launch'){'电子羊表格 '+$editionName}else{'卸载电子羊表格 '+$editionName}
  $shortcutTarget=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
  $shortcutArguments='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $homePath 'install.ps1')+'" -Edition '+$Edition+' -Mode '+$operation
  [DHDesktopShortcut]::Write((Join-Path $desktop ($label+'.lnk')),$shortcutTarget,$shortcutArguments,'电子羊表格',(Join-Path $homePath 'app.ico'))
  # Remove only an old shortcut whose target and arguments match this installation.
  $oldLabel=if($operation -eq 'Launch'){'DataHelper 1.0 '+$name}else{'Uninstall DataHelper '+$name}
  $oldShortcut=Join-Path $desktop ($oldLabel+'.lnk')
  if(Test-Path -LiteralPath $oldShortcut){
   try{$oldLink=[DHDesktopShortcut]::Read($oldShortcut);if($oldLink[0] -ieq $shortcutTarget -and $oldLink[1] -ceq $shortcutArguments){Remove-Item -LiteralPath $oldShortcut}}catch{Write-Warning '旧快捷方式归属无法确认，已保留。'}
  }
 }
}
SetFrontend
@{Version='1.0.0';Edition=$Edition;Mode=$Mode;CoreSHA256=(PayloadHash $engine);Time=(Get-Date -Format o);PublicRelease=$true}|ConvertTo-Json|Set-Content (Join-Path $homePath 'last-install.json') -Encoding UTF8
if(-not $NoLaunch){LaunchExcel}
'SUCCESS '+$Edition+' '+$Mode
