param([switch]$NoLaunch)

$ErrorActionPreference = 'Stop'
$clsid = '{88AB88AB-CDFB-4C68-9C3A-F10B75A5BC61}'
$source = Join-Path $PSScriptRoot 'OneMoreAddIn'
$target = Join-Path $env:LOCALAPPDATA 'OneMoreComments\OneMoreAddIn'

if (-not (Test-Path -LiteralPath (Join-Path $source 'River.OneMoreAddIn.dll'))) {
  throw 'The OneMoreAddIn payload is missing. Extract the complete release ZIP first.'
}

$oneNote = Get-Process ONENOTE -ErrorAction SilentlyContinue
if ($oneNote) {
  $oneNote | ForEach-Object { [void]$_.CloseMainWindow() }
  $deadline = (Get-Date).AddSeconds(15)
  while ((Get-Process ONENOTE -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }
  if (Get-Process ONENOTE -ErrorAction SilentlyContinue) { throw 'OneNote is still running. Close it manually and run this installer again.' }
}

$tray = Get-Process OneMoreTray -ErrorAction SilentlyContinue
foreach ($process in $tray) {
  try {
    if ($process.Path -and $process.Path.StartsWith($target, [StringComparison]::OrdinalIgnoreCase)) { $process | Stop-Process -Force }
  } catch {}
}
Start-Sleep -Milliseconds 500

New-Item -ItemType Directory -Path $target -Force | Out-Null
Get-ChildItem -LiteralPath $source -Force | Copy-Item -Destination $target -Recurse -Force

$classes = 'Registry::HKEY_CURRENT_USER\Software\Classes'
$base = "$classes\CLSID\$clsid"
$inproc = "$base\InprocServer32"
$version = "$inproc\7.2.0"
$assembly = 'River.OneMoreAddIn, Version=7.2.0, Culture=neutral, PublicKeyToken=null'
$codeBase = Join-Path $target 'River.OneMoreAddIn.dll'

New-Item -Path "$classes\River.OneMoreAddIn\CLSID", $version, "$base\ProgID", "$base\Programmable", "$base\TypeLib", "$base\VersionIndependentProgID", "$base\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null
Set-Item -LiteralPath "$classes\River.OneMoreAddIn" -Value 'River.OneMoreAddIn.AddIn'
Set-Item -LiteralPath "$classes\River.OneMoreAddIn\CLSID" -Value $clsid
Set-Item -LiteralPath $base -Value 'River.OneMoreAddIn.AddIn'
Set-Item -LiteralPath $inproc -Value 'mscoree.dll'
foreach ($path in @($inproc, $version)) {
  Set-ItemProperty -LiteralPath $path -Name Assembly -Value $assembly
  Set-ItemProperty -LiteralPath $path -Name Class -Value 'River.OneMoreAddIn.AddIn'
  Set-ItemProperty -LiteralPath $path -Name RuntimeVersion -Value 'v4.0.30319'
  Set-ItemProperty -LiteralPath $path -Name CodeBase -Value $codeBase
}
Set-ItemProperty -LiteralPath $inproc -Name ThreadingModel -Value Both
Set-Item -LiteralPath "$base\ProgID" -Value 'River.OneMoreAddIn'
Set-Item -LiteralPath "$base\VersionIndependentProgID" -Value 'River.OneMoreAddIn'
Set-Item -LiteralPath "$base\TypeLib" -Value $clsid

$addin = 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\River.OneMoreAddIn'
New-Item -Path $addin -Force | Out-Null
Set-ItemProperty -LiteralPath $addin -Name LoadBehavior -Type DWord -Value 3
Set-ItemProperty -LiteralPath $addin -Name FriendlyName -Value 'Comments for OneNote (OneMore)'
Set-ItemProperty -LiteralPath $addin -Name Description -Value 'Embedded movable comments for OneNote, powered by OneMore'

Write-Host "Comments for OneNote installed to $target" -ForegroundColor Green
if (-not $NoLaunch) { Start-Process -FilePath 'onenote.exe' }
