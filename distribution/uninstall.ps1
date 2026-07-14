param([switch]$NoLaunch)

$ErrorActionPreference = 'Stop'
$clsid = '{88AB88AB-CDFB-4C68-9C3A-F10B75A5BC61}'
$targetRoot = Join-Path $env:LOCALAPPDATA 'OneMoreComments'
$target = Join-Path $targetRoot 'OneMoreAddIn'
$classes = 'Registry::HKEY_CURRENT_USER\Software\Classes'
$base = "$classes\CLSID\$clsid"
$inproc = "$base\InprocServer32"

$oneNote = Get-Process ONENOTE -ErrorAction SilentlyContinue
if ($oneNote) {
  $oneNote | ForEach-Object { [void]$_.CloseMainWindow() }
  $deadline = (Get-Date).AddSeconds(15)
  while ((Get-Process ONENOTE -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }
  if (Get-Process ONENOTE -ErrorAction SilentlyContinue) { throw 'OneNote is still running. Close it manually and run this uninstaller again.' }
}

$ours = $false
if (Test-Path -LiteralPath $inproc) {
  $codeBase = (Get-ItemProperty -LiteralPath $inproc -ErrorAction SilentlyContinue).CodeBase
  $ours = $codeBase -and $codeBase.StartsWith($target, [StringComparison]::OrdinalIgnoreCase)
}
if ($ours) {
  Remove-Item -LiteralPath $base -Recurse -Force
  $progId = "$classes\River.OneMoreAddIn"
  if (Test-Path -LiteralPath $progId) { Remove-Item -LiteralPath $progId -Recurse -Force }
} elseif (Test-Path -LiteralPath $base) {
  throw 'A different user-level OneMore COM override is present; it was not removed.'
}

$tray = Get-Process OneMoreTray -ErrorAction SilentlyContinue
foreach ($process in $tray) {
  try {
    if ($process.Path -and $process.Path.StartsWith($target, [StringComparison]::OrdinalIgnoreCase)) { $process | Stop-Process -Force }
  } catch {}
}
Start-Sleep -Milliseconds 500

$resolvedRoot = [IO.Path]::GetFullPath($targetRoot)
$expectedRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'OneMoreComments'))
if ($resolvedRoot -eq $expectedRoot -and (Test-Path -LiteralPath $resolvedRoot)) { Remove-Item -LiteralPath $resolvedRoot -Recurse -Force }

$addin = 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\River.OneMoreAddIn'
if (Test-Path -LiteralPath $addin) {
  Set-ItemProperty -LiteralPath $addin -Name LoadBehavior -Type DWord -Value 3
  Set-ItemProperty -LiteralPath $addin -Name FriendlyName -Value 'OneMoreAddIn'
  Set-ItemProperty -LiteralPath $addin -Name Description -Value 'Add-in for OneNote'
}
Write-Host 'OneMore Comments removed. The official OneMore registration is active again.' -ForegroundColor Green
if (-not $NoLaunch) { Start-Process -FilePath 'onenote.exe' }
