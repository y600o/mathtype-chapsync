$ErrorActionPreference = 'SilentlyContinue'

if (Get-Process WINWORD) { Write-Host 'Close Word first, then run again.'; exit 1 }

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$file = Join-Path $startup 'MathTypeChapSync.dotm'
if (Test-Path $file) { Remove-Item $file -Force }
Remove-Item 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\LocationMTChapSync' -Recurse -Force
Remove-Item 'HKCU:\Software\VB and VBA Program Settings\MTChapSync' -Recurse -Force

Write-Host 'Uninstalled. The anchors already written into documents are harmless'
Write-Host '(they are invisible SEQ fields and can be deleted with Alt+F9 if needed).'
