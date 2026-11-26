param(
    [string]$Source = (Join-Path $PSScriptRoot 'MathTypeChapSync.dotm')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Source)) { throw "Template not found: $Source" }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close Word first.' }

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup | Out-Null }
Copy-Item $Source (Join-Path $startup 'MathTypeChapSync.dotm') -Force

$key = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\LocationMTChapSync'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name Path -Value ($startup + '\') -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name AllowSubfolders -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $key -Name Description -Value 'MathType chapter number sync add-in' -PropertyType String -Force | Out-Null

Write-Host ''
Write-Host 'Installed:' (Join-Path $startup 'MathTypeChapSync.dotm')
Write-Host 'Trusted location:' $startup
Write-Host ''
Write-Host 'Next: start Word, open a document, press Ctrl+S once.'
Write-Host 'Equations in chapter N should then read (N.1), (N.2) ...'
