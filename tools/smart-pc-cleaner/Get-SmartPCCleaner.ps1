<#
    Smart PC Cleaner - one-line web installer.
    Downloads the app from GitHub, installs it to %LOCALAPPDATA%\Programs\SmartPCCleaner,
    creates the desktop icon and launches the app. No admin rights needed to install.

    Run from any PowerShell prompt:
      irm https://raw.githubusercontent.com/evkiinc/evkiinc.github.io/claude/pc-performance-cleaner-lsnfs6/tools/smart-pc-cleaner/Get-SmartPCCleaner.ps1 | iex
#>

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

$baseUrl = 'https://raw.githubusercontent.com/evkiinc/evkiinc.github.io/claude/pc-performance-cleaner-lsnfs6/tools/smart-pc-cleaner'
$destDir = Join-Path $env:LOCALAPPDATA 'Programs\SmartPCCleaner'

Write-Host ''
Write-Host '  Smart PC Cleaner - installing...' -ForegroundColor Cyan
Write-Host ("  Destination: {0}" -f $destDir)

if (-not (Test-Path -LiteralPath $destDir)) {
    $null = New-Item -ItemType Directory -Path $destDir -Force
}

$files = @('SmartPCCleaner.ps1', 'SmartPCCleaner.ico', 'Install-SmartPCCleaner.bat', 'README.md')
foreach ($f in $files) {
    Write-Host ("  Downloading {0} ..." -f $f)
    Invoke-WebRequest -Uri ("{0}/{1}" -f $baseUrl, $f) -OutFile (Join-Path $destDir $f) -UseBasicParsing
}

$app = Join-Path $destDir 'SmartPCCleaner.ps1'
if (-not (Test-Path -LiteralPath $app)) { throw 'Download failed - SmartPCCleaner.ps1 not found after download.' }

Write-Host '  Creating desktop icon and launching the app...' -ForegroundColor Cyan
Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Sta -File `"$app`" -InstallShortcut"

Write-Host ''
Write-Host '  Done. The Smart PC Cleaner icon is now on your desktop - use it from now on.' -ForegroundColor Green
Write-Host ''
