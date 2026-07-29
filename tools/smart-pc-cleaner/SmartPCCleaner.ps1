<#
    Smart PC Cleaner
    ----------------
    A safety-first Windows tune-up and organizer app (single-file PowerShell + WinForms, no dependencies).

    What it does:
      * System Cleanup   - previews then removes temp files, caches, dumps, update leftovers (skips anything in use or recent)
      * AI Tools         - finds Claude / Claude Code / ChatGPT / Codex / Cursor / VS Code data, tells you what is safe to
                           clean, what to keep (configs & credentials), and what is big enough to move to another drive
      * Move to G:       - finds large folders/files, checks whether they are safe to move, lets you rename them and
                           organize them into category/project folders on a target drive (copy -> verify -> delete)
      * Registry Care    - conservative: only clearly-orphaned entries, every change exported to a .reg backup first,
                           with one-click System Restore Point creation
      * Performance      - startup manager (reversible disable), background memory trim, DNS flush, leftover
                           agent-process cleanup, power plan switch

    Design rules (why this app will not hurt Windows):
      1. Nothing is deleted without an explicit scan -> checkbox -> confirmation flow.
      2. Files newer than a configurable age (default 24h) are never auto-deleted, and locked/in-use files are skipped,
         so running apps and AI agents are never tripped up mid-task.
      3. Registry changes are limited to a tiny, well-understood set of orphaned entries; every key is exported to a
         .reg file before deletion and a System Restore Point can be created first. MSI-registered apps are never touched.
      4. Moves are copy -> verify (file count + bytes) -> delete, logged to a CSV, with an optional shortcut left behind.
      5. Hard denylist: Windows, Program Files, ProgramData, AppData and installed-application folders are never movable.

    Usage:  right-click > Run with PowerShell, or run Install-SmartPCCleaner.bat once to get a desktop icon.
#>

[CmdletBinding()]
param(
    [switch]$InstallShortcut,
    [switch]$Elevated
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ============================================================================
#  App constants / folders
# ============================================================================
$script:AppName    = 'Smart PC Cleaner'
$script:AppVersion = '1.0.0'
$script:AppDir     = Join-Path $env:LOCALAPPDATA 'SmartPCCleaner'
$script:LogDir     = Join-Path $script:AppDir 'logs'
$script:BackupDir  = Join-Path $script:AppDir 'registry-backups'
$script:DisabledStartupDir = Join-Path $script:AppDir 'disabled-startup'
$script:MoveLogCsv = Join-Path $script:AppDir 'move-history.csv'
$script:LogFile    = Join-Path $script:LogDir ("session_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

foreach ($d in @($script:AppDir, $script:LogDir, $script:BackupDir, $script:DisabledStartupDir)) {
    if (-not (Test-Path -LiteralPath $d)) { $null = New-Item -ItemType Directory -Path $d -Force -ErrorAction SilentlyContinue }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================================
#  Core helpers
# ============================================================================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try { Add-Content -LiteralPath $script:LogFile -Value $line -ErrorAction SilentlyContinue } catch { }
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} B' -f $Bytes)
}

function Get-PathSize {
    # Size in bytes of a file or a whole folder tree. Fast COM path first, PowerShell fallback.
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return [long]0 }
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        if (-not $item.PSIsContainer) { return [long]$item.Length }
    } catch { return [long]0 }
    try {
        $fso = New-Object -ComObject Scripting.FileSystemObject
        return [long]$fso.GetFolder($Path).Size
    } catch {
        try {
            $m = Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum
            if ($m -and $m.Sum) { return [long]$m.Sum }
        } catch { }
        return [long]0
    }
}

function Test-DangerousPath {
    # Absolute guard used before ANY delete: refuse roots and ultra-short paths.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $true }
    $p = $Path.TrimEnd('\')
    if ($p.Length -lt 8) { return $true }                     # e.g. "C:", "C:\Users" style roots
    if ($p -match '^[A-Za-z]:$') { return $true }
    if ($p -ieq $env:windir) { return $true }
    if ($p -ieq $env:USERPROFILE) { return $true }
    if ($p -ieq ${env:ProgramFiles}) { return $true }
    if ($p -ieq ${env:ProgramFiles(x86)}) { return $true }
    if ($p -ieq $env:ProgramData) { return $true }
    return $false
}

function Remove-ContentsSafe {
    <#  Deletes files inside $Path (not the folder itself).
        - Skips files newer than $MinAgeHours (protects active app/agent state)
        - Skips locked / permission-denied files silently
        - Optional $Include name patterns (e.g. thumbcache_*.db)
        - Removes only sub-directories that end up empty                       #>
    param(
        [string]$Path,
        [double]$MinAgeHours = 24,
        [string[]]$Include = $null
    )
    $result = [pscustomobject]@{ FreedBytes = [long]0; Deleted = 0; Skipped = 0 }
    if (Test-DangerousPath $Path) { Write-Log "REFUSED unsafe clean path: $Path" 'WARN'; return $result }
    if (-not (Test-Path -LiteralPath $Path)) { return $result }

    $cutoff = (Get-Date).AddHours(-1 * $MinAgeHours)
    $files = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
    if ($Include -and $Include.Count -gt 0) {
        $files = @($files | Where-Object {
            $name = $_.Name
            @($Include | Where-Object { $name -like $_ }).Count -gt 0
        })
    }
    foreach ($f in $files) {
        if ($f.LastWriteTime -gt $cutoff) { $result.Skipped++; continue }
        $len = [long]$f.Length
        try {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            $result.FreedBytes += $len
            $result.Deleted++
        } catch { $result.Skipped++ }
    }
    # sweep empty sub-folders, deepest first (never the root we were given)
    try {
        $dirs = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -Directory -ErrorAction SilentlyContinue |
                  Sort-Object { $_.FullName.Length } -Descending)
        foreach ($d in $dirs) {
            try {
                if (-not (Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue)) {
                    Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
                }
            } catch { }
        }
    } catch { }
    Write-Log ("Cleaned '{0}': freed {1}, deleted {2}, skipped {3}" -f $Path, (Format-Bytes $result.FreedBytes), $result.Deleted, $result.Skipped)
    return $result
}

function New-SafetyRestorePoint {
    if (-not (Test-IsAdmin)) {
        [System.Windows.Forms.MessageBox]::Show('Creating a System Restore Point requires running as Administrator.',
            $script:AppName, 'OK', 'Warning') | Out-Null
        return $false
    }
    try {
        Checkpoint-Computer -Description 'Smart PC Cleaner' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        Write-Log 'System Restore Point created.'
        [System.Windows.Forms.MessageBox]::Show('System Restore Point created successfully.', $script:AppName, 'OK', 'Information') | Out-Null
        return $true
    } catch {
        Write-Log ("Restore point failed: {0}" -f $_.Exception.Message) 'WARN'
        $msg = ("Windows did not create a new restore point:`n{0}`n`nNote: Windows only allows one restore point every 24 hours " +
                "(a recent one still protects you), and System Protection must be enabled on drive C:.")
        [System.Windows.Forms.MessageBox]::Show(($msg -f $_.Exception.Message), $script:AppName, 'OK', 'Warning') | Out-Null
        return $false
    }
}

function New-DesktopShortcut {
    try {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $lnkPath = Join-Path $desktop ($script:AppName + '.lnk')
        $ws = New-Object -ComObject WScript.Shell
        $sc = $ws.CreateShortcut($lnkPath)
        $sc.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $sc.Arguments  = "-NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$PSCommandPath`""
        $sc.WorkingDirectory = Split-Path -Parent $PSCommandPath
        $icon = "$env:SystemRoot\System32\cleanmgr.exe"
        if (Test-Path -LiteralPath $icon) { $sc.IconLocation = "$icon,0" }
        else { $sc.IconLocation = "$env:SystemRoot\System32\shell32.dll,21" }
        $sc.Description = 'Smart PC Cleaner - safe cleanup, AI-tool tidy-up, mover and performance care'
        $sc.Save()
        Write-Log "Desktop shortcut created: $lnkPath"
        return $true
    } catch {
        Write-Log ("Shortcut creation failed: {0}" -f $_.Exception.Message) 'ERROR'
        return $false
    }
}

# ============================================================================
#  Elevation / install-shortcut startup paths
# ============================================================================
if ($InstallShortcut) {
    $ok = New-DesktopShortcut
    if ($ok) {
        [System.Windows.Forms.MessageBox]::Show(
            "Smart PC Cleaner is installed.`n`nA desktop icon has been created - use it any time.`nThe app will open now.",
            $script:AppName, 'OK', 'Information') | Out-Null
    }
}

if (-not (Test-IsAdmin) -and -not $Elevated -and -not $InstallShortcut) {
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("Run as Administrator?`n`nAdministrator mode unlocks: Windows temp / update cache cleanup, restore points, " +
         "machine-wide startup entries and registry care.`n`nChoose No to run with user-level features only."),
        $script:AppName, 'YesNo', 'Question')
    if ($answer -eq 'Yes') {
        try {
            Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Sta -File `"$PSCommandPath`" -Elevated" -Verb RunAs | Out-Null
            exit
        } catch {
            Write-Log 'Elevation declined at UAC prompt; continuing unelevated.' 'WARN'
        }
    }
}
$script:IsAdmin = Test-IsAdmin
Write-Log ("{0} v{1} started. Admin={2}" -f $script:AppName, $script:AppVersion, $script:IsAdmin)

# ============================================================================
#  Cleanup target definitions (System Cleanup tab)
# ============================================================================
function Get-CleanupTargets {
    $t = @()
    $t += [pscustomobject]@{ Key='UserTemp';  Name='User temp files';            Admin=$false; Special=$null; MinAgeHours=24;
        Paths=@($env:TEMP); Include=$null;
        Desc='Files in your %TEMP% older than 24h. In-use and recent files are skipped so nothing running breaks.' }
    $t += [pscustomobject]@{ Key='WinTemp';   Name='Windows temp files';         Admin=$true;  Special=$null; MinAgeHours=24;
        Paths=@("$env:windir\Temp"); Include=$null;
        Desc='System-wide temp folder (needs Administrator).' }
    $t += [pscustomobject]@{ Key='Thumbs';    Name='Thumbnail / icon cache';     Admin=$false; Special=$null; MinAgeHours=0;
        Paths=@("$env:LOCALAPPDATA\Microsoft\Windows\Explorer"); Include=@('thumbcache_*.db','iconcache_*.db');
        Desc='Explorer rebuilds these automatically. Files locked by Explorer are skipped.' }
    $t += [pscustomobject]@{ Key='DXCache';   Name='DirectX / GPU shader caches'; Admin=$false; Special=$null; MinAgeHours=0;
        Paths=@("$env:LOCALAPPDATA\D3DSCache", "$env:LOCALAPPDATA\NVIDIA\DXCache", "$env:LOCALAPPDATA\NVIDIA\GLCache",
                "$env:LOCALAPPDATA\AMD\DxCache"); Include=$null;
        Desc='Shader caches are rebuilt on demand; games may load a little slower once.' }
    $t += [pscustomobject]@{ Key='CrashDumps'; Name='Crash dumps & minidumps';   Admin=$true;  Special=$null; MinAgeHours=0;
        Paths=@("$env:LOCALAPPDATA\CrashDumps", "$env:windir\Minidump"); Include=$null;
        Desc='Old crash reports. Keep them only if you are actively debugging a crash.' }
    $t += [pscustomobject]@{ Key='WER';       Name='Windows Error Reporting queue'; Admin=$true; Special=$null; MinAgeHours=0;
        Paths=@("$env:ProgramData\Microsoft\Windows\WER\ReportQueue", "$env:ProgramData\Microsoft\Windows\WER\ReportArchive",
                "$env:ProgramData\Microsoft\Windows\WER\Temp"); Include=$null;
        Desc='Queued/archived error reports.' }
    $t += [pscustomobject]@{ Key='CBSLogs';   Name='Old Windows servicing logs (30d+)'; Admin=$true; Special=$null; MinAgeHours=720;
        Paths=@("$env:windir\Logs\CBS"); Include=$null;
        Desc='Component servicing logs older than 30 days.' }
    $t += [pscustomobject]@{ Key='EdgeCache'; Name='Microsoft Edge cache';       Admin=$false; Special=$null; MinAgeHours=0;
        Paths=@("$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\Cache_Data",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"); Include=$null;
        Desc='Browser cache only - never history, passwords or cookies. Close Edge first for a fuller clean.' }
    $t += [pscustomobject]@{ Key='ChromeCache'; Name='Google Chrome cache';      Admin=$false; Special=$null; MinAgeHours=0;
        Paths=@("$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\Cache_Data",
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache"); Include=$null;
        Desc='Browser cache only - never history, passwords or cookies. Close Chrome first for a fuller clean.' }
    $t += [pscustomobject]@{ Key='WUCache';   Name='Windows Update download cache'; Admin=$true; Special='WU'; MinAgeHours=0;
        Paths=@("$env:windir\SoftwareDistribution\Download"); Include=$null;
        Desc='Already-installed update payloads. The update service is stopped, cleaned, then restarted.' }
    $t += [pscustomobject]@{ Key='DO';        Name='Delivery Optimization cache'; Admin=$true; Special='DO'; MinAgeHours=0;
        Paths=@(); Include=$null;
        Desc='Peer-to-peer update cache, removed via the official Windows command.' }
    $t += [pscustomobject]@{ Key='RecycleBin'; Name='Recycle Bin';               Admin=$false; Special='RecycleBin'; MinAgeHours=0;
        Paths=@(); Include=$null;
        Desc='Empties the Recycle Bin. This is permanent - review the bin first if unsure.' }
    return $t
}

function Get-RecycleBinSize {
    try {
        $shell = New-Object -ComObject Shell.Application
        $rb = $shell.Namespace(0xA)
        $size = [long]0
        foreach ($item in @($rb.Items())) {
            try { $size += [long]$item.Size } catch { }
        }
        return $size
    } catch { return [long]0 }
}

function Clear-WindowsUpdateCache {
    $r = [pscustomobject]@{ FreedBytes = [long]0; Deleted = 0; Skipped = 0 }
    if (-not $script:IsAdmin) { return $r }
    try { Stop-Service -Name wuauserv, bits -Force -ErrorAction SilentlyContinue } catch { }
    Start-Sleep -Milliseconds 800
    $r = Remove-ContentsSafe -Path "$env:windir\SoftwareDistribution\Download" -MinAgeHours 0
    try { Start-Service -Name bits, wuauserv -ErrorAction SilentlyContinue } catch { }
    return $r
}

# ============================================================================
#  AI tool data map (AI Tools tab)
#     Action: Clean  = cache/log data, safe to remove (respecting age)
#             Review = user data (sessions, models) - keep or MOVE, never auto-delete
#             Keep   = configs & credentials - shown so you know NOT to touch them
# ============================================================================
function Get-AITargets {
    $home_ = $env:USERPROFILE
    $rows = @()

    # --- Claude Code (CLI/agent)
    $rows += [pscustomobject]@{ Tool='Claude Code'; Item='Session transcripts (projects)'; Path="$home_\.claude\projects";        Action='Review'; MinAgeHours=0;   Note='Chat/session history per project. Big over time - move or archive, do not blind-delete.' }
    $rows += [pscustomobject]@{ Tool='Claude Code'; Item='Shell snapshots';                Path="$home_\.claude\shell-snapshots"; Action='Clean';  MinAgeHours=72;  Note='Temporary shell state snapshots older than 3 days.' }
    $rows += [pscustomobject]@{ Tool='Claude Code'; Item='Telemetry cache (statsig)';      Path="$home_\.claude\statsig";         Action='Clean';  MinAgeHours=24;  Note='Feature-flag/telemetry cache, safely rebuilt.' }
    $rows += [pscustomobject]@{ Tool='Claude Code'; Item='Old todo lists';                 Path="$home_\.claude\todos";           Action='Clean';  MinAgeHours=720; Note='Per-session todo state older than 30 days.' }
    $rows += [pscustomobject]@{ Tool='Claude Code'; Item='Config + credentials';           Path="$home_\.claude.json";            Action='Keep';   MinAgeHours=0;   Note='NEVER delete - holds your settings and login.' }

    # --- Claude Desktop
    $rows += [pscustomobject]@{ Tool='Claude Desktop'; Item='Renderer cache';   Path="$env:APPDATA\Claude\Cache";      Action='Clean'; MinAgeHours=24; Note='Electron cache, rebuilt automatically. Close Claude first.' }
    $rows += [pscustomobject]@{ Tool='Claude Desktop'; Item='Code cache';       Path="$env:APPDATA\Claude\Code Cache"; Action='Clean'; MinAgeHours=24; Note='Compiled JS cache.' }
    $rows += [pscustomobject]@{ Tool='Claude Desktop'; Item='GPU cache';        Path="$env:APPDATA\Claude\GPUCache";   Action='Clean'; MinAgeHours=24; Note='GPU shader cache.' }
    $rows += [pscustomobject]@{ Tool='Claude Desktop'; Item='Settings & login'; Path="$env:APPDATA\Claude\Local Storage"; Action='Keep'; MinAgeHours=0; Note='Keeps you signed in - do not delete.' }

    # --- ChatGPT desktop (Store app)
    $pkgRoot = "$env:LOCALAPPDATA\Packages"
    if (Test-Path -LiteralPath $pkgRoot) {
        foreach ($pkg in @(Get-ChildItem -LiteralPath $pkgRoot -Directory -Filter 'OpenAI.ChatGPT*' -ErrorAction SilentlyContinue)) {
            $rows += [pscustomobject]@{ Tool='ChatGPT Desktop'; Item='Local cache';         Path=(Join-Path $pkg.FullName 'LocalCache'); Action='Clean'; MinAgeHours=24; Note='App cache, rebuilt automatically. Close ChatGPT first.' }
            $rows += [pscustomobject]@{ Tool='ChatGPT Desktop'; Item='App state & login';   Path=(Join-Path $pkg.FullName 'LocalState'); Action='Keep';  MinAgeHours=0;  Note='Sign-in and app state - do not delete.' }
        }
    }
    $rows += [pscustomobject]@{ Tool='ChatGPT Desktop'; Item='Electron cache'; Path="$env:APPDATA\ChatGPT\Cache"; Action='Clean'; MinAgeHours=24; Note='Cache for the non-Store desktop build.' }

    # --- OpenAI Codex CLI
    $rows += [pscustomobject]@{ Tool='Codex CLI'; Item='Session history';  Path="$home_\.codex\sessions"; Action='Review'; MinAgeHours=0;   Note='Past agent sessions - move or archive rather than delete.' }
    $rows += [pscustomobject]@{ Tool='Codex CLI'; Item='Logs';             Path="$home_\.codex\log";      Action='Clean';  MinAgeHours=168; Note='Log files older than 7 days.' }
    $rows += [pscustomobject]@{ Tool='Codex CLI'; Item='Auth & config';    Path="$home_\.codex\auth.json"; Action='Keep';  MinAgeHours=0;   Note='Login token - do not delete.' }

    # --- Cursor
    $rows += [pscustomobject]@{ Tool='Cursor'; Item='Cache';        Path="$env:APPDATA\Cursor\Cache";       Action='Clean'; MinAgeHours=24; Note='Editor cache, rebuilt automatically.' }
    $rows += [pscustomobject]@{ Tool='Cursor'; Item='Cached data';  Path="$env:APPDATA\Cursor\CachedData";  Action='Clean'; MinAgeHours=24; Note='Update/workbench cache.' }
    $rows += [pscustomobject]@{ Tool='Cursor'; Item='Code cache';   Path="$env:APPDATA\Cursor\Code Cache";  Action='Clean'; MinAgeHours=24; Note='Compiled JS cache.' }
    $rows += [pscustomobject]@{ Tool='Cursor'; Item='GPU cache';    Path="$env:APPDATA\Cursor\GPUCache";    Action='Clean'; MinAgeHours=24; Note='GPU shader cache.' }
    $rows += [pscustomobject]@{ Tool='Cursor'; Item='User settings'; Path="$env:APPDATA\Cursor\User";       Action='Keep';  MinAgeHours=0;  Note='Your settings, keybindings, extensions state.' }

    # --- VS Code (agents like Copilot / Claude Code extension run here)
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='Cache';                Path="$env:APPDATA\Code\Cache";                 Action='Clean'; MinAgeHours=24;  Note='Rebuilt automatically. Close VS Code first.' }
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='Cached data';          Path="$env:APPDATA\Code\CachedData";            Action='Clean'; MinAgeHours=24;  Note='Workbench cache.' }
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='Code cache';           Path="$env:APPDATA\Code\Code Cache";            Action='Clean'; MinAgeHours=24;  Note='Compiled JS cache.' }
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='Downloaded VSIX cache'; Path="$env:APPDATA\Code\CachedExtensionVSIXs"; Action='Clean'; MinAgeHours=0;   Note='Extension installers already applied.' }
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='Old logs';             Path="$env:APPDATA\Code\logs";                  Action='Clean'; MinAgeHours=168; Note='Log sessions older than 7 days.' }
    $rows += [pscustomobject]@{ Tool='VS Code'; Item='User settings';        Path="$env:APPDATA\Code\User";                  Action='Keep';  MinAgeHours=0;   Note='Your settings and snippets.' }

    # --- GitHub Copilot
    $rows += [pscustomobject]@{ Tool='GitHub Copilot'; Item='Auth & versions'; Path="$env:LOCALAPPDATA\github-copilot"; Action='Keep'; MinAgeHours=0; Note='Holds Copilot sign-in - do not delete.' }

    # --- Developer package caches (agents fill these fast)
    $rows += [pscustomobject]@{ Tool='Dev caches'; Item='npm cache';        Path="$env:LOCALAPPDATA\npm-cache";  Action='Clean'; MinAgeHours=720; Note='Package cache older than 30 days; npm re-downloads on demand.' }
    $rows += [pscustomobject]@{ Tool='Dev caches'; Item='pip cache';        Path="$env:LOCALAPPDATA\pip\cache";  Action='Clean'; MinAgeHours=720; Note='Python wheel cache; pip re-downloads on demand.' }
    $rows += [pscustomobject]@{ Tool='Dev caches'; Item='Yarn cache';       Path="$env:LOCALAPPDATA\Yarn\Cache"; Action='Clean'; MinAgeHours=720; Note='Yarn package cache.' }
    $rows += [pscustomobject]@{ Tool='Dev caches'; Item='uv cache';         Path="$env:LOCALAPPDATA\uv\cache";   Action='Clean'; MinAgeHours=720; Note='uv package cache.' }
    $rows += [pscustomobject]@{ Tool='Dev caches'; Item='Playwright browsers'; Path="$env:LOCALAPPDATA\ms-playwright"; Action='Review'; MinAgeHours=0; Note='Browser binaries used by test agents. Delete only if you re-install via "playwright install".' }

    # --- Large AI model stores (move candidates, never auto-delete)
    $rows += [pscustomobject]@{ Tool='AI models'; Item='Ollama models';       Path="$home_\.ollama\models";      Action='Review'; MinAgeHours=0; Note='Local LLM weights - often many GB. Good candidate to MOVE to another drive.' }
    $rows += [pscustomobject]@{ Tool='AI models'; Item='HuggingFace cache';   Path="$home_\.cache\huggingface";  Action='Review'; MinAgeHours=0; Note='Downloaded model weights. Movable; re-downloads if missing.' }
    $rows += [pscustomobject]@{ Tool='AI models'; Item='Torch hub cache';     Path="$home_\.cache\torch";        Action='Review'; MinAgeHours=0; Note='PyTorch model cache.' }

    # keep only rows that exist on this machine
    return @($rows | Where-Object { Test-Path -LiteralPath $_.Path })
}

# ============================================================================
#  Mover (Move to G:) helpers
# ============================================================================
function Get-InstalledAppLocations {
    if ($script:InstalledLocations) { return $script:InstalledLocations }
    $locs = New-Object System.Collections.Generic.List[string]
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in @(Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            try {
                $il = (Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue).InstallLocation
                if ($il -and $il.Trim()) { $locs.Add($il.Trim().TrimEnd('\')) }
            } catch { }
        }
    }
    $script:InstalledLocations = $locs
    return $locs
}

function Test-MoveSafety {
    # Returns @{ Movable = bool; Reason = string }
    param([string]$Path)
    try {
        $p = $Path.TrimEnd('\')
        $protected = @($env:windir, ${env:ProgramFiles}, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:APPDATA, $env:LOCALAPPDATA)
        foreach ($root in $protected) {
            if ($root -and ($p -ieq $root -or $p.StartsWith(($root + '\'), [System.StringComparison]::OrdinalIgnoreCase))) {
                return @{ Movable = $false; Reason = 'Inside a system/application area (Windows, Program Files, AppData)' }
            }
        }
        $item = Get-Item -LiteralPath $p -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            return @{ Movable = $false; Reason = 'Link/junction - moving it would break the link target' }
        }
        $attrVal = [int]$item.Attributes
        if ((($attrVal -band 0x00400000) -ne 0) -or (($attrVal -band 0x00040000) -ne 0) -or
            (($item.Attributes -band [IO.FileAttributes]::Offline) -ne 0)) {
            return @{ Movable = $false; Reason = 'Cloud placeholder (OneDrive online-only) - download it first' }
        }
        foreach ($loc in (Get-InstalledAppLocations)) {
            if ($p -ieq $loc) {
                return @{ Movable = $false; Reason = 'Registered installed application folder' }
            }
        }
        if ($item.PSIsContainer) {
            $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
                $_.Path -and $_.Path.StartsWith(($p + '\'), [System.StringComparison]::OrdinalIgnoreCase)
            })
            if ($running.Count -gt 0) {
                return @{ Movable = $false; Reason = ('A program inside is currently running ({0})' -f $running[0].ProcessName) }
            }
        }
        return @{ Movable = $true; Reason = 'OK to move' }
    } catch {
        return @{ Movable = $false; Reason = ('Could not inspect: {0}' -f $_.Exception.Message) }
    }
}

function Get-SuggestedCategory {
    param([string]$Path)
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $name = $item.Name
        if ($name -match '(?i)claude|chatgpt|openai|codex|anthropic|copilot') { return 'AI-Archives' }
        if (-not $item.PSIsContainer) {
            $ext = $item.Extension.ToLowerInvariant()
            if ($ext -in @('.mp4','.mov','.mkv','.avi','.wmv','.webm')) { return 'Media\Videos' }
            if ($ext -in @('.jpg','.jpeg','.png','.gif','.bmp','.heic','.raw','.tiff','.psd')) { return 'Media\Images' }
            if ($ext -in @('.mp3','.wav','.flac','.m4a','.aac')) { return 'Media\Audio' }
            if ($ext -in @('.zip','.7z','.rar','.iso','.tar','.gz')) { return 'Archives' }
            if ($ext -in @('.exe','.msi','.msix')) { return 'Installers' }
            if ($ext -in @('.pdf','.docx','.doc','.xlsx','.pptx','.txt','.md','.csv')) { return 'Documents' }
            return 'Misc'
        }
        # project markers
        $markers = @('.git','package.json','pyproject.toml','requirements.txt','*.sln','*.csproj','Cargo.toml','go.mod')
        foreach ($m in $markers) {
            if (@(Get-ChildItem -LiteralPath $Path -Filter $m -Force -ErrorAction SilentlyContinue).Count -gt 0) { return 'Projects' }
        }
        # majority extension vote over a sample of files
        $sample = @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue | Select-Object -First 400)
        if ($sample.Count -gt 0) {
            $video  = @($sample | Where-Object { $_.Extension -match '(?i)^\.(mp4|mov|mkv|avi|wmv|webm)$' }).Count
            $image  = @($sample | Where-Object { $_.Extension -match '(?i)^\.(jpg|jpeg|png|gif|bmp|heic|raw|tiff|psd)$' }).Count
            $audio  = @($sample | Where-Object { $_.Extension -match '(?i)^\.(mp3|wav|flac|m4a|aac)$' }).Count
            $docs   = @($sample | Where-Object { $_.Extension -match '(?i)^\.(pdf|docx|doc|xlsx|pptx|txt|md|csv)$' }).Count
            $arch   = @($sample | Where-Object { $_.Extension -match '(?i)^\.(zip|7z|rar|iso|tar|gz)$' }).Count
            $half = [Math]::Max(1, [int]($sample.Count / 2))
            if ($video -ge $half) { return 'Media\Videos' }
            if ($image -ge $half) { return 'Media\Images' }
            if ($audio -ge $half) { return 'Media\Audio' }
            if ($docs  -ge $half) { return 'Documents' }
            if ($arch  -ge $half) { return 'Archives' }
        }
        return 'Misc'
    } catch { return 'Misc' }
}

function Get-CleanName {
    param([string]$Name)
    $clean = [Regex]::Replace($Name, '[<>:"/\\|?*\x00-\x1f]', ' ')
    $clean = ($clean -replace '\s{2,}', ' ').Trim().TrimEnd('.')
    if (-not $clean) { $clean = 'Untitled' }
    return $clean
}

function Invoke-SafeMove {
    <#  Copy -> verify (count + bytes) -> delete source -> optional shortcut, logged to CSV.
        Returns @{ Success; Message; Dest }                                          #>
    param(
        [string]$Source,
        [string]$DestFolder,      # full destination folder that will CONTAIN the item
        [string]$NewName,
        [bool]$LeaveShortcut = $true
    )
    try {
        if (-not (Test-Path -LiteralPath $Source)) { return @{ Success=$false; Message='Source no longer exists'; Dest='' } }
        $safety = Test-MoveSafety -Path $Source
        if (-not $safety.Movable) { return @{ Success=$false; Message=$safety.Reason; Dest='' } }

        if (-not (Test-Path -LiteralPath $DestFolder)) {
            $null = New-Item -ItemType Directory -Path $DestFolder -Force -ErrorAction Stop
        }
        $destDrive = [IO.Path]::GetPathRoot($DestFolder)
        $srcSize = Get-PathSize -Path $Source
        $free = (New-Object IO.DriveInfo($destDrive)).AvailableFreeSpace
        if ($free -lt ($srcSize + 200MB)) {
            return @{ Success=$false; Message=("Not enough free space on {0} ({1} needed)" -f $destDrive, (Format-Bytes $srcSize)); Dest='' }
        }

        $dest = Join-Path $DestFolder $NewName
        $n = 1
        while (Test-Path -LiteralPath $dest) { $dest = Join-Path $DestFolder ("{0} ({1})" -f $NewName, $n); $n++ }

        $srcItem = Get-Item -LiteralPath $Source -Force
        if ($srcItem.PSIsContainer) {
            # copy with robocopy (junctions excluded), then verify
            $rcLog = Join-Path $script:LogDir 'robocopy-last.log'
            $rcArgs = @(('"{0}"' -f $Source), ('"{0}"' -f $dest), '/E', '/COPY:DAT', '/DCOPY:DAT', '/R:2', '/W:2', '/XJ',
                        '/NFL', '/NDL', '/NJH', '/NJS', ('/LOG:"{0}"' -f $rcLog))
            $proc = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList $rcArgs -Wait -PassThru -WindowStyle Hidden
            if ($proc.ExitCode -ge 8) {
                return @{ Success=$false; Message=("Copy failed (robocopy code {0}). Source untouched." -f $proc.ExitCode); Dest=$dest }
            }
            $srcFiles  = @(Get-ChildItem -LiteralPath $Source -Recurse -Force -File -ErrorAction SilentlyContinue)
            $destFiles = @(Get-ChildItem -LiteralPath $dest   -Recurse -Force -File -ErrorAction SilentlyContinue)
            $srcSum  = ($srcFiles  | Measure-Object -Property Length -Sum).Sum
            $destSum = ($destFiles | Measure-Object -Property Length -Sum).Sum
            if ($null -eq $srcSum)  { $srcSum  = 0 }
            if ($null -eq $destSum) { $destSum = 0 }
            if (($destFiles.Count -lt $srcFiles.Count) -or ($destSum -lt $srcSum)) {
                return @{ Success=$false;
                          Message=("Verification failed ({0}/{1} files copied). BOTH copies kept - nothing deleted." -f $destFiles.Count, $srcFiles.Count);
                          Dest=$dest }
            }
            Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop
        } else {
            Copy-Item -LiteralPath $Source -Destination $dest -Force -ErrorAction Stop
            $a = (Get-Item -LiteralPath $Source -Force).Length
            $b = (Get-Item -LiteralPath $dest -Force).Length
            if ($b -ne $a) {
                return @{ Success=$false; Message='Verification failed (size mismatch). Both copies kept.'; Dest=$dest }
            }
            Remove-Item -LiteralPath $Source -Force -ErrorAction Stop
        }

        if ($LeaveShortcut) {
            try {
                $ws = New-Object -ComObject WScript.Shell
                $lnk = $ws.CreateShortcut(($Source + '.lnk'))
                $lnk.TargetPath = $dest
                $lnk.Description = 'Moved by Smart PC Cleaner'
                $lnk.Save()
            } catch { }
        }

        # move history for undo reference
        try {
            if (-not (Test-Path -LiteralPath $script:MoveLogCsv)) {
                Set-Content -LiteralPath $script:MoveLogCsv -Value '"Timestamp","Source","Destination","Size"'
            }
            $line = '"{0}","{1}","{2}","{3}"' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Source, $dest, (Format-Bytes $srcSize)
            Add-Content -LiteralPath $script:MoveLogCsv -Value $line
        } catch { }

        Write-Log ("Moved '{0}' -> '{1}' ({2})" -f $Source, $dest, (Format-Bytes $srcSize))
        return @{ Success=$true; Message=('Moved ({0})' -f (Format-Bytes $srcSize)); Dest=$dest }
    } catch {
        return @{ Success=$false; Message=$_.Exception.Message; Dest='' }
    }
}

# ============================================================================
#  Registry care helpers (conservative by design)
# ============================================================================
function ConvertTo-PSRegPath {
    param([string]$KeyName)   # 'HKEY_LOCAL_MACHINE\Sub\Key' -> 'HKLM:\Sub\Key'
    $map = @{ 'HKEY_LOCAL_MACHINE' = 'HKLM:'; 'HKEY_CURRENT_USER' = 'HKCU:'; 'HKEY_CLASSES_ROOT' = 'HKCR:'; 'HKEY_USERS' = 'HKU:' }
    foreach ($k in $map.Keys) {
        if ($KeyName.StartsWith($k)) { return ($map[$k] + $KeyName.Substring($k.Length)) }
    }
    return $KeyName
}

function Backup-RegistryKey {
    param([string]$KeyName)   # native form: HKEY_LOCAL_MACHINE\...
    try {
        $safe = ($KeyName -replace '[\\/:*?"<>| ]', '_')
        if ($safe.Length -gt 110) { $safe = $safe.Substring($safe.Length - 110) }
        $file = Join-Path $script:BackupDir ("{0}_{1}.reg" -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $safe)
        $null = & "$env:SystemRoot\System32\reg.exe" export $KeyName $file /y 2>&1
        if (Test-Path -LiteralPath $file) { return $file }
        return $null
    } catch { return $null }
}

function Get-ExeFromCommand {
    param([string]$Command)
    if ([string]::IsNullOrWhiteSpace($Command)) { return $null }
    $c = [Environment]::ExpandEnvironmentVariables($Command.Trim())
    $m = [Regex]::Match($c, '^\s*"([^"]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return ($c -split '\s+')[0]
}

function Get-OrphanedAppPaths {
    $found = @()
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths'
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in @(Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            try {
                $props = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $default = $null
                if ($props) { $default = $props.'(default)' }
                if (-not $default) { continue }
                $exe = [Environment]::ExpandEnvironmentVariables(([string]$default).Trim('"').Trim())
                if ($exe -and -not (Test-Path -LiteralPath $exe)) {
                    $found += [pscustomobject]@{
                        Type   = 'Orphaned App Path'
                        KeyName= $k.Name
                        Detail = ("Points to missing program: {0}" -f $exe)
                    }
                }
            } catch { }
        }
    }
    return $found
}

function Get-OrphanedUninstallEntries {
    $found = @()
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in @(Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            try {
                $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                if (-not $p) { continue }
                if (-not ($p.PSObject.Properties['DisplayName'] -and $p.DisplayName)) { continue }
                if ($p.PSObject.Properties['SystemComponent'] -and $p.SystemComponent -eq 1) { continue }
                if (-not ($p.PSObject.Properties['UninstallString'] -and $p.UninstallString)) { continue }
                if ($p.UninstallString -match '(?i)msiexec') { continue }   # MSI-registered: never touch
                $exe = Get-ExeFromCommand -Command $p.UninstallString
                if (-not $exe) { continue }
                $exeMissing = -not (Test-Path -LiteralPath $exe)
                $il = $null
                if ($p.PSObject.Properties['InstallLocation']) { $il = $p.InstallLocation }
                $ilMissing = $true
                if ($il -and $il.Trim()) { $ilMissing = -not (Test-Path -LiteralPath $il.Trim()) }
                if ($exeMissing -and $ilMissing) {
                    $found += [pscustomobject]@{
                        Type    = 'Orphaned uninstall entry'
                        KeyName = $k.Name
                        Detail  = ("'{0}' - uninstaller and install folder both missing" -f $p.DisplayName)
                    }
                }
            } catch { }
        }
    }
    return $found
}

function Remove-RegistryKeySafe {
    param([string]$KeyName)
    $backup = Backup-RegistryKey -KeyName $KeyName
    if (-not $backup) {
        Write-Log ("SKIPPED registry delete (backup failed): {0}" -f $KeyName) 'WARN'
        return @{ Success = $false; Message = 'Backup export failed - key NOT deleted' }
    }
    try {
        $ps = ConvertTo-PSRegPath -KeyName $KeyName
        Remove-Item -Path $ps -Recurse -Force -ErrorAction Stop
        Write-Log ("Registry key removed (backup: {0}): {1}" -f (Split-Path -Leaf $backup), $KeyName)
        return @{ Success = $true; Message = ('Removed (backup: {0})' -f (Split-Path -Leaf $backup)) }
    } catch {
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

function Clear-ExplorerMRU {
    # Privacy/no-risk cleanup: recent-file lists. Backed up first; Windows recreates the keys.
    $keys = @(
        'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs',
        'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU'
    )
    $done = 0
    foreach ($k in $keys) {
        $ps = ConvertTo-PSRegPath -KeyName $k
        if (-not (Test-Path $ps)) { continue }
        $backup = Backup-RegistryKey -KeyName $k
        if (-not $backup) { continue }
        try {
            Remove-Item -Path $ps -Recurse -Force -ErrorAction Stop
            $done++
        } catch { }
    }
    return $done
}

# ============================================================================
#  Performance helpers
# ============================================================================
$script:StartupBackupKey = 'HKCU:\Software\SmartPCCleaner\DisabledStartup'

function Get-StartupEntries {
    $rows = @()
    $regSources = @(
        @{ Source='HKCU Run';   Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run';               Admin=$false },
        @{ Source='HKLM Run';   Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run';               Admin=$true  },
        @{ Source='HKLM Run32'; Path='HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run';   Admin=$true  }
    )
    foreach ($src in $regSources) {
        if (-not (Test-Path $src.Path)) { continue }
        try {
            $props = Get-ItemProperty -Path $src.Path -ErrorAction SilentlyContinue
            if (-not $props) { continue }
            foreach ($prop in @($props.PSObject.Properties)) {
                if ($prop.Name -in @('PSPath','PSParentPath','PSChildName','PSDrive','PSProvider')) { continue }
                $rows += [pscustomobject]@{
                    Source = $src.Source; Name = $prop.Name; Command = [string]$prop.Value
                    Enabled = $true; Kind = 'Registry'; Meta = $src.Path; NeedsAdmin = $src.Admin
                }
            }
        } catch { }
    }
    $folders = @(
        @{ Source='Startup folder (user)';   Path=[Environment]::GetFolderPath('Startup');       Admin=$false },
        @{ Source='Startup folder (common)'; Path=[Environment]::GetFolderPath('CommonStartup'); Admin=$true  }
    )
    foreach ($f in $folders) {
        if (-not ($f.Path -and (Test-Path -LiteralPath $f.Path))) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $f.Path -File -ErrorAction SilentlyContinue)) {
            if ($file.Name -ieq 'desktop.ini') { continue }
            $rows += [pscustomobject]@{
                Source = $f.Source; Name = $file.BaseName; Command = $file.FullName
                Enabled = $true; Kind = 'File'; Meta = $file.FullName; NeedsAdmin = $f.Admin
            }
        }
    }
    # entries we previously disabled (so they can be re-enabled)
    if (Test-Path $script:StartupBackupKey) {
        foreach ($k in @(Get-ChildItem $script:StartupBackupKey -ErrorAction SilentlyContinue)) {
            try {
                $p = Get-ItemProperty -Path $k.PSPath -ErrorAction SilentlyContinue
                $rows += [pscustomobject]@{
                    Source = [string]$p.Source; Name = [string]$p.EntryName; Command = [string]$p.Data
                    Enabled = $false; Kind = [string]$p.Kind; Meta = [string]$p.Meta; NeedsAdmin = ([string]$p.Source -like 'HKLM*' -or [string]$p.Source -like '*common*')
                }
            } catch { }
        }
    }
    return $rows
}

function Disable-StartupEntry {
    param($Row)
    try {
        if ($Row.NeedsAdmin -and -not $script:IsAdmin) { return @{ Success=$false; Message='Needs Administrator' } }
        if (-not (Test-Path $script:StartupBackupKey)) { $null = New-Item -Path $script:StartupBackupKey -Force }
        $id = Get-CleanName -Name ("{0}__{1}" -f $Row.Source, $Row.Name)
        $bk = Join-Path $script:StartupBackupKey $id
        if (-not (Test-Path $bk)) { $null = New-Item -Path $bk -Force }
        Set-ItemProperty -Path $bk -Name 'Source'    -Value $Row.Source
        Set-ItemProperty -Path $bk -Name 'EntryName' -Value $Row.Name
        Set-ItemProperty -Path $bk -Name 'Kind'      -Value $Row.Kind

        if ($Row.Kind -eq 'Registry') {
            Set-ItemProperty -Path $bk -Name 'Data' -Value $Row.Command
            Set-ItemProperty -Path $bk -Name 'Meta' -Value $Row.Meta
            Remove-ItemProperty -Path $Row.Meta -Name $Row.Name -Force -ErrorAction Stop
        } else {
            $target = Join-Path $script:DisabledStartupDir (Split-Path -Leaf $Row.Meta)
            Move-Item -LiteralPath $Row.Meta -Destination $target -Force -ErrorAction Stop
            Set-ItemProperty -Path $bk -Name 'Data' -Value $target
            Set-ItemProperty -Path $bk -Name 'Meta' -Value $Row.Meta
        }
        Write-Log ("Startup entry disabled: {0} ({1})" -f $Row.Name, $Row.Source)
        return @{ Success=$true; Message='Disabled (reversible)' }
    } catch { return @{ Success=$false; Message=$_.Exception.Message } }
}

function Enable-StartupEntry {
    param($Row)
    try {
        if ($Row.NeedsAdmin -and -not $script:IsAdmin) { return @{ Success=$false; Message='Needs Administrator' } }
        $id = Get-CleanName -Name ("{0}__{1}" -f $Row.Source, $Row.Name)
        $bk = Join-Path $script:StartupBackupKey $id
        if ($Row.Kind -eq 'Registry') {
            Set-ItemProperty -Path $Row.Meta -Name $Row.Name -Value $Row.Command -ErrorAction Stop
        } else {
            Move-Item -LiteralPath $Row.Command -Destination $Row.Meta -Force -ErrorAction Stop
        }
        if (Test-Path $bk) { Remove-Item -Path $bk -Recurse -Force -ErrorAction SilentlyContinue }
        Write-Log ("Startup entry re-enabled: {0} ({1})" -f $Row.Name, $Row.Source)
        return @{ Success=$true; Message='Re-enabled' }
    } catch { return @{ Success=$false; Message=$_.Exception.Message } }
}

# working-set trim P/Invoke
try {
    Add-Type -Namespace SPC -Name Native -MemberDefinition @'
[DllImport("psapi.dll", SetLastError=true)]
public static extern bool EmptyWorkingSet(System.IntPtr hProcess);
'@ -ErrorAction SilentlyContinue
} catch { }

$script:TrimDenylist = @(
    'Idle','System','Registry','Memory Compression','MemCompression','smss','csrss','wininit','winlogon','services',
    'lsass','svchost','dwm','fontdrvhost','audiodg','MsMpEng','SecurityHealthService','SecurityHealthSystray',
    'explorer','conhost','sihost','ctfmon','SearchHost','StartMenuExperienceHost','ShellExperienceHost','RuntimeBroker',
    # never trim active dev/agent tooling - the whole point is to leave agents alone
    'node','python','pythonw','Code','cursor','claude','ollama','pwsh','powershell','WindowsTerminal','ssh','git','dotnet'
)

function Invoke-MemoryTrim {
    $trimmed = 0
    foreach ($p in @(Get-Process -ErrorAction SilentlyContinue)) {
        if ($p.Id -eq $PID) { continue }
        if ($script:TrimDenylist -contains $p.ProcessName) { continue }
        try {
            $h = $p.Handle
            if ($h -and [SPC.Native]::EmptyWorkingSet($h)) { $trimmed++ }
        } catch { }
    }
    Write-Log ("Working-set trim done: {0} processes" -f $trimmed)
    return $trimmed
}

function Get-FreeMemoryMB {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        return [int]($os.FreePhysicalMemory / 1024)
    } catch { return 0 }
}

function Get-OrphanAgentProcesses {
    # Leftover helper processes from coding agents: windowless, running > 3h, parent gone.
    $names = @('node','python','pythonw','esbuild','MSBuild','java','deno','bun','tsserver','rust-analyzer')
    $rows = @()
    $cutoff = (Get-Date).AddHours(-3)
    $alive = @{}
    foreach ($p in @(Get-Process -ErrorAction SilentlyContinue)) { $alive[$p.Id] = $true }
    $cim = @{}
    try {
        foreach ($c in @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)) { $cim[[int]$c.ProcessId] = $c }
    } catch { }
    foreach ($p in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $names -contains $_.ProcessName })) {
        try {
            if ($p.MainWindowHandle -ne 0) { continue }
            if (-not $p.StartTime -or $p.StartTime -gt $cutoff) { continue }
            $parentId = $null
            if ($cim.ContainsKey($p.Id)) { $parentId = [int]$cim[$p.Id].ParentProcessId }
            $parentGone = $true
            if ($parentId -and $alive.ContainsKey($parentId)) { $parentGone = $false }
            if (-not $parentGone) { continue }
            $rows += [pscustomobject]@{
                Pid = $p.Id; Name = $p.ProcessName
                MemMB = [int]($p.WorkingSet64 / 1MB)
                Started = $p.StartTime
                Cmd = ''
            }
        } catch { }
    }
    return $rows
}

# ============================================================================
#  GUI construction
# ============================================================================
$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = "$($script:AppName) v$($script:AppVersion)" + $(if ($script:IsAdmin) { '  [Administrator]' } else { '  [Standard user - some features limited]' })
$script:Form.Size = New-Object System.Drawing.Size(1180, 760)
$script:Form.MinimumSize = New-Object System.Drawing.Size(980, 640)
$script:Form.StartPosition = 'CenterScreen'
$script:Form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$script:Status = New-Object System.Windows.Forms.StatusStrip
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Text = 'Ready. Everything is preview-first: nothing is changed until you confirm.'
$null = $script:Status.Items.Add($script:StatusLabel)
$script:Form.Controls.Add($script:Status)

$script:Tabs = New-Object System.Windows.Forms.TabControl
$script:Tabs.Dock = 'Fill'
$script:Form.Controls.Add($script:Tabs)
$script:Tabs.BringToFront()

function Set-Status {
    param([string]$Text)
    $script:StatusLabel.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function New-Btn {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 170, [int]$H = 30, [scriptblock]$OnClick)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    if ($OnClick) { $b.Add_Click($OnClick) }
    return $b
}

function New-Lbl {
    param([string]$Text, [int]$X, [int]$Y, [int]$W = 700, [int]$H = 18, [bool]$Bold = $false)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, $H)
    if ($Bold) { $l.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold) }
    return $l
}

function New-CheckedListView {
    param([string[]]$Columns, [int[]]$Widths, [bool]$CheckBoxes = $true)
    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.CheckBoxes = $CheckBoxes
    $lv.Dock = 'Fill'
    $lv.HideSelection = $false
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $null = $lv.Columns.Add($Columns[$i], $Widths[$i])
    }
    return $lv
}

# ----------------------------------------------------------------------------
#  TAB 1: Dashboard
# ----------------------------------------------------------------------------
$tabDash = New-Object System.Windows.Forms.TabPage
$tabDash.Text = ' Dashboard '
$null = $script:Tabs.TabPages.Add($tabDash)

$dashInfo = New-Object System.Windows.Forms.TextBox
$dashInfo.Multiline = $true
$dashInfo.ReadOnly = $true
$dashInfo.ScrollBars = 'Vertical'
$dashInfo.Dock = 'Fill'
$dashInfo.Font = New-Object System.Drawing.Font('Consolas', 10)

$dashTop = New-Object System.Windows.Forms.Panel
$dashTop.Dock = 'Top'
$dashTop.Height = 92

$dashTop.Controls.Add((New-Btn -Text 'Refresh system info' -X 10 -Y 10 -OnClick { Update-Dashboard }))
$dashTop.Controls.Add((New-Btn -Text 'Run Safe Quick Clean' -X 190 -Y 10 -W 190 -OnClick { Invoke-QuickClean }))
$dashTop.Controls.Add((New-Btn -Text 'Create Restore Point' -X 390 -Y 10 -OnClick { $null = New-SafetyRestorePoint }))
$dashTop.Controls.Add((New-Btn -Text 'Add desktop icon' -X 570 -Y 10 -OnClick {
    if (New-DesktopShortcut) {
        [System.Windows.Forms.MessageBox]::Show('Desktop icon created.', $script:AppName, 'OK', 'Information') | Out-Null
    }
}))
$dashTop.Controls.Add((New-Btn -Text 'Open log folder' -X 750 -Y 10 -W 150 -OnClick {
    Start-Process explorer.exe -ArgumentList $script:LogDir
}))
$dashTop.Controls.Add((New-Lbl -Text 'Quick Clean only touches always-safe items: aged user temp files, thumbnail/shader caches and error-report queues. Everything else lives in its own tab with previews.' -X 12 -Y 52 -W 1100 -H 34))

$tabDash.Controls.Add($dashInfo)
$tabDash.Controls.Add($dashTop)

function Update-Dashboard {
    Set-Status 'Reading system information...'
    $sb = New-Object System.Text.StringBuilder
    try {
        $os  = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $cs  = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        $null = $sb.AppendLine('SYSTEM')
        $null = $sb.AppendLine('------')
        if ($os) {
            $null = $sb.AppendLine(("  OS        : {0} (build {1})" -f $os.Caption, $os.BuildNumber))
            $up = (Get-Date) - $os.LastBootUpTime
            $null = $sb.AppendLine(("  Uptime    : {0}d {1}h {2}m" -f [int]$up.TotalDays, $up.Hours, $up.Minutes))
            $totalMB = [int]($os.TotalVisibleMemorySize / 1024)
            $freeMB  = [int]($os.FreePhysicalMemory / 1024)
            $null = $sb.AppendLine(("  Memory    : {0:N0} MB free of {1:N0} MB ({2:N0}% free)" -f $freeMB, $totalMB, (100.0 * $freeMB / [Math]::Max(1,$totalMB))))
        }
        if ($cpu) { $null = $sb.AppendLine(("  CPU       : {0}" -f $cpu.Name.Trim())) }
        if ($cs)  { $null = $sb.AppendLine(("  Machine   : {0}" -f $cs.Model)) }
        $null = $sb.AppendLine('')
        $null = $sb.AppendLine('DRIVES')
        $null = $sb.AppendLine('------')
        foreach ($d in [IO.DriveInfo]::GetDrives()) {
            if (-not $d.IsReady) { continue }
            if ($d.DriveType -notin @([IO.DriveType]::Fixed, [IO.DriveType]::Removable)) { continue }
            $pctFree = 0
            if ($d.TotalSize -gt 0) { $pctFree = [int](100.0 * $d.TotalFreeSpace / $d.TotalSize) }
            $barLen = 30
            $usedBars = [int]($barLen * (100 - $pctFree) / 100)
            $bar = ('#' * $usedBars).PadRight($barLen, '-')
            $warn = ''
            if ($pctFree -lt 10) { $warn = '   << LOW SPACE' }
            $null = $sb.AppendLine(("  {0}  [{1}] {2,3}% free   {3} free of {4}{5}" -f $d.Name, $bar, $pctFree,
                (Format-Bytes $d.TotalFreeSpace), (Format-Bytes $d.TotalSize), $warn))
        }
        $null = $sb.AppendLine('')
        $null = $sb.AppendLine('SAFETY')
        $null = $sb.AppendLine('------')
        $null = $sb.AppendLine("  Log file          : $($script:LogFile)")
        $null = $sb.AppendLine("  Registry backups  : $($script:BackupDir)")
        $null = $sb.AppendLine("  Move history      : $($script:MoveLogCsv)")
        $null = $sb.AppendLine('')
        $null = $sb.AppendLine('  Every destructive action in this app: previews first, backs up what it can, skips')
        $null = $sb.AppendLine('  files that are recent or in use, and refuses system-critical locations outright.')
    } catch {
        $null = $sb.AppendLine("Error reading system info: $($_.Exception.Message)")
    }
    $dashInfo.Text = $sb.ToString()
    Set-Status 'Ready.'
}

# ----------------------------------------------------------------------------
#  TAB 2: System Cleanup
# ----------------------------------------------------------------------------
$tabClean = New-Object System.Windows.Forms.TabPage
$tabClean.Text = ' System Cleanup '
$null = $script:Tabs.TabPages.Add($tabClean)

$script:CleanLV = New-CheckedListView -Columns @('Cleanup item', 'Reclaimable', 'Notes') -Widths @(260, 110, 700)
$cleanTop = New-Object System.Windows.Forms.Panel
$cleanTop.Dock = 'Top'
$cleanTop.Height = 78

$cleanTop.Controls.Add((New-Btn -Text '1. Scan (preview sizes)' -X 10 -Y 10 -W 180 -OnClick { Invoke-CleanScan }))
$cleanTop.Controls.Add((New-Btn -Text '2. Clean checked items' -X 200 -Y 10 -W 180 -OnClick { Invoke-CleanRun }))
$cleanTop.Controls.Add((New-Lbl -Text 'Scan is read-only. Cleaning skips files in use and files newer than each item''s safety age, so active apps and AI agents are never disturbed.' -X 12 -Y 48 -W 1100 -H 24))
$tabClean.Controls.Add($script:CleanLV)
$tabClean.Controls.Add($cleanTop)

$script:CleanLV.Add_ItemCheck({
    param($sender, $e)
    $item = $script:CleanLV.Items[$e.Index]
    if ($item.Tag -and $item.Tag.Admin -and -not $script:IsAdmin) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
        Set-Status ("'{0}' needs Administrator - restart the app and choose Yes at the prompt." -f $item.Text)
    }
})

function Invoke-CleanScan {
    $script:CleanLV.Items.Clear()
    $targets = Get-CleanupTargets
    $total = [long]0
    foreach ($t in $targets) {
        Set-Status ("Scanning: {0} ..." -f $t.Name)
        $size = [long]0
        if ($t.Special -eq 'RecycleBin') {
            $size = Get-RecycleBinSize
        } elseif ($t.Special -eq 'DO') {
            $size = -1   # size not cheaply known
        } else {
            foreach ($p in $t.Paths) {
                if (Test-Path -LiteralPath $p) { $size += Get-PathSize -Path $p }
            }
        }
        $sizeText = if ($size -lt 0) { 'n/a' } else { Format-Bytes $size }
        if ($size -gt 0) { $total += $size }
        $li = New-Object System.Windows.Forms.ListViewItem($t.Name)
        $null = $li.SubItems.Add($sizeText)
        $notes = $t.Desc
        if ($t.Admin -and -not $script:IsAdmin) { $notes = '[needs Admin] ' + $notes }
        $null = $li.SubItems.Add($notes)
        $li.Tag = $t
        if ($t.Admin -and -not $script:IsAdmin) { $li.ForeColor = [System.Drawing.Color]::Gray }
        $null = $script:CleanLV.Items.Add($li)
    }
    Set-Status ("Scan complete. Up to {0} reclaimable (actual amount depends on what is in use). Tick items, then Clean." -f (Format-Bytes $total))
}

function Invoke-CleanRun {
    $checked = @($script:CleanLV.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Run a scan and tick at least one item first.', $script:AppName, 'OK', 'Information') | Out-Null
        return
    }
    $names = ($checked | ForEach-Object { ' - ' + $_.Text }) -join "`n"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Clean the following?`n`n$names`n`nIn-use and recent files are skipped automatically.",
        $script:AppName, 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }

    $freed = [long]0
    foreach ($li in $checked) {
        $t = $li.Tag
        Set-Status ("Cleaning: {0} ..." -f $t.Name)
        try {
            if ($t.Special -eq 'RecycleBin') {
                Clear-RecycleBin -Force -ErrorAction SilentlyContinue
                $li.SubItems[1].Text = 'emptied'
            } elseif ($t.Special -eq 'DO') {
                try { Delete-DeliveryOptimizationCache -Force -ErrorAction Stop; $li.SubItems[1].Text = 'done' }
                catch { $li.SubItems[1].Text = 'n/a on this system' }
            } elseif ($t.Special -eq 'WU') {
                $r = Clear-WindowsUpdateCache
                $freed += $r.FreedBytes
                $li.SubItems[1].Text = ('freed ' + (Format-Bytes $r.FreedBytes))
            } else {
                $itemFreed = [long]0
                foreach ($p in $t.Paths) {
                    $r = Remove-ContentsSafe -Path $p -MinAgeHours $t.MinAgeHours -Include $t.Include
                    $itemFreed += $r.FreedBytes
                }
                $freed += $itemFreed
                $li.SubItems[1].Text = ('freed ' + (Format-Bytes $itemFreed))
            }
            $li.Checked = $false
        } catch {
            Write-Log ("Clean error on {0}: {1}" -f $t.Name, $_.Exception.Message) 'ERROR'
        }
    }
    Set-Status ("Cleanup finished. Freed {0} (plus Recycle Bin / Delivery Optimization where selected)." -f (Format-Bytes $freed))
    [System.Windows.Forms.MessageBox]::Show(("Cleanup finished.`nFreed: {0}" -f (Format-Bytes $freed)), $script:AppName, 'OK', 'Information') | Out-Null
}

function Invoke-QuickClean {
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("Safe Quick Clean will remove:`n" +
         " - user temp files older than 24h`n - thumbnail & shader caches`n - error-report queues`n - crash dumps (if Administrator)`n`n" +
         "Nothing in use, nothing recent, no registry, no user documents. Continue?"),
        $script:AppName, 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }
    $keys = @('UserTemp','Thumbs','DXCache','WER','CrashDumps')
    $freed = [long]0
    foreach ($t in (Get-CleanupTargets | Where-Object { $keys -contains $_.Key })) {
        if ($t.Admin -and -not $script:IsAdmin) { continue }
        Set-Status ("Quick clean: {0} ..." -f $t.Name)
        foreach ($p in $t.Paths) {
            $r = Remove-ContentsSafe -Path $p -MinAgeHours $t.MinAgeHours -Include $t.Include
            $freed += $r.FreedBytes
        }
    }
    Set-Status ("Quick Clean done - freed {0}." -f (Format-Bytes $freed))
    [System.Windows.Forms.MessageBox]::Show(("Quick Clean done.`nFreed: {0}" -f (Format-Bytes $freed)), $script:AppName, 'OK', 'Information') | Out-Null
}

# ----------------------------------------------------------------------------
#  TAB 3: AI Tools (Claude / ChatGPT / Codex / Cursor / VS Code)
# ----------------------------------------------------------------------------
$tabAI = New-Object System.Windows.Forms.TabPage
$tabAI.Text = ' AI Tools '
$null = $script:Tabs.TabPages.Add($tabAI)

$script:AILV = New-CheckedListView -Columns @('Tool', 'Item', 'Size', 'Verdict', 'Notes', 'Path') -Widths @(120, 200, 100, 90, 380, 260)
$aiTop = New-Object System.Windows.Forms.Panel
$aiTop.Dock = 'Top'
$aiTop.Height = 78

$aiTop.Controls.Add((New-Btn -Text '1. Scan AI tool data' -X 10 -Y 10 -W 170 -OnClick { Invoke-AIScan }))
$aiTop.Controls.Add((New-Btn -Text '2. Clean checked caches' -X 190 -Y 10 -W 180 -OnClick { Invoke-AIClean }))
$aiTop.Controls.Add((New-Btn -Text 'Send checked to Mover ->' -X 380 -Y 10 -W 190 -OnClick { Invoke-AISendToMover }))
$aiTop.Controls.Add((New-Lbl -Text 'Verdicts:  CLEAN = safe cache/log data (tickable).  MOVE = user data such as sessions & model weights - send it to the Mover tab instead of deleting.  KEEP = configs & credentials, never touch (cannot be ticked).' -X 12 -Y 48 -W 1140 -H 26))
$tabAI.Controls.Add($script:AILV)
$tabAI.Controls.Add($aiTop)

$script:AILV.Add_ItemCheck({
    param($sender, $e)
    $item = $script:AILV.Items[$e.Index]
    if ($item.Tag -and $item.Tag.Action -eq 'Keep') {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
        Set-Status 'That item holds configuration or credentials - the app will not let it be selected.'
    }
})

function Invoke-AIScan {
    $script:AILV.Items.Clear()
    Set-Status 'Scanning AI tool folders (Claude, ChatGPT, Codex, Cursor, VS Code, model caches)...'
    $rows = Get-AITargets
    $cleanTotal = [long]0
    $moveTotal = [long]0
    foreach ($r in $rows) {
        Set-Status ("Sizing: {0} - {1} ..." -f $r.Tool, $r.Item)
        $size = Get-PathSize -Path $r.Path
        $li = New-Object System.Windows.Forms.ListViewItem($r.Tool)
        $null = $li.SubItems.Add($r.Item)
        $null = $li.SubItems.Add((Format-Bytes $size))
        $verdict = switch ($r.Action) { 'Clean' { 'CLEAN' } 'Review' { 'MOVE' } default { 'KEEP' } }
        $null = $li.SubItems.Add($verdict)
        $null = $li.SubItems.Add($r.Note)
        $null = $li.SubItems.Add($r.Path)
        $li.Tag = $r
        switch ($r.Action) {
            'Clean'  { $li.ForeColor = [System.Drawing.Color]::FromArgb(20, 110, 20); $cleanTotal += $size }
            'Review' { $li.ForeColor = [System.Drawing.Color]::FromArgb(160, 90, 0); $moveTotal += $size }
            default  { $li.ForeColor = [System.Drawing.Color]::Gray }
        }
        $null = $script:AILV.Items.Add($li)
    }
    if ($rows.Count -eq 0) {
        Set-Status 'No AI tool data found on this machine.'
    } else {
        Set-Status ("AI scan done. Cleanable caches: ~{0}.  Movable user data (sessions/models): ~{1}." -f (Format-Bytes $cleanTotal), (Format-Bytes $moveTotal))
    }
}

function Invoke-AIClean {
    $checked = @($script:AILV.Items | Where-Object { $_.Checked -and $_.Tag.Action -eq 'Clean' })
    $reviewChecked = @($script:AILV.Items | Where-Object { $_.Checked -and $_.Tag.Action -eq 'Review' })
    if ($reviewChecked.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Some checked items are MOVE items (session history / model weights). Those are never deleted here - use "Send checked to Mover" for them. Only CLEAN items will be cleaned.',
            $script:AppName, 'OK', 'Information') | Out-Null
    }
    if ($checked.Count -eq 0) { Set-Status 'No CLEAN items ticked.'; return }
    $names = ($checked | ForEach-Object { ' - ' + $_.Tag.Tool + ': ' + $_.Tag.Item }) -join "`n"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Clean these AI caches?`n`n$names`n`nClose the matching apps first for the fullest clean. In-use files are skipped.",
        $script:AppName, 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }
    $freed = [long]0
    foreach ($li in $checked) {
        $r = $li.Tag
        Set-Status ("Cleaning {0}: {1} ..." -f $r.Tool, $r.Item)
        $res = Remove-ContentsSafe -Path $r.Path -MinAgeHours $r.MinAgeHours
        $freed += $res.FreedBytes
        $li.SubItems[2].Text = ('freed ' + (Format-Bytes $res.FreedBytes))
        $li.Checked = $false
    }
    Set-Status ("AI cache clean finished - freed {0}." -f (Format-Bytes $freed))
    [System.Windows.Forms.MessageBox]::Show(("AI cache clean finished.`nFreed: {0}" -f (Format-Bytes $freed)), $script:AppName, 'OK', 'Information') | Out-Null
}

function Invoke-AISendToMover {
    $checked = @($script:AILV.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 0) { Set-Status 'Tick items to send to the Mover first.'; return }
    $sent = 0
    foreach ($li in $checked) {
        if ($li.Tag.Action -eq 'Keep') { continue }
        Add-MoveCandidate -Path $li.Tag.Path -SuggestedCategory 'AI-Archives'
        $li.Checked = $false
        $sent++
    }
    $script:Tabs.SelectedTab = $script:TabMove
    Set-Status ("{0} item(s) added to the Mover list. Review, rename and choose a destination there." -f $sent)
}

# ----------------------------------------------------------------------------
#  TAB 4: Move to G:  (large-item mover/organizer)
# ----------------------------------------------------------------------------
$script:TabMove = New-Object System.Windows.Forms.TabPage
$script:TabMove.Text = ' Move to G: '
$null = $script:Tabs.TabPages.Add($script:TabMove)

$script:MoveLV = New-CheckedListView -Columns @('Name (double-click to rename)', 'Category', 'Size', 'Movable?', 'Why / where', 'Current location') -Widths @(240, 110, 100, 80, 300, 320)
$moveTop = New-Object System.Windows.Forms.Panel
$moveTop.Dock = 'Top'
$moveTop.Height = 122

$moveTop.Controls.Add((New-Btn -Text '1. Scan for large items' -X 10 -Y 10 -W 170 -OnClick { Invoke-MoveScan }))
$moveTop.Controls.Add((New-Lbl -Text 'Min size (MB):' -X 190 -Y 16 -W 85 -H 20))
$script:MoveMinSize = New-Object System.Windows.Forms.NumericUpDown
$script:MoveMinSize.Location = New-Object System.Drawing.Point(278, 13)
$script:MoveMinSize.Size = New-Object System.Drawing.Size(70, 24)
$script:MoveMinSize.Minimum = 10
$script:MoveMinSize.Maximum = 100000
$script:MoveMinSize.Value = 200
$moveTop.Controls.Add($script:MoveMinSize)

$moveTop.Controls.Add((New-Btn -Text 'Add another folder to scan...' -X 360 -Y 10 -W 200 -OnClick {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Pick an extra folder to scan for large items'
    if ($dlg.ShowDialog() -eq 'OK') {
        if (-not $script:ExtraScanRoots) { $script:ExtraScanRoots = @() }
        $script:ExtraScanRoots += $dlg.SelectedPath
        Set-Status ("Added scan root: {0}. Run the scan again." -f $dlg.SelectedPath)
    }
}))
$moveTop.Controls.Add((New-Btn -Text 'Open move history' -X 570 -Y 10 -W 150 -OnClick {
    if (Test-Path -LiteralPath $script:MoveLogCsv) { Start-Process notepad.exe -ArgumentList $script:MoveLogCsv }
    else { Set-Status 'No moves logged yet.' }
}))

$moveTop.Controls.Add((New-Lbl -Text 'Destination drive:' -X 12 -Y 52 -W 105 -H 20))
$script:MoveDrive = New-Object System.Windows.Forms.ComboBox
$script:MoveDrive.Location = New-Object System.Drawing.Point(120, 48)
$script:MoveDrive.Size = New-Object System.Drawing.Size(70, 24)
$script:MoveDrive.DropDownStyle = 'DropDownList'
foreach ($d in [IO.DriveInfo]::GetDrives()) {
    if ($d.IsReady -and $d.DriveType -in @([IO.DriveType]::Fixed, [IO.DriveType]::Removable)) {
        $null = $script:MoveDrive.Items.Add($d.Name)
    }
}
for ($i = 0; $i -lt $script:MoveDrive.Items.Count; $i++) {
    if ($script:MoveDrive.Items[$i] -like 'G:*') { $script:MoveDrive.SelectedIndex = $i }
}
if ($script:MoveDrive.SelectedIndex -lt 0 -and $script:MoveDrive.Items.Count -gt 0) { $script:MoveDrive.SelectedIndex = $script:MoveDrive.Items.Count - 1 }
$moveTop.Controls.Add($script:MoveDrive)

$moveTop.Controls.Add((New-Lbl -Text 'Base folder:' -X 200 -Y 52 -W 75 -H 20))
$script:MoveBase = New-Object System.Windows.Forms.TextBox
$script:MoveBase.Location = New-Object System.Drawing.Point(278, 48)
$script:MoveBase.Size = New-Object System.Drawing.Size(160, 24)
$script:MoveBase.Text = 'Organized'
$moveTop.Controls.Add($script:MoveBase)

$script:MoveAutoCat = New-Object System.Windows.Forms.CheckBox
$script:MoveAutoCat.Text = 'Auto-sort into category folders (Projects, Media, Documents, AI-Archives...)'
$script:MoveAutoCat.Location = New-Object System.Drawing.Point(450, 50)
$script:MoveAutoCat.Size = New-Object System.Drawing.Size(440, 22)
$script:MoveAutoCat.Checked = $true
$moveTop.Controls.Add($script:MoveAutoCat)

$moveTop.Controls.Add((New-Lbl -Text 'Or one project folder for all:' -X 12 -Y 84 -W 160 -H 20))
$script:MoveProject = New-Object System.Windows.Forms.TextBox
$script:MoveProject.Location = New-Object System.Drawing.Point(175, 80)
$script:MoveProject.Size = New-Object System.Drawing.Size(180, 24)
$moveTop.Controls.Add($script:MoveProject)

$script:MoveShortcut = New-Object System.Windows.Forms.CheckBox
$script:MoveShortcut.Text = 'Leave a shortcut at the old location'
$script:MoveShortcut.Location = New-Object System.Drawing.Point(370, 82)
$script:MoveShortcut.Size = New-Object System.Drawing.Size(240, 22)
$script:MoveShortcut.Checked = $true
$moveTop.Controls.Add($script:MoveShortcut)

$moveTop.Controls.Add((New-Btn -Text '2. Move checked items' -X 620 -Y 78 -W 180 -OnClick { Invoke-MoveRun }))
$script:TabMove.Controls.Add($script:MoveLV)
$script:TabMove.Controls.Add($moveTop)

$script:MoveLV.Add_ItemCheck({
    param($sender, $e)
    $item = $script:MoveLV.Items[$e.Index]
    if ($item.Tag -and -not $item.Tag.Movable) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
        Set-Status ("Cannot move: {0}" -f $item.Tag.Reason)
    }
})

$script:MoveLV.Add_MouseDoubleClick({
    if ($script:MoveLV.SelectedItems.Count -eq 0) { return }
    $li = $script:MoveLV.SelectedItems[0]
    if (-not $li.Tag.Movable) { return }
    $new = [Microsoft.VisualBasic.Interaction]::InputBox(
        'New name for this item at its destination:', 'Rename', $li.Tag.NewName)
    if ($new -and $new.Trim()) {
        $li.Tag.NewName = Get-CleanName -Name $new
        $li.Text = $li.Tag.NewName
        Set-Status ("Will be moved as: {0}" -f $li.Tag.NewName)
    }
})

function Add-MoveCandidate {
    param([string]$Path, [string]$SuggestedCategory = $null)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    foreach ($existing in @($script:MoveLV.Items)) {
        if ($existing.Tag.Path -ieq $Path) { return }   # no duplicates
    }
    Set-Status ("Inspecting: {0} ..." -f $Path)
    $size = Get-PathSize -Path $Path
    $safety = Test-MoveSafety -Path $Path
    $cat = $SuggestedCategory
    if (-not $cat) { $cat = Get-SuggestedCategory -Path $Path }
    $name = Split-Path -Leaf $Path
    $tag = [pscustomobject]@{
        Path = $Path; NewName = (Get-CleanName -Name $name); Category = $cat
        Size = $size; Movable = [bool]$safety.Movable; Reason = [string]$safety.Reason
    }
    $li = New-Object System.Windows.Forms.ListViewItem($tag.NewName)
    $null = $li.SubItems.Add($cat)
    $null = $li.SubItems.Add((Format-Bytes $size))
    $null = $li.SubItems.Add($(if ($tag.Movable) { 'Yes' } else { 'NO' }))
    $null = $li.SubItems.Add($tag.Reason)
    $null = $li.SubItems.Add($Path)
    $li.Tag = $tag
    if (-not $tag.Movable) { $li.ForeColor = [System.Drawing.Color]::Gray }
    $null = $script:MoveLV.Items.Add($li)
}

function Invoke-MoveScan {
    $script:MoveLV.Items.Clear()
    $minBytes = [long]$script:MoveMinSize.Value * 1MB
    $roots = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('MyDocuments'),
        (Join-Path $env:USERPROFILE 'Downloads'),
        [Environment]::GetFolderPath('MyVideos'),
        [Environment]::GetFolderPath('MyPictures'),
        [Environment]::GetFolderPath('MyMusic')
    )
    if ($script:ExtraScanRoots) { $roots += $script:ExtraScanRoots }
    $roots = @($roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)
    foreach ($root in $roots) {
        Set-Status ("Scanning {0} for items over {1} MB ..." -f $root, $script:MoveMinSize.Value)
        foreach ($entry in @(Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue)) {
            try {
                if ($entry.Name -like '*.lnk') { continue }
                $size = Get-PathSize -Path $entry.FullName
                if ($size -ge $minBytes) { Add-MoveCandidate -Path $entry.FullName }
            } catch { }
        }
    }
    $count = $script:MoveLV.Items.Count
    $movable = @($script:MoveLV.Items | Where-Object { $_.Tag.Movable }).Count
    Set-Status ("Found {0} large item(s), {1} movable. Tick items, adjust names/destination, then Move." -f $count, $movable)
}

function Invoke-MoveRun {
    $checked = @($script:MoveLV.Items | Where-Object { $_.Checked -and $_.Tag.Movable })
    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one movable item first.', $script:AppName, 'OK', 'Information') | Out-Null
        return
    }
    if ($script:MoveDrive.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show('Pick a destination drive.', $script:AppName, 'OK', 'Warning') | Out-Null
        return
    }
    $drive = [string]$script:MoveDrive.SelectedItem
    $base = Get-CleanName -Name $script:MoveBase.Text
    $project = $script:MoveProject.Text.Trim()
    $totalSize = [long]0
    foreach ($li in $checked) { $totalSize += $li.Tag.Size }

    $planLines = foreach ($li in $checked) {
        $cat = if ($project) { Get-CleanName -Name $project } elseif ($script:MoveAutoCat.Checked) { $li.Tag.Category } else { '' }
        $destFolder = Join-Path $drive $base
        if ($cat) { $destFolder = Join-Path $destFolder $cat }
        (' - {0}  ->  {1}\{2}' -f $li.Tag.Path, $destFolder.TrimEnd('\'), $li.Tag.NewName)
    }
    $confirmMsg = ("Move {0} item(s), {1} total?`n`n{2}`n`nEach item is COPIED first, verified (file count + bytes), and the original " +
                   "is deleted only after verification passes. A CSV log records every move.")
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ($confirmMsg -f $checked.Count, (Format-Bytes $totalSize), ($planLines -join "`n")),
        $script:AppName, 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }

    $moved = 0
    $failed = 0
    foreach ($li in $checked) {
        $cat = if ($project) { Get-CleanName -Name $project } elseif ($script:MoveAutoCat.Checked) { $li.Tag.Category } else { '' }
        $destFolder = Join-Path $drive $base
        if ($cat) { $destFolder = Join-Path $destFolder $cat }
        $destFolder = $destFolder.TrimEnd('\')
        Set-Status ("Moving {0} ({1}) ..." -f $li.Tag.NewName, (Format-Bytes $li.Tag.Size))
        $res = Invoke-SafeMove -Source $li.Tag.Path -DestFolder $destFolder -NewName $li.Tag.NewName -LeaveShortcut $script:MoveShortcut.Checked
        if ($res.Success) {
            $moved++
            $li.SubItems[3].Text = 'MOVED'
            $li.SubItems[4].Text = $res.Dest
            $li.ForeColor = [System.Drawing.Color]::FromArgb(20, 110, 20)
            $li.Checked = $false
            $li.Tag.Movable = $false
            $li.Tag.Reason = 'Already moved'
        } else {
            $failed++
            $li.SubItems[4].Text = ('FAILED: ' + $res.Message)
            $li.ForeColor = [System.Drawing.Color]::Firebrick
            Write-Log ("Move failed for {0}: {1}" -f $li.Tag.Path, $res.Message) 'ERROR'
        }
    }
    Set-Status ("Move complete: {0} moved, {1} failed. History: {2}" -f $moved, $failed, $script:MoveLogCsv)
    [System.Windows.Forms.MessageBox]::Show(
        ("Move complete.`nMoved: {0}`nFailed: {1}`n`nA log of every move is kept at:`n{2}" -f $moved, $failed, $script:MoveLogCsv),
        $script:AppName, 'OK', 'Information') | Out-Null
}

# ----------------------------------------------------------------------------
#  TAB 5: Registry Care
# ----------------------------------------------------------------------------
$tabReg = New-Object System.Windows.Forms.TabPage
$tabReg.Text = ' Registry Care '
$null = $script:Tabs.TabPages.Add($tabReg)

$script:RegLV = New-CheckedListView -Columns @('Issue type', 'Detail', 'Registry key') -Widths @(170, 420, 480)
$regTop = New-Object System.Windows.Forms.Panel
$regTop.Dock = 'Top'
$regTop.Height = 104

$regTop.Controls.Add((New-Btn -Text 'Create Restore Point' -X 10 -Y 10 -OnClick { $null = New-SafetyRestorePoint }))
$regTop.Controls.Add((New-Btn -Text 'Scan for safe issues' -X 190 -Y 10 -OnClick { Invoke-RegScan }))
$regTop.Controls.Add((New-Btn -Text 'Fix checked (auto-backup)' -X 370 -Y 10 -W 190 -OnClick { Invoke-RegFix }))
$regTop.Controls.Add((New-Btn -Text 'Clear recent-file lists' -X 570 -Y 10 -W 170 -OnClick {
    $answer = [System.Windows.Forms.MessageBox]::Show(
        'Clear Explorer recent-documents and Run-box history? This is a privacy cleanup, backed up first; Windows rebuilds the lists automatically.',
        $script:AppName, 'YesNo', 'Question')
    if ($answer -eq 'Yes') {
        $n = Clear-ExplorerMRU
        Set-Status ("Recent-file lists cleared ({0} key(s), backups saved)." -f $n)
    }
}))
$regTop.Controls.Add((New-Btn -Text 'Open backup folder' -X 750 -Y 10 -W 160 -OnClick {
    Start-Process explorer.exe -ArgumentList $script:BackupDir
}))
$regTop.Controls.Add((New-Lbl -Text 'Deliberately conservative: this scan only finds entries pointing at programs that no longer exist (orphaned App Paths and non-MSI uninstall leftovers). It never touches drivers, services, file associations or MSI-installed apps - the categories where "registry cleaners" cause damage.' -X 12 -Y 46 -W 1140 -H 30))
$regTop.Controls.Add((New-Lbl -Text 'Every key is exported to a .reg backup BEFORE deletion - double-click any backup file to restore it. Create a Restore Point first for belt-and-braces safety.' -X 12 -Y 78 -W 1140 -H 20 -Bold $true))
$tabReg.Controls.Add($script:RegLV)
$tabReg.Controls.Add($regTop)

function Invoke-RegScan {
    $script:RegLV.Items.Clear()
    Set-Status 'Scanning registry for orphaned entries (read-only)...'
    $issues = @()
    $issues += Get-OrphanedAppPaths
    $issues += Get-OrphanedUninstallEntries
    foreach ($i in $issues) {
        $li = New-Object System.Windows.Forms.ListViewItem($i.Type)
        $null = $li.SubItems.Add($i.Detail)
        $null = $li.SubItems.Add($i.KeyName)
        $li.Tag = $i
        $null = $script:RegLV.Items.Add($li)
    }
    if ($issues.Count -eq 0) {
        Set-Status 'Registry scan complete - no orphaned entries found. That is a good thing; nothing needs fixing.'
    } else {
        Set-Status ("Registry scan complete - {0} orphaned entr{1} found. Review, tick, then Fix (each is backed up first)." -f $issues.Count, $(if ($issues.Count -eq 1) { 'y' } else { 'ies' }))
    }
}

function Invoke-RegFix {
    $checked = @($script:RegLV.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 0) { Set-Status 'Scan and tick at least one issue first.'; return }
    $hklm = @($checked | Where-Object { $_.Tag.KeyName -like 'HKEY_LOCAL_MACHINE*' })
    if ($hklm.Count -gt 0 -and -not $script:IsAdmin) {
        [System.Windows.Forms.MessageBox]::Show('Some checked keys are machine-wide (HKLM) and need Administrator. Restart the app as Administrator, or untick those.', $script:AppName, 'OK', 'Warning') | Out-Null
        return
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("Remove {0} orphaned registry entr{1}?`n`nEach key is exported to a .reg backup first. If a backup export fails, that key is NOT deleted.`n`nTip: create a Restore Point first if you have not today." -f $checked.Count, $(if ($checked.Count -eq 1) { 'y' } else { 'ies' })),
        $script:AppName, 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }
    $fixed = 0
    $failedCount = 0
    foreach ($li in $checked) {
        Set-Status ("Fixing: {0} ..." -f $li.Tag.KeyName)
        $res = Remove-RegistryKeySafe -KeyName $li.Tag.KeyName
        if ($res.Success) {
            $fixed++
            $li.SubItems[1].Text = $res.Message
            $li.ForeColor = [System.Drawing.Color]::FromArgb(20, 110, 20)
            $li.Checked = $false
        } else {
            $failedCount++
            $li.SubItems[1].Text = ('NOT removed: ' + $res.Message)
            $li.ForeColor = [System.Drawing.Color]::Firebrick
        }
    }
    Set-Status ("Registry fix done: {0} removed (with backups), {1} skipped. Backups: {2}" -f $fixed, $failedCount, $script:BackupDir)
}

# ----------------------------------------------------------------------------
#  TAB 6: Performance
# ----------------------------------------------------------------------------
$tabPerf = New-Object System.Windows.Forms.TabPage
$tabPerf.Text = ' Performance '
$null = $script:Tabs.TabPages.Add($tabPerf)

$perfSplit = New-Object System.Windows.Forms.SplitContainer
$perfSplit.Dock = 'Fill'
$perfSplit.Orientation = 'Horizontal'

$script:StartupLV = New-CheckedListView -Columns @('Startup entry', 'State', 'Source', 'Command') -Widths @(230, 80, 170, 600) -CheckBoxes $false
$startupPanel = New-Object System.Windows.Forms.Panel
$startupPanel.Dock = 'Top'
$startupPanel.Height = 44
$startupPanel.Controls.Add((New-Btn -Text 'Refresh startup list' -X 10 -Y 7 -OnClick { Update-StartupList }))
$startupPanel.Controls.Add((New-Btn -Text 'Disable selected' -X 190 -Y 7 -W 140 -OnClick { Invoke-StartupToggle -Disable $true }))
$startupPanel.Controls.Add((New-Btn -Text 'Enable selected' -X 340 -Y 7 -W 140 -OnClick { Invoke-StartupToggle -Disable $false }))
$startupPanel.Controls.Add((New-Lbl -Text 'Disabling is fully reversible - entries are stored, not deleted. Leave security software (e.g. SecurityHealth) enabled.' -X 495 -Y 12 -W 640 -H 22))
$perfSplit.Panel1.Controls.Add($script:StartupLV)
$perfSplit.Panel1.Controls.Add($startupPanel)

$perfBottom = New-Object System.Windows.Forms.Panel
$perfBottom.Dock = 'Fill'
$perfBottom.Controls.Add((New-Lbl -Text 'Memory & agent hygiene' -X 10 -Y 8 -W 300 -H 20 -Bold $true))
$script:MemLabel = New-Lbl -Text '' -X 320 -Y 10 -W 600 -H 20
$perfBottom.Controls.Add($script:MemLabel)

$perfBottom.Controls.Add((New-Btn -Text 'Trim background memory' -X 10 -Y 34 -W 185 -OnClick {
    Set-Status 'Trimming working sets of idle background processes (dev tools & agents are excluded)...'
    $before = Get-FreeMemoryMB
    $n = Invoke-MemoryTrim
    Start-Sleep -Milliseconds 700
    $after = Get-FreeMemoryMB
    $gain = [Math]::Max(0, $after - $before)
    Update-MemLabel
    Set-Status ("Trimmed {0} background processes; ~{1:N0} MB returned to the free pool. Apps reload pages on demand - no harm done." -f $n, $gain)
}))
$perfBottom.Controls.Add((New-Btn -Text 'Flush DNS cache' -X 205 -Y 34 -W 140 -OnClick {
    try { Clear-DnsClientCache -ErrorAction Stop; Set-Status 'DNS cache flushed - fixes stale lookups that can stall agents and browsers.' }
    catch { & "$env:SystemRoot\System32\ipconfig.exe" /flushdns | Out-Null; Set-Status 'DNS cache flushed.' }
}))
$perfBottom.Controls.Add((New-Btn -Text 'Restart Explorer' -X 355 -Y 34 -W 140 -OnClick {
    $answer = [System.Windows.Forms.MessageBox]::Show('Restart Windows Explorer? Your taskbar/desktop reload; open apps are unaffected.', $script:AppName, 'YesNo', 'Question')
    if ($answer -eq 'Yes') {
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }
        Set-Status 'Explorer restarted.'
    }
}))
$perfBottom.Controls.Add((New-Btn -Text 'High Performance plan' -X 505 -Y 34 -W 165 -OnClick {
    & "$env:SystemRoot\System32\powercfg.exe" /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 2>&1 | Out-Null
    Set-Status 'High Performance power plan requested (on laptops this uses more battery).'
}))
$perfBottom.Controls.Add((New-Btn -Text 'Balanced plan' -X 680 -Y 34 -W 120 -OnClick {
    & "$env:SystemRoot\System32\powercfg.exe" /setactive 381b4222-f694-41f0-9685-ff5bb260df2e 2>&1 | Out-Null
    Set-Status 'Balanced power plan restored.'
}))

$perfBottom.Controls.Add((New-Lbl -Text 'Leftover agent processes (windowless node/python/build helpers whose parent app has exited - these hold memory and file locks that can trip up new agent runs):' -X 10 -Y 76 -W 1120 -H 20))
$script:OrphanLV = New-CheckedListView -Columns @('PID', 'Process', 'Memory', 'Running since') -Widths @(80, 160, 100, 220)
$script:OrphanLV.Dock = 'None'
$script:OrphanLV.Location = New-Object System.Drawing.Point(10, 100)
$script:OrphanLV.Size = New-Object System.Drawing.Size(700, 160)
$script:OrphanLV.Anchor = 'Top,Left,Right,Bottom'
$perfBottom.Controls.Add($script:OrphanLV)
$perfBottom.Controls.Add((New-Btn -Text 'Scan for leftovers' -X 730 -Y 100 -W 160 -OnClick { Update-OrphanList }))
$btnKillOrphans = New-Btn -Text 'End checked leftovers' -X 730 -Y 140 -W 160 -OnClick { Invoke-KillOrphans }
$btnKillOrphans.Anchor = 'Top,Right'
$perfBottom.Controls.Add($btnKillOrphans)

$perfSplit.Panel2.Controls.Add($perfBottom)
$tabPerf.Controls.Add($perfSplit)
try { $perfSplit.SplitterDistance = 300 } catch { }

function Update-MemLabel {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $totalMB = [int]($os.TotalVisibleMemorySize / 1024)
        $freeMB  = [int]($os.FreePhysicalMemory / 1024)
        $script:MemLabel.Text = ("RAM: {0:N0} MB free of {1:N0} MB" -f $freeMB, $totalMB)
    } catch { $script:MemLabel.Text = '' }
}

function Update-StartupList {
    $script:StartupLV.Items.Clear()
    Set-Status 'Reading startup entries...'
    foreach ($row in (Get-StartupEntries | Sort-Object Enabled -Descending)) {
        $li = New-Object System.Windows.Forms.ListViewItem($row.Name)
        $null = $li.SubItems.Add($(if ($row.Enabled) { 'Enabled' } else { 'Disabled' }))
        $null = $li.SubItems.Add($row.Source)
        $null = $li.SubItems.Add($row.Command)
        $li.Tag = $row
        if (-not $row.Enabled) { $li.ForeColor = [System.Drawing.Color]::Gray }
        if ($row.Name -match '(?i)security|defender|MsMpEng') { $li.ForeColor = [System.Drawing.Color]::Firebrick }
        $null = $script:StartupLV.Items.Add($li)
    }
    Set-Status ("{0} startup entries. Red = security software, leave enabled. Select one, then Disable/Enable." -f $script:StartupLV.Items.Count)
}

function Invoke-StartupToggle {
    param([bool]$Disable)
    if ($script:StartupLV.SelectedItems.Count -eq 0) { Set-Status 'Select a startup entry first.'; return }
    foreach ($li in @($script:StartupLV.SelectedItems)) {
        $row = $li.Tag
        if ($Disable -and -not $row.Enabled) { continue }
        if (-not $Disable -and $row.Enabled) { continue }
        if ($Disable -and $row.Name -match '(?i)security|defender|MsMpEng') {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                ("'{0}' looks like security software. Disabling it is NOT recommended. Disable anyway?" -f $row.Name),
                $script:AppName, 'YesNo', 'Warning')
            if ($answer -ne 'Yes') { continue }
        }
        $res = if ($Disable) { Disable-StartupEntry -Row $row } else { Enable-StartupEntry -Row $row }
        if ($res.Success) { Set-Status ("{0}: {1}" -f $row.Name, $res.Message) }
        else {
            Set-Status ("{0}: FAILED - {1}" -f $row.Name, $res.Message)
            [System.Windows.Forms.MessageBox]::Show(("Could not change '{0}': {1}" -f $row.Name, $res.Message), $script:AppName, 'OK', 'Warning') | Out-Null
        }
    }
    Update-StartupList
}

function Update-OrphanList {
    $script:OrphanLV.Items.Clear()
    Set-Status 'Looking for leftover windowless agent helper processes (parent exited, running 3h+)...'
    $rows = Get-OrphanAgentProcesses
    foreach ($r in $rows) {
        $li = New-Object System.Windows.Forms.ListViewItem([string]$r.Pid)
        $null = $li.SubItems.Add($r.Name)
        $null = $li.SubItems.Add(("{0:N0} MB" -f $r.MemMB))
        $null = $li.SubItems.Add($r.Started.ToString('yyyy-MM-dd HH:mm'))
        $li.Tag = $r
        $null = $script:OrphanLV.Items.Add($li)
    }
    if ($rows.Count -eq 0) { Set-Status 'No leftover agent processes found - clean slate for your agents.' }
    else { Set-Status ("{0} leftover process(es) found. These are candidates only - tick and end them if you recognize them as stale." -f $rows.Count) }
}

function Invoke-KillOrphans {
    $checked = @($script:OrphanLV.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 0) { Set-Status 'Tick leftover processes to end first.'; return }
    $names = ($checked | ForEach-Object { ' - PID {0}: {1} ({2:N0} MB)' -f $_.Tag.Pid, $_.Tag.Name, $_.Tag.MemMB }) -join "`n"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "End these processes?`n`n$names`n`nOnly do this if nothing important is mid-task in a terminal somewhere.",
        $script:AppName, 'YesNo', 'Warning')
    if ($answer -ne 'Yes') { return }
    $ended = 0
    foreach ($li in $checked) {
        try {
            Stop-Process -Id $li.Tag.Pid -Force -ErrorAction Stop
            $ended++
            Write-Log ("Ended leftover process PID {0} ({1})" -f $li.Tag.Pid, $li.Tag.Name)
        } catch { }
    }
    Update-OrphanList
    Set-Status ("Ended {0} leftover process(es)." -f $ended)
}

# ============================================================================
#  Launch
# ============================================================================
$script:Form.Add_Shown({
    Update-Dashboard
    Update-MemLabel
})

try {
    [System.Windows.Forms.Application]::Run($script:Form)
} catch {
    Write-Log ("Fatal UI error: {0}" -f $_.Exception.Message) 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(("Smart PC Cleaner hit an unexpected error and will close:`n{0}`n`nDetails are in:`n{1}" -f $_.Exception.Message, $script:LogFile), $script:AppName, 'OK', 'Error') | Out-Null
}
Write-Log 'Session ended.'
