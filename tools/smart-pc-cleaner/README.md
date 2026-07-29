# Smart PC Cleaner

A safety-first Windows tune-up and organizer app in a single PowerShell file — no installation,
no dependencies, nothing phoned home. Built for machines that run AI coding agents
(Claude Code, Codex, Cursor, Copilot) and have accumulated caches, session logs and large
project folders that need cleaning up or moving to another drive.

**v2.0** — crisp Google Material-style interface: left navigation sidebar, the four-color
accent, flat blue/white buttons with hover states, hairline-bordered white cards, stat tiles
on the dashboard, and a status bar narrating every step. Plus a full **Registry — Repair &
Optimization** page (see below).

## Quick start

**Fastest — one line, no zip needed.** Open PowerShell (Start menu → type `powershell` → Enter) and paste:

```powershell
irm https://raw.githubusercontent.com/evkiinc/evkiinc.github.io/claude/pc-performance-cleaner-lsnfs6/tools/smart-pc-cleaner/Get-SmartPCCleaner.ps1 | iex
```

It downloads the app to `%LOCALAPPDATA%\Programs\SmartPCCleaner`, puts the logo icon on the
desktop and opens the app. No admin rights needed to install.

**Or manually:**

1. Copy the `smart-pc-cleaner` folder anywhere on the PC (e.g. `C:\Tools\smart-pc-cleaner`).
2. Double-click **`Install-SmartPCCleaner.bat`** once.
   It creates a **desktop icon** and opens the app.
3. From then on, launch it from the desktop icon.
   At startup it offers to run as Administrator — say **Yes** for the full feature set
   (Windows temp, update cache, restore points, machine-wide startup entries), or **No**
   to run with user-level features only.

Requires Windows 10 or 11 with the built-in PowerShell 5.1 (present on every machine).

## What each tab does

| Tab | What it does | Safety mechanism |
|---|---|---|
| **Dashboard** | System/drive overview, Safe Quick Clean, restore point, desktop icon, logs | Quick Clean touches only always-safe items |
| **System Cleanup** | Temp files, thumbnail/shader caches, crash dumps, error reports, browser caches, Windows Update leftovers, Recycle Bin | Scan first (read-only preview with sizes) → tick → confirm. Files in use or newer than each item's safety age are skipped |
| **AI Tools** | Finds Claude Code, Claude Desktop, ChatGPT, Codex CLI, Cursor, VS Code, Copilot, npm/pip caches, Ollama/HuggingFace model stores | Three verdicts: **CLEAN** (safe caches, tickable) / **MOVE** (sessions & models — sent to the Mover, never deleted) / **KEEP** (configs & credentials — the UI refuses to select them) |
| **Move to G:** | Finds large folders/files in Desktop, Documents, Downloads, Media folders (plus any folder you add), checks movability, lets you rename each item and organize into category or project folders on the target drive | Copy → verify (file count + bytes) → only then delete the original. Optional shortcut left behind. Every move logged to CSV. Windows/Program Files/AppData/installed apps/OneDrive placeholders/junctions are refused outright |
| **Registry — Repair & Optimization** | **Repair:** five scanners for verifiably-dead entries — orphaned App Paths, non-MSI uninstall leftovers, broken startup (Run) entries, stale Explorer display-name cache (MuiCache), broken shared-DLL reference counts. **Optimization:** six documented, user-level responsiveness tweaks (snappier menus, no startup-app delay, no window/taskbar animations, Game DVR off, faster sign-out timeouts, local-only Start search) — each applied with its original values backed up and revertible in one click. Plus recent-file-list privacy clear, full HKCU\Software backup, and a hive-size report | Every key exported to a `.reg` backup **before** any change (restore by double-clicking the backup). One-click System Restore Point. MSI apps, drivers, services, file associations are never scanned or touched. No "registry defrag" — Windows compacts hives itself at boot, and the report says so |
| **Performance** | Reversible startup manager, background memory trim, DNS flush, Explorer restart, power plans, and cleanup of leftover windowless agent helper processes (node/python/build servers whose parent exited) | Startup disables are stored, not deleted (one click to re-enable). Memory trim excludes system processes *and* all dev/agent tooling. Leftover processes are listed for review, never auto-killed |

## Design principles (why this won't damage Windows)

1. **Preview-first everywhere.** Scans are read-only; nothing changes without a checkbox and a
   confirmation dialog.
2. **Agent-friendly by default.** Files newer than each item's safety age (24 h for temp files)
   are never deleted, and locked/in-use files are skipped silently — a running Claude/Codex
   session, build, or download is never tripped up. Memory trim and process cleanup explicitly
   exclude active dev tooling.
3. **Registry cleaning is deliberately minimal.** Broad registry cleaners are the classic cause
   of broken Windows installs, and the measurable speed benefit is near zero — so this app only
   removes entries that point at programs that verifiably no longer exist, backs up every key
   first, and can create a System Restore Point in one click.
4. **Moves can't lose data.** Copy → verify → delete, with a CSV history and optional shortcuts
   at the old location. If verification fails, both copies are kept and you're told.
5. **Hard denylist.** `C:\Windows`, `Program Files`, `ProgramData`, `AppData`, registered
   application folders, junctions and OneDrive online-only placeholders can never be moved, and
   root/short paths can never be cleaned — enforced in code, not just by convention.
6. **Everything is logged** to `%LOCALAPPDATA%\SmartPCCleaner\logs`, registry backups to
   `%LOCALAPPDATA%\SmartPCCleaner\registry-backups`, moves to
   `%LOCALAPPDATA%\SmartPCCleaner\move-history.csv`.

## Undo / recovery

- **Registry repairs:** double-click the `.reg` file in the backup folder (button on the
  Registry page), or use the System Restore Point.
- **Registry optimizations:** tick the tweak and click *Revert checked* — original values
  are restored exactly from the stored backup.
- **Startup entries:** select the disabled entry and click *Enable selected*.
- **Moves:** the move history CSV records source → destination for every item; move it back with
  Explorer if needed (a shortcut at the old location points to the new one).
- **Cleaned files:** temp/cache data is rebuilt automatically by Windows and the apps.

## Prior art this design draws on

- **BleachBit** (open source) — the whitelist-of-known-locations approach to cleaning.
- **Microsoft PC Manager / Storage Sense / cleanmgr** — which locations Microsoft itself
  considers safe to purge (temp, WER, thumbnails, update cache, Delivery Optimization).
- **Sysinternals Autoruns** — startup management done reversibly.
- **WizTree / WinDirStat** — surfacing the biggest folders first when freeing disk space.
- **Robocopy** — the battle-tested engine used for verified folder moves.
- The long history of *aggressive* registry cleaners breaking Windows is exactly why the
  Registry Care tab is scoped to orphaned-reference removal with mandatory backups.

## Files

```
smart-pc-cleaner/
├── SmartPCCleaner.ps1          # the whole app (Material-style GUI, 6 pages)
├── SmartPCCleaner.ico          # designed app logo (desktop icon, window & taskbar)
├── Get-SmartPCCleaner.ps1      # one-line web installer (irm ... | iex)
├── Install-SmartPCCleaner.bat  # one-time: desktop icon + first launch
└── README.md
```
