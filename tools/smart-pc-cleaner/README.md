# Smart PC Cleaner

A safety-first Windows tune-up and organizer app in a single PowerShell file — no installation,
no dependencies, nothing phoned home. Built for machines that run AI coding agents
(Claude Code, Codex, Cursor, Copilot) and have accumulated caches, session logs and large
project folders that need cleaning up or moving to another drive.

## Quick start

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
| **Registry Care** | Removes only clearly-orphaned entries: App Paths pointing at missing programs, and non-MSI uninstall leftovers whose uninstaller *and* install folder are both gone. Also a privacy clear of recent-file lists | Every key exported to a `.reg` backup **before** deletion (restore by double-clicking the backup). One-click System Restore Point. MSI apps, drivers, services, file associations are never scanned or touched |
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

- **Registry:** double-click the `.reg` file in the backup folder (button on the Registry tab),
  or use the System Restore Point.
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
├── SmartPCCleaner.ps1          # the whole app (GUI, ~6 tabs)
├── Install-SmartPCCleaner.bat  # one-time: desktop icon + first launch
└── README.md
```
