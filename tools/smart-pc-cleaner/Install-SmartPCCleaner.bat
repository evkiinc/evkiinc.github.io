@echo off
rem ---------------------------------------------------------------
rem  Smart PC Cleaner - one-time setup
rem  Creates a desktop icon and opens the app. Safe to re-run.
rem ---------------------------------------------------------------
title Smart PC Cleaner Setup
echo.
echo  Installing Smart PC Cleaner (creating desktop icon)...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%~dp0SmartPCCleaner.ps1" -InstallShortcut
if errorlevel 1 (
    echo.
    echo  Something went wrong. Try right-clicking SmartPCCleaner.ps1
    echo  and choosing "Run with PowerShell" instead.
    pause
)
