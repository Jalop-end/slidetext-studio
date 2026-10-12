@echo off
chcp 65001 >nul
title SlideText Studio Launcher
setlocal EnableDelayedExpansion

cd /d "%~dp0"

where powershell.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [CHYBA] PowerShell nebyl nalezen v systemu Windows.
    echo Pro spousteni aplikace je vyzadovan PowerShell.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\launcher.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

if %EXIT_CODE% NEQ 0 (
    echo.
    echo [CHYBA] Aplikace se ukoncila s chybou [kod: %EXIT_CODE%].
    pause
)
