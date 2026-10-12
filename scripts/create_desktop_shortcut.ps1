<#
.SYNOPSIS
    Vytvori zastupce aplikace SlideText Studio na plose uzivatele (Desktop).
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$RootPath = (Resolve-Path "$PSScriptRoot\..").Path
$BatPath = Join-Path $RootPath "start.bat"
$IcoPath = Join-Path $RootPath "app.ico"

# Ujistit se, ze existuje app.ico
if (-not (Test-Path $IcoPath)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\make_ico.ps1"
}

$DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$ShortcutPath = Join-Path $DesktopPath "SlideText Studio.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $BatPath
$Shortcut.WorkingDirectory = $RootPath
$Shortcut.Description = "SlideText Studio - PPTX Editor & AI Translator"

if (Test-Path $IcoPath) {
    $Shortcut.IconLocation = "$IcoPath,0"
}

$Shortcut.Save()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null

Write-Host "Zastupce uspesne vytvoren na plose:" -ForegroundColor Green
Write-Host "  -> $ShortcutPath" -ForegroundColor Cyan
Write-Host "Cil: $BatPath" -ForegroundColor Gray
Write-Host "Ikona: $IcoPath" -ForegroundColor Gray
