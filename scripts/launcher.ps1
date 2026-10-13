<#
.SYNOPSIS
    SlideText Studio - Inteligentni Launcher a Inicializator
.DESCRIPTION
    Pri prvnim spusteni automaticky detekuje a nainstaluje chybejici behove prostredi
    (Node.js / Bun), zavislosti npm a inicializuje konfiguraci.
    Pri dalsim spusteni okamzite nastartuje server a otevre aplikaci v prohlizeci.
#>

param(
    [switch]$Reinstall,
    [switch]$NoBrowser,
    [int]$Port = 3000
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Continue"
$RootPath = (Resolve-Path "$PSScriptRoot\..").Path
Set-Location $RootPath

function Write-Color([string]$text, [ConsoleColor]$color = [ConsoleColor]::White) {
    Write-Host $text -ForegroundColor $color
}

function Write-Banner {
    Clear-Host
    Write-Color "=================================================================" ([ConsoleColor]::Cyan)
    Write-Color "               🚀 SlideText Studio Launcher                     " ([ConsoleColor]::Yellow)
    Write-Color "  PowerPoint & Word Editor & AI Translator (100% Zachovani stylu)" ([ConsoleColor]::White)
    Write-Color "=================================================================" ([ConsoleColor]::Cyan)
    Write-Host ""
}

function Refresh-EnvPath {
    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $extraPaths = @(
        "C:\Program Files\nodejs",
        "C:\Program Files (x86)\nodejs",
        "C:\Program Files\Git\cmd",
        "C:\Program Files\Git\bin",
        "$env:USERPROFILE\.bun\bin",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links",
        "$env:APPDATA\npm"
    )
    $combined = "$userPath;$machinePath;" + ($extraPaths -join ";")
    $env:PATH = $combined
}

function Find-Command([string]$cmdName) {
    Refresh-EnvPath
    $cmd = Get-Command $cmdName -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Test-PortOpen([string]$hostAddress, [int]$portNumber) {
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $client.BeginConnect($hostAddress, $portNumber, $null, $null)
        $success = $asyncResult.AsyncWaitHandle.WaitOne(500, $false)
        if ($success) {
            $client.EndConnect($asyncResult)
            $client.Close()
            return $true
        }
        $client.Close()
        return $false
    } catch {
        return $false
    }
}

Write-Banner

# ---------------------------------------------------------------------------
# 1. KONTROLA PRVNIHO SPUSTENI
# ---------------------------------------------------------------------------
$modulesExist = Test-Path "$RootPath\node_modules"
$readyMarker = Test-Path "$RootPath\.launcher-ready"
$isFirstRun = (-not $modulesExist) -or (-not $readyMarker) -or $Reinstall

if ($isFirstRun) {
    Write-Color "🛠️  PRVNI SPUSTENI: Inicializace a kontrola potrebnych nastroju..." ([ConsoleColor]::Magenta)
    Write-Color "-----------------------------------------------------------------" ([ConsoleColor]::DarkGray)
} else {
    Write-Color "⚡ RYCHLY START: Vsechny nastroje a zavislosti jsou pripraveny." ([ConsoleColor]::Green)
    Write-Color "-----------------------------------------------------------------" ([ConsoleColor]::DarkGray)
}

# ---------------------------------------------------------------------------
# 2. DETEKCE A INSTALACE RUNTIME (BUN / NODE.JS)
# ---------------------------------------------------------------------------
Refresh-EnvPath

$bunCmd = Find-Command "bun"
$nodeCmd = Find-Command "node"
$npmCmd = Find-Command "npm"

if (-not $bunCmd -and -not $nodeCmd) {
    Write-Color "[!] Behove prostredi (Node.js nebo Bun) nebylo nalezeno v systemu." ([ConsoleColor]::Yellow)
    Write-Color "--> Zahajuji automatickou instalaci..." ([ConsoleColor]::Cyan)

    $wingetCmd = Find-Command "winget"
    $installed = $false

    if ($wingetCmd) {
        Write-Color "   [1/2] Instaluji Node.js LTS pres Windows Package Manager (winget)..." ([ConsoleColor]::Cyan)
        try {
            $proc = Start-Process -FilePath "winget" -ArgumentList "install --id OpenJS.NodeJS.LTS -e --silent --accept-package-agreements --accept-source-agreements" -NoNewWindow -PassThru -Wait
            if ($proc.ExitCode -eq 0) {
                $installed = $true
                Write-Color "   [OK] Node.js LTS byl uspesne nainstalovan!" ([ConsoleColor]::Green)
            }
        } catch {
            Write-Color "   WinGet instalace selhala." ([ConsoleColor]::DarkYellow)
        }
    }

    if (-not $installed) {
        Write-Color "   [2/2] Instaluji Bun runtime..." ([ConsoleColor]::Cyan)
        try {
            Invoke-RestMethod -Uri "https://bun.sh/install.ps1" | Invoke-Expression
            $installed = $true
            Write-Color "   [OK] Bun byl uspesne nainstalovan!" ([ConsoleColor]::Green)
        } catch {
            Write-Color "   Instalace Bun selhala." ([ConsoleColor]::DarkYellow)
        }
    }

    Refresh-EnvPath
    $bunCmd = Find-Command "bun"
    $nodeCmd = Find-Command "node"
    $npmCmd = Find-Command "npm"

    if (-not $bunCmd -and -not $nodeCmd) {
        Write-Color "[X] Nepodarilo se automaticky nainstalovat Node.js ani Bun." ([ConsoleColor]::Red)
        Write-Color "--> Prosim stahnete a nainstalujte Node.js LTS rucne z: https://nodejs.org" ([ConsoleColor]::Yellow)
        Write-Host ""
        Read-Host "Stisknete Enter pro ukonceni"
        exit 1
    }
}

if ($bunCmd) {
    $bunVer = & $bunCmd --version 2>$null
    Write-Color "   [OK] Runtime: Bun ($bunVer) detekovan" ([ConsoleColor]::Green)
} elseif ($nodeCmd) {
    $nodeVer = & $nodeCmd --version 2>$null
    Write-Color "   [OK] Runtime: Node.js ($nodeVer) detekovan" ([ConsoleColor]::Green)
}

# ---------------------------------------------------------------------------
# 3. KONTROLA KONFIGURACE (.env)
# ---------------------------------------------------------------------------
if (-not (Test-Path "$RootPath\.env") -and (Test-Path "$RootPath\.env.example")) {
    Write-Color "   [..] Vytvarim vychozi soubor .env z .env.example..." ([ConsoleColor]::Cyan)
    Copy-Item "$RootPath\.env.example" "$RootPath\.env"
    Write-Color "   [OK] Soubor .env byl vytvoren" ([ConsoleColor]::Green)
}

# ---------------------------------------------------------------------------
# 4. INSTALACE ZAVISLOSTÍ (POUZE PRI PRVNIM SPUSTENI NEBO -Reinstall)
# ---------------------------------------------------------------------------
if ($isFirstRun) {
    Write-Host ""
    Write-Color "📦 Instaluji projektove knihovny a zavislosti..." ([ConsoleColor]::Cyan)

    if ($bunCmd) {
        & $bunCmd install
    } else {
        & $npmCmd install
        if ($LASTEXITCODE -ne 0) {
            Write-Color "   [!] Zkousim instalaci s --legacy-peer-deps..." ([ConsoleColor]::Yellow)
            & $npmCmd install --legacy-peer-deps
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Color "[X] Chyba pri instalaci zavislosti." ([ConsoleColor]::Red)
        Read-Host "Stisknete Enter pro ukonceni"
        exit 1
    }

    $initTimestamp = (Get-Date).ToString("o")
    Set-Content -Path "$RootPath\.launcher-ready" -Value "Initialized at $initTimestamp" -Encoding UTF8
    Write-Color "   [OK] Vsechny knihovny byly uspesne nainstalovany." ([ConsoleColor]::Green)
}

# ---------------------------------------------------------------------------
# 4b. ZASTUPCE NA PLOSE (DESKTOP SHORTCUT)
# ---------------------------------------------------------------------------
$desktopDir = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$shortcutFile = Join-Path $desktopDir "SlideText Studio.lnk"
if (-not (Test-Path $shortcutFile)) {
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\create_desktop_shortcut.ps1" | Out-Null
        Write-Color "   [OK] Vytvoren zastupce na plose: SlideText Studio.lnk" ([ConsoleColor]::Green)
    } catch {}
}

# ---------------------------------------------------------------------------
# 5. KONTROLA VOLITELNYCH NASTROJU
# ---------------------------------------------------------------------------
$hasPowerPoint = $false
try {
    $ppt = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    $hasPowerPoint = $true
} catch {
    $hasPowerPoint = $false
}

if ($hasPowerPoint) {
    Write-Color "   [OK] Microsoft PowerPoint: Detekovan (Full HD 1080p nativni export snimku aktivni)" ([ConsoleColor]::Green)
} else {
    Write-Color "   [i]  Microsoft PowerPoint: Neni detekovan (pouzije se webovy generator nahledu)" ([ConsoleColor]::DarkGray)
}

$hasOllama = $false
try {
    $null = Invoke-RestMethod -Uri "http://127.0.0.1:11434/api/tags" -TimeoutSec 1 -ErrorAction Stop
    $hasOllama = $true
} catch {
    $hasOllama = $false
}

if ($hasOllama) {
    Write-Color "   [OK] Lokalni AI (Ollama): Bezi (Gemma 4 k dispozici)" ([ConsoleColor]::Green)
} else {
    Write-Color "   [i]  Lokalni AI (Ollama): Nebezi (volitelne - online prekladac funguje 100% bez ni)" ([ConsoleColor]::DarkGray)
}

Write-Host ""

# ---------------------------------------------------------------------------
# 6. SPUSTENI SERVERU A OTEVRENI PROHLIZECE
# ---------------------------------------------------------------------------
$appUrl = "http://localhost:$Port"
$alreadyRunning = Test-PortOpen "127.0.0.1" $Port

if ($alreadyRunning) {
    Write-Color "=================================================================" ([ConsoleColor]::Green)
    Write-Color "✨ SlideText Studio jiz bezi na adrese: $appUrl" ([ConsoleColor]::Green)
    Write-Color "=================================================================" ([ConsoleColor]::Green)
    Write-Host ""
    if (-not $NoBrowser) {
        Write-Color "🌐 Oteviram aplikaci ve vychozim prohlizeci..." ([ConsoleColor]::Cyan)
        Start-Process $appUrl
    }
    Write-Color "Stisknete Enter pro zavreni tohoto okna." ([ConsoleColor]::DarkGray)
    Read-Host
    exit 0
}

Write-Color "=================================================================" ([ConsoleColor]::Cyan)
Write-Color "🌐 Spoustim SlideText Studio server na $appUrl" ([ConsoleColor]::Yellow)
Write-Color "👉 Prohlizec se automaticky otevre po nabehnuti serveru." ([ConsoleColor]::Cyan)
Write-Color "🛑 Pro ukonceni stisknete kdykoliv Ctrl+C" ([ConsoleColor]::DarkGray)
Write-Color "=================================================================" ([ConsoleColor]::Cyan)
Write-Host ""

if (-not $NoBrowser) {
    $browserJob = Start-Job -ScriptBlock {
        param($targetUrl, $portNum)
        $maxTries = 40
        for ($i = 0; $i -lt $maxTries; $i++) {
            Start-Sleep -Milliseconds 400
            try {
                $client = New-Object System.Net.Sockets.TcpClient
                $asyncResult = $client.BeginConnect("127.0.0.1", $portNum, $null, $null)
                $success = $asyncResult.AsyncWaitHandle.WaitOne(400, $false)
                if ($success) {
                    $client.EndConnect($asyncResult)
                    $client.Close()
                    Start-Process $targetUrl
                    break
                }
                $client.Close()
            } catch {
                # Ceka na port
            }
        }
    } -ArgumentList $appUrl, $Port
}

try {
    if ($bunCmd) {
        & $bunCmd server.ts
    } else {
        & $npmCmd run dev
    }
} finally {
    if ($browserJob) {
        Stop-Job $browserJob -ErrorAction SilentlyContinue
        Remove-Job $browserJob -ErrorAction SilentlyContinue
    }
}
