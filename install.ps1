<# 
    Install + Pin Script (Neu/Überarbeitet)
    ---------------------------------------
    Funktionen:
      - Selbst-Elevation zu Administrator
      - Download von syspin.exe (mehrere Quellen + Fallback + BITS)
      - Taskleiste leeren / zurücksetzen
      - Installation definierter Anwendungen (direkter Download oder winget)
      - Silent-Parameter pro App
      - Taskleisten-Pinning in definierter Reihenfolge (Explorer zuerst)
      - WhatsApp (Store/winget) + dynamische Shortcut-Suche
      - Aufräumen temporärer Dateien
      - Optionale Protokollierung
      - Parameter zur Auswahl von Apps / SkipPinning / SkipUnpin

    Aufrufbeispiele:
      powershell -ExecutionPolicy Bypass -File .\install.ps1
      powershell -File .\install.ps1 -Apps Firefox,AnyDesk -Verbose
      irm https://raw.githubusercontent.com/lorige55/PSSSA/refs/heads/main/install.ps1 | iex

    Getestet unter: Windows 10/11, PowerShell 5.1 / 7.x

#>

[CmdletBinding()]
param(
    [string[]]$Apps = @("Firefox","AnyDesk","PDF24","AdobeReader","WhatsApp"),
    [switch]$SkipPinning,
    [switch]$SkipUnpin,
    [switch]$KeepInstallers,
    [string]$LogFile
)

# -----------------------------
# Grund-Konfiguration
# -----------------------------
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($LogFile) {
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $LogFile) | Out-Null
        Start-Transcript -Path $LogFile -Append | Out-Null
    } catch {
        Write-Warning ("Konnte Logging nicht starten: {0}" -f $_.Exception.Message)
    }
}

function Write-Section {
    param([string]$Text)
    Write-Host "`n==== $Text ====" -ForegroundColor Cyan
}

# -----------------------------
# Selbst-Elevation
# -----------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Starte erneut als Administrator..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $rawArgs = @()
    if ($MyInvocation.BoundParameters.Count -gt 0) {
        foreach ($k in $MyInvocation.BoundParameters.Keys) {
            $v = $MyInvocation.BoundParameters[$k]
            if ($v -is [System.Array]) {
                $rawArgs += "-$k", ($v -join ",")
            } elseif ($v -is [switch]) {
                if ($v.IsPresent) { $rawArgs += "-$k" }
            } else {
                $rawArgs += "-$k", $v
            }
        }
    }
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $($rawArgs -join ' ')"
    $psi.Verb = "runas"
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

# -----------------------------
# TLS 1.2 sicherstellen
# -----------------------------
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "Konnte TLS 1.2 nicht setzen. Fahre fort..."
}

# -----------------------------
# Pfade
# -----------------------------
$TempRoot       = Join-Path $env:TEMP "AppInstallers"
$InstallerPath  = $TempRoot
$SyspinPath     = Join-Path $TempRoot "syspin.exe"
$LocalSyspin1   = Join-Path $PSScriptRoot "syspin.exe"
$LocalSyspin2   = Join-Path (Get-Location) "syspin.exe"
$SyspinDownloaded = $false

New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

# -----------------------------
# syspin Quellen
# -----------------------------
$SyspinUrls = @(
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/raw/master/syspin.exe",
  "https://raw.githubusercontent.com/PFCKrutonium/Windows-10-Taskbar-Pinning/master/syspin.exe",
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/releases/latest/download/syspin.exe"
)

function Get-Syspin {
    param([string[]]$Urls)

    foreach ($p in @($LocalSyspin1, $LocalSyspin2)) {
        if (Test-Path $p -PathType Leaf) {
            Write-Host "Verwende lokales syspin.exe: $p"
            return $p
        }
    }

    foreach ($u in $Urls) {
        try {
            Write-Host "Lade syspin.exe von $u ..."
            Invoke-WebRequest -Uri $u -OutFile $SyspinPath -UseBasicParsing -MaximumRedirection 5 -ErrorAction Stop
            if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
                $script:SyspinDownloaded = $true
                return $SyspinPath
            }
        } catch {
            Write-Warning ("syspin Download fehlgeschlagen von {0}: {1}" -f $u, $_.Exception.Message)
        }
    }

    try {
        Write-Host "Versuche BITS Transfer für syspin..."
        Start-BitsTransfer -Source $Urls[0] -Destination $SyspinPath -ErrorAction Stop
        if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
            $script:SyspinDownloaded = $true
            return $SyspinPath
        }
    } catch {
        Write-Warning ("BITS Transfer fehlgeschlagen: {0}" -f $_.Exception.Message)
    }

    return $null
}

$SyspinExe = if ($SkipPinning) { $null } else { Get-Syspin -Urls $SyspinUrls }
$PinningEnabled = [bool]$SyspinExe -and -not $SkipPinning
if (-not $PinningEnabled) {
    Write-Warning "syspin.exe nicht verfügbar oder Pinning übersprungen. Es wird NICHT gepinnt."
}

# -----------------------------
# Hilfsfunktionen
# -----------------------------
function Pin-ToTaskbar {
    param(
        [Parameter(Mandatory)][string]$PathToExeOrLnk,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not $PinningEnabled) { return }

    if (Test-Path $PathToExeOrLnk) {
        Write-Host ("Pinne {0} ..." -f $Label)
        try {
            Start-Process -FilePath $SyspinExe -ArgumentList "`"$PathToExeOrLnk`"", "5386" -Wait
        } catch {
            Write-Warning ("Pinning {0} fehlgeschlagen: {1}" -f $Label, $_.Exception.Message)
        }
    } else {
        Write-Warning ("Konnte Pfad zum Pinnen nicht finden für {0}: {1}" -f $Label, $PathToExeOrLnk)
    }
}

function Download-And-Install {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$OutFileName,
        [string]$SilentArgs = "",
        [string]$ExeToPin
    )

    $dst = Join-Path $InstallerPath $OutFileName
    try {
        Write-Host ("Lade {0} herunter..." -f $Name)
        Invoke-WebRequest -Uri $Url -OutFile $dst -UseBasicParsing -ErrorAction Stop
        Write-Host ("Installiere {0}..." -f $Name)
        Start-Process -FilePath $dst -ArgumentList $SilentArgs -Wait -NoNewWindow
    } catch {
        Write-Warning ("{0} Download/Installation fehlgeschlagen: {1}" -f $Name, $_.Exception.Message)
    } finally {
        if (-not $KeepInstallers -and (Test-Path $dst)) {
            Remove-Item $dst -Force -ErrorAction SilentlyContinue
        }
    }
    if ($ExeToPin) { Pin-ToTaskbar -PathToExeOrLnk $ExeToPin -Label $Name }
}

function Install-WhatsApp {
    Write-Host "Installiere WhatsApp (winget)..."
    try {
        Start-Process "winget" -ArgumentList "install -e --id 9NKSQGP7F2NH --accept-package-agreements --accept-source-agreements" -Wait
    } catch {
        Write-Warning ("winget WhatsApp Installation fehlgeschlagen: {0}" -f $_.Exception.Message)
    }

    $waCandidates = @(
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk",
        "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk"
    ) + (Get-ChildItem "$env:APPDATA\Microsoft\Windows\Start Menu\Programs" -Filter "*WhatsApp*.lnk" -ErrorAction SilentlyContinue |
         Select-Object -Expand FullName)

    foreach ($wa in $waCandidates | Where-Object { $_ -and (Test-Path $_) }) {
        Pin-ToTaskbar -PathToExeOrLnk $wa -Label "WhatsApp"
        break
    }
}

# -----------------------------
# Taskleiste leeren
# -----------------------------
if (-not $SkipUnpin) {
    Write-Section "Alte Taskleisten-Symbole entfernen"
    try {
        $taskbarPins = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
        if (Test-Path $taskbarPins) {
            Get-ChildItem $taskbarPins -Include *.lnk -Force | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        Stop-Process -Name explorer -Force
        Start-Process explorer
    } catch {
        Write-Warning ("Konnte Taskleiste nicht vollständig zurücksetzen: {0}" -f $_.Exception.Message)
    }
} else {
    Write-Host "Überspringe Unpin-Vorgang (--SkipUnpin)."
}

# -----------------------------
# Explorer zuerst pinnen
# -----------------------------
Pin-ToTaskbar -PathToExeOrLnk "$env:WINDIR\explorer.exe" -Label "File Explorer"

# -----------------------------
# App-Definitionen
# -----------------------------
$AppCatalog = @(
    @{
        Key        = "Firefox"
        Name       = "Firefox"
        Type       = "Direct"
        Url        = "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US"
        OutFile    = "Firefox.exe"
        Silent     = "/S"
        PinPath    = "C:\Program Files\Mozilla Firefox\firefox.exe"
    }
    @{
        Key        = "AnyDesk"
        Name       = "AnyDesk"
        Type       = "Direct"
        Url        = "https://download.anydesk.com/AnyDesk.exe"
        OutFile    = "AnyDesk.exe"
        Silent     = "--install"
        PinPath    = "C:\Program Files (x86)\AnyDesk\AnyDesk.exe"
    }
    @{
        Key        = "PDF24"
        Name       = "PDF24 Creator"
        Type       = "Direct"
        Url        = "https://tools.pdf24.org/static/builds/pdf24-creator.exe"
        OutFile    = "PDF24.exe"
        Silent     = "/VERYSILENT"
        PinPath    = "C:\Program Files\PDF24\pdf24.exe"
    }
    @{
        Key        = "AdobeReader"
        Name       = "Adobe Acrobat Reader"
        Type       = "Direct"
        Url        = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820419/AcroRdrDC2300820419_en_US.exe"
        OutFile    = "AdobeReader.exe"
        Silent     = "/sAll /rs /rps /msi EULA_ACCEPT=YES"
        PinPath    = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    }
    @{
        Key        = "WhatsApp"
        Name       = "WhatsApp"
        Type       = "Winget"
        Url        = ""
        OutFile    = ""
        Silent     = ""
        PinPath    = ""
    }
)

# -----------------------------
# Installation
# -----------------------------
Write-Section "Installationen starten"

foreach ($appKey in $Apps) {
    $entry = $AppCatalog | Where-Object { $_.Key -ieq $appKey }
    if (-not $entry) {
        Write-Warning ("Unbekannter App-Key: {0} (übersprungen)" -f $appKey)
        continue
    }

    switch ($entry.Type) {
        "Direct" {
            Download-And-Install -Name $entry.Name -Url $entry.Url -OutFileName $entry.OutFile -SilentArgs $entry.Silent -ExeToPin $entry.PinPath
        }
        "Winget" {
            if ($entry.Key -eq "WhatsApp") {
                Install-WhatsApp
            } else {
                Write-Warning ("Winget-App nicht implementiert: {0}" -f $entry.Name)
            }
        }
        default {
            Write-Warning ("Unbekannter Installations-Typ für {0}: {1}" -f $entry.Name, $entry.Type)
        }
    }
}

# -----------------------------
# Aufräumen
# -----------------------------
Write-Section "Aufräumen"
if ($SyspinDownloaded -and (Test-Path $SyspinExe) -and -not $KeepInstallers) {
    Write-Host "Entferne syspin.exe ..."
    Remove-Item $SyspinExe -Force -ErrorAction SilentlyContinue
}
if ((Test-Path $TempRoot) -and -not $KeepInstallers) {
    Remove-Item $TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n✅ Fertig: Apps installiert und (falls möglich) in Reihenfolge gepinnt."

if ($LogFile) {
    try { Stop-Transcript | Out-Null } catch {}
}
