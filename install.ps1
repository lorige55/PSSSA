# =========================
# Install + Pin Script
# =========================

#--- Self-elevate to Administrator ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Re-launching as Administrator..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $psi.Verb = "runas"
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

# --- Use TLS 1.2 for GitHub downloads ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Paths ---
$TempPath       = Join-Path $env:TEMP "AppInstallers"
$InstallerPath  = $TempPath   # keep installers here
$SyspinPath     = Join-Path $TempPath "syspin.exe"
$LocalSyspin1   = Join-Path $PSScriptRoot "syspin.exe"   # local fallback
$LocalSyspin2   = Join-Path (Get-Location) "syspin.exe"  # local fallback
$SyspinDownloaded = $false

New-Item -ItemType Directory -Force -Path $TempPath | Out-Null

# --- Robust syspin getter (direct URLs + BITS + local fallback) ---
$SyspinUrls = @(
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/raw/master/syspin.exe",
  "https://raw.githubusercontent.com/PFCKrutonium/Windows-10-Taskbar-Pinning/master/syspin.exe",
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/releases/latest/download/syspin.exe"
)

function Get-Syspin {
    param([string[]]$Urls)

    foreach ($p in @($LocalSyspin1, $LocalSyspin2)) {
        if (Test-Path $p -PathType Leaf) {
            Write-Host "Using local syspin.exe at: $p"
            return $p
        }
    }

    foreach ($u in $Urls) {
        try {
            Write-Host "Downloading syspin.exe from $u ..."
            Invoke-WebRequest -Uri $u -OutFile $SyspinPath -UseBasicParsing -MaximumRedirection 5 -ErrorAction Stop
            if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
                $script:SyspinDownloaded = $true
                return $SyspinPath
            }
        } catch {
            Write-Warning ("syspin download failed from {0}: {1}" -f $u, $_.Exception.Message)
        }
    }

    try {
        Write-Host "Trying BITS transfer for syspin..."
        Start-BitsTransfer -Source $Urls[0] -Destination $SyspinPath -ErrorAction Stop
        if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
            $script:SyspinDownloaded = $true
            return $SyspinPath
        }
    } catch {
        Write-Warning ("BITS transfer failed: {0}" -f $_.Exception.Message)
    }

    return $null
}

$SyspinExe = Get-Syspin -Urls $SyspinUrls
$PinningEnabled = [bool]$SyspinExe
if (-not $PinningEnabled) {
    Write-Warning "syspin.exe unavailable. I'll install apps but skip pinning."
}

# --- Helpers ---
function Pin-ToTaskbar {
    param([string]$PathToExeOrLnk, [string]$Label)

    if ($PinningEnabled -and (Test-Path $PathToExeOrLnk)) {
        Write-Host ("Pinning {0} ..." -f $Label)
        Start-Process -FilePath $SyspinExe -ArgumentList "`"$PathToExeOrLnk`"", "5386" -Wait
    } elseif ($PinningEnabled) {
        Write-Warning ("Could not find path to pin for {0}: {1}" -f $Label, $PathToExeOrLnk)
    }
}

function Install-App {
    param(
        [string]$Name,
        [string]$Url,
        [string]$OutFileName,
        [string]$SilentArgs,
        [string]$ExeToPin
    )
    $dst = Join-Path $InstallerPath $OutFileName
    try {
        Write-Host ("Downloading {0} ..." -f $Name)
        Invoke-WebRequest -Uri $Url -OutFile $dst -UseBasicParsing -ErrorAction Stop
        Write-Host ("Installing {0} ..." -f $Name)
        Start-Process -FilePath $dst -ArgumentList $SilentArgs -Wait -NoNewWindow
    } catch {
        Write-Warning ("{0} download/install failed: {1}" -f $Name, $_.Exception.Message)
    } finally {
        if (Test-Path $dst) { Remove-Item $dst -Force -ErrorAction SilentlyContinue }
    }
    if ($ExeToPin) { Pin-ToTaskbar -PathToExeOrLnk $ExeToPin -Label $Name }
}

# --- Unpin everything currently on the taskbar ---
try {
    Write-Host "Unpinning current taskbar items..."
    $taskbarPins = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    if (Test-Path $taskbarPins) {
        Get-ChildItem $taskbarPins -Include *.lnk -Force | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    Stop-Process -Name explorer -Force
    Start-Process explorer
} catch { Write-Warning ("Could not fully reset taskbar: {0}" -f $_.Exception.Message) }

# --- Pin File Explorer first ---
Pin-ToTaskbar -PathToExeOrLnk "$env:WINDIR\explorer.exe" -Label "File Explorer"

# --- Install & pin in order ---

# 1) Firefox
Install-App -Name "Firefox" `
    -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" `
    -OutFileName "Firefox.exe" `
    -SilentArgs "/S" `
    -ExeToPin "C:\Program Files\Mozilla Firefox\firefox.exe"

# 2) AnyDesk
Install-App -Name "AnyDesk" `
    -Url "https://download.anydesk.com/AnyDesk.exe" `
    -OutFileName "AnyDesk.exe" `
    -SilentArgs "--install" `
    -ExeToPin "C:\Program Files (x86)\AnyDesk\AnyDesk.exe"

# 3) PDF24 Creator
Install-App -Name "PDF24 Creator" `
    -Url "https://tools.pdf24.org/static/builds/pdf24-creator.exe" `
    -OutFileName "PDF24.exe" `
    -SilentArgs "/VERYSILENT" `
    -ExeToPin "C:\Program Files\PDF24\pdf24.exe"

# 4) Adobe Acrobat Reader
Install-App -Name "Adobe Acrobat Reader" `
    -Url "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820419/AcroRdrDC2300820419_en_US.exe" `
    -OutFileName "AdobeReader.exe" `
    -SilentArgs "/sAll /rs /rps /msi EULA_ACCEPT=YES" `
    -ExeToPin "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

# 5) WhatsApp (Store)
Write-Host "Installing WhatsApp (winget)..."
try {
    Start-Process "winget" -ArgumentList "install -e --id 9NKSQGP7F2NH --accept-package-agreements --accept-source-agreements" -Wait
} catch { Write-Warning ("winget WhatsApp install failed: {0}" -f $_.Exception.Message) }

# Try to locate a WhatsApp shortcut to pin
$waCandidates = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk",
    "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk"
) + (Get-ChildItem "$env:APPDATA\Microsoft\Windows\Start Menu\Programs" -Filter "*WhatsApp*.lnk" -ErrorAction SilentlyContinue |
     Select-Object -Expand FullName)

foreach ($wa in $waCandidates | Where-Object { $_ -and (Test-Path $_) }) {
    Pin-ToTaskbar -PathToExeOrLnk $wa -Label "WhatsApp"
    break
}

# --- Cleanup ---
if ($SyspinDownloaded -and (Test-Path $SyspinExe)) {
    Write-Host "Cleaning up syspin.exe ..."
    Remove-Item $SyspinExe -Force -ErrorAction SilentlyContinue
}
if (Test-Path $TempPath) {
    Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`n✅ Done: apps installed and pinned in order (Explorer, then each app)."