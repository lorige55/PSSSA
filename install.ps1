# --- Force TLS 1.2 for GitHub downloads ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Temp folder
$TempPath = Join-Path $env:TEMP "AppInstallers"
New-Item -ItemType Directory -Force -Path $TempPath | Out-Null

# Where we want syspin to be (if downloaded)
$SyspinPath   = Join-Path $TempPath "syspin.exe"
$LocalSyspin1 = Join-Path $PSScriptRoot "syspin.exe"   # fallback: same folder as the script
$LocalSyspin2 = Join-Path (Get-Location) "syspin.exe"  # fallback: current working dir
$SyspinDownloaded = $false

# Direct URLs (no redirect) then a releases link as last resort
$SyspinUrls = @(
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/raw/master/syspin.exe",
  "https://raw.githubusercontent.com/PFCKrutonium/Windows-10-Taskbar-Pinning/master/syspin.exe",
  "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/releases/latest/download/syspin.exe"
)

function Get-Syspin {
    param([string[]]$Urls)

    # 1) Try local copies first (useful on locked-down networks)
    foreach ($p in @($LocalSyspin1, $LocalSyspin2)) {
        if (Test-Path $p -PathType Leaf) {
            Write-Host "Using local syspin.exe at: $p"
            return $p
        }
    }

    # 2) Try to download from the list of URLs
    foreach ($u in $Urls) {
        try {
            Write-Host "Downloading syspin.exe from $u ..."
            Invoke-WebRequest -Uri $u -OutFile $SyspinPath -UseBasicParsing -MaximumRedirection 5 -ErrorAction Stop
            if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
                $script:SyspinDownloaded = $true
                return $SyspinPath
            }
        } catch {
            Write-Warning "Download failed from $u : $($_.Exception.Message)"
        }
    }

    # 3) Last-ditch: BITS (often works behind some proxies)
    try {
        Write-Host "Trying BITS transfer for the first URL..."
        Start-BitsTransfer -Source $Urls[0] -Destination $SyspinPath -ErrorAction Stop
        if ((Test-Path $SyspinPath) -and ((Get-Item $SyspinPath).Length -gt 0)) {
            $script:SyspinDownloaded = $true
            return $SyspinPath
        }
    } catch {
        Write-Warning "BITS transfer also failed: $($_.Exception.Message)"
    }

    return $null
}

$SyspinExe = Get-Syspin -Urls $SyspinUrls
if (-not $SyspinExe) {
    Write-Warning "Could not obtain syspin.exe. Pinning will be skipped, installs will continue."
    $PinningEnabled = $false
} else {
    $PinningEnabled = $true
}

# Helper to pin (only if we have syspin)
function Pin-ToTaskbar {
    param([string]$PathToExeOrLnk, [string]$Label)
    if ($PinningEnabled -and (Test-Path $PathToExeOrLnk)) {
        Write-Host "Pinning $Label ..."
        Start-Process -FilePath $SyspinExe -ArgumentList "`"$PathToExeOrLnk`"", "5386" -Wait
    } elseif ($PinningEnabled) {
        Write-Warning "Could not find path to pin for $Label: $PathToExeOrLnk"
    }
}

# Example usage in your script:
# 1) unpin current items (your existing logic)
# 2) Pin File Explorer first:
if ($PinningEnabled) {
    Pin-ToTaskbar -PathToExeOrLnk "$env:WINDIR\explorer.exe" -Label "File Explorer"
}

# Later, after each app installs, call:
# Pin-ToTaskbar -PathToExeOrLnk "C:\Program Files\Mozilla Firefox\firefox.exe" -Label "Firefox"
# Pin-ToTaskbar -PathToExeOrLnk "C:\Program Files (x86)\AnyDesk\AnyDesk.exe" -Label "AnyDesk"
# Pin-ToTaskbar -PathToTaskbar "C:\Program Files\PDF24\pdf24.exe" -Label "PDF24 Creator"
# Pin-ToTaskbar -PathToExeOrLnk "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe" -Label "Adobe Reader"
# For WhatsApp after winget install, try the Start Menu link:
# Pin-ToTaskbar -PathToExeOrLnk "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk" -Label "WhatsApp"

# --- Clean up downloaded syspin when finished ---
if ($SyspinDownloaded -and (Test-Path $SyspinExe)) {
    Write-Host "Cleaning up syspin.exe ..."
    Remove-Item $SyspinExe -Force -ErrorAction SilentlyContinue
}