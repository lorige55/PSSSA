# Ensure admin
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator."
    exit
}

# Variables
$TempPath = "$env:TEMP\AppInstallers"
$SyspinUrl = "https://github.com/PFCKrutonium/Windows-10-Taskbar-Pinning/releases/latest/download/syspin.exe"
$SyspinPath = "$TempPath\syspin.exe"

# Create temp directory
New-Item -ItemType Directory -Force -Path $TempPath | Out-Null

# Download syspin.exe
Write-Host "Downloading syspin.exe..."
Invoke-WebRequest -Uri $SyspinUrl -OutFile $SyspinPath -UseBasicParsing

if (!(Test-Path $SyspinPath)) {
    Write-Error "Failed to download syspin.exe. Aborting."
    exit
}

# Function to install and pin apps
function Install-App {
    param (
        [string]$Name,
        [string]$Url,
        [string]$InstallerPath,
        [string]$SilentArgs,
        [string]$ExeToPin
    )

    Write-Host "Downloading $Name..."
    Invoke-WebRequest -Uri $Url -OutFile $InstallerPath -UseBasicParsing

    if (Test-Path $InstallerPath) {
        Write-Host "Installing $Name..."
        Start-Process -FilePath $InstallerPath -ArgumentList $SilentArgs -Wait -NoNewWindow
        Remove-Item $InstallerPath -Force
        Write-Host "$Name installed.`n"

        if ($ExeToPin) {
            Write-Host "Pinning $Name to taskbar..."
            Start-Process -FilePath $SyspinPath -ArgumentList "`"$ExeToPin`"", "5386" -Wait
        }
    } else {
        Write-Warning "Failed to download $Name."
    }
}

# --- Unpin all current taskbar items ---
Write-Host "Unpinning current taskbar items..."
$taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
if (Test-Path $taskbarPath) {
    Get-ChildItem $taskbarPath -Include *.lnk -Force | Remove-Item -Force
}
Stop-Process -Name explorer -Force
Start-Process explorer

# --- Pin File Explorer first ---
$explorer = "$env:WINDIR\explorer.exe"
Write-Host "Pinning File Explorer..."
Start-Process -FilePath $SyspinPath -ArgumentList "`"$explorer`"", "5386" -Wait

# --- 1. Firefox ---
Install-App -Name "Firefox" `
    -Url "https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US" `
    -InstallerPath "$TempPath\Firefox.exe" `
    -SilentArgs "/S" `
    -ExeToPin "C:\Program Files\Mozilla Firefox\firefox.exe"

# --- 2. AnyDesk ---
Install-App -Name "AnyDesk" `
    -Url "https://download.anydesk.com/AnyDesk.exe" `
    -InstallerPath "$TempPath\AnyDesk.exe" `
    -SilentArgs "--install" `
    -ExeToPin "C:\Program Files (x86)\AnyDesk\AnyDesk.exe"

# --- 3. PDF24 Creator ---
Install-App -Name "PDF24 Creator" `
    -Url "https://tools.pdf24.org/static/builds/pdf24-creator.exe" `
    -InstallerPath "$TempPath\PDF24.exe" `
    -SilentArgs "/VERYSILENT" `
    -ExeToPin "C:\Program Files\PDF24\pdf24.exe"

# --- 4. Adobe Acrobat Reader ---
Install-App -Name "Adobe Acrobat Reader" `
    -Url "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820419/AcroRdrDC2300820419_en_US.exe" `
    -InstallerPath "$TempPath\AdobeReader.exe" `
    -SilentArgs "/sAll /rs /rps /msi EULA_ACCEPT=YES" `
    -ExeToPin "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"

# --- 5. WhatsApp via winget ---
Write-Host "Installing WhatsApp (via winget)..."
Start-Process "winget" -ArgumentList "install -e --id 9NKSQGP7F2NH --accept-package-agreements --accept-source-agreements" -Wait

# Try to pin WhatsApp if the .lnk exists
$whatsappLnk = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\WhatsApp.lnk"
if (Test-Path $whatsappLnk) {
    Write-Host "Pinning WhatsApp..."
    Start-Process -FilePath $SyspinPath -ArgumentList "`"$whatsappLnk`"", "5386" -Wait
} else {
    Write-Warning "WhatsApp shortcut not found for pinning."
}

# --- Clean up ---
Write-Host "Cleaning up..."
Remove-Item $SyspinPath -Force -ErrorAction SilentlyContinue
Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n✅ All programs installed and pinned. Script complete."