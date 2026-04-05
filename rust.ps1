# ==============================================================================
# Deploy-RustDesk.ps1  —  TEBIN IT / RustDesk One-Command Deploy
#
# Run as Administrator:
#   powershell -ExecutionPolicy Bypass -File Deploy-RustDesk.ps1
#
# Requirements:
#   - Windows 10/11 or Server 2016+  (no extra tools needed)
#   - Administrator privileges
#   - Internet access
# ==============================================================================

#Requires -RunAsAdministrator

$ErrorActionPreference = 'SilentlyContinue'

# ==== Settings ====
$secretsUrl  = "https://raw.githubusercontent.com/4aykas/rust/main/secrets.zip"
$archivePath = "C:\Temp\rd_secrets.zip"
$extractPath = "C:\Temp\rd_secrets_tmp"
$logFile     = "C:\Temp\rustdesk_deploy.log"
$installTemp = "C:\Temp\rustdesk_setup.exe"
$rustdeskExe = "C:\Program Files\RustDesk\rustdesk.exe"
$userConfig  = "C:\Users\$env:USERNAME\AppData\Roaming\RustDesk\config\RustDesk2.toml"
$svcConfig   = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk2.toml"
$rdLogDir    = "$env:APPDATA\RustDesk\log"

# ==== Helpers ====
function Write-Log {
    param([string]$msg, [ValidateSet('INFO','WARN','ERROR')][string]$lvl = 'INFO')
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts [$lvl] $msg" | Out-File -Append -FilePath $logFile
}
function Show {
    param([string]$msg, [string]$color = 'Cyan')
    Write-Host "  $msg" -ForegroundColor $color
}
function Cleanup {
    Remove-Item $archivePath  -Force -ErrorAction SilentlyContinue
    Remove-Item $extractPath  -Recurse -Force -ErrorAction SilentlyContinue
}

# ==== Banner ====
Clear-Host
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════╗" -ForegroundColor DarkCyan
Write-Host "  ║      RustDesk One-Command Deploy              ║" -ForegroundColor DarkCyan
Write-Host "  ║      TEBIN IT                                 ║" -ForegroundColor DarkCyan
Write-Host "  ╚═══════════════════════════════════════════════╝" -ForegroundColor DarkCyan
Write-Host ""

# ==== Init ====
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory "C:\Temp" -Force | Out-Null }
Write-Log "==== Deploy Start ==== Host: $env:COMPUTERNAME | User: $env:USERNAME"

# ==== Download secrets.zip ====
Show "Downloading secrets archive from GitHub..."
Write-Log "Fetching: $secretsUrl"
try {
    Invoke-WebRequest -Uri $secretsUrl -OutFile $archivePath -UseBasicParsing -ErrorAction Stop
    Write-Log "Archive downloaded."
} catch {
    Write-Log "Download failed: $_" 'ERROR'
    Show "ERROR: Could not download secrets archive. Check network." 'Red'
    exit 1
}

# ==== Prompt Password ====
Write-Host ""
$securePass = Read-Host "  Enter archive password" -AsSecureString
$plainPass  = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass))
Write-Host ""

# ==== Extract ZIP with password via Shell.Application (native Windows COM) ====
Write-Log "Extracting password-protected ZIP..."

$rdServer = $null; $rdKey = $null; $rdPass = $null

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    New-Item -ItemType Directory $extractPath -Force | Out-Null

    # Open ZIP with password using .NET ZipArchive
    $stream  = [IO.File]::OpenRead($archivePath)
    $zipArgs = [IO.Compression.ZipArchiveMode]::Read
    $zip     = New-Object IO.Compression.ZipArchive($stream, $zipArgs)

    $entry = $zip.Entries | Where-Object { $_.Name -eq "rustdesk-secrets.txt" } | Select-Object -First 1

    if (-not $entry) {
        throw "rustdesk-secrets.txt not found inside archive."
    }

    # .NET ZipArchive doesn't support passwords natively — use a memory trick:
    # Read raw entry bytes then decrypt ZipCrypto manually, OR
    # Use Shell.Application COM which supports standard ZIP passwords on Windows
    $zip.Dispose(); $stream.Dispose()

    # Shell.Application approach — works natively on all Windows versions
    $shell    = New-Object -ComObject Shell.Application
    $zipObj   = $shell.NameSpace($archivePath)
    $destObj  = $shell.NameSpace($extractPath)

    # Set password via Shell (this is how Windows Explorer handles it)
    # FOF_SILENT(4) + FOF_NOCONFIRMATION(16) + FOF_NOERRORUI(1024) = 1044
    $destObj.CopyHere($zipObj.Items(), 1044)
    Start-Sleep -Seconds 2

    $secretsFile = Join-Path $extractPath "rustdesk-secrets.txt"

    # If Shell.Application extracted without password prompt (wrong pwd = empty file or no file)
    if (-not (Test-Path $secretsFile) -or (Get-Item $secretsFile).Length -eq 0) {
        # Shell.Application on modern Windows may prompt interactively for password
        # Fallback: use a .NET workaround with ICSharpCode or direct byte manipulation
        # Best cross-version approach: write a temp VBScript to handle password
        $vbs = @"
Dim oShell, oZip, oDest, oItems
Set oShell = CreateObject("Shell.Application")
Set oZip   = oShell.NameSpace("$archivePath")
Set oDest  = oShell.NameSpace("$extractPath")
oZip.Self.InvokeVerbEx "open", "$plainPass"
oDest.CopyHere oZip.Items(), 1044
WScript.Sleep 2000
"@
        $vbsPath = "C:\Temp\rd_extract.vbs"
        Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII
        Start-Process "cscript.exe" -ArgumentList "//NoLogo `"$vbsPath`"" -Wait -WindowStyle Hidden
        Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }

    if (-not (Test-Path $secretsFile)) {
        throw "Extraction failed — wrong password or corrupt archive."
    }

    # Parse secrets
    Get-Content $secretsFile | ForEach-Object {
        $l = $_.Trim()
        if ($l -match '^\s*#|^$') { return }
        if ($l -match '^SERVER\s*=\s*(.+)$')   { $rdServer = $Matches[1].Trim() }
        if ($l -match '^KEY\s*=\s*(.+)$')       { $rdKey    = $Matches[1].Trim() }
        if ($l -match '^PASSWORD\s*=\s*(.+)$')  { $rdPass   = $Matches[1].Trim() }
    }

    Write-Log "Secrets parsed."

} catch {
    $plainPass = $null; [GC]::Collect()
    Cleanup
    Write-Log "Decryption/extraction failed: $_" 'ERROR'
    Show "✖  Incorrect password or corrupt archive. Nothing deployed." 'Red'
    Write-Host ""; exit 1
}

# Wipe archive and extracted files from disk immediately
Cleanup
$plainPass = $null; [GC]::Collect()

# ==== Validate Secrets ====
$missing = @()
if (-not $rdServer) { $missing += 'SERVER' }
if (-not $rdKey)    { $missing += 'KEY' }
if (-not $rdPass)   { $missing += 'PASSWORD' }
if ($missing.Count -gt 0) {
    Write-Log "Missing fields: $($missing -join ', ')" 'ERROR'
    Show "ERROR: Secrets incomplete. Missing: $($missing -join ', ')" 'Red'
    exit 1
}

Show "✔  Password accepted." 'Green'
Write-Log "Secrets validated. Server: $rdServer"

# ==== Get Latest RustDesk Version ====
Show "Checking latest RustDesk version..."
Write-Log "Querying GitHub API..."
try {
    $release     = Invoke-RestMethod -Uri "https://api.github.com/repos/rustdesk/rustdesk/releases/latest" -UseBasicParsing -ErrorAction Stop
    $latestVer   = $release.tag_name -replace '^v', ''
    $asset       = $release.assets | Where-Object { $_.name -match "x86_64\.exe$" } | Select-Object -First 1
    $downloadUrl = $asset.browser_download_url
    Show "Latest version: $latestVer" 'White'
    Write-Log "Latest: v$latestVer"
} catch {
    $latestVer   = "1.4.3"
    $downloadUrl = "https://github.com/rustdesk/rustdesk/releases/download/1.4.3/rustdesk-1.4.3-x86_64.exe"
    Show "GitHub API unavailable — fallback v$latestVer" 'Yellow'
    Write-Log "API failed, fallback v$latestVer" 'WARN'
}

# ==== Check Installed Version ====
$skipInstall = $false
if (Test-Path $rustdeskExe) {
    try {
        $installedVer = (Get-Command $rustdeskExe).FileVersionInfo.ProductVersion
        Write-Log "Installed: v$installedVer"
        if ([version]$installedVer -ge [version]$latestVer) {
            Show "RustDesk v$installedVer already up to date." 'Green'
            Write-Log "Up to date. Skipping install."
            $skipInstall = $true
        } else {
            Show "Updating v$installedVer → v$latestVer" 'Yellow'
            Write-Log "Updating v$installedVer → v$latestVer"
        }
    } catch {
        Write-Log "Cannot read installed version. Reinstalling." 'WARN'
    }
} else {
    Show "RustDesk not found. Installing v$latestVer..."
    Write-Log "Fresh install."
}

# ==== Download & Install ====
if (-not $skipInstall) {
    Show "Downloading RustDesk v$latestVer..."
    Write-Log "Downloading installer..."
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installTemp -UseBasicParsing -ErrorAction Stop
        Write-Log "Installer downloaded."
    } catch {
        Write-Log "Installer download failed: $_" 'ERROR'
        Show "ERROR: Could not download RustDesk installer." 'Red'
        exit 1
    }

    Show "Installing silently (please wait)..."
    Write-Log "Running silent install..."
    Start-Process -FilePath $installTemp -ArgumentList "--silent-install" -PassThru | Wait-Process
    Write-Log "Install complete."
    Start-Sleep -Seconds 5

    Write-Log "Stopping RustDesk before config..."
    net stop rustdesk 2>&1 | Out-Null
    Get-Process rustdesk -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Remove-Item $installTemp -Force -ErrorAction SilentlyContinue
}

# ==== Write TOML Config ====
$toml = @"
rendezvous_server = '${rdServer}:21116'
nat_type = 1
serial = 0

[options]
custom-rendezvous-server = '$rdServer'
key = '$rdKey'
whitelist = ''
direct-server = 'Y'
direct-access-port = '21118'
"@

foreach ($cfgPath in @($userConfig, $svcConfig)) {
    $dir = Split-Path $cfgPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
    Set-Content -Path $cfgPath -Value $toml -Encoding UTF8
    Write-Log "Config written: $cfgPath"
}

# ==== Start Service ====
Write-Log "Starting RustDesk service..."
net start rustdesk 2>&1 | Out-Null
Start-Sleep -Seconds 5

# ==== Apply Password ====
Write-Log "Setting access password..."
Start-Process -FilePath $rustdeskExe -ArgumentList "--password", $rdPass -Wait
Start-Sleep -Seconds 3

$rdPass = $null; $rdKey = $null; [GC]::Collect()

# ==== Get Machine ID ====
Write-Log "Retrieving RustDesk machine ID..."
$machineId = $null

# Method 1: RustDesk.toml
$idConfig = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk.toml"
if (Test-Path $idConfig) {
    $idLine = Get-Content $idConfig | Where-Object { $_ -match '^id\s*=' }
    if ($idLine -match "=\s*'?([0-9]+)'?") { $machineId = $Matches[1].Trim() }
}

# Method 2: Scan logs
if (-not $machineId -and (Test-Path $rdLogDir)) {
    $logs = Get-ChildItem $rdLogDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    foreach ($log in $logs) {
        $match = Select-String -Path $log.FullName -Pattern "My\s+ID:\s*(\d+)" | Select-Object -Last 1
        if ($match) {
            $machineId = ($match.Line -replace '.*My\s+ID:\s*(\d+).*', '$1').Trim()
            break
        }
    }
}

# Method 3: --get-id flag
if (-not $machineId) {
    $idOutput = & $rustdeskExe --get-id 2>&1
    if ($idOutput -match '(\d{9,})') { $machineId = $Matches[1] }
}

Write-Log "Machine ID: $machineId"

# ==== Final Output ====
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "  ║   ✔  Deployment Complete                      ║" -ForegroundColor Green
Write-Host "  ╚═══════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Server  : $rdServer"  -ForegroundColor Gray
Write-Host "  Version : $latestVer" -ForegroundColor Gray

if ($machineId) {
    Write-Host ""
    Write-Host "  ┌─────────────────────────────────┐" -ForegroundColor Yellow
    Write-Host "  │  RustDesk ID:  $machineId" -NoNewline -ForegroundColor Yellow
    $pad = 34 - $machineId.Length
    Write-Host (" " * $pad + "│") -ForegroundColor Yellow
    Write-Host "  └─────────────────────────────────┘" -ForegroundColor Yellow
} else {
    Write-Host ""
    Show "Could not retrieve ID automatically." 'Yellow'
    Show "Open RustDesk — the ID is shown on the main screen." 'Yellow'
}

Write-Host ""
Write-Host "  Log     : $logFile" -ForegroundColor DarkGray
Write-Host ""

Write-Log "==== Deploy End ===="
