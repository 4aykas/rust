# ==============================================================================
# Deploy-RustDesk.ps1  — RustDesk Deploy
# Run as Administrator:
#   powershell -ExecutionPolicy Bypass -File Deploy-RustDesk.ps1
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

# ==== Helpers ====
function Write-Log {
    param([string]$msg, [ValidateSet('INFO','WARN','ERROR')][string]$lvl = 'INFO')
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts [$lvl] $msg" | Out-File -Append -FilePath $logFile
}
function Show {
    param([string]$msg, [string]$color = 'White')
    Write-Host $msg -ForegroundColor $color
}
function Cleanup {
    Remove-Item $archivePath -Force -ErrorAction SilentlyContinue
    Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
}

# ==== Inline C# ZipCrypto extractor (no 7zip needed) ====
$csCode = @'
using System;
using System.IO;
using System.IO.Compression;
using System.Text;

public class ZipExtractor {
    // ZipCrypto keys
    static uint[] keys = new uint[3];

    static void UpdateKeys(byte b) {
        keys[0] = Crc32(keys[0], b);
        keys[1] = (keys[1] + (keys[0] & 0xFF)) * 134775813 + 1;
        keys[2] = Crc32(keys[2], (byte)(keys[1] >> 24));
    }

    static byte DecryptByte() {
        uint temp = (keys[2] | 2) & 0xFFFF;
        return (byte)((temp * (temp ^ 1)) >> 8);
    }

    static uint[] crcTable;
    static uint Crc32(uint crc, byte b) {
        if (crcTable == null) {
            crcTable = new uint[256];
            for (uint i = 0; i < 256; i++) {
                uint c = i;
                for (int j = 0; j < 8; j++)
                    c = (c & 1) != 0 ? (0xEDB88320 ^ (c >> 1)) : (c >> 1);
                crcTable[i] = c;
            }
        }
        return crcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
    }

    static void InitKeys(string password) {
        keys[0] = 305419896;
        keys[1] = 591751049;
        keys[2] = 878082192;
        foreach (char c in password) UpdateKeys((byte)c);
    }

    public static string ExtractTextFile(string zipPath, string password) {
        byte[] zipBytes = File.ReadAllBytes(zipPath);
        int pos = 0;

        // Find local file header signature 0x04034b50
        while (pos < zipBytes.Length - 4) {
            if (zipBytes[pos] == 0x50 && zipBytes[pos+1] == 0x4B &&
                zipBytes[pos+2] == 0x03 && zipBytes[pos+3] == 0x04) break;
            pos++;
        }
        if (pos >= zipBytes.Length - 4) throw new Exception("No local file header found.");

        pos += 4; // skip signature
        // general purpose bit flag at offset +2 from after signature (offset 6 from header start)
        pos += 2; // version needed
        int flags         = zipBytes[pos] | (zipBytes[pos+1] << 8); pos += 2;
        int compression   = zipBytes[pos] | (zipBytes[pos+1] << 8); pos += 2;
        pos += 4; // mod time + date
        pos += 4; // crc32
        int compSize   = (int)(zipBytes[pos] | (zipBytes[pos+1] << 8) |
                               (zipBytes[pos+2] << 16) | (zipBytes[pos+3] << 24)); pos += 4;
        pos += 4; // uncompressed size
        int fnLen      = zipBytes[pos] | (zipBytes[pos+1] << 8); pos += 2;
        int extraLen   = zipBytes[pos] | (zipBytes[pos+1] << 8); pos += 2;
        pos += fnLen + extraLen; // skip filename + extra

        // Decrypt 12-byte encryption header
        InitKeys(password);
        byte[] encHeader = new byte[12];
        for (int i = 0; i < 12; i++) {
            byte c = (byte)(zipBytes[pos+i] ^ DecryptByte());
            UpdateKeys(c);
            encHeader[i] = c;
        }
        pos += 12;
        compSize -= 12;

        // Validate: last byte of decrypted header should match high byte of CRC or mod time
        // (for standard ZipCrypto check byte 11)
        // We just attempt decryption and see if deflate succeeds

        // Decrypt compressed data
        byte[] compData = new byte[compSize];
        for (int i = 0; i < compSize; i++) {
            byte c = (byte)(zipBytes[pos+i] ^ DecryptByte());
            UpdateKeys(c);
            compData[i] = c;
        }

        // Decompress (deflate) or raw store
        byte[] plainBytes;
        if (compression == 8) {
            using (var ms = new MemoryStream(compData))
            using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
            using (var outMs = new MemoryStream()) {
                ds.CopyTo(outMs);
                plainBytes = outMs.ToArray();
            }
        } else {
            plainBytes = compData;
        }

        return Encoding.UTF8.GetString(plainBytes);
    }
}
'@

Add-Type -TypeDefinition $csCode -Language CSharp

# ==== Banner ====
Clear-Host
Write-Host ""
Show "  RustDesk Deploy — TEBIN IT" 'Cyan'
Show "  --------------------------------" 'DarkGray'
Write-Host ""

# ==== Init ====
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory "C:\Temp" -Force | Out-Null }
Write-Log "==== Deploy Start ==== Host: $env:COMPUTERNAME | User: $env:USERNAME"

# ==== Download secrets.zip ====
Show "  Downloading secrets archive..." 'Cyan'
Write-Log "Fetching: $secretsUrl"
try {
    Invoke-WebRequest -Uri $secretsUrl -OutFile $archivePath -UseBasicParsing -ErrorAction Stop
    Write-Log "Archive downloaded."
} catch {
    Write-Log "Download failed: $_" 'ERROR'
    Show "  ERROR: Could not download secrets archive." 'Red'
    exit 1
}

# ==== Prompt Password ====
Write-Host ""
$securePass = Read-Host "  Enter archive password" -AsSecureString
$plainPass  = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass))
Write-Host ""

# ==== Extract & Parse Secrets ====
Write-Log "Decrypting archive..."
$rdServer = $null; $rdKey = $null; $rdPass = $null

try {
    $plainText = [ZipExtractor]::ExtractTextFile($archivePath, $plainPass)

    $plainText -split "`n" | ForEach-Object {
        $l = $_.Trim()
        if ($l -match '^\s*#|^$') { return }
        if ($l -match '^SERVER\s*=\s*(.+)$')   { $rdServer = $Matches[1].Trim() }
        if ($l -match '^KEY\s*=\s*(.+)$')       { $rdKey    = $Matches[1].Trim() }
        if ($l -match '^PASSWORD\s*=\s*(.+)$')  { $rdPass   = $Matches[1].Trim() }
    }
    $plainText = $null

    # Validate
    if (-not $rdServer -or -not $rdKey -or -not $rdPass) {
        throw "One or more required fields (SERVER, KEY, PASSWORD) missing in secrets file."
    }

} catch {
    $plainPass = $null; [GC]::Collect()
    Cleanup
    Write-Log "Extraction failed: $_" 'ERROR'
    Show "  ERROR: Incorrect password or corrupt archive. Nothing deployed." 'Red'
    Write-Host ""
    exit 1
}

$plainPass = $null; [GC]::Collect()
Cleanup
Write-Log "Secrets loaded. Server: $rdServer"
Show "  Password accepted." 'Green'

# ==== Get Latest RustDesk Version ====
Show "  Checking latest RustDesk version..." 'Cyan'
Write-Log "Querying GitHub API..."
try {
    $release     = Invoke-RestMethod -Uri "https://api.github.com/repos/rustdesk/rustdesk/releases/latest" -UseBasicParsing -ErrorAction Stop
    $latestVer   = $release.tag_name -replace '^v', ''
    $asset       = $release.assets | Where-Object { $_.name -match "x86_64\.exe$" } | Select-Object -First 1
    $downloadUrl = $asset.browser_download_url
    Show "  Latest version: $latestVer" 'White'
    Write-Log "Latest: v$latestVer"
} catch {
    $latestVer   = "1.4.3"
    $downloadUrl = "https://github.com/rustdesk/rustdesk/releases/download/1.4.3/rustdesk-1.4.3-x86_64.exe"
    Show "  GitHub API unavailable, using fallback v$latestVer" 'Yellow'
    Write-Log "API failed, fallback v$latestVer" 'WARN'
}

# ==== Check Installed Version ====
$skipInstall = $false
if (Test-Path $rustdeskExe) {
    try {
        $installedVer = (Get-Command $rustdeskExe).FileVersionInfo.ProductVersion
        Write-Log "Installed: v$installedVer"
        if ([version]$installedVer -ge [version]$latestVer) {
            Show "  RustDesk v$installedVer is already up to date." 'Green'
            Write-Log "Up to date. Skipping install."
            $skipInstall = $true
        } else {
            Show "  Updating v$installedVer to v$latestVer..." 'Yellow'
            Write-Log "Updating v$installedVer to v$latestVer"
        }
    } catch {
        Write-Log "Cannot read installed version. Reinstalling." 'WARN'
    }
} else {
    Show "  RustDesk not found. Installing v$latestVer..." 'Cyan'
    Write-Log "Fresh install."
}

# ==== Download & Install ====
if (-not $skipInstall) {
    Show "  Downloading RustDesk v$latestVer..." 'Cyan'
    Write-Log "Downloading installer..."
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installTemp -UseBasicParsing -ErrorAction Stop
        Write-Log "Installer downloaded."
    } catch {
        Write-Log "Installer download failed: $_" 'ERROR'
        Show "  ERROR: Could not download RustDesk installer." 'Red'
        exit 1
    }

    Show "  Installing silently, please wait..." 'Cyan'
    Write-Log "Running silent install..."
    Start-Process -FilePath $installTemp -ArgumentList "--silent-install" -PassThru | Wait-Process
    Write-Log "Install complete."
    Start-Sleep -Seconds 5

    Write-Log "Stopping RustDesk after install..."
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

# Stop service before writing config regardless of install path
Write-Log "Stopping RustDesk before config..."
net stop rustdesk 2>&1 | Out-Null
Get-Process rustdesk -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
function Write-Config {
    foreach ($cfgPath in @($userConfig, $svcConfig)) {
        $dir = Split-Path $cfgPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
        [System.IO.File]::WriteAllText($cfgPath, $toml, $utf8NoBom)
        Write-Log "Config written: $cfgPath"
    }
}
Write-Config

# ==== Start Service (first run may overwrite config — restart to re-apply) ====
Write-Log "Starting RustDesk service (first run)..."
net start rustdesk 2>&1 | Out-Null
Start-Sleep -Seconds 6

Write-Log "Re-applying config after first-run init..."
net stop rustdesk 2>&1 | Out-Null
Get-Process rustdesk -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Config

Write-Log "Starting RustDesk service..."
net start rustdesk 2>&1 | Out-Null
Start-Sleep -Seconds 5

# ==== Apply Password ====
Write-Log "Setting access password..."
Start-Process -FilePath $rustdeskExe -ArgumentList "--password", $rdPass -Wait
$rdPass = $null; $rdKey = $null; [GC]::Collect()

# ==== Ensure service is running ====
Write-Log "Starting RustDesk service (final)..."
net start rustdesk 2>&1 | Out-Null
Start-Sleep -Seconds 6

# ==== Get Machine ID ====
Write-Log "Retrieving RustDesk machine ID..."
$machineId = $null

$idTomlPaths = @(
    "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk.toml",
    "C:\Windows\System32\config\systemprofile\AppData\Roaming\RustDesk\config\RustDesk.toml",
    "$env:APPDATA\RustDesk\config\RustDesk.toml"
)

# Retry up to 5 times — RustDesk may not have written the ID file immediately
for ($attempt = 1; $attempt -le 5 -and -not $machineId; $attempt++) {
    foreach ($idConfig in $idTomlPaths) {
        if (-not (Test-Path $idConfig)) { continue }
        $idLine = Get-Content $idConfig -ErrorAction SilentlyContinue |
                  Where-Object { $_ -match '^id\s*=' } | Select-Object -First 1
        if ($idLine -match "=\s*'?([A-Za-z0-9]{6,})'?") {
            $machineId = $Matches[1].Trim()
            Write-Log "ID found in $idConfig (attempt $attempt)"
            break
        }
    }
    if (-not $machineId -and $attempt -lt 5) { Start-Sleep -Seconds 3 }
}

if (-not $machineId) {
    $idTmp = "$env:TEMP\rd_getid.txt"
    try {
        Start-Process -FilePath $rustdeskExe -ArgumentList '--get-id' `
            -RedirectStandardOutput $idTmp -WindowStyle Hidden -Wait -ErrorAction Stop
        $idOut = (Get-Content $idTmp -ErrorAction SilentlyContinue) -join '' | ForEach-Object { $_.Trim() }
        if ($idOut -match '([A-Za-z0-9]{6,})') { $machineId = $Matches[1] }
    } catch {}
    Remove-Item $idTmp -Force -ErrorAction SilentlyContinue
}

Write-Log "Machine ID: $machineId"

# ==== Final Output ====
Write-Host ""
Show "  --------------------------------" 'DarkGray'
Show "  Deployment complete" 'Green'
Show "  --------------------------------" 'DarkGray'
Write-Host ""
Show "  Server  : $rdServer" 'Gray'
Show "  Version : $latestVer" 'Gray'
Write-Host ""

if ($machineId) {
    Show "  RustDesk ID  >>  $machineId" 'Yellow'
} else {
    Show "  Could not retrieve ID automatically." 'Yellow'
    Show "  Open RustDesk to see the ID on the main screen." 'Gray'
}

Write-Host ""
Show "  Log: $logFile" 'DarkGray'
Write-Host ""

Write-Log "==== Deploy End ===="
