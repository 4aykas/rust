# ==============================================================================
# rust.ps1 - RustDesk Deploy (TEBIN IT)
# Run in PowerShell as Administrator:
#   irm https://tebin.pro/rust | iex
# or from a local copy:
#   powershell -ExecutionPolicy Bypass -File rust.ps1
# Asks for a deploy token (fetches server + key from the TEBIN config service)
# and for the unattended access password to set on this machine.
# Keep this file pure ASCII: Windows PowerShell 5.1 reads BOM-less files in the
# ANSI code page, where some UTF-8 bytes turn into quote characters.
# ==============================================================================

# No #Requires -RunAsAdministrator: Invoke-Expression ignores it, and it would stop
# the tests from loading the helpers. Invoke-Deploy checks elevation itself.
# `irm https://tebin.pro/rust | iex` runs this in the admin's own session, so
# nothing here may change session-wide preferences (see Invoke-Deploy).

# Windows PowerShell 5.1 on older systems defaults to TLS 1.0/1.1, which GitHub refuses.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ==== Settings ====
$configUrl   = "https://rd.tebin.pro/config"
$rustdeskExe = "C:\Program Files\RustDesk\rustdesk.exe"
$svcConfig   = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\RustDesk\config\RustDesk2.toml"
# The admin's own temp folder: other users cannot swap the installer in it
# (C:\Temp may be owned by whoever created it first).
$workDir     = Join-Path $env:TEMP "RustDeskDeploy"
$logFile     = Join-Path $workDir "rustdesk_deploy.log"
$installTemp = Join-Path $workDir "rustdesk_setup.exe"

# ==== Pure helpers (covered by tests\rust.Tests.ps1) ====

# Returns @{ top = @{key=value}; options = @{key=value} } from RustDesk2.toml text.
function Read-RustDeskToml {
    param([string]$Text)
    $result = @{ top = @{}; options = @{} }
    $section = ''
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '^\s*\[([^\]]+)\]\s*$') { $section = $Matches[1].Trim(); continue }
        if ($line -match "^\s*([A-Za-z0-9_\-]+)\s*=\s*(?:'([^']*)'|`"([^`"]*)`"|(\S+))\s*$") {
            $value = if ($Matches[2] -ne $null) { $Matches[2] } elseif ($Matches[3] -ne $null) { $Matches[3] } else { $Matches[4] }
            if ($section -eq '') { $result.top[$Matches[1]] = $value }
            elseif ($section -eq 'options') { $result.options[$Matches[1]] = $value }
        }
    }
    return $result
}

# Sets the given keys in RustDesk2.toml text and keeps every other line as it is
# (2FA, trusted devices, unlock PIN and other user settings live in the same file).
function Set-RustDeskToml {
    param([string]$Text, [hashtable]$Top, [hashtable]$Options)
    foreach ($v in @($Top.Values) + @($Options.Values)) {
        if ("$v" -match "['`r`n]") { throw "Config value contains a quote or line break: $v" }
    }
    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $section = ''
    $hasOptions = $false
    $emitMissing = {
        param($table)
        foreach ($k in $table.Keys) { if (-not $seen.ContainsKey("$section/$k")) { $out.Add("$k = '$($table[$k])'") } }
    }
    $lines = if ($Text) { $Text -split "`r?`n" } else { @() }
    # Drop the trailing empty element that a final newline produces.
    if ($lines.Count -gt 0 -and $lines[-1] -eq '') { $lines = $lines[0..($lines.Count - 2)] }
    foreach ($line in $lines) {
        if ($line -match '^\s*\[([^\]]+)\]\s*$') {
            if ($section -eq '') { & $emitMissing $Top }
            elseif ($section -eq 'options') { & $emitMissing $Options }
            $section = $Matches[1].Trim()
            if ($section -eq 'options') { $hasOptions = $true }
            $out.Add($line)
            continue
        }
        $table = if ($section -eq '') { $Top } elseif ($section -eq 'options') { $Options } else { $null }
        if ($table -and $line -match '^\s*([A-Za-z0-9_\-]+)\s*=' -and $table.ContainsKey($Matches[1])) {
            $k = $Matches[1]
            $out.Add("$k = '$($table[$k])'")
            $seen["$section/$k"] = $true
            continue
        }
        $out.Add($line)
    }
    if ($section -eq '') { & $emitMissing $Top }
    elseif ($section -eq 'options') { & $emitMissing $Options }
    if (-not $hasOptions) {
        $out.Add('')
        $out.Add('[options]')
        $section = 'options'
        & $emitMissing $Options
    }
    return (($out -join "`n") + "`n")
}

# True when every wanted key has the wanted value in the parsed file.
function Test-RustDeskToml {
    param([string]$Text, [hashtable]$Top, [hashtable]$Options)
    $cfg = Read-RustDeskToml $Text
    foreach ($k in $Top.Keys) { if ($cfg.top[$k] -ne $Top[$k]) { return $false } }
    foreach ($k in $Options.Keys) { if ($cfg.options[$k] -ne $Options[$k]) { return $false } }
    return $true
}

# "1.5.0+68" -> [version]1.5.0; $null when it cannot be read.
function ConvertTo-RustDeskVersion {
    param([string]$Text)
    $clean = ($Text -replace '^v', '' -replace '[+\-].*$', '').Trim()
    try { return [version]$clean } catch { return $null }
}

# ==== Logging / output ====
function Write-Log {
    param([string]$msg, [string]$lvl = 'INFO')
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts [$lvl] $msg" | Out-File -Append -FilePath $logFile -Encoding utf8
}
function Show {
    param([string]$msg, [string]$color = 'White')
    Write-Host $msg -ForegroundColor $color
}
function ConvertFrom-Secure {
    param([Security.SecureString]$Secure)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# ==== Service helpers ====
# RustDesk 1.5.x starts slower on a fresh machine; fixed sleeps let --password run
# before the service IPC is up, and the service could overwrite config written too early.
function Stop-RD {
    Stop-Service rustdesk -Force -ErrorAction SilentlyContinue
    Get-Process rustdesk -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt 15 -and (Get-Process rustdesk -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Seconds 1 }
}

function Start-RD {
    # Returns $true once the service runs and its --server process is up.
    Start-Service rustdesk -ErrorAction SilentlyContinue
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        $svc = Get-Service rustdesk -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne 'Running') { Start-Service rustdesk -ErrorAction SilentlyContinue }
        $server = Get-CimInstance Win32_Process -Filter "Name='rustdesk.exe'" -ErrorAction SilentlyContinue |
                  Where-Object { $_.CommandLine -match '--server' }
        if ($svc.Status -eq 'Running' -and $server) {
            Start-Sleep -Seconds 3   # let the IPC listener settle
            return $true
        }
        Start-Sleep -Seconds 2
    }
    return $false
}

# Runs rustdesk.exe with arguments and returns its stdout.
function Invoke-RD {
    param([string[]]$Arguments)
    $tmp = Join-Path $workDir ("rd_out_{0}.txt" -f [guid]::NewGuid())
    try {
        Start-Process -FilePath $rustdeskExe -ArgumentList $Arguments -RedirectStandardOutput $tmp `
            -WindowStyle Hidden -Wait -ErrorAction Stop
        return ((Get-Content $tmp -ErrorAction SilentlyContinue) -join "`n").Trim()
    } catch {
        return ''
    } finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ==== Main ====
function Invoke-Deploy {
    # Function scope: applies to the helpers it calls, not to the caller's session.
    $ErrorActionPreference = 'SilentlyContinue'

    # #Requires is ignored when the script runs through Invoke-Expression (the README one-liner).
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Show "  ERROR: Run PowerShell as Administrator." 'Red'
        return
    }
    if ($PSVersionTable.PSVersion.Major -lt 4) {
        Show "  ERROR: PowerShell 4.0 or later is required (found $($PSVersionTable.PSVersion))." 'Red'
        return
    }

    New-Item -ItemType Directory $workDir -Force | Out-Null

    Clear-Host
    Write-Host ""
    Show "  RustDesk Deploy - TEBIN IT" 'Cyan'
    Show "  --------------------------------" 'DarkGray'
    Write-Host ""
    Write-Log "==== Deploy Start ==== Host: $env:COMPUTERNAME | User: $env:USERNAME"

    # ---- Server and key from the config service ----
    $token = ConvertFrom-Secure (Read-Host "  Deploy token" -AsSecureString)
    try {
        $remote = Invoke-RestMethod -Uri $configUrl -Headers @{ Authorization = "Bearer $token" } `
            -UseBasicParsing -ErrorAction Stop
    } catch {
        $status = $null
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        Write-Log "Config fetch failed (HTTP $status): $($_.Exception.Message)" 'ERROR'
        $msg = switch ($status) {
            401     { "Deploy token rejected." }
            429     { "Too many attempts. Wait a minute and try again." }
            default { "Could not reach the config service." }
        }
        Show "  ERROR: $msg Nothing deployed." 'Red'
        return
    } finally {
        $token = $null
    }
    $rdServer = "$($remote.server)".Trim()
    $rdKey    = "$($remote.key)".Trim()
    if (-not $rdServer -or -not $rdKey) {
        Write-Log "Config service returned no server or key." 'ERROR'
        Show "  ERROR: Config service returned incomplete settings. Nothing deployed." 'Red'
        return
    }
    Write-Log "Config received. Server: $rdServer"
    Show "  Token accepted. Server: $rdServer" 'Green'

    # ---- Access password, typed by the admin ----
    Write-Host ""
    while ($true) {
        $p1 = ConvertFrom-Secure (Read-Host "  Access password for this PC (min 8 chars)" -AsSecureString)
        $p2 = ConvertFrom-Secure (Read-Host "  Repeat access password" -AsSecureString)
        if ($p1.Length -lt 8) { Show "  Too short." 'Yellow'; continue }
        if ($p1 -match '"') { Show "  Double quotes are not supported." 'Yellow'; continue }
        if ($p1 -cne $p2) { Show "  Passwords do not match." 'Yellow'; continue }
        break
    }
    $rdPass = $p1; $p1 = $null; $p2 = $null
    Write-Host ""

    # ---- Latest RustDesk version ----
    Show "  Checking latest RustDesk version..." 'Cyan'
    try {
        $release     = Invoke-RestMethod -Uri "https://api.github.com/repos/rustdesk/rustdesk/releases/latest" -UseBasicParsing -ErrorAction Stop
        $latestVer   = $release.tag_name -replace '^v', ''
        $downloadUrl = ($release.assets | Where-Object { $_.name -match "x86_64\.exe$" } | Select-Object -First 1).browser_download_url
        Write-Log "Latest: v$latestVer"
    } catch {
        $latestVer   = "1.5.1"
        $downloadUrl = "https://github.com/rustdesk/rustdesk/releases/download/1.5.1/rustdesk-1.5.1-x86_64.exe"
        Show "  GitHub API unavailable, using fallback v$latestVer" 'Yellow'
        Write-Log "API failed, fallback v$latestVer" 'WARN'
    }
    $wantVer = ConvertTo-RustDeskVersion $latestVer

    # ---- Install or update ----
    $installedVer = $null
    if (Test-Path $rustdeskExe) {
        $installedVer = ConvertTo-RustDeskVersion (Get-Item $rustdeskExe).VersionInfo.ProductVersion
        Write-Log "Installed: v$installedVer"
    }
    $installOk = $true
    if ($installedVer -and $wantVer -and $installedVer -ge $wantVer) {
        Show "  RustDesk v$installedVer is already up to date." 'Green'
    } else {
        Show "  Installing RustDesk v$latestVer..." 'Cyan'
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $installTemp -UseBasicParsing -ErrorAction Stop
        } catch {
            Write-Log "Installer download failed: $_" 'ERROR'
            Show "  ERROR: Could not download RustDesk installer." 'Red'
            return
        }
        $p = Start-Process -FilePath $installTemp -ArgumentList "--silent-install" -PassThru
        $p | Wait-Process -Timeout 300 -ErrorAction SilentlyContinue
        Remove-Item $installTemp -Force -ErrorAction SilentlyContinue

        # The 1.5.x installer hands the copy to an elevated cmd and exits before
        # rustdesk.exe is replaced, so wait for the new version (and the service),
        # not for the installer process.
        $nowVer = $null
        $deadline = (Get-Date).AddSeconds(120)
        while ($true) {
            $nowVer = $null
            if (Test-Path $rustdeskExe) { $nowVer = ConvertTo-RustDeskVersion (Get-Item $rustdeskExe).VersionInfo.ProductVersion }
            $updated = $nowVer -and (-not $wantVer -or $nowVer -ge $wantVer)
            if (($updated -and (Get-Service rustdesk -ErrorAction SilentlyContinue)) -or (Get-Date) -gt $deadline) { break }
            Start-Sleep -Seconds 2
        }
        if (-not $nowVer) {
            Write-Log "rustdesk.exe not found after install." 'ERROR'
            Show "  ERROR: RustDesk installation failed." 'Red'
            return
        }
        if ($wantVer -and $nowVer -lt $wantVer) {
            # An update that failed leaves the old exe in place; do not call that success.
            $installOk = $false
            Write-Log "Install did not update: have v$nowVer, wanted v$latestVer." 'ERROR'
            Show "  WARNING: Update failed, still v$nowVer (wanted v$latestVer). Continuing with config." 'Red'
        } else {
            Write-Log "Install complete: v$nowVer"
        }
    }

    # ---- Config targets ----
    # The service reads the LocalService profile. The tray/UI reads the profile of the
    # user who is logged on, which is not always the admin account running this script.
    $targets = @($svcConfig, "$env:APPDATA\RustDesk\config\RustDesk2.toml")
    try {
        $consoleUser = (Get-CimInstance Win32_ComputerSystem).UserName
        if ($consoleUser) {
            $sid = ([Security.Principal.NTAccount]$consoleUser).Translate([Security.Principal.SecurityIdentifier]).Value
            $profilePath = (Get-CimInstance Win32_UserProfile -Filter "SID='$sid'").LocalPath
            if ($profilePath) { $targets += "$profilePath\AppData\Roaming\RustDesk\config\RustDesk2.toml" }
        }
    } catch {
        Write-Log "Cannot resolve logged-on user profile: $_" 'WARN'
    }
    $targets = @($targets | Select-Object -Unique)

    $top     = @{ 'rendezvous_server' = "${rdServer}:21116" }
    $options = @{
        'custom-rendezvous-server' = $rdServer
        'key'                      = $rdKey
        'direct-server'            = 'Y'
        'direct-access-port'       = '21118'
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    $writeConfig = {
        foreach ($path in $targets) {
            $dir = Split-Path $path
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
            $old = if (Test-Path $path) { [IO.File]::ReadAllText($path) } else { '' }
            [IO.File]::WriteAllText($path, (Set-RustDeskToml $old $top $options), $utf8NoBom)
            Write-Log "Config written: $path"
        }
    }
    $testConfig = {
        $text = if (Test-Path $svcConfig) { [IO.File]::ReadAllText($svcConfig) } else { '' }
        Test-RustDeskToml $text $top $options
    }

    # First start lets RustDesk create its own files; then write config with everything stopped.
    if (-not (Start-RD)) { Write-Log "Service not ready after initial start." 'WARN' }
    $configOk = $false
    for ($round = 1; $round -le 3 -and -not $configOk; $round++) {
        Stop-RD
        & $writeConfig
        if (-not (Start-RD)) { Write-Log "Service not ready after config round $round." 'WARN' }
        Start-Sleep -Seconds 5   # give the service time to overwrite it, if it is going to
        $configOk = & $testConfig
        Write-Log "Config check (round $round): $configOk"
    }
    if ($configOk) { Show "  Server settings applied." 'Green' }
    else { Show "  WARNING: Server settings did not persist in the service config." 'Red' }

    # ---- Access password ----
    $passwordOk = $false
    for ($attempt = 1; $attempt -le 5 -and -not $passwordOk; $attempt++) {
        $out = Invoke-RD @('--password', "`"$rdPass`"")
        Write-Log "--password attempt ${attempt}: '$out'"
        if ($out -match 'Done') {
            $passwordOk = $true
        } else {
            Start-Sleep -Seconds 5
            if ($attempt -eq 3) { Stop-RD; Start-RD | Out-Null }   # IPC stuck: restart once
        }
    }
    $rdPass = $null
    if ($passwordOk) { Show "  Access password set." 'Green' }
    else { Show "  WARNING: Could not confirm the access password was set." 'Red' }

    # Password change must not reset server settings.
    if (-not (& $testConfig)) {
        Write-Log "Config lost after password; re-applying." 'WARN'
        Stop-RD; & $writeConfig; Start-RD | Out-Null
        $configOk = & $testConfig
    }

    # ---- Machine ID ----
    $machineId = $null
    for ($attempt = 1; $attempt -le 5 -and -not $machineId; $attempt++) {
        $out = Invoke-RD @('--get-id')
        $line = ($out -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Last 1)
        if ($line -match '^[A-Za-z0-9_\-]{6,}$') { $machineId = $line } else { Start-Sleep -Seconds 3 }
    }
    Write-Log "Machine ID: $machineId"

    # ---- Summary ----
    $allOk = $configOk -and $passwordOk -and $installOk
    Write-Host ""
    Show "  --------------------------------" 'DarkGray'
    if ($allOk) { Show "  Deployment complete" 'Green' } else { Show "  Deployment finished with warnings" 'Yellow' }
    Show "  --------------------------------" 'DarkGray'
    Write-Host ""
    Show "  Server  : $rdServer" 'Gray'
    Show "  Version : $((Get-Item $rustdeskExe).VersionInfo.ProductVersion)$(if (-not $installOk) {' (UPDATE FAILED)'})" 'Gray'
    Show "  Config  : $(if ($configOk) {'OK'} else {'NOT APPLIED'})" 'Gray'
    Show "  Password: $(if ($passwordOk) {'OK'} else {'NOT CONFIRMED'})" 'Gray'
    Write-Host ""
    if ($machineId) {
        Show "  RustDesk ID  >>  $machineId" 'Yellow'
    } else {
        Show "  Could not retrieve ID automatically. Open RustDesk to see it." 'Yellow'
    }
    Write-Host ""
    Show "  Log: $logFile" 'DarkGray'
    Write-Host ""
    Write-Log "==== Deploy End ==== install=$installOk config=$configOk password=$passwordOk"
}

# Tests load the helpers above without running the deploy.
if (-not $env:RD_DEPLOY_NO_RUN) { Invoke-Deploy }
