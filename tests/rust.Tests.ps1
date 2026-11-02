# Plain PowerShell tests for the pure helpers in rust.ps1 (no Pester: Windows ships
# Pester 3.4, whose syntax Pester 5 rejects). Run:
#   powershell -NoProfile -File tests\rust.Tests.ps1
#   pwsh -NoProfile -File tests\rust.Tests.ps1
# Exit code 0 = all passed.

$env:RD_DEPLOY_NO_RUN = '1'
. (Join-Path $PSScriptRoot '..\rust.ps1')
Remove-Item Env:\RD_DEPLOY_NO_RUN
$ErrorActionPreference = 'Continue'   # rust.ps1 sets SilentlyContinue

$script:failed = 0
function Check([string]$name, [bool]$ok) {
    if ($ok) { Write-Host "PASS  $name" -ForegroundColor Green }
    else { Write-Host "FAIL  $name" -ForegroundColor Red; $script:failed++ }
}

$top = @{ 'rendezvous_server' = 'rd.example.org:21116' }
$opt = @{
    'custom-rendezvous-server' = 'rd.example.org'
    'key'                      = 'AbC+/9xyz='
    'direct-server'            = 'Y'
    'direct-access-port'       = '21118'
}

# A file RustDesk 1.5.x wrote, with settings the deploy must not touch.
$existing = @"
rendezvous_server = 'old.example.org:21116'
nat_type = 1
serial = 0
unlock_pin = 'pin-enc'
trusted_devices = 'dev-enc'

[options]
2fa = 'totp-secret-enc'
approve-mode = 'password'
custom-rendezvous-server = 'old.example.org'
key = 'OLDKEY'
whitelist = '10.0.0.0/8'
"@

$merged = Set-RustDeskToml $existing $top $opt
$parsed = Read-RustDeskToml $merged

Check 'merge sets server and key'          ((Test-RustDeskToml $merged $top $opt))
Check 'merge keeps 2fa'                    ($parsed.options['2fa'] -eq 'totp-secret-enc')
Check 'merge keeps approve-mode'           ($parsed.options['approve-mode'] -eq 'password')
Check 'merge keeps whitelist'              ($parsed.options['whitelist'] -eq '10.0.0.0/8')
Check 'merge keeps trusted devices + pin'  ($parsed.top['trusted_devices'] -eq 'dev-enc' -and $parsed.top['unlock_pin'] -eq 'pin-enc')
Check 'merge does not duplicate keys'      (([regex]::Matches($merged, "(?m)^key = ")).Count -eq 1)
Check 'merge keeps server line in top'     ($parsed.top['rendezvous_server'] -eq 'rd.example.org:21116' -and -not $parsed.options.ContainsKey('rendezvous_server'))
Check 'merge is idempotent'                ((Set-RustDeskToml $merged $top $opt) -eq $merged)

$fresh = Set-RustDeskToml '' $top $opt
Check 'empty file gets all keys'           ((Test-RustDeskToml $fresh $top $opt))
Check 'empty file gets [options] section'  ($fresh -match '(?m)^\[options\]$')

$noOptions = Set-RustDeskToml "rendezvous_server = 'x:21116'`r`nnat_type = 1`r`n" $top $opt
Check 'file without [options] (CRLF)'      ((Test-RustDeskToml $noOptions $top $opt) -and (Read-RustDeskToml $noOptions).top['nat_type'] -eq '1')

$otherSection = Set-RustDeskToml "[options]`nkey = 'OLD'`n[other]`nkey = 'keep'`n" $top $opt
Check 'key in another section untouched'   ($otherSection -match "(?m)^\[other\]\nkey = 'keep'")

Check 'test fails on wrong key'            (-not (Test-RustDeskToml $existing $top $opt))
Check 'test reads double-quoted values'    ((Read-RustDeskToml "[options]`nkey = `"Q`"").options['key'] -eq 'Q')

$threw = $false
try { Set-RustDeskToml '' $top @{ key = "bad'quote" } | Out-Null } catch { $threw = $true }
Check 'quote in value is rejected'         $threw

Check 'version 1.5.0+68'                   ((ConvertTo-RustDeskVersion '1.5.0+68') -eq [version]'1.5.0')
Check 'version v1.5.1'                     ((ConvertTo-RustDeskVersion 'v1.5.1') -eq [version]'1.5.1')
Check 'version garbage is null'            ($null -eq (ConvertTo-RustDeskVersion 'nightly'))
Check 'old build is lower'                 ((ConvertTo-RustDeskVersion '1.5.0+68') -lt (ConvertTo-RustDeskVersion '1.5.1'))

if ($script:failed) { Write-Host "$script:failed failed" -ForegroundColor Red; exit 1 }
Write-Host "All passed" -ForegroundColor Green
