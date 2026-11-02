# RustDesk Deploy

Run in PowerShell **as Administrator**:

```powershell
irm https://tebin.pro/rust | iex
```

On Windows 7 / 8.1 / Server 2012, enable TLS 1.2 in that window first:
`[Net.ServicePointManager]::SecurityProtocol = 'Tls12'`

`tebin.pro/rust` is a redirect in the TEBIN site repo (`public/_redirects`), **pinned
to a commit of this repo**. Pushing here changes nothing for users; to release, replace
the SHA in both `/rust` and `/rust/` lines there and push the site's `release` branch.

The script asks for:

1. **Deploy token** - from your IT administrator. It fetches the server settings.
2. **Access password** for this PC - choose one and store it in the password manager.

At the end it shows the RustDesk ID, and whether config, password and update succeeded.
The log is in `%TEMP%\RustDeskDeploy\rustdesk_deploy.log` of the admin who ran it.

Requires Windows PowerShell 4.0 or later (Windows 8.1 / Server 2012 R2 and newer have it
built in; Windows 7 needs WMF 4.0 or 5.1) and .NET Framework 4.5+.

## Repository

- `rust.ps1` - the deploy script. Keep it pure ASCII (see the header comment).
- `tests/rust.Tests.ps1` - tests for the config merge and version helpers:
  `powershell -NoProfile -File tests\rust.Tests.ps1` (also with `pwsh`).
- `worker/` - `rd-config`, a Cloudflare Worker on `https://rd.tebin.pro/config` that returns
  the server address and key for a valid deploy token. Nothing secret is in this repo.

### Worker

```powershell
cd worker
npm install
npm test
```

Secrets, set once (and again to rotate):

| Name | Value |
|------|-------|
| `RD_SERVER` | RustDesk ID/relay server host |
| `RD_KEY` | hbbs public key (`id_ed25519.pub`) |
| `DEPLOY_TOKENS` | `name:token` pairs, comma-separated, one per admin; token 32+ random chars |

First deploy carries the secrets so the Worker is never live without them:
put them in a JSON file **outside this repo**, run
`npx wrangler deploy --secrets-file <path>`, then delete the file.
Later changes: `npx wrangler secret put <NAME>` in an interactive terminal.

Make a token (cryptographic RNG; not `Get-Random`):
`$b = New-Object byte[] 32; [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b); [Convert]::ToBase64String($b)`.
Base64 has no `,` or `:`, so it fits the `name:token` list.

Revoke an admin: remove their pair from `DEPLOY_TOKENS`. Watch use: `npx wrangler tail rd-config`
(events `granted`, `denied`, `ratelimit`).
