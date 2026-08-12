# SharePoint Online File Scanner (POC)

A proof of concept that builds a read-only inventory of every file in your SharePoint
Online tenant, as the discovery step for an AI-driven data classification effort.

You run everything below yourself, in your own tenant, from this repo. The result is
a local CSV listing every file (site, library, path, name, extension, size, timestamps).

## What's in this repo

| Path | Purpose | Who runs it |
|---|---|---|
| `scripts/New-SpoScannerApp.ps1` | One-time setup: creates the scanner's identity — a single-tenant Entra ID app registration with certificate-only auth and one read-only Graph permission, admin-consented in the same script. | Tenant admin, once |
| `scripts/Get-SpoFileInventory.ps1` | The scan: enumerates all sites → document libraries → files via Graph delta queries and writes `output/FileInventory.csv`. Saves per-library delta state so `-Incremental` reruns return only new/changed files. | Anyone with the app's ClientId + certificate (no user sign-in) |
| `scripts/Grant-SitesSelectedSweep.ps1` | Fallback only — see [Alternative: per-site grants](#alternative-per-site-grants-sitesselected). | SharePoint admin, recurring |

## Security model

- The app registration is created **in your tenant, by you**. You own it, you hold the
  certificate, and you can revoke all access at any time by deleting the one app
  registration in Entra ID.
- **Certificate auth only.** No client secrets, no service account, no MFA or
  Conditional Access exceptions, no license consumed.
- **One permission, read-only** (`Sites.Read.All`). Every Graph call the scanner makes
  is a GET, and all activity is visible in your sign-in and audit logs.
- Everything the app can do is visible in the scripts — nothing else is consented.
- The same app-only certificate auth carries unchanged to the production path
  (containerized Python: same ClientId, with a certificate or federated workload
  identity).

## Prerequisites

- A Windows machine with PowerShell 7+ (Windows PowerShell 5.1 also works). The setup
  script uses `New-SelfSignedCertificate` and the `Cert:\CurrentUser\My` store, so it
  is Windows-only.
- Graph PowerShell modules:
  ```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser
  ```
- For step 1 only: an account that can create app registrations and grant admin
  consent (Global Administrator, or Application Administrator + Privileged Role
  Administrator).

## Step 1 — Create the scanner app (one time, as admin)

```powershell
.\scripts\New-SpoScannerApp.ps1 -TenantId <your-tenant-guid-or-domain>
```

`-TenantId` accepts your tenant's GUID or a domain name — e.g. `contoso.onmicrosoft.com`
or a verified custom domain like `contoso.com`. It is not a URL (no `https://`, and not
the `contoso.sharepoint.com` address).

This will:

1. Generate a self-signed certificate in your `Cert:\CurrentUser\My` store (the
   private key never leaves your machine; the public `.cer` is exported to `output/`).
2. Prompt you to **sign in** — the script automatically opens a browser window at the
   Microsoft sign-in page; just complete it as the admin (your normal MFA/Conditional
   Access applies) and return to the terminal. This session is used only to create the
   app and grant consent, and it ends when the script finishes. It is the only user
   sign-in in the whole process; the scan itself runs without one.

   > If no sign-in window appears, or you get an error mentioning *"A window handle
   > must be configured"* (this happens in embedded terminals such as VS Code), rerun
   > with `-UseDeviceCode`. Instead of a popup, the script prints a URL and a
   > one-time code — open the URL on any device, enter the code, and sign in there.
3. Create a single-tenant app registration named `SPO-File-Scanner`.
4. Add the single `Sites.Read.All` application permission and grant admin consent.
5. Print the **TenantId**, **ClientId**, and **certificate thumbprint** — note these,
   the scan script needs all three. The output looks like:

   ```
   TenantId              : contoso.onmicrosoft.com
   ClientId              : 1a2b3c4d-...-9f8e
   CertificateThumbprint : A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0
   ...
   Next step:
     .\Get-SpoFileInventory.ps1 -TenantId ... -ClientId ... -CertificateThumbprint ...
   ```

   The easiest path: **copy the whole `Next step` line** — it is the step 2 command
   with all three values already filled in.

Wait 1–2 minutes after it completes for the consent to propagate.

## Step 2 — Run the scan

No user sign-in — the script authenticates as the app with the certificate:

```powershell
.\scripts\Get-SpoFileInventory.ps1 -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
```

The three values come straight from the end of step 1 — simplest is to paste the
`Next step` command it printed. A few things to know:

- The **thumbprint** identifies the certificate step 1 created; it's a 40-character
  string of letters and digits (no spaces).
- Run this on the **same Windows machine and as the same user** that ran step 1 —
  that's where the certificate's private key lives.
- If you've lost the thumbprint, look it up from the certificate store:

  ```powershell
  Get-ChildItem Cert:\CurrentUser\My | Where-Object Subject -eq 'CN=SPO-File-Scanner'
  ```

  The `Thumbprint` column of the output is the value to use. (The ClientId can
  likewise be recovered any time from Entra ID → App registrations →
  `SPO-File-Scanner` → "Application (client) ID".)

Output:

- `output/FileInventory.csv` — one row per file across every site and document library.
- `output/delta-state.json` — per-library delta links used for incremental reruns.

Useful switches:

- `-MaxSites 5` — limit the scan to the first N sites for a quick smoke test.
- `-OutputCsv <path>` / `-StateFile <path>` — override the default output locations.

The script honors Graph throttling (429/503) with Retry-After backoff, but a first
full crawl of a large tenant will take a while — consider running it off-hours.
Progress prints per site, and delta state is saved as it goes, so an interrupted run
resumes cheaply.

## Step 3 — Incremental reruns (detecting new files)

```powershell
.\scripts\Get-SpoFileInventory.ps1 -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint> -Incremental
```

Only files added or changed since the last run are appended to the CSV. Sites or
libraries created after the initial scan are picked up automatically (any library
with no saved delta state gets a fresh full walk).

To verify: upload a new file anywhere in SharePoint, rerun with `-Incremental`, and
confirm only that file appears.

## Revoking / cleanup

- **Revoke all access instantly:** delete the `SPO-File-Scanner` app registration (or
  its service principal) in Entra ID.
- Delete the certificate from `Cert:\CurrentUser\My` on the machine that ran step 1.
- Delete the `output/` folder if you no longer need the CSV and delta state.

## Alternative: per-site grants (Sites.Selected)

If tenant-wide read (`Sites.Read.All`) is not acceptable, the app can instead be
created with the `Sites.Selected` permission, which grants **no access** until each
site is explicitly granted:

```powershell
.\scripts\New-SpoScannerApp.ps1 -TenantId <tenant> -PermissionMode Selected
.\scripts\Grant-SitesSelectedSweep.ps1 -TenantId <tenant> -ClientId <appId> -AppDisplayName SPO-File-Scanner
```

Be aware of the trade-offs before choosing this model (they are also documented in
the sweep script itself):

- The sweep must run under a **higher-privileged** delegated session
  (`Sites.FullControl.All` as a SharePoint admin) — the least privilege moves from
  the standing app into a recurring admin task.
- Sites created after a sweep are invisible to the scanner until the next sweep, so
  it must be re-run on a schedule.
- One grant call per site: for thousands of sites expect throttling and a long runtime.
- Revocation is also per-site, versus deleting one app registration in the
  `Sites.Read.All` model.

For "discover all files, including future sites," `Sites.Read.All` is the more honest
least privilege; the sweep pattern is provided as a fallback.

## Scope notes

- `getAllSites` includes OneDrive personal sites. If those are out of scope, filter
  on `webUrl` (e.g. exclude `-my.sharepoint.com`) in `Get-SpoFileInventory.ps1`.
- The scanner reads file **metadata only** — it does not download or open file contents.
