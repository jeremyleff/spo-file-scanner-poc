<#
.SYNOPSIS
    FALLBACK ONLY: grants a Sites.Selected app read access to every site in the tenant,
    one grant at a time.

.DESCRIPTION
    Use this only if the tenant admin declines Sites.Read.All and insists on per-site
    grants. Be aware of what this trades away:

      * This script itself must run with FAR HIGHER privilege than the scanner app:
        a delegated Sites.FullControl.All session under an admin account. The
        "least privilege" moves from the standing app into a recurring admin task.
      * New sites created after a sweep are invisible to the scanner until the next
        sweep - this MUST be re-run on a schedule or wired to a site-creation event.
      * ~9000 sites means ~9000 grant calls; expect throttling and a long runtime.
      * Revocation is also a sweep (delete the permission per site), vs. deleting
        one app registration in the Sites.Read.All model.

    Enumeration note: delegated contexts cannot call /sites/getAllSites (app-only
    endpoint), so this uses /sites?search=* which can lag for very new sites -
    another reason this pattern drifts.

.NOTES
    Run as: SharePoint Administrator (or Global Admin).
    Requires: Microsoft.Graph.Authentication
    Idempotent: skips sites where the app already holds a permission.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$TenantId,
    [Parameter(Mandatory)] [string]$ClientId,      # the Sites.Selected scanner app
    [Parameter(Mandatory)] [string]$AppDisplayName,
    [ValidateSet('read', 'write')] [string]$Role = 'read',
    [int]$MaxSites = 0                              # safety valve for testing; 0 = all
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication

Connect-MgGraph -TenantId $TenantId -Scopes 'Sites.FullControl.All' -NoWelcome

# ---- Enumerate sites (delegated: search-based) -----------------------------------------
$sites = [System.Collections.Generic.List[object]]::new()
$uri = "https://graph.microsoft.com/v1.0/sites?search=*&`$select=id,webUrl"
while ($uri) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $uri
    foreach ($s in $page.value) { $sites.Add($s) }
    $uri = $page.'@odata.nextLink'
}
Write-Host "Found $($sites.Count) sites."
if ($MaxSites -gt 0) { $sites = $sites | Select-Object -First $MaxSites }

# ---- Grant per site ---------------------------------------------------------------------
$granted = 0; $skipped = 0; $failed = 0
$body = @{
    roles               = @($Role)
    grantedToIdentities = @(@{ application = @{ id = $ClientId; displayName = $AppDisplayName } })
}

foreach ($site in $sites) {
    try {
        $existing = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions"
        $already = $existing.value | Where-Object {
            $_.grantedToIdentities.application.id -contains $ClientId
        }
        if ($already) { $skipped++; continue }

        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" `
            -Body ($body | ConvertTo-Json -Depth 5) -ContentType 'application/json' | Out-Null
        $granted++
        Write-Host "Granted $Role on $($site.webUrl)"
    }
    catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch {}
        if ($status -eq 429) {
            Start-Sleep -Seconds 30
            $sites.Add($site)   # retry at the end of the queue
            continue
        }
        $failed++
        Write-Warning "FAILED $($site.webUrl): $($_.Exception.Message)"
    }
}

Disconnect-MgGraph | Out-Null
Write-Host "Sweep complete. Granted: $granted  Already granted: $skipped  Failed: $failed"
Write-Host "Remember: sites created after this run have NO grant until the next sweep."
