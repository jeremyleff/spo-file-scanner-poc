<#
.SYNOPSIS
    Enumerates files across SharePoint Online via Microsoft Graph (app-only, read-only)
    and writes a file inventory to a local CSV.

.DESCRIPTION
    This is the "scan" script. It authenticates as the app created by
    New-SpoScannerApp.ps1 (certificate, no user, no secrets) and:

      1. Lists all site collections   (GET /sites/getAllSites)
      2. Lists document libraries     (GET /sites/{id}/drives)
      3. Walks every file per library (GET /drives/{id}/root/delta)
      4. Appends one CSV row per file: site, library, path, name, extension,
         size, created/modified timestamps, and IDs for stable joins.

    Delta enumeration is used deliberately: the final response of each walk
    includes a deltaLink, which this script saves to a state file. Re-running
    with -Incremental replays only changes since the last run - this is the
    "identify new files" mechanism, and it also picks up sites/libraries
    created after the initial scan (a fresh full walk is started for any
    library with no saved state).

    Read-only guarantees: every request in this file is a GET; the app's only
    permission is Sites.Read.All (or Sites.Selected). Graph throttling (429/503)
    is honored via Retry-After.

.NOTES
    Requires: Microsoft.Graph.Authentication
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$TenantId,
    [Parameter(Mandatory)] [string]$ClientId,
    [Parameter(Mandatory)] [string]$CertificateThumbprint,

    [string]$OutputCsv = (Join-Path (Split-Path -Parent $PSScriptRoot) 'output\FileInventory.csv'),

    # Per-library deltaLink state; enables -Incremental
    [string]$StateFile = (Join-Path (Split-Path -Parent $PSScriptRoot) 'output\delta-state.json'),

    # Replay only changes since the last run instead of a full walk
    [switch]$Incremental,

    # Safety valve for dev-tenant testing; 0 = no limit
    [int]$MaxSites = 0
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication

# ---- Graph GET with throttling/backoff ------------------------------------------------
function Invoke-GraphGet {
    param([Parameter(Mandatory)][string]$Uri, [int]$MaxRetries = 6)
    for ($attempt = 0; ; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri
        }
        catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch {}
            if ($attempt -lt $MaxRetries -and $status -in 429, 503, 504) {
                $wait = 10 * ($attempt + 1)
                try {
                    $ra = $_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1
                    if ($ra) { $wait = [Math]::Max([int]$ra, 5) }
                } catch {}
                Write-Warning "Throttled ($status). Waiting $wait s..."
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

# ---- Auth: app-only, certificate ------------------------------------------------------
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
Write-Host "Connected app-only as $ClientId"

New-Item -ItemType Directory -Force (Split-Path -Parent $OutputCsv) | Out-Null
$state = @{}
if ($Incremental -and (Test-Path $StateFile)) {
    ($stateJson = Get-Content $StateFile -Raw | ConvertFrom-Json).PSObject.Properties |
        ForEach-Object { $state[$_.Name] = $_.Value }
    Write-Host "Loaded delta state for $($state.Count) libraries."
}

# ---- 1. All site collections ----------------------------------------------------------
$sites = [System.Collections.Generic.List[object]]::new()
$uri = "https://graph.microsoft.com/v1.0/sites/getAllSites?`$select=id,webUrl,displayName"
while ($uri) {
    $page = Invoke-GraphGet $uri
    foreach ($s in $page.value) { $sites.Add($s) }
    $uri = $page.'@odata.nextLink'
}
Write-Host "Found $($sites.Count) sites."
if ($MaxSites -gt 0) { $sites = $sites | Select-Object -First $MaxSites }

# ---- 2-3. Libraries and files ---------------------------------------------------------
$select = '$select=id,name,size,file,folder,deleted,parentReference,webUrl,createdDateTime,lastModifiedDateTime'
$fileCount = 0
$siteIndex = 0
$skippedSites = 0

foreach ($site in $sites) {
    $siteIndex++
    Write-Host ("[{0}/{1}] {2}" -f $siteIndex, $sites.Count, $site.webUrl)

    $drivesUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,driveType"
    $drives = @()
    try {
        while ($drivesUri) {
            $page = Invoke-GraphGet $drivesUri
            $drives += $page.value
            $drivesUri = $page.'@odata.nextLink'
        }
    }
    catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch {}
        if ($status -eq 423) {   # site collection is locked/archived (e.g. "resourceLocked")
            Write-Warning "Skipping locked site: $($site.webUrl)"
            $skippedSites++
            continue
        }
        throw
    }

    foreach ($drive in $drives) {
        # Incremental: resume from the saved deltaLink; otherwise start a full walk.
        $uri = if ($Incremental -and $state.ContainsKey($drive.id)) { $state[$drive.id] }
               else { "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/delta?$select" }

        $rows = [System.Collections.Generic.List[object]]::new()
        while ($uri) {
            $page = Invoke-GraphGet $uri
            foreach ($item in $page.value) {
                if (-not $item.file -or $item.deleted) { continue }   # files only; skip folders/tombstones
                $folderPath = ($item.parentReference.path -replace '^/drives/[^/]+/root:?', '')
                $rows.Add([pscustomobject]@{
                    SiteId       = $site.id
                    SiteUrl      = $site.webUrl
                    SiteName     = $site.displayName
                    DriveId      = $drive.id
                    LibraryName  = $drive.name
                    ItemId       = $item.id
                    FileName     = $item.name
                    FolderPath   = $folderPath
                    Extension    = [System.IO.Path]::GetExtension($item.name).TrimStart('.').ToLowerInvariant()
                    SizeBytes    = $item.size
                    CreatedUtc   = $item.createdDateTime
                    ModifiedUtc  = $item.lastModifiedDateTime
                    WebUrl       = $item.webUrl
                    ScannedUtc   = (Get-Date).ToUniversalTime().ToString('o')
                })
            }
            if ($page.'@odata.deltaLink') { $state[$drive.id] = $page.'@odata.deltaLink' }
            $uri = $page.'@odata.nextLink'
        }

        if ($rows.Count -gt 0) {
            $rows | Export-Csv -Path $OutputCsv -Append -NoTypeInformation -Encoding UTF8
            $fileCount += $rows.Count
        }
    }

    # Persist state as we go so an interrupted run resumes cheaply
    $state | ConvertTo-Json | Set-Content $StateFile -Encoding UTF8
}

Disconnect-MgGraph | Out-Null
if ($skippedSites -gt 0) { Write-Warning "$skippedSites locked site(s) were skipped." }
Write-Host "Done. $fileCount file rows written to $OutputCsv"
Write-Host "Delta state saved to $StateFile - rerun with -Incremental to pick up only new/changed files."
