<#
.SYNOPSIS
    Creates the least-privilege identity used to scan SharePoint Online: an Entra ID
    app registration with certificate-only auth and a single read-only Graph permission.

.DESCRIPTION
    This is the "account generation" script. It is meant to be run BY A TENANT ADMIN,
    in their own tenant, so they own the app and can revoke it at any time. It:

      1. Generates a self-signed certificate (no client secrets, ever).
      2. Creates a single-tenant app registration named $AppName.
      3. Adds ONE Microsoft Graph application permission:
           -PermissionMode ReadAll   -> Sites.Read.All  (read-only, all sites, incl. future ones)
           -PermissionMode Selected  -> Sites.Selected  (no access until a site is explicitly granted;
                                        pair with Grant-SitesSelectedSweep.ps1)
      4. Creates the service principal and grants admin consent for that one permission.
      5. Prints the TenantId / ClientId / certificate thumbprint the scan script needs.

    Everything the app can do is visible in this file; nothing else is consented.

.NOTES
    Run as: a user able to create applications and grant admin consent
            (Global Administrator or Privileged Role Administrator + Application Administrator).
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Applications
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [string]$AppName = 'SPO-File-Scanner',

    [ValidateSet('ReadAll', 'Selected')]
    [string]$PermissionMode = 'ReadAll',

    [int]$CertValidYears = 1,

    # Sign in with a device code instead of a browser popup. Needed in embedded
    # terminals (VS Code, CI) where the default Windows (WAM) sign-in cannot
    # attach to a window.
    [switch]$UseDeviceCode,

    # Folder where the public .cer is exported (for the admin's records)
    [string]$OutputFolder = (Join-Path (Split-Path -Parent $PSScriptRoot) 'output')
)

$ErrorActionPreference = 'Stop'
$graphAppId = '00000003-0000-0000-c000-000000000000'  # Microsoft Graph (same in every tenant)
$roleValue  = if ($PermissionMode -eq 'ReadAll') { 'Sites.Read.All' } else { 'Sites.Selected' }

Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications

# --- 1. Certificate (private key stays in the admin's CurrentUser store) -------------
Write-Host "Creating self-signed certificate CN=$AppName ..."
$cert = New-SelfSignedCertificate `
    -Subject "CN=$AppName" `
    -CertStoreLocation 'Cert:\CurrentUser\My' `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears($CertValidYears)

New-Item -ItemType Directory -Force $OutputFolder | Out-Null
$cerPath = Join-Path $OutputFolder "$AppName.cer"
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

# --- 2. Sign in and resolve the Graph app role by name -------------------------------
# Scopes here are for THIS admin session only; they are not granted to the new app.
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All' -UseDeviceCode:$UseDeviceCode -NoWelcome

$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
$appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq $roleValue -and $_.AllowedMemberTypes -contains 'Application' }
if (-not $appRole) { throw "Could not resolve Graph application permission '$roleValue'." }

# --- 3. App registration: single-tenant, cert credential, one permission -------------
Write-Host "Creating app registration '$AppName' with $roleValue ..."
$app = New-MgApplication `
    -DisplayName $AppName `
    -SignInAudience 'AzureADMyOrg' `
    -RequiredResourceAccess @(
        @{
            resourceAppId  = $graphAppId
            resourceAccess = @(@{ id = $appRole.Id; type = 'Role' })
        }
    ) `
    -KeyCredentials @(
        @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = $cert.RawData
            displayName = "$AppName certificate"
        }
    )

# --- 4. Service principal + admin consent for the single permission ------------------
$sp = New-MgServicePrincipal -AppId $app.AppId

New-MgServicePrincipalAppRoleAssignment `
    -ServicePrincipalId $sp.Id `
    -PrincipalId $sp.Id `
    -ResourceId $graphSp.Id `
    -AppRoleId $appRole.Id | Out-Null

Disconnect-MgGraph | Out-Null

# --- 5. Hand back what the scan script needs ------------------------------------------
$result = [pscustomobject]@{
    TenantId              = $TenantId
    ClientId              = $app.AppId
    AppObjectId           = $app.Id
    CertificateThumbprint = $cert.Thumbprint
    CertificateExpires    = $cert.NotAfter
    PublicCertPath        = $cerPath
    GrantedPermission     = $roleValue
}
$result | Format-List

Write-Host @"

Done. To revoke this app at any time: delete the app registration (or its service
principal) in Entra ID - access stops immediately.

Next step:
  .\Get-SpoFileInventory.ps1 -TenantId $TenantId -ClientId $($app.AppId) -CertificateThumbprint $($cert.Thumbprint)
"@
$result
