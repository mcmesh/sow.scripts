#Requires -Version 7.0
<#
.SYNOPSIS
    Search for SharePoint sites whose display name contains a given filter string,
    and optionally check for (or create) a named folder in the Documents library.

.DESCRIPTION
    Uses the Microsoft Graph Search API (POST /v1.0/search/query) with entityTypes
    "site" to find SharePoint sites. Results are additionally filtered client-side
    so only sites whose display name contains -NameFilter are returned.

    When -SiteUrl is supplied, the script also:
      Phase 1 — Checks whether -FolderName exists at the root of the site's
                Documents document library.
      Phase 2 — Creates the folder if it doesn't exist, but only when
                -CreateIfMissing is also passed.

    Credentials must be supplied via environment variables — do NOT hardcode secrets.
    Set $env:GRAPH_CLIENT_SECRET before running.

.PARAMETER NameFilter
    Required. Case-insensitive substring to match against site display names.

.PARAMETER MaxResults
    Optional. Maximum number of matching sites to return. Default: 200.

.PARAMETER SiteUrl
    Optional. Full URL of a SharePoint site to check for -FolderName.
    When omitted the script runs search-only and exits normally.

.PARAMETER FolderName
    Optional. Exact display name of the folder to look for (or create).
    Default: "Call Transcripts".

.PARAMETER CreateIfMissing
    Optional switch. When present and the folder is not found, the script
    creates the folder. Requires Files.ReadWrite.All on the app registration.

.EXAMPLE
    pwsh ./Search-SharePointSites.ps1 -NameFilter "SOW"

.EXAMPLE
    $sites = ./Search-SharePointSites.ps1 -NameFilter "Project" -MaxResults 50
    $sites | Select-Object DisplayName, WebUrl

.EXAMPLE
    # Check whether "Call Transcripts" exists on a specific site
    pwsh ./Search-SharePointSites.ps1 -NameFilter "AB InBev" `
        -SiteUrl "https://meshsystemsus.sharepoint.com/sites/BUD-ABInbev-PerfectDraft-COM"

.EXAMPLE
    # Check and create the folder if it's missing
    pwsh ./Search-SharePointSites.ps1 -NameFilter "AB InBev" `
        -SiteUrl "https://meshsystemsus.sharepoint.com/sites/BUD-ABInbev-PerfectDraft-COM" `
        -CreateIfMissing
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string] $NameFilter,

    [Parameter(Mandatory = $false)]
    [int] $MaxResults = 200,

    [Parameter(Mandatory = $false)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $false)]
    [string] $FolderName = "Call Transcripts",

    [Parameter(Mandatory = $false)]
    [switch] $CreateIfMissing
)

# ── Config ────────────────────────────────────────────────────────────────────
$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf"
$ClientId     = "5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b"
$ClientSecret = $env:GRAPH_CLIENT_SECRET
$Region       = "NAM"
$PageSize     = 25
# ─────────────────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "ERROR: Set the environment variable GRAPH_CLIENT_SECRET before running." -ForegroundColor Red
    exit 1
}

# ── 1. Get Token ──────────────────────────────────────────────────────────────
Write-Host "Authenticating..." -ForegroundColor Cyan

try {
    $tokenResponse = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Method Post `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
    $token = $tokenResponse.access_token
    Write-Host "Token acquired." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to get token." -ForegroundColor Red
    Write-Host $_.ErrorDetails.Message
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

# ── 2. Search — page through results up to $MaxResults ────────────────────────
Write-Host "`nSearching for sites matching '$NameFilter'..." -ForegroundColor Cyan

$allSites = [System.Collections.Generic.List[PSCustomObject]]::new()
$from     = 0
$total    = $null

do {
    $searchBody = @{
        requests = @(
            @{
                entityTypes = @("site")
                query       = @{ queryString = $NameFilter }
                fields      = @("displayName", "webUrl", "createdDateTime", "description")
                from        = $from
                size        = $PageSize
                region      = $Region
            }
        )
    } | ConvertTo-Json -Depth 5

    try {
        $result = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/search/query" `
            -Headers $headers `
            -Method Post `
            -Body $searchBody
    }
    catch {
        Write-Host "ERROR: Search API call failed." -ForegroundColor Red
        Write-Host $_.ErrorDetails.Message
        exit 1
    }

    $container = $result.value[0].hitsContainers[0]

    if ($null -eq $total) {
        $total = $container.total
        Write-Host "Graph reports approximately $total site(s) matching the query." -ForegroundColor Green

        if ($total -eq 0) {
            Write-Host "No sites matched. Exiting." -ForegroundColor Yellow
            exit 0
        }
    }

    foreach ($hit in $container.hits) {
        $displayName = $hit.resource.displayName

        # Client-side "contains" filter — enforces exact substring match
        if ($displayName -match [regex]::Escape($NameFilter)) {
            $allSites.Add([PSCustomObject]@{
                DisplayName     = $displayName
                WebUrl          = $hit.resource.webUrl
                CreatedDateTime = $hit.resource.createdDateTime
                Description     = $hit.resource.description
            })
        }
    }

    $from += $PageSize
    Write-Host "  Scanned $from of ~$total, matched $($allSites.Count) so far..." -ForegroundColor Gray

} while ($from -lt $total -and $allSites.Count -lt $MaxResults)

# ── 3. Output — formatted table ───────────────────────────────────────────────
if ($allSites.Count -eq 0) {
    Write-Host "`nNo sites found whose display name contains '$NameFilter'." -ForegroundColor Yellow
}
else {
    Write-Host "`nSites whose display name contains '$NameFilter':`n" -ForegroundColor Cyan

    $nameWidth = [Math]::Max(
        ($allSites | ForEach-Object { $_.DisplayName.Length } | Measure-Object -Maximum).Maximum,
        12
    )
    $nameWidth = [Math]::Min($nameWidth, 50)

    $header = "{0,-4}  {1,-$nameWidth}  {2}" -f "#", "Display Name", "URL"
    $divider = "{0,-4}  {1,-$nameWidth}  {2}" -f "--", ("-" * $nameWidth), ("-" * 60)

    Write-Host $header   -ForegroundColor White
    Write-Host $divider  -ForegroundColor DarkGray

    $i = 0
    foreach ($site in $allSites) {
        $i++
        $truncName = if ($site.DisplayName.Length -gt $nameWidth) {
            $site.DisplayName.Substring(0, $nameWidth - 1) + "…"
        } else {
            $site.DisplayName
        }
        Write-Host ("{0,-4}  {1,-$nameWidth}  {2}" -f $i, $truncName, $site.WebUrl)
    }
}

# ── 4. Summary ────────────────────────────────────────────────────────────────
Write-Host "`n── Summary ───────────────────────────────" -ForegroundColor Cyan
Write-Host "  Filter     : '$NameFilter'"
Write-Host "  Graph total: ~$total"
Write-Host "  Matched    : $($allSites.Count) site(s)"
Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan

# Write objects to the pipeline so callers can capture them:
#   $sites = ./Search-SharePointSites.ps1 -NameFilter "SOW"
$allSites

# ── Phase 1 — Folder check ────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
    exit 0
}

Write-Host "`n── Folder Check ──────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host "  Site       : $SiteUrl"
Write-Host "  Folder     : $FolderName"

# 1.1 — Resolve site ID from URL
# Parse "https://hostname/sites/path" → hostname + server-relative path
$uri = [System.Uri]$SiteUrl
$hostname   = $uri.Host                          # meshsystemsus.sharepoint.com
$serverPath = $uri.AbsolutePath                  # /sites/BUD-ABInbev-PerfectDraft-COM

Write-Host "`nResolving site ID..." -ForegroundColor Cyan
try {
    $siteResource = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:${serverPath}" `
        -Headers $headers `
        -Method Get
    $siteId = $siteResource.id
    Write-Host "  Site ID    : $siteId" -ForegroundColor Green
}
catch {
    $status = $_.Exception.Response.StatusCode.value__
    if ($status -eq 404) {
        Write-Host "ERROR: Site not found — check that -SiteUrl is correct." -ForegroundColor Red
    }
    else {
        Write-Host "ERROR: Failed to resolve site ($status)." -ForegroundColor Red
        Write-Host $_.ErrorDetails.Message
    }
    exit 1
}

# 1.2 — Get the Documents document library drive
Write-Host "`nLooking for Documents library..." -ForegroundColor Cyan
try {
    $drivesResponse = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" `
        -Headers $headers `
        -Method Get
}
catch {
    Write-Host "ERROR: Failed to list drives ($($_.Exception.Response.StatusCode.value__))." -ForegroundColor Red
    Write-Host $_.ErrorDetails.Message
    exit 1
}

$drive = $drivesResponse.value |
    Where-Object { $_.driveType -eq "documentLibrary" -and $_.name -eq "Documents" } |
    Select-Object -First 1

if ($null -eq $drive) {
    # Fallback: first documentLibrary on the site
    $drive = $drivesResponse.value |
        Where-Object { $_.driveType -eq "documentLibrary" } |
        Select-Object -First 1
}

if ($null -eq $drive) {
    Write-Host "ERROR: No document library found on this site." -ForegroundColor Red
    exit 1
}

$driveId = $drive.id
Write-Host "  Drive      : $($drive.name)  ($driveId)" -ForegroundColor Green

# 1.3 — Walk root children pages looking for the folder (exact name, case-insensitive)
Write-Host "`nChecking root of '$($drive.name)' for folder '$FolderName'..." -ForegroundColor Cyan

$folderFound  = $false
$folderWebUrl = $null
$nextLink     = "https://graph.microsoft.com/v1.0/drives/$driveId/root/children"

do {
    try {
        $childrenResponse = Invoke-RestMethod `
            -Uri $nextLink `
            -Headers $headers `
            -Method Get
    }
    catch {
        Write-Host "ERROR: Failed to list drive root ($($_.Exception.Response.StatusCode.value__))." -ForegroundColor Red
        Write-Host $_.ErrorDetails.Message
        exit 1
    }

    foreach ($item in $childrenResponse.value) {
        # Match folders only, exact name, case-insensitive
        if ($null -ne $item.folder -and $item.name -ieq $FolderName) {
            $folderFound  = $true
            $folderWebUrl = $item.webUrl
            break
        }
    }

    $nextLink = $childrenResponse.'@odata.nextLink'

} while (-not $folderFound -and $null -ne $nextLink)

if ($folderFound) {
    Write-Host "`n[FOUND] '$FolderName' exists in the $($drive.name) library." -ForegroundColor Green
    Write-Host "  URL : $folderWebUrl"
}
else {
    Write-Host "`n[NOT FOUND] No folder named '$FolderName' at the root of $($drive.name)." -ForegroundColor Yellow
}

# ── Phase 2 — Create folder if missing ────────────────────────────────────────
if ($folderFound) {
    # Nothing to do — folder already exists
    exit 0
}

if (-not $CreateIfMissing) {
    Write-Host "`n[INFO] Run with -CreateIfMissing to create the folder automatically." -ForegroundColor Gray
    exit 0
}

Write-Host "`nCreating folder '$FolderName'..." -ForegroundColor Cyan

$newFolderBody = @{
    name   = $FolderName
    folder = @{}
    "@microsoft.graph.conflictBehavior" = "fail"
} | ConvertTo-Json -Depth 3

try {
    $created = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/items/root/children" `
        -Headers $headers `
        -Method Post `
        -Body $newFolderBody
    Write-Host "`n[CREATED] Folder '$FolderName' successfully created." -ForegroundColor Green
    Write-Host "  URL : $($created.webUrl)"
}
catch {
    $status = $_.Exception.Response.StatusCode.value__
    if ($status -eq 409) {
        Write-Host "`n[INFO] Folder '$FolderName' already exists (conflict returned by API — likely a race condition)." -ForegroundColor Yellow
    }
    elseif ($status -eq 403) {
        Write-Host "`n[ERROR] Permission denied creating folder." -ForegroundColor Red
        Write-Host "  Ensure 'Files.ReadWrite.All' (application) has been added to app registration"
        Write-Host "  $ClientId and admin consent has been granted."
    }
    else {
        Write-Host "`n[ERROR] Failed to create folder ($status)." -ForegroundColor Red
        Write-Host $_.ErrorDetails.Message
    }
    exit 1
}
