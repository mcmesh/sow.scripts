#Requires -Version 7.0
<#
.SYNOPSIS
    Search for SharePoint sites whose display name contains a given filter string.

.DESCRIPTION
    Uses the Microsoft Graph Search API (POST /v1.0/search/query) with entityTypes
    "site" to find SharePoint sites. Results are additionally filtered client-side
    so only sites whose display name contains -NameFilter are returned.

    Credentials must be supplied via environment variables — do NOT hardcode secrets.
    Set $env:GRAPH_CLIENT_SECRET before running.

.PARAMETER NameFilter
    Required. Case-insensitive substring to match against site display names.

.PARAMETER MaxResults
    Optional. Maximum number of matching sites to return. Default: 200.

.EXAMPLE
    pwsh ./Search-SharePointSites.ps1 -NameFilter "SOW"

.EXAMPLE
    $sites = ./Search-SharePointSites.ps1 -NameFilter "Project" -MaxResults 50
    $sites | Select-Object DisplayName, WebUrl
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string] $NameFilter,

    [Parameter(Mandatory = $false)]
    [int] $MaxResults = 200
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
