# ── Config ────────────────────────────────────────────────────────────────────
# Store your client secret in an environment variable before running, e.g.:
#   $env:GRAPH_CLIENT_SECRET = "your_secret_here"
$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf"
$ClientId     = "5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b"
$ClientSecret = $env:GRAPH_CLIENT_SECRET
$DownloadDir  = "C:\Temp\Pricing_Files_Spreadsheets"
$ZipOutput    = "C:\Temp\Pricing_Files_Spreadsheets.zip"
$SearchQuery  = "FileName:*Pricing* AND (FileType:xlsx OR FileType:xls)"
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

$downloadHeaders = @{
    "Authorization" = "Bearer $token"
}

# ── 2. Search — page through ALL results ──────────────────────────────────────
Write-Host "`nSearching for SOW files..." -ForegroundColor Cyan

$allFiles = [System.Collections.Generic.List[hashtable]]::new()
$from     = 0
$total    = $null

do {
    $searchBody = @{
        requests = @(
            @{
                entityTypes = @("driveItem")
                query       = @{ queryString = $SearchQuery }
                fields      = @("name", "webUrl", "createdDateTime")
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
        Write-Host "Total files found: $total" -ForegroundColor Green

        if ($total -eq 0) {
            Write-Host "No files matched the search query. Exiting." -ForegroundColor Yellow
            exit 0
        }
    }

    foreach ($hit in $container.hits) {
        $allFiles.Add(@{
            Name   = $hit.resource.name
            WebUrl = $hit.resource.webUrl
        })
    }

    $from += $PageSize
    Write-Host "  Retrieved $($allFiles.Count) of $total..." -ForegroundColor Gray

} while ($from -lt $total)

Write-Host "Search complete. $($allFiles.Count) files found." -ForegroundColor Green

# ── 3. Download Files ─────────────────────────────────────────────────────────
New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null
Write-Host "`nDownloading to $DownloadDir..." -ForegroundColor Cyan

$count      = 0
$errorCount = 0

foreach ($file in $allFiles) {
    try {
        # Encode webUrl as a base64 sharing token
        $encodedUrl = [Convert]::ToBase64String(
            [System.Text.Encoding]::UTF8.GetBytes($file.WebUrl)
        ).TrimEnd('=').Replace('+', '-').Replace('/', '_')

        $sharingUrl  = "u!$encodedUrl"
        $downloadUrl = "https://graph.microsoft.com/v1.0/shares/$sharingUrl/driveItem/content"

        # Handle duplicate filenames
        $localPath = Join-Path $DownloadDir $file.Name
        if (Test-Path $localPath) {
            $base = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $ext  = [System.IO.Path]::GetExtension($file.Name)
            $i    = 1
            while (Test-Path $localPath) {
                $localPath = Join-Path $DownloadDir "${base}_${i}${ext}"
                $i++
            }
        }

        Invoke-WebRequest `
            -Uri $downloadUrl `
            -Headers $downloadHeaders `
            -Method Get `
            -OutFile $localPath `
            -MaximumRedirection 5

        $count++
        Write-Host "  [$count/$($allFiles.Count)] $($file.Name)" -ForegroundColor Gray
    }
    catch {
        $errorCount++
        Write-Host "  ERROR: $($file.Name) — $($_.ErrorDetails.Message)" -ForegroundColor Red
    }
}

# ── 4. Zip ────────────────────────────────────────────────────────────────────
if ($count -gt 0) {
    Write-Host "`nCreating zip at $ZipOutput..." -ForegroundColor Cyan
    if (Test-Path $ZipOutput) { Remove-Item $ZipOutput -Force }
    Compress-Archive -Path "$DownloadDir\*" -DestinationPath $ZipOutput -Force
    Write-Host "Done!" -ForegroundColor Green
}
else {
    Write-Host "No files downloaded — skipping zip." -ForegroundColor Yellow
}

# ── 5. Summary ────────────────────────────────────────────────────────────────
Write-Host "`n── Summary ───────────────────────────────" -ForegroundColor Cyan
Write-Host "  Found      : $total files"
Write-Host "  Downloaded : $count files"
Write-Host "  Errors     : $errorCount files"
if ($count -gt 0) {
    Write-Host "  Zip        : $ZipOutput"
}
Write-Host "──────────────────────────────────────────" -ForegroundColor Cyan