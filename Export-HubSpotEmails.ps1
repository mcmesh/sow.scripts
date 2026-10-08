# ══════════════════════════════════════════════════════════════════════════════
#  Export-HubSpotEmails.ps1
#
#  Pulls CRM email activity records from the HubSpot CRM Search API and writes
#  one slim JSON file per email to a local directory, plus a CSV/JSON manifest,
#  an incremental run bookmark, and a zip archive of the output folder.
#
#  Auth: HubSpot Private App token — store in $env:HUBSPOT_API_KEY.
#        Sent as "Authorization: Bearer <token>" on every request.
#
#  Required scopes: crm.objects.emails.read, sales-email-read,
#                   crm.objects.deals.read, crm.objects.companies.read,
#                   crm.objects.contacts.read
#
#  API reference:
#    https://developers.hubspot.com/docs/api-reference/latest/crm/activities/emails/get-emails
#    POST https://api.hubapi.com/crm/v3/objects/emails/search
#
#  Pagination: HubSpot uses cursor-based paging (paging.next.after token).
#
#  10 000-record cap: The HubSpot Search API returns at most 10 000 results per
#  query.  When a date range is expected to exceed that the script automatically
#  splits the range into calendar-month chunks, runs a separate query per chunk,
#  and union-deduplicates — the same strategy used in SOW_docx_Exporter.ps1.
#
#  Usage examples:
#    # Export everything (first run — no bookmark yet)
#    .\Export-HubSpotEmails.ps1
#
#    # Export from a specific date onward (also seeds the incremental bookmark)
#    .\Export-HubSpotEmails.ps1 -BeginDate "2025-01-01"
#
#    # Export a specific date window
#    .\Export-HubSpotEmails.ps1 -BeginDate "2025-01-01" -EndDate "2025-03-31"
#
#    # Re-run later — automatically picks up from last successful run's bookmark
#    .\Export-HubSpotEmails.ps1
#
#    # Filter by subject keyword
#    .\Export-HubSpotEmails.ps1 -SearchQuery "SOW"
# ══════════════════════════════════════════════════════════════════════════════

param(
    # Optional: only return emails created on or after this date (YYYY-MM-DD).
    # If omitted, defaults to the bookmark written by the previous successful
    # run (_lastRun.json in -OutputDir), or all-time if no bookmark exists yet.
    # Explicitly passing this always overrides the bookmark.
    [string]$BeginDate    = "",

    # Optional: only return emails created on or before this date (YYYY-MM-DD).
    [string]$EndDate      = "",

    # Where individual JSON files are written before zipping
    [string]$OutputDir    = "C:\Temp\HubSpotEmails",

    # Final zip archive path
    [string]$ZipOutput    = "C:\Temp\HubSpotEmails.zip",

    # Results per Search API page — HubSpot max is 100
    [int]$PageSize        = 50,

    # Max retry attempts on transient failures (exponential back-off)
    [int]$MaxRetries      = 5,

    # Optional: substring filter applied to hs_email_subject
    [string]$SearchQuery  = ""
)

# ── Config ────────────────────────────────────────────────────────────────────
# Store your Private App token in an environment variable, e.g.:
#   $env:HUBSPOT_API_KEY = "pat-na1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$ApiKey  = $env:HUBSPOT_API_KEY
$BaseUrl = "https://api.hubapi.com"
# ─────────────────────────────────────────────────────────────────────────────

# ── Validate prerequisites ────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Host "ERROR: No API key found." -ForegroundColor Red
    Write-Host "       Set the environment variable HUBSPOT_API_KEY before running:" -ForegroundColor Yellow
    Write-Host '       $env:HUBSPOT_API_KEY = "pat-na1-..."' -ForegroundColor Yellow
    exit 1
}

foreach ($dateParam in @(@{Name="-BeginDate"; Value=$BeginDate}, @{Name="-EndDate"; Value=$EndDate})) {
    if ($dateParam.Value -ne "" -and -not ($dateParam.Value -match '^\d{4}-\d{2}-\d{2}$')) {
        Write-Host "ERROR: $($dateParam.Name) must be in YYYY-MM-DD format (e.g. '2025-01-01')." -ForegroundColor Red
        exit 1
    }
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: Invoke with exponential back-off retry ────────────────────────────
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 5
    )
    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # Do not retry on auth/permission failures — fail fast and show the response body
            if ($statusCode -eq 401 -or $statusCode -eq 403) {
                $label = if ($statusCode -eq 401) { "401 Unauthorized" } else { "403 Forbidden" }
                Write-Host "ERROR: $label — check your HUBSPOT_API_KEY and Private App scopes." -ForegroundColor Red
                # Print HubSpot's error detail from the response body
                if (-not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
                    Write-Host "HubSpot error detail: $($_.ErrorDetails.Message)" -ForegroundColor Red
                }
                throw
            }

            # HubSpot rate limit — honour Retry-After when present
            if ($statusCode -eq 429) {
                $retryAfter = 10
                try {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                } catch { }
                Write-Host "    Rate limited (429). Waiting ${retryAfter}s before retry..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $retryAfter
                continue
            }

            if ($attempt -ge $MaxAttempts) { throw }

            $wait = [math]::Pow(2, $attempt)   # 2, 4, 8, 16 … seconds
            Write-Host "    Transient error (attempt $attempt/$MaxAttempts). Retrying in ${wait}s..." `
                -ForegroundColor DarkYellow
            Start-Sleep -Seconds $wait
        }
    }
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: Flatten email body to plain text ──────────────────────────────────
# HubSpot returns hs_email_text (plain) and/or hs_email_html.
# Prefer plain text; strip HTML tags as fallback.
function ConvertTo-FlatEmailBody {
    param(
        [string]$PlainText,
        [string]$HtmlText
    )

    if (-not [string]::IsNullOrWhiteSpace($PlainText)) {
        return $PlainText.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($HtmlText)) {
        # Strip HTML tags and collapse whitespace
        $stripped = $HtmlText -replace '<[^>]+>', ' '
        $stripped = [System.Net.WebUtility]::HtmlDecode($stripped)
        $stripped = $stripped -replace '\s{2,}', ' '
        return $stripped.Trim()
    }

    return ""
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: Build a HubSpot Search API request body ──────────────────────────
function New-SearchBody {
    param(
        [long]$AfterMs   = 0,      # hs_createdate GTE (Unix ms); 0 = no lower bound
        [long]$BeforeMs  = 0,      # hs_createdate LTE (Unix ms); 0 = no upper bound
        [string]$Cursor  = "",     # paging.next.after token
        [string]$Subject = ""      # optional subject substring
    )

    $filters = [System.Collections.Generic.List[object]]::new()

    if ($AfterMs -gt 0) {
        $filters.Add(@{ propertyName = "hs_createdate"; operator = "GTE"; value = "$AfterMs" })
    }
    if ($BeforeMs -gt 0) {
        $filters.Add(@{ propertyName = "hs_createdate"; operator = "LTE"; value = "$BeforeMs" })
    }
    if (-not [string]::IsNullOrWhiteSpace($Subject)) {
        $filters.Add(@{ propertyName = "hs_email_subject"; operator = "CONTAINS_TOKEN"; value = $Subject })
    }

    $body = [ordered]@{
        filterGroups = @(@{ filters = $filters.ToArray() })
        properties   = @(
            "hs_email_subject",
            "hs_email_html",
            "hs_email_text",
            "hs_email_from_email",
            "hs_email_to_email",
            "hs_email_cc",
            "hs_email_bcc",
            "hs_email_direction",
            "hs_createdate",
            "hs_timestamp"
        )
        associations = @("contacts", "deals", "companies")
        limit        = $PageSize
    }

    if (-not [string]::IsNullOrWhiteSpace($Cursor)) {
        $body["after"] = $Cursor
    }

    return $body
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: Fetch all pages for one date-range chunk ─────────────────────────
function Get-EmailChunk {
    param(
        [long]$AfterMs,
        [long]$BeforeMs,
        [string]$Subject,
        [hashtable]$Headers,
        [System.Collections.Generic.HashSet[string]]$SeenIds,
        [System.Collections.Generic.List[object]]$Results,
        [ref]$ErrorCount
    )

    $cursor    = ""
    $pageIndex = 0
    $chunkNew  = 0

    do {
        $pageIndex++
        $body    = New-SearchBody -AfterMs $AfterMs -BeforeMs $BeforeMs -Cursor $cursor -Subject $Subject
        $bodyJson = $body | ConvertTo-Json -Depth 10 -Compress

        try {
            $response = Invoke-WithRetry -MaxAttempts $MaxRetries -ScriptBlock {
                Invoke-RestMethod `
                    -Uri     "$BaseUrl/crm/v3/objects/emails/search" `
                    -Headers $Headers `
                    -Method  Post `
                    -Body    $bodyJson
            }
        }
        catch {
            $ErrorCount.Value++
            Write-Host "  ERROR on page $pageIndex — $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  Aborting this chunk." -ForegroundColor DarkRed
            return $chunkNew
        }

        $items = $response.results
        if ($null -eq $items -or $items.Count -eq 0) {
            Write-Host "  Page $pageIndex returned no items — chunk complete." -ForegroundColor DarkGray
            break
        }

        foreach ($email in $items) {
            if ($SeenIds.Add($email.id)) {
                $Results.Add($email)
                $chunkNew++
            }
        }

        Write-Host "  Page $pageIndex — $($items.Count) items ($($Results.Count) unique total)" -ForegroundColor Gray

        $cursor = $response.paging.next.after

    } while (-not [string]::IsNullOrWhiteSpace($cursor))

    return $chunkNew
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: Batch-read CRM objects by ID (max 100 per call) ──────────────────
# Returns a hashtable keyed by object ID -> properties object.
# Silently skips the enrichment (returns empty table) if the token lacks the
# required scope — a warning is printed but the export continues.
function Get-CrmObjectsBatch {
    param(
        [string]$ObjectType,           # e.g. "deals", "companies", "contacts"
        [string[]]$Ids,
        [string[]]$Properties,
        [hashtable]$Headers
    )

    $lookup = @{}
    if ($Ids.Count -eq 0) { return $lookup }

    $batchSize = 100
    for ($i = 0; $i -lt $Ids.Count; $i += $batchSize) {
        $chunk = $Ids[$i..([math]::Min($i + $batchSize - 1, $Ids.Count - 1))]
        $body  = @{
            inputs     = @($chunk | ForEach-Object { @{ id = $_ } })
            properties = $Properties
        } | ConvertTo-Json -Depth 5 -Compress

        try {
            $response = Invoke-WithRetry -MaxAttempts $MaxRetries -ScriptBlock {
                Invoke-RestMethod `
                    -Uri     "$BaseUrl/crm/v3/objects/$ObjectType/batch/read" `
                    -Headers $Headers `
                    -Method  Post `
                    -Body    $body
            }
            foreach ($obj in @($response.results)) {
                if ($obj.id) { $lookup[$obj.id] = $obj.properties }
            }
        }
        catch {
            $sc = $null
            if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
            if ($sc -eq 403) {
                Write-Host "  WARNING: 403 on $ObjectType batch read — token may lack crm.objects.$ObjectType.read scope. Skipping enrichment." -ForegroundColor DarkYellow
                if (-not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
                    Write-Host "  Detail: $($_.ErrorDetails.Message)" -ForegroundColor DarkYellow
                }
                return $lookup   # return whatever we have so far; caller gets partial/empty data gracefully
            }
            Write-Host "  WARNING: Could not batch-fetch $ObjectType — $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    return $lookup
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Build shared auth headers ─────────────────────────────────────────────────
$Headers = @{
    "Authorization" = "Bearer $ApiKey"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Resolve incremental bookmark ──────────────────────────────────────────────
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$StateFile = Join-Path $OutputDir "_lastRun.json"

$effectiveBeginDate = $BeginDate
$usedBookmark       = $false
if ($effectiveBeginDate -eq "" -and (Test-Path $StateFile)) {
    try {
        $state = Get-Content $StateFile -Raw | ConvertFrom-Json
        if ($state.lastCreatedAt) {
            $effectiveBeginDate = ([datetime]$state.lastCreatedAt).ToString("yyyy-MM-dd")
            $usedBookmark = $true
        }
    }
    catch {
        Write-Host "WARNING: Could not read bookmark at $StateFile — ignoring. $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

# Convert dates to Unix milliseconds for HubSpot filters
function ConvertTo-UnixMs ([string]$DateStr, [bool]$EndOfDay = $false) {
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return 0 }
    $dt = [datetime]::ParseExact($DateStr, "yyyy-MM-dd", $null)
    if ($EndOfDay) { $dt = $dt.AddHours(23).AddMinutes(59).AddSeconds(59) }
    return [long]([datetimeoffset]::new($dt, [System.TimeSpan]::Zero).ToUnixTimeMilliseconds())
}

$beginMs = ConvertTo-UnixMs -DateStr $effectiveBeginDate
$endMs   = ConvertTo-UnixMs -DateStr $EndDate -EndOfDay $true
# ─────────────────────────────────────────────────────────────────────────────

# ── Announce run parameters ───────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  HubSpot Email Exporter"                  -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  API Base    : $BaseUrl"
Write-Host "  Begin Date  : $(if ($effectiveBeginDate) { $effectiveBeginDate } else { '(all time)' })$(if ($usedBookmark) { ' (from bookmark)' })"
Write-Host "  End Date    : $(if ($EndDate) { $EndDate } else { '(none)' })"
Write-Host "  Page Size   : $PageSize"
Write-Host "  Output Dir  : $OutputDir"
Write-Host "  Zip Output  : $ZipOutput"
if ($SearchQuery -ne "") { Write-Host "  Subject Filter : $SearchQuery" }
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
# ─────────────────────────────────────────────────────────────────────────────

# ── 1. Page through emails (with auto-chunking past 10 000) ───────────────────
Write-Host "Fetching emails from HubSpot..." -ForegroundColor Cyan

$allEmails  = [System.Collections.Generic.List[object]]::new()
$fetchedIds = [System.Collections.Generic.HashSet[string]]::new()
$errorCount = 0

# Determine whether to chunk by month (guard against the 10 000-record cap).
# Strategy: attempt a first pass over the full range. If results hit exactly
# $PageSize * 200 (= 10 000) we assume truncation and fall back to monthly chunks.
# This avoids an extra count query on every run.

$needsChunking = $false

if ($beginMs -gt 0) {
    # Check total by issuing a minimal first-page probe
    $probeBody = New-SearchBody -AfterMs $beginMs -BeforeMs $endMs -Subject $SearchQuery
    $probeBody["limit"] = 1
    try {
        $probe = Invoke-WithRetry -MaxAttempts $MaxRetries -ScriptBlock {
            Invoke-RestMethod `
                -Uri     "$BaseUrl/crm/v3/objects/emails/search" `
                -Headers $Headers `
                -Method  Post `
                -Body    ($probeBody | ConvertTo-Json -Depth 10 -Compress)
        }
        $estimatedTotal = $probe.total
        Write-Host "  Estimated email count : $estimatedTotal" -ForegroundColor Green
        if ($estimatedTotal -ge 10000) {
            $needsChunking = $true
            Write-Host "  Count >= 10 000 — enabling monthly chunking." -ForegroundColor DarkYellow
        }
    }
    catch {
        Write-Host "  WARNING: Could not probe total count — will attempt single-pass fetch. $($_.Exception.Message)" `
            -ForegroundColor DarkYellow
    }
}

if (-not $needsChunking) {
    # Single-pass over the full (possibly open-ended) range
    Write-Host "  Running single-pass fetch..." -ForegroundColor DarkCyan
    $null = Get-EmailChunk `
        -AfterMs  $beginMs `
        -BeforeMs $endMs `
        -Subject  $SearchQuery `
        -Headers  $Headers `
        -SeenIds  $fetchedIds `
        -Results  $allEmails `
        -ErrorCount ([ref]$errorCount)
}
else {
    # Monthly chunking: iterate from beginDate month to endDate (or today) month
    $chunkStart = if ($beginMs -gt 0) {
        [datetimeoffset]::FromUnixTimeMilliseconds($beginMs).UtcDateTime
    } else {
        [datetime]::new(2010, 1, 1)
    }
    $chunkEnd = if ($endMs -gt 0) {
        [datetimeoffset]::FromUnixTimeMilliseconds($endMs).UtcDateTime
    } else {
        [datetime]::UtcNow
    }

    $monthCursor = [datetime]::new($chunkStart.Year, $chunkStart.Month, 1)

    while ($monthCursor -le $chunkEnd) {
        $monthStart = $monthCursor
        $monthEnd   = $monthCursor.AddMonths(1).AddSeconds(-1)
        if ($monthEnd -gt $chunkEnd) { $monthEnd = $chunkEnd }

        $chunkAfterMs  = [long]([datetimeoffset]::new($monthStart, [System.TimeSpan]::Zero).ToUnixTimeMilliseconds())
        $chunkBeforeMs = [long]([datetimeoffset]::new($monthEnd,   [System.TimeSpan]::Zero).ToUnixTimeMilliseconds())

        Write-Host "  Chunk: $($monthStart.ToString('yyyy-MM')) ..." -ForegroundColor DarkCyan

        $null = Get-EmailChunk `
            -AfterMs  $chunkAfterMs `
            -BeforeMs $chunkBeforeMs `
            -Subject  $SearchQuery `
            -Headers  $Headers `
            -SeenIds  $fetchedIds `
            -Results  $allEmails `
            -ErrorCount ([ref]$errorCount)

        $monthCursor = $monthCursor.AddMonths(1)
    }
}

if ($allEmails.Count -eq 0) {
    Write-Host "`nNo emails found for the given filters. Exiting." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nFetch complete. $($allEmails.Count) unique emails retrieved." -ForegroundColor Green
# ─────────────────────────────────────────────────────────────────────────────

# ── 1b. Enrich: batch-fetch deal and company details ─────────────────────────
Write-Host "`nEnriching with deal and company details..." -ForegroundColor Cyan

# Collect all unique association IDs across every email
$allDealIds    = [System.Collections.Generic.HashSet[string]]::new()
$allCompanyIds = [System.Collections.Generic.HashSet[string]]::new()
$allContactIds = [System.Collections.Generic.HashSet[string]]::new()

foreach ($email in $allEmails) {
    if ($email.associations) {
        foreach ($a in @($email.associations.deals.results))    { if ($a.id) { [void]$allDealIds.Add($a.id) } }
        foreach ($a in @($email.associations.companies.results)) { if ($a.id) { [void]$allCompanyIds.Add($a.id) } }
        foreach ($a in @($email.associations.contacts.results)) { if ($a.id) { [void]$allContactIds.Add($a.id) } }
    }
}

Write-Host "  Unique deals    : $($allDealIds.Count)" -ForegroundColor Gray
Write-Host "  Unique companies: $($allCompanyIds.Count)" -ForegroundColor Gray
Write-Host "  Unique contacts : $($allContactIds.Count)" -ForegroundColor Gray

$dealLookup    = Get-CrmObjectsBatch -ObjectType "deals"    -Ids @($allDealIds)    -Properties @("dealname","dealstage","closedate","amount") -Headers $Headers
$companyLookup = Get-CrmObjectsBatch -ObjectType "companies" -Ids @($allCompanyIds) -Properties @("name","domain","industry")                  -Headers $Headers
$contactLookup = Get-CrmObjectsBatch -ObjectType "contacts"  -Ids @($allContactIds) -Properties @("firstname","lastname","email","company")    -Headers $Headers

Write-Host "  Deals resolved    : $($dealLookup.Count)" -ForegroundColor Green
Write-Host "  Companies resolved: $($companyLookup.Count)" -ForegroundColor Green
Write-Host "  Contacts resolved : $($contactLookup.Count)" -ForegroundColor Green
# ─────────────────────────────────────────────────────────────────────────────


# ── 2. Write each email as a slim JSON file ────────────────────────────────────
Write-Host "`nWriting email files to $OutputDir..." -ForegroundColor Cyan

$savedCount   = 0
$saveErrors   = 0
$seenIds      = [System.Collections.Generic.HashSet[string]]::new()
$manifestRows = [System.Collections.Generic.List[object]]::new()
$maxCreatedMs = [long]0

foreach ($email in $allEmails) {
    try {
        $props  = $email.properties
        $emailId = $email.id

        if ([string]::IsNullOrWhiteSpace($emailId)) {
            Write-Host "  SKIP: email with no id" -ForegroundColor DarkGray
            continue
        }

        if (-not $seenIds.Add($emailId)) {
            Write-Host "  SKIP (duplicate): $emailId" -ForegroundColor DarkGray
            continue
        }

        # hs_createdate comes back as Unix ms string
        $createdMs = [long]0
        $createdAt = ""
        $datePart  = "unknown-date"
        if (-not [string]::IsNullOrWhiteSpace($props.hs_createdate)) {
            try {
                $createdMs = [long]$props.hs_createdate
                $createdAt = [datetimeoffset]::FromUnixTimeMilliseconds($createdMs).UtcDateTime.ToString("o")
                $datePart  = [datetimeoffset]::FromUnixTimeMilliseconds($createdMs).UtcDateTime.ToString("yyyy-MM-dd")
                if ($createdMs -gt $maxCreatedMs) { $maxCreatedMs = $createdMs }
            } catch { }
        }

        $subject = if (-not [string]::IsNullOrWhiteSpace($props.hs_email_subject)) {
            $props.hs_email_subject
        } else { "(no subject)" }

        $safeSubject = ($subject -replace '[\\/:*?"<>|]', '_')
        $safeSubject = $safeSubject.Substring(0, [math]::Min($safeSubject.Length, 60))

        $body = ConvertTo-FlatEmailBody -PlainText $props.hs_email_text -HtmlText $props.hs_email_html

        # Extract associated IDs and enrich with looked-up details
        $contactIds  = [System.Collections.Generic.List[string]]::new()
        $dealIds     = [System.Collections.Generic.List[string]]::new()
        $companyIds  = [System.Collections.Generic.List[string]]::new()
        if ($email.associations) {
            if ($email.associations.contacts) {
                foreach ($a in @($email.associations.contacts.results)) {
                    if ($a.id) { [void]$contactIds.Add($a.id) }
                }
            }
            if ($email.associations.deals) {
                foreach ($a in @($email.associations.deals.results)) {
                    if ($a.id) { [void]$dealIds.Add($a.id) }
                }
            }
            if ($email.associations.companies) {
                foreach ($a in @($email.associations.companies.results)) {
                    if ($a.id) { [void]$companyIds.Add($a.id) }
                }
            }
        }

        # Build enriched detail objects from lookups
        $deals = @($dealIds | ForEach-Object {
            $p = $dealLookup[$_]
            if ($p) {
                [PSCustomObject]@{ id = $_; name = $p.dealname; stage = $p.dealstage; closeDate = $p.closedate; amount = $p.amount }
            } else {
                [PSCustomObject]@{ id = $_ }
            }
        })
        $companies = @($companyIds | ForEach-Object {
            $p = $companyLookup[$_]
            if ($p) {
                [PSCustomObject]@{ id = $_; name = $p.name; domain = $p.domain; industry = $p.industry }
            } else {
                [PSCustomObject]@{ id = $_ }
            }
        })
        $contacts = @($contactIds | ForEach-Object {
            $p = $contactLookup[$_]
            if ($p) {
                [PSCustomObject]@{ id = $_; firstName = $p.firstname; lastName = $p.lastname; email = $p.email; company = $p.company }
            } else {
                [PSCustomObject]@{ id = $_ }
            }
        })

        # Primary company/deal name for manifest (first resolved value)
        $primaryCompany = ($companies | Where-Object { $_.name } | Select-Object -First 1).name
        $primaryDeal    = ($deals    | Where-Object { $_.name } | Select-Object -First 1).name

        $record = [PSCustomObject]@{
            id           = $emailId
            subject      = $subject
            body         = $body
            from         = $props.hs_email_from_email
            to           = $props.hs_email_to_email
            cc           = $props.hs_email_cc
            bcc          = $props.hs_email_bcc
            direction    = $props.hs_email_direction
            createdAt    = $createdAt
            associations = [PSCustomObject]@{
                contacts  = $contacts
                deals     = $deals
                companies = $companies
            }
        }

        $fileName  = "${datePart}_${safeSubject}_${emailId}.json"
        $localPath = Join-Path $OutputDir $fileName

        $json = $record | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($localPath, $json, [System.Text.Encoding]::UTF8)

        $manifestRows.Add([PSCustomObject]@{
            id        = $emailId
            filename  = $fileName
            subject   = $subject
            from      = $props.hs_email_from_email
            direction = $props.hs_email_direction
            createdAt = $createdAt
            company   = $primaryCompany
            deal      = $primaryDeal
        })

        $savedCount++
        Write-Host "  [$savedCount] $fileName" -ForegroundColor Gray
    }
    catch {
        $saveErrors++
        Write-Host "  ERROR saving email $($email.id): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nFile writing complete. $savedCount files saved." -ForegroundColor Green
# ─────────────────────────────────────────────────────────────────────────────

# ── 3. Write manifest (JSON + CSV) ────────────────────────────────────────────
try {
    $manifestJsonPath = Join-Path $OutputDir "_manifest.json"
    [System.IO.File]::WriteAllText(
        $manifestJsonPath,
        ($manifestRows | ConvertTo-Json -Depth 5),
        [System.Text.Encoding]::UTF8
    )
    Write-Host "Manifest written: $manifestJsonPath" -ForegroundColor DarkGray

    $manifestCsvPath = Join-Path $OutputDir "_manifest.csv"
    $manifestRows | Export-Csv -Path $manifestCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Manifest written: $manifestCsvPath" -ForegroundColor DarkGray
}
catch {
    Write-Host "WARNING: Could not write manifest — $($_.Exception.Message)" -ForegroundColor DarkYellow
}
# ─────────────────────────────────────────────────────────────────────────────

# ── 4. Update incremental bookmark ────────────────────────────────────────────
# Only advance the bookmark on a clean run (no fetch errors) and only when we
# actually saw a newer createdAt — avoids silently skipping data on a partial failure.
if ($errorCount -eq 0 -and $maxCreatedMs -gt 0) {
    try {
        $lastCreatedAt = [datetimeoffset]::FromUnixTimeMilliseconds($maxCreatedMs).UtcDateTime.ToString("o")
        $state = @{ lastCreatedAt = $lastCreatedAt }
        [System.IO.File]::WriteAllText($StateFile, ($state | ConvertTo-Json), [System.Text.Encoding]::UTF8)
        Write-Host "Bookmark updated: $StateFile -> $lastCreatedAt" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "WARNING: Could not write bookmark — $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}
elseif ($errorCount -gt 0) {
    Write-Host "Bookmark NOT updated — run had $errorCount fetch error(s)." -ForegroundColor DarkYellow
}
# ─────────────────────────────────────────────────────────────────────────────

# ── 5. Zip the output directory ───────────────────────────────────────────────
if ($savedCount -gt 0) {
    Write-Host "`nCreating zip archive at $ZipOutput..." -ForegroundColor Cyan

    if (Test-Path $ZipOutput) {
        Remove-Item $ZipOutput -Force
        Write-Host "  Removed existing zip." -ForegroundColor DarkGray
    }

    try {
        Compress-Archive -Path "$OutputDir\*" -DestinationPath $ZipOutput -Force
        $zipSize = (Get-Item $ZipOutput).Length / 1MB
        Write-Host "  Zip created — $([math]::Round($zipSize, 2)) MB" -ForegroundColor Green
    }
    catch {
        Write-Host "  ERROR creating zip: $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
    }
}
else {
    Write-Host "`nNo files to zip — skipping archive step." -ForegroundColor Yellow
}
# ─────────────────────────────────────────────────────────────────────────────

# ── 6. Summary ────────────────────────────────────────────────────────────────
$byDirection = $manifestRows | Group-Object direction | Sort-Object Count -Descending

Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SUMMARY"                                  -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Emails fetched       : $($allEmails.Count)"
Write-Host "  Deals enriched       : $($dealLookup.Count)"
Write-Host "  Companies enriched   : $($companyLookup.Count)"
Write-Host "  Contacts enriched    : $($contactLookup.Count)"
Write-Host "  JSON files saved     : $savedCount"
Write-Host "  Save errors          : $saveErrors"
Write-Host "  API/fetch errors     : $errorCount"
Write-Host "  Direction breakdown:"
foreach ($g in $byDirection) {
    $label = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(unknown)" } else { $g.Name }
    Write-Host "    ${label}: $($g.Count)"
}
if ($savedCount -gt 0) {
    Write-Host "  Zip archive          : $ZipOutput"
}
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
