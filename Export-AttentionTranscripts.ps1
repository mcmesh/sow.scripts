# ══════════════════════════════════════════════════════════════════════════════
#  Export-AttentionTranscripts.ps1
#
#  Pulls conversations from the Attention API (attention.tech), same
#  auth/pagination/retry scaffolding as Export-AttentionConversations.ps1, but
#  writes slim transcript-only files tagged with best-guess customer/project
#  metadata plus a manifest — for a downstream consumer (human or script) to
#  sort/move into per-customer/per-project folders. This script does NOT sort
#  or move files itself.
#
#  Why "best-guess"? There is no customer/project field in the API payload,
#  and no server-side filter resolves a customer/deal NAME to an ID (confirmed
#  via Test-AttentionCrmFilters.ps1 — see Attention-Customer-Project-Filtering-Plan.md).
#  Metadata is derived client-side using a priority-ordered heuristic:
#
#    1. externalOpportunity.title split on "_" or " - " -> customer / project
#       (~47% of conversations have this field; ~95% follow the convention)
#    2. External participant email domain -> friendly customer name
#       (~48% coverage; catches many conversations with no opportunity link)
#    3. Raw conversation title, flagged low-confidence, for manual sorting
#    4. "Unknown" when none of the above resolve anything (internal-only
#       syncs with no external participants land here — a hard floor, not a
#       bug to fix)
#
#  Each record is tagged with customerSource/projectSource
#  (opportunity/domain/title/none) so the consumer knows how much to trust it.
#
#  Auth:  https://docs.attention.com/api-authentication
#  Route: https://docs.attention.com/api-reference/conversation/list-conversations-optimized
#
#  Usage examples:
#    # Export everything
#    .\Export-AttentionTranscripts.ps1
#
#    # Export from a specific date onward (also seeds the incremental bookmark)
#    .\Export-AttentionTranscripts.ps1 -BeginDate "2025-01-01"
#
#    # Re-run later with no args — automatically picks up from the last
#    # successful run's bookmark (see _lastRun.json in -OutputDir)
#    .\Export-AttentionTranscripts.ps1
#
#    # Targeted re-fetch of a known CRM opportunity/deal
#    .\Export-AttentionTranscripts.ps1 -OpportunityIds "55649281914"
#
#    # Override the default domain -> friendly-name guess
#    .\Export-AttentionTranscripts.ps1 -DomainMap "C:\Temp\domain-map.json"
# ══════════════════════════════════════════════════════════════════════════════

param(
    # Optional: only return conversations on or after this date (ISO 8601 date, e.g. "2025-01-01").
    # If omitted, defaults to the bookmark left by the previous successful run (see _lastRun.json),
    # or all-time if no bookmark exists yet. Explicitly passing this always overrides the bookmark.
    [string]$BeginDate         = "",

    # Where individual JSON files are written before zipping
    [string]$OutputDir         = "C:\Temp\AttentionTranscripts",

    # Final zip archive path
    [string]$ZipOutput         = "C:\Temp\AttentionTranscripts.zip",

    # Items per page — Attention API default is 20; keep <=100 to be safe
    [int]$PageSize             = 50,

    # Max retry attempts on transient failures (exponential back-off)
    [int]$MaxRetries           = 5,

    # Optional: passthrough to documented filter[title] (substring match).
    # Accepts one or more terms; results are unioned and deduplicated so you can
    # cover alternate spellings, e.g. -TitleContains "AB In","ABIn"
    [string[]]$TitleContains   = @(),

    # Optional: passthrough to documented filter[participants.email] (one or more addresses)
    [string[]]$ParticipantEmails = @(),

    # Optional: passthrough to confirmed-working filter[external_opportunity.id],
    # for a targeted re-fetch of one or more known deals
    [string[]]$OpportunityIds  = @(),

    # Optional: path to a JSON file mapping domain -> friendly customer name
    # (e.g. { "bunn.com": "Bunn", "gill.group": "Gill Group" }). Overrides the
    # default "second-to-last label, capitalized" guess for any domain present.
    [string]$DomainMap         = "",

    # Domains treated as internal (never used as a customer signal). Defaults
    # cover both spellings seen in real tenant data.
    [string[]]$InternalDomains = @("mesh-systems.com", "meshsystems.com")
)

# ── Config ────────────────────────────────────────────────────────────────────
# Store your key in an environment variable and reference it here, e.g.:
#   $env:ATTENTION_API_KEY = "your_key_here"   (set once in your session)
$ApiKey  = $env:ATTENTION_API_KEY
$BaseUrl = "https://api.attention.tech/v2"
# ─────────────────────────────────────────────────────────────────────────────

# ── Validate prerequisites ────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Host "ERROR: No API key found." -ForegroundColor Red
    Write-Host "       Set the environment variable ATTENTION_API_KEY before running:" -ForegroundColor Yellow
    Write-Host '       $env:ATTENTION_API_KEY = "your_key_here"' -ForegroundColor Yellow
    exit 1
}

if ($BeginDate -ne "" -and -not ($BeginDate -match '^\d{4}-\d{2}-\d{2}')) {
    Write-Host "ERROR: -BeginDate must be in YYYY-MM-DD format (e.g. '2025-01-01')." -ForegroundColor Red
    exit 1
}

$domainMapTable = @{}
if ($DomainMap -ne "") {
    if (-not (Test-Path $DomainMap)) {
        Write-Host "ERROR: -DomainMap file not found: $DomainMap" -ForegroundColor Red
        exit 1
    }
    try {
        $rawMap = Get-Content $DomainMap -Raw | ConvertFrom-Json
        $rawMap.PSObject.Properties | ForEach-Object {
            $domainMapTable[$_.Name.ToLowerInvariant()] = $_.Value
        }
    }
    catch {
        Write-Host "ERROR: Could not parse -DomainMap JSON: $($_.Exception.Message)" -ForegroundColor Red
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

            # Do not retry on auth failures — fail fast
            if ($statusCode -eq 401) {
                Write-Host "ERROR: 401 Unauthorized — check your ATTENTION_API_KEY." -ForegroundColor Red
                throw
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

# ── Helper: Normalize transcript to flat "Speaker:\nText" string ─────────────
# See Export-AttentionConversations.ps1 for full rationale — the API
# inconsistently returns transcript as a plain string or a word-level array.
function ConvertTo-FlatTranscript {
    param([object]$Transcript)

    if ($Transcript -is [string]) {
        return $Transcript
    }

    if ($null -eq $Transcript) {
        return ""
    }

    if ($Transcript -is [System.Collections.IEnumerable]) {
        $lines = [System.Text.StringBuilder]::new()
        foreach ($turn in $Transcript) {
            $speakerName = $turn.speaker.name
            if ([string]::IsNullOrWhiteSpace($speakerName)) {
                $speakerName = $turn.speaker.email
            }

            $sentence = ($turn.words | ForEach-Object { $_.text }) -join ""
            $sentence = $sentence.Trim()

            if (-not [string]::IsNullOrWhiteSpace($sentence)) {
                if ($lines.Length -gt 0) { [void]$lines.Append("`n") }
                [void]$lines.Append("${speakerName}:`n${sentence}")
            }
        }
        return $lines.ToString()
    }

    return ($Transcript | ConvertTo-Json -Depth 10 -Compress)
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: friendly customer name from an email domain ───────────────────────
# Default guess takes the second-to-last label (e.g. "uk.spiraxsarco.com" and
# "spiraxsarco.com" both -> "Spiraxsarco"), capitalized. -DomainMap overrides
# any specific domain with an explicit friendly name.
function Get-FriendlyDomainName {
    param([string]$Domain)

    $key = $Domain.ToLowerInvariant()
    if ($domainMapTable.ContainsKey($key)) {
        return $domainMapTable[$key]
    }

    $labels = $Domain.Split('.')
    $core = if ($labels.Count -ge 2) { $labels[$labels.Count - 2] } else { $labels[0] }
    if ($core.Length -eq 0) { return $Domain }
    return $core.Substring(0,1).ToUpperInvariant() + $core.Substring(1)
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Helper: derive customer/project metadata for one conversation ────────────
function Get-CustomerProjectInfo {
    param(
        [object]$Attributes,
        [string[]]$InternalDomains
    )

    $internalSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]($InternalDomains | ForEach-Object { $_.ToLowerInvariant() })
    )

    # External participant domains (excludes internal domains)
    $externalDomains = [System.Collections.Generic.List[string]]::new()
    $seenDomains = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($p in @($Attributes.participants)) {
        $email = $p.email
        if ($email -and $email.Contains("@")) {
            $domain = $email.Substring($email.LastIndexOf("@") + 1).ToLowerInvariant()
            if (-not $internalSet.Contains($domain) -and $seenDomains.Add($domain)) {
                $externalDomains.Add($domain)
            }
        }
    }

    $customer       = $null
    $project        = $null
    $customerSource = "none"
    $projectSource  = "none"

    # 1. Primary: split externalOpportunity.title on its delimiter
    $oppTitle = $Attributes.externalOpportunity.title
    if (-not [string]::IsNullOrWhiteSpace($oppTitle)) {
        $splitIndex = $oppTitle.IndexOf("_")
        $delimLength = 1
        if ($splitIndex -lt 0) {
            $splitIndex = $oppTitle.IndexOf(" - ")
            $delimLength = 3
        }

        if ($splitIndex -gt 0) {
            $customer       = $oppTitle.Substring(0, $splitIndex).Trim()
            $project        = $oppTitle.Substring($splitIndex + $delimLength).Trim()
            $customerSource = "opportunity"
            $projectSource  = "opportunity"
        }
    }

    # 2. Fallback: external attendee domain -> friendly name
    if (-not $customer -and $externalDomains.Count -gt 0) {
        $customer = ($externalDomains | ForEach-Object { Get-FriendlyDomainName -Domain $_ }) -join " + "
        $customerSource = "domain"
    }

    # 3. Last resort: raw conversation title, flagged low-confidence
    if (-not $customer -and -not [string]::IsNullOrWhiteSpace($Attributes.title)) {
        $customer       = $Attributes.title
        $customerSource = "title"
    }

    if (-not $customer) { $customer = "Unknown"; $customerSource = "none" }
    if (-not $project)  { $project  = "Unknown";  $projectSource  = "none" }

    return [PSCustomObject]@{
        customer         = $customer
        project           = $project
        customerSource    = $customerSource
        projectSource     = $projectSource
        externalDomains   = $externalDomains
    }
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
# ─────────────────────────────────────────────────────────────────────────────

# ── Announce run parameters ───────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Attention Transcript Exporter"           -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  API Base   : $BaseUrl"
Write-Host "  Begin Date : $(if ($effectiveBeginDate) { $effectiveBeginDate }` else { '(all time)' })$(if ($usedBookmark) { ' (from bookmark)' })"
Write-Host "  Page Size  : $PageSize"
Write-Host "  Output Dir : $OutputDir"
Write-Host "  Zip Output : $ZipOutput"
if ($TitleContains.Count -gt 0) { Write-Host "  Title Contains      : $($TitleContains -join ', ')" }
if ($ParticipantEmails.Count -gt 0) { Write-Host "  Participant Emails  : $($ParticipantEmails -join ', ')" }
if ($OpportunityIds.Count -gt 0)    { Write-Host "  Opportunity Ids     : $($OpportunityIds -join ', ')" }
if ($DomainMap -ne "")          { Write-Host "  Domain Map          : $DomainMap" }
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
# ─────────────────────────────────────────────────────────────────────────────

# ── 1. Page through all conversations ────────────────────────────────────────
Write-Host "Fetching conversations..." -ForegroundColor Cyan

$allConversations  = [System.Collections.Generic.List[object]]::new()
$fetchedIds        = [System.Collections.Generic.HashSet[string]]::new()
$errorCount        = 0

# Build the list of title-filter passes: one per term, or a single pass with no
# title filter when -TitleContains was not supplied.
$titlePasses = if ($TitleContains.Count -gt 0) { $TitleContains } else { @($null) }

foreach ($titleTerm in $titlePasses) {

    if ($titleTerm) {
        Write-Host "  Querying title term: '$titleTerm'" -ForegroundColor DarkCyan
    }

    $pageNumber   = 1
    $totalRecords = $null
    $totalPages   = $null

    do {
        $queryParams = [System.Collections.Generic.List[string]]::new()
        $queryParams.Add("page=$pageNumber")
        $queryParams.Add("size=$PageSize")

        if ($effectiveBeginDate -ne "") {
            $fromDateTime = "${effectiveBeginDate}T00:00:00Z"
            $queryParams.Add("fromDateTime=$([System.Uri]::EscapeDataString($fromDateTime))")
        }

        if ($titleTerm) {
            $queryParams.Add("filter[title]=$([System.Uri]::EscapeDataString($titleTerm))")
        }

        if ($ParticipantEmails.Count -gt 0) {
            $queryParams.Add("filter[participants.email]=$([System.Uri]::EscapeDataString(($ParticipantEmails -join ',')))")
        }

        if ($OpportunityIds.Count -gt 0) {
            $queryParams.Add("filter[external_opportunity.id]=$([System.Uri]::EscapeDataString(($OpportunityIds -join ',')))")
        }

        $uri = "$BaseUrl/conversations/list?$($queryParams -join '&')"

        try {
            $response = Invoke-WithRetry -MaxAttempts $MaxRetries -ScriptBlock {
                Invoke-RestMethod `
                    -Uri $uri `
                    -Headers $Headers `
                    -Method Get
            }
        }
        catch {
            $errorCount++
            Write-Host "  ERROR on page $pageNumber — $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  Aborting further pagination." -ForegroundColor DarkRed
            break
        }

        if ($null -eq $totalRecords) {
            $totalRecords = $response.meta.totalRecords
            $totalPages   = $response.meta.pageCount
            Write-Host "  Total conversations : $totalRecords" -ForegroundColor Green
            Write-Host "  Total pages         : $totalPages"  -ForegroundColor Green

            if ($totalRecords -eq 0) {
                Write-Host "  No conversations found for this term — continuing." -ForegroundColor Yellow
                break
            }
        }

        $pageItems = $response.data
        if ($null -eq $pageItems -or $pageItems.Count -eq 0) {
            Write-Host "  Page $pageNumber returned no items — stopping." -ForegroundColor DarkYellow
            break
        }

        foreach ($conv in $pageItems) {
            $cid = if ($conv.id) { $conv.id } else { $conv.attributes.uuid }
            if ($fetchedIds.Add($cid)) {
                $allConversations.Add($conv)
            }
        }

        Write-Host "  Page $pageNumber/$totalPages — $($allConversations.Count) unique conversations so far" `
            -ForegroundColor Gray

        $pageNumber++

    } while ($pageNumber -le $totalPages)
}

if ($allConversations.Count -eq 0) {
    Write-Host "`nNo conversations found for the given filters. Exiting." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nPagination complete. $($allConversations.Count) unique conversations retrieved." -ForegroundColor Green
# ─────────────────────────────────────────────────────────────────────────────

# ── 2. Write each conversation as a slim, tagged JSON file ────────────────────
Write-Host "`nWriting transcript files to $OutputDir..." -ForegroundColor Cyan

$savedCount    = 0
$saveErrors    = 0
$seenIds       = [System.Collections.Generic.HashSet[string]]::new()
$manifestRows  = [System.Collections.Generic.List[object]]::new()
$maxCreatedAt  = $null

foreach ($conv in $allConversations) {
    try {
        $attrs  = $conv.attributes
        $convId = $conv.id
        if ([string]::IsNullOrWhiteSpace($convId)) { $convId = $attrs.uuid }
        if ([string]::IsNullOrWhiteSpace($convId)) { $convId = [System.Guid]::NewGuid().ToString() }

        if (-not $seenIds.Add($convId)) {
            Write-Host "  SKIP (duplicate): $convId" -ForegroundColor DarkGray
            continue
        }

        $createdAt = $attrs.createdAt
        if ($createdAt) {
            try {
                $createdDt = [datetime]$createdAt
                if ($null -eq $maxCreatedAt -or $createdDt -gt $maxCreatedAt) {
                    $maxCreatedAt = $createdDt
                }
            } catch { }
        }

        $datePart = if ($createdAt) {
            try { ([datetime]$createdAt).ToString("yyyy-MM-dd") } catch { "unknown-date" }
        } else { "unknown-date" }

        $safeTitle = if ($attrs.title) {
            ($attrs.title -replace '[\\/:*?"<>|]', '_').Substring(0, [math]::Min($attrs.title.Length, 60))
        } else { "untitled" }

        $info = Get-CustomerProjectInfo -Attributes $attrs -InternalDomains $InternalDomains

        $flatTranscript = ConvertTo-FlatTranscript -Transcript $attrs.transcript

        # Slim record: only what a customer/project-sorting consumer needs
        $record = [PSCustomObject]@{
            id              = $convId
            title           = $attrs.title
            createdAt       = $createdAt
            transcript      = $flatTranscript
            customer        = $info.customer
            project         = $info.project
            customerSource  = $info.customerSource
            projectSource   = $info.projectSource
            externalDomains = $info.externalDomains
        }

        $safeCustomer = ($info.customer -replace '[\\/:*?"<>|]', '_').Substring(0, [math]::Min($info.customer.Length, 40))
        $fileName  = "${datePart}_${safeCustomer}_${safeTitle}_${convId}.json"
        $localPath = Join-Path $OutputDir $fileName

        $json = $record | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($localPath, $json, [System.Text.Encoding]::UTF8)

        $manifestRows.Add([PSCustomObject]@{
            id              = $convId
            filename        = $fileName
            customer        = $info.customer
            project         = $info.project
            customerSource  = $info.customerSource
            projectSource   = $info.projectSource
            createdAt       = $createdAt
            externalDomains = ($info.externalDomains -join ";")
        })

        $savedCount++
        Write-Host "  [$savedCount] $fileName" -ForegroundColor Gray
    }
    catch {
        $saveErrors++
        Write-Host "  ERROR saving conversation $($conv.id): $($_.Exception.Message)" -ForegroundColor Red
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
# actually saw a newer createdAt — otherwise leave the existing bookmark alone
# so a partial failure doesn't silently skip data on the next run.
if ($errorCount -eq 0 -and $maxCreatedAt) {
    try {
        $state = @{ lastCreatedAt = $maxCreatedAt.ToString("o") }
        [System.IO.File]::WriteAllText($StateFile, ($state | ConvertTo-Json), [System.Text.Encoding]::UTF8)
        Write-Host "Bookmark updated: $StateFile -> $($maxCreatedAt.ToString('o'))" -ForegroundColor DarkGray
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
$bySource = $manifestRows | Group-Object customerSource | Sort-Object Count -Descending

Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SUMMARY"                                  -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  API total reported   : $totalRecords"
Write-Host "  Title terms searched : $(if ($TitleContains.Count -gt 0) { "$($TitleContains.Count) ($($TitleContains -join ', '))" } else { '(none — all titles)' })"
Write-Host "  Conversations fetched: $($allConversations.Count)"
Write-Host "  JSON files saved     : $savedCount"
Write-Host "  Save errors          : $saveErrors"
Write-Host "  API/fetch errors     : $errorCount"
Write-Host "  Customer source breakdown:"
foreach ($g in $bySource) {
    Write-Host "    $($g.Name): $($g.Count)"
}
if ($savedCount -gt 0) {
    Write-Host "  Zip archive          : $ZipOutput"
}
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
