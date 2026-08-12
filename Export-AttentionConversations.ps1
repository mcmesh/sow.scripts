# ══════════════════════════════════════════════════════════════════════════════
#  Export-AttentionConversations.ps1
#
#  Pulls all call conversations from the Attention API (attention.tech) and
#  saves each conversation as a JSON file, then zips the output directory.
#
#  Auth:  https://docs.attention.com/api-authentication
#  Route: https://docs.attention.com/api-reference/conversation/list-conversations
#
#  Usage examples:
#    # Export everything
#    .\Export-AttentionConversations.ps1
#
#    # Export from a specific date onward
#    .\Export-AttentionConversations.ps1 -BeginDate "2025-01-01"
#
#    # Override output paths and page size
#    .\Export-AttentionConversations.ps1 -BeginDate "2024-06-01" `
#        -OutputDir "D:\AttentionExport" -ZipOutput "D:\AttentionExport.zip"
# ══════════════════════════════════════════════════════════════════════════════

param(
    # Optional: only return conversations on or after this date (ISO 8601 date, e.g. "2025-01-01")
    [string]$BeginDate   = "",

    # Where individual JSON files are written before zipping
    [string]$OutputDir   = "C:\Temp\AttentionConversations",

    # Final zip archive path
    [string]$ZipOutput   = "C:\Temp\AttentionConversations.zip",

    # Items per page — Attention API default is 20; keep ≤100 to be safe
    [int]$PageSize        = 50,

    # Max retry attempts on transient failures (exponential back-off)
    [int]$MaxRetries      = 5
)

# ── Config ────────────────────────────────────────────────────────────────────
# Store your key in an environment variable and reference it here, e.g.:
#   $env:ATTENTION_API_KEY = "your_key_here"   (set once in your session)
# Or replace the line below with a literal string for ad-hoc use — but never
# commit a real key to source control.
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
#
#  The Attention API sometimes returns transcript as:
#    (a) A plain string  — "Jerry Jackson:\nHey, Will…"          (Postman default)
#    (b) An array of     — [{ speaker:{name}, words:[{text}…] }] (detailedTranscript)
#
#  This function always produces the (a) format so output is consistent
#  regardless of which shape the API decided to return on a given call.
#
function ConvertTo-FlatTranscript {
    param([object]$Transcript)

    # Already a plain string — nothing to do
    if ($Transcript -is [string]) {
        return $Transcript
    }

    # Null / empty — return empty string
    if ($null -eq $Transcript) {
        return ""
    }

    # Array of speaker-turn objects  [{ speaker: {name}, words: [{text}] }]
    if ($Transcript -is [System.Collections.IEnumerable]) {
        $lines = [System.Text.StringBuilder]::new()
        foreach ($turn in $Transcript) {
            $speakerName = $turn.speaker.name
            if ([string]::IsNullOrWhiteSpace($speakerName)) {
                $speakerName = $turn.speaker.email   # fall back to email if name missing
            }

            # Concatenate every word's text, preserving the trailing space each
            # word already carries (e.g. "Hey, " "Will, " "you " "there? ")
            $sentence = ($turn.words | ForEach-Object { $_.text }) -join ""
            $sentence = $sentence.Trim()

            if (-not [string]::IsNullOrWhiteSpace($sentence)) {
                if ($lines.Length -gt 0) { [void]$lines.Append("`n") }
                [void]$lines.Append("${speakerName}:`n${sentence}")
            }
        }
        return $lines.ToString()
    }

    # Unknown shape — serialise as-is so we don't silently lose data
    return ($Transcript | ConvertTo-Json -Depth 10 -Compress)
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Build shared auth headers ─────────────────────────────────────────────────
$Headers = @{
    "Authorization" = "Bearer $ApiKey"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}
# ─────────────────────────────────────────────────────────────────────────────

# ── Announce run parameters ───────────────────────────────────────────────────
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Attention Conversation Exporter"          -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  API Base   : $BaseUrl"
Write-Host "  Begin Date : $(if ($BeginDate) { $BeginDate } else { '(all time)' })"
Write-Host "  Page Size  : $PageSize"
Write-Host "  Output Dir : $OutputDir"
Write-Host "  Zip Output : $ZipOutput"
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
# ─────────────────────────────────────────────────────────────────────────────

# ── 1. Create output directory ────────────────────────────────────────────────
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
# ─────────────────────────────────────────────────────────────────────────────

# ── 2. Page through all conversations ────────────────────────────────────────
Write-Host "Fetching conversations..." -ForegroundColor Cyan

$allConversations = [System.Collections.Generic.List[object]]::new()
$pageNumber       = 1
$totalRecords     = $null
$totalPages       = $null
$errorCount       = 0

do {
    # Build query string parameters
    $queryParams = [System.Collections.Generic.List[string]]::new()
    $queryParams.Add("page=$pageNumber")
    $queryParams.Add("size=$PageSize")

    # Only add date filter when a BeginDate was supplied
    if ($BeginDate -ne "") {
        # API expects ISO 8601 — append midnight UTC if no time provided
        $fromDateTime = "${BeginDate}T00:00:00Z"
        $queryParams.Add("fromDateTime=$([System.Uri]::EscapeDataString($fromDateTime))")
    }

    $uri = "$BaseUrl/conversations?$($queryParams -join '&')"

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

    # Capture totals from the first page's meta block
    if ($null -eq $totalRecords) {
        $totalRecords = $response.meta.totalRecords
        $totalPages   = $response.meta.pageCount
        Write-Host "  Total conversations : $totalRecords" -ForegroundColor Green
        Write-Host "  Total pages         : $totalPages"  -ForegroundColor Green

        if ($totalRecords -eq 0) {
            Write-Host "`nNo conversations found for the given filters. Exiting." -ForegroundColor Yellow
            exit 0
        }
    }

    $pageItems = $response.data
    if ($null -eq $pageItems -or $pageItems.Count -eq 0) {
        Write-Host "  Page $pageNumber returned no items — stopping." -ForegroundColor DarkYellow
        break
    }

    foreach ($conv in $pageItems) {
        $allConversations.Add($conv)
    }

    Write-Host "  Page $pageNumber/$totalPages — retrieved $($allConversations.Count) of $totalRecords conversations" `
        -ForegroundColor Gray

    $pageNumber++

} while ($pageNumber -le $totalPages)

Write-Host "`nPagination complete. $($allConversations.Count) conversations retrieved." -ForegroundColor Green
# ─────────────────────────────────────────────────────────────────────────────

# ── 3. Write each conversation to a JSON file ─────────────────────────────────
Write-Host "`nWriting JSON files to $OutputDir..." -ForegroundColor Cyan

$savedCount   = 0
$saveErrors   = 0
$seenIds      = [System.Collections.Generic.HashSet[string]]::new()

foreach ($conv in $allConversations) {
    try {
        # Use the conversation id (or uuid from attributes as fallback)
        $convId = $conv.id
        if ([string]::IsNullOrWhiteSpace($convId)) {
            $convId = $conv.attributes.uuid
        }
        if ([string]::IsNullOrWhiteSpace($convId)) {
            $convId = [System.Guid]::NewGuid().ToString()
        }

        # Deduplicate (API should not return dupes, but guard anyway)
        if (-not $seenIds.Add($convId)) {
            Write-Host "  SKIP (duplicate): $convId" -ForegroundColor DarkGray
            continue
        }

        # Build a safe filename: "YYYY-MM-DD_<uuid>.json"
        $createdAt  = $conv.attributes.createdAt
        $datePart   = if ($createdAt) {
            try { ([datetime]$createdAt).ToString("yyyy-MM-dd") } catch { "unknown-date" }
        } else { "unknown-date" }

        $safeTitle  = if ($conv.attributes.title) {
            # Strip characters that are invalid in Windows filenames
            ($conv.attributes.title -replace '[\\/:*?"<>|]', '_').Substring(
                0, [math]::Min($conv.attributes.title.Length, 60)
            )
        } else { "untitled" }

        $fileName  = "${datePart}_${safeTitle}_${convId}.json"
        $localPath = Join-Path $OutputDir $fileName

        # ── Normalize transcript to flat string regardless of API response shape ──
        # The API intermittently returns transcript as a word-level array even
        # when detailedTranscript=false.  We always flatten it here so the saved
        # JSON matches the plain-string format returned by Postman.
        if ($null -ne $conv.attributes.transcript) {
            $flatTranscript = ConvertTo-FlatTranscript -Transcript $conv.attributes.transcript
            # PSCustomObject properties are read-only after deserialization, so
            # we must re-attach via Add-Member (Force overwrites if already set)
            $conv.attributes | Add-Member -MemberType NoteProperty `
                -Name "transcript" -Value $flatTranscript -Force
        }
        # ──────────────────────────────────────────────────────────────────────

        # Serialise the full conversation object to pretty-printed JSON.
        # NOTE: We use [System.IO.File]::WriteAllText instead of Set-Content
        # because Set-Content -Encoding UTF8 throws "parameter cannot be found
        # that matches parameter name 'Encoding'" on Windows PowerShell 5.x
        # when the serialized string exceeds ~32 KB (large transcripts hit this).
        # The .NET method is version-agnostic and handles any string size cleanly.
        $json = $conv | ConvertTo-Json -Depth 20
        [System.IO.File]::WriteAllText($localPath, $json, [System.Text.Encoding]::UTF8)

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

# ── 4. Also write a single combined manifest JSON ─────────────────────────────
try {
    $manifestPath = Join-Path $OutputDir "_manifest.json"
    $manifest = @{
        exportedAt      = (Get-Date -Format "o")
        beginDateFilter = if ($BeginDate) { $BeginDate } else { $null }
        totalRecords    = $totalRecords
        retrievedCount  = $allConversations.Count
        savedCount      = $savedCount
    }
    $manifestJson = $manifest | ConvertTo-Json
    [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.Encoding]::UTF8)
    Write-Host "Manifest written: $manifestPath" -ForegroundColor DarkGray
}
catch {
    Write-Host "WARNING: Could not write manifest — $($_.Exception.Message)" -ForegroundColor DarkYellow
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
Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SUMMARY"                                  -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  API total reported : $totalRecords"
Write-Host "  Conversations fetched : $($allConversations.Count)"
Write-Host "  JSON files saved   : $savedCount"
Write-Host "  Save errors        : $saveErrors"
Write-Host "  API/fetch errors   : $errorCount"
if ($savedCount -gt 0) {
    Write-Host "  Zip archive        : $ZipOutput"
}
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""