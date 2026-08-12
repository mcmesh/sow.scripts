# ══════════════════════════════════════════════════════════════════════════════
#  Test-AttentionCrmFilters.ps1
#
#  DIAGNOSTIC / EXPLORATORY SCRIPT — not a data exporter.
#
#  Answers two open questions about the Attention.tech API before we build a
#  customer/project-aware export script:
#
#    1) Does `withCrmRecords=true` return human-readable CRM company/deal
#       NAMES, or only bare IDs (as the plain conversation payload does)?
#
#    2) Does `filter[crm_field.entity_code]` + `filter[crm_field.field_name]`
#       + `filter[crm_field.values]` actually narrow results server-side for
#       any guessed field name (the docs don't say what field names your org's
#       CRM sync uses)?
#
#  A sanity-check (Test 0) using the DOCUMENTED `filter[participants.email]`
#  array filter runs first, so if every crm_field guess in Test 2 "fails" you
#  can tell whether that's because the field-name guesses are wrong, or
#  because the request-building mechanism itself is broken.
#
#  Test 3 sanity-checks a DIFFERENT documented CRM filter —
#  `filter[external_opportunity.id]` — against a known-real opportunity ID, so
#  you have at least one confirmed-working ID-based CRM filter even if every
#  crm_field.field_name guess in Test 2 comes up empty.
#
#  This script only reads data and prints to the console — it does not write
#  any files.
#
#  Usage examples:
#    # Run with defaults (uses known-good IDs pulled from your own sample export)
#    .\Test-AttentionCrmFilters.ps1
#
#    # Point Test 1/3 at a different known opportunity you have data for
#    .\Test-AttentionCrmFilters.ps1 -KnownOpportunityId "<some-other-opportunity-id>"
#
#    # Try your own field-name guesses / test value
#    .\Test-AttentionCrmFilters.ps1 -FieldNameCandidates @("name","company") -TestValue "Bunn"
# ══════════════════════════════════════════════════════════════════════════════

param(
    # A known attendee email from a real conversation — used only to prove the
    # documented array-filter mechanism works (Test 0 sanity check).
    [string]$KnownParticipantEmail = "josh.craig@bunn.com",

    # A known real CRM company id (externalAccounts[].id) from your sample
    # export (Bunn) — used as a second, ID-shaped guess for crm_field.values in
    # Test 2, alongside the free-text -TestValue.
    [string]$KnownCompanyId = "18717127232",

    # A known real CRM opportunity/deal id (externalOpportunity.uuid) from your
    # sample export (Bunn: "BUNN_Gateway Emulator") — used in Test 3 to sanity-
    # check the documented filter[external_opportunity.id] filter.
    [string]$KnownOpportunityId = "55649281914",

    # CRM entity to test filtering against (see docs enum: companies, contacts,
    # deals, owners, pipelines, users, account, contact, lead, opportunity, user)
    [string]$EntityCode = "companies",

    # Field-name guesses to try for filter[crm_field.field_name]. The docs say
    # this comes from your org's org_crm_records.record_data shape, which isn't
    # documented — so we brute-force common HubSpot/Salesforce-style names.
    [string[]]$FieldNameCandidates = @(
        "id", "name", "company_name", "companyName", "account_name", "accountName",
        "dealname", "deal_name", "title", "domain"
    ),

    # Free-text value to filter for — pick a customer name you know appears in
    # your data (e.g. one of the domains/customers seen in the sample export:
    # Bunn, GOJO, RenewAire, Eaton, Gill Group, ...). Each field-name candidate
    # in Test 2 is tried against BOTH this value and -KnownCompanyId, since
    # exact-match semantics mean an ID-shaped field needs an ID-shaped value.
    [string]$TestValue = "Bunn",

    # Page size for the filter-guess calls — kept small since we only care
    # about totalRecords / a sanity peek at the first result.
    [int]$Size = 5
)

# ── Config ────────────────────────────────────────────────────────────────────
$ApiKey  = $env:ATTENTION_API_KEY
$BaseUrl = "https://api.attention.tech/v2"
# ─────────────────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Host "ERROR: No API key found." -ForegroundColor Red
    Write-Host "       Set the environment variable ATTENTION_API_KEY before running:" -ForegroundColor Yellow
    Write-Host '       $env:ATTENTION_API_KEY = "your_key_here"' -ForegroundColor Yellow
    exit 1
}

$Headers = @{
    "Authorization" = "Bearer $ApiKey"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}

# ── Helper: call the API, tolerate 4xx/5xx, and surface the response body ────
function Invoke-AttentionApi {
    param([string]$Uri)

    try {
        $resp = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get
        return @{ Ok = $true; Body = $resp; StatusCode = 200; Error = $null }
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        # PowerShell 7's Invoke-RestMethod populates ErrorDetails.Message with
        # the raw response body on non-2xx responses — this is where the API's
        # actual error text (often naming valid fields) will show up.
        $errorBody = $_.ErrorDetails.Message
        if (-not $errorBody) { $errorBody = $_.Exception.Message }

        return @{ Ok = $false; Body = $null; StatusCode = $statusCode; Error = $errorBody }
    }
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
}

$summary = [System.Collections.Generic.List[object]]::new()

# ── Test 0: baseline count + sanity-check the documented array filter ────────
Write-Section "TEST 0 — Baseline count + sanity-check filter[participants.email]"

$baselineResult = Invoke-AttentionApi -Uri "$BaseUrl/conversations?page=1&size=1"
if (-not $baselineResult.Ok) {
    Write-Host "  ERROR fetching baseline count: $($baselineResult.Error)" -ForegroundColor Red
    exit 1
}
$baselineTotal = $baselineResult.Body.meta.totalRecords
Write-Host "  Baseline total records (no filter): $baselineTotal" -ForegroundColor Green

$sanityUri = "$BaseUrl/conversations?page=1&size=1&filter[participants.email]=$([System.Uri]::EscapeDataString($KnownParticipantEmail))"
$sanityResult = Invoke-AttentionApi -Uri $sanityUri

if ($sanityResult.Ok) {
    $sanityTotal = $sanityResult.Body.meta.totalRecords
    Write-Host "  filter[participants.email]=$KnownParticipantEmail -> totalRecords: $sanityTotal" -ForegroundColor Green

    if ($sanityTotal -lt $baselineTotal -and $sanityTotal -gt 0) {
        Write-Host "  MECHANISM CHECK: PASS — array filter narrowed results as expected." -ForegroundColor Green
        $mechanismWorks = $true
    }
    else {
        Write-Host "  MECHANISM CHECK: SUSPECT — expected a smaller-than-baseline, non-zero count." -ForegroundColor Yellow
        Write-Host "  (Check that -KnownParticipantEmail is a real attendee email in your data.)" -ForegroundColor Yellow
        $mechanismWorks = $false
    }
}
else {
    Write-Host "  ERROR: $($sanityResult.Error)" -ForegroundColor Red
    Write-Host "  MECHANISM CHECK: FAIL — could not validate array-filter syntax at all." -ForegroundColor Red
    $mechanismWorks = $false
}

# ── Test 1: does withCrmRecords=true expose real names, or just IDs? ─────────
Write-Section "TEST 1 — withCrmRecords=true: names or just IDs?"
Write-Host "  NOTE: GET /conversations/{id} does NOT document withCrmRecords as a supported" -ForegroundColor DarkGray
Write-Host "  param (only GET /conversations/list does), and every conversation id we tried" -ForegroundColor DarkGray
Write-Host "  against the single-record endpoint 404'd regardless of source (archived sample" -ForegroundColor DarkGray
Write-Host "  id, fresh live-list id). So this test now uses the proven-working /conversations/list" -ForegroundColor DarkGray
Write-Host "  route, filtered to -KnownOpportunityId (which Test 3 already confirmed returns data)." -ForegroundColor DarkGray
Write-Host ""

$crmUri = "$BaseUrl/conversations/list?page=1&size=1&withCrmRecords=true" +
    "&filter[external_opportunity.id]=$([System.Uri]::EscapeDataString($KnownOpportunityId))"
$crmResult = Invoke-AttentionApi -Uri $crmUri

if (-not $crmResult.Ok) {
    Write-Host "  ERROR calling /conversations/list?withCrmRecords=true: $($crmResult.Error)" -ForegroundColor Red
    Write-Host "  (If this 404s/400s too, /conversations/list may not exist at this API version.)" -ForegroundColor Yellow
}
else {
    if ($crmResult.Body.data.Count -eq 0) {
        Write-Host "  No results for -KnownOpportunityId=$KnownOpportunityId — pick a different known opportunity id." -ForegroundColor Yellow
    }
    else {
        $attrs = $crmResult.Body.data[0].attributes

        $fieldsToShow = @(
            "linkedCrmRecords", "externalAccounts", "externalContacts",
            "externalOpportunity", "dealUUID", "externalIntelligenceOpportunityID"
        )

        foreach ($f in $fieldsToShow) {
            $val = $attrs.$f
            Write-Host "  ${f}:" -ForegroundColor Gray
            Write-Host "    $($val | ConvertTo-Json -Depth 6 -Compress)"
        }

        if ($attrs.extractedIntelligence -or $attrs.confirmedExtractedIntelligence) {
            Write-Host "  extractedIntelligence / confirmedExtractedIntelligence: PRESENT" -ForegroundColor Gray
            Write-Host "    (these carry AI-derived CRM field values, not raw CRM record data — see entries above)"
        }
        else {
            Write-Host "  extractedIntelligence / confirmedExtractedIntelligence: absent on this record" -ForegroundColor DarkGray
        }

        # Heuristic: flag any value under externalAccounts/externalContacts/linkedCrmRecords
        # that is NOT a bare numeric/UUID id — that would indicate a real name leaked through.
        $idLike = { param($s) $s -match '^\d+$' -or $s -match '^[0-9a-fA-F-]{8,}$' }
        $sawNameLike = $false
        foreach ($f in @("externalAccounts", "externalContacts", "linkedCrmRecords")) {
            foreach ($item in @($attrs.$f)) {
                if ($item -and $item.id -and -not (& $idLike $item.id)) {
                    $sawNameLike = $true
                    Write-Host "  Possible name-like value found in $f -> $($item.id)" -ForegroundColor Green
                }
            }
        }

        Write-Host ""
        if ($sawNameLike) {
            Write-Host "  CONCLUSION: withCrmRecords=true DOES appear to expose human-readable names." -ForegroundColor Green
        }
        else {
            Write-Host "  CONCLUSION: withCrmRecords=true still only exposes bare IDs for accounts/contacts/deals" -ForegroundColor Yellow
            Write-Host "              on this record. Only externalOpportunity.title (a free-text CRM deal name)" -ForegroundColor Yellow
            Write-Host "              is human-readable. Re-run against a few more -KnownOpportunityId values" -ForegroundColor Yellow
            Write-Host "              before ruling this out entirely." -ForegroundColor Yellow
        }
    }
}

# ── Test 2: does filter[crm_field.*] actually narrow results server-side? ────
Write-Section "TEST 2 — filter[crm_field.entity_code/field_name/values] guesses"
Write-Host "  entity_code      : $EntityCode"
Write-Host "  test value (text): $TestValue"
Write-Host "  test value (id)  : $KnownCompanyId  (real externalAccounts id from your sample export)"
Write-Host "  baseline         : $baselineTotal total records (unfiltered)"
Write-Host ""

# Try each field-name guess against BOTH a free-text value and a known-real ID
# value — exact-match semantics mean an id-shaped field will never match text
# like "Bunn", and a name-shaped field will never match a raw numeric id.
$testValues = @(
    @{ Label = "text"; Value = $TestValue },
    @{ Label = "id";   Value = $KnownCompanyId }
)

foreach ($candidate in $FieldNameCandidates) {
    foreach ($tv in $testValues) {
        $uri = "$BaseUrl/conversations" +
            "?page=1&size=$Size" +
            "&filter[crm_field.entity_code]=$([System.Uri]::EscapeDataString($EntityCode))" +
            "&filter[crm_field.field_name]=$([System.Uri]::EscapeDataString($candidate))" +
            "&filter[crm_field.values]=$([System.Uri]::EscapeDataString($tv.Value))"

        $result = Invoke-AttentionApi -Uri $uri
        $label = "$candidate (" + $tv.Label + ")"

        if (-not $result.Ok) {
            Write-Host "  [$label] ERROR ($($result.StatusCode)): $($result.Error)" -ForegroundColor Red
            $summary.Add([PSCustomObject]@{ FieldName = $label; Status = "ERROR"; TotalRecords = $null; Note = $result.Error })
            continue
        }

        $total = $result.Body.meta.totalRecords
        $firstTitle = if ($result.Body.data.Count -gt 0) { $result.Body.data[0].attributes.title } else { "(none)" }

        if ($total -eq $baselineTotal) {
            Write-Host "  [$label] totalRecords=$total  -> SAME AS BASELINE (filter likely ignored / no-op)" -ForegroundColor Yellow
            $summary.Add([PSCustomObject]@{ FieldName = $label; Status = "NO-OP"; TotalRecords = $total; Note = "matches baseline" })
        }
        elseif ($total -gt 0) {
            Write-Host "  [$label] totalRecords=$total  -> NARROWED. First result: '$firstTitle'" -ForegroundColor Green
            $summary.Add([PSCustomObject]@{ FieldName = $label; Status = "POSSIBLE MATCH"; TotalRecords = $total; Note = "first: $firstTitle" })
        }
        else {
            Write-Host "  [$label] totalRecords=0  -> valid syntax but no matches" -ForegroundColor DarkYellow
            $summary.Add([PSCustomObject]@{ FieldName = $label; Status = "ZERO RESULTS"; TotalRecords = 0; Note = "" })
        }
    }
}

# ── Test 3: sanity-check filter[external_opportunity.id] against a KNOWN id ──
Write-Section "TEST 3 — filter[external_opportunity.id] with a known-real id"
Write-Host "  This is a DOCUMENTED filter (unlike crm_field.field_name, which is a guess)." -ForegroundColor DarkGray
Write-Host "  Using -KnownOpportunityId=$KnownOpportunityId (from your sample export: 'BUNN_Gateway Emulator')" -ForegroundColor DarkGray
Write-Host ""

$oppUri = "$BaseUrl/conversations?page=1&size=$Size&filter[external_opportunity.id]=$([System.Uri]::EscapeDataString($KnownOpportunityId))"
$oppResult = Invoke-AttentionApi -Uri $oppUri

if (-not $oppResult.Ok) {
    Write-Host "  ERROR ($($oppResult.StatusCode)): $($oppResult.Error)" -ForegroundColor Red
    $opportunityFilterWorks = $false
}
else {
    $oppTotal = $oppResult.Body.meta.totalRecords
    $oppFirstTitle = if ($oppResult.Body.data.Count -gt 0) { $oppResult.Body.data[0].attributes.title } else { "(none)" }
    Write-Host "  totalRecords=$oppTotal, first result title: '$oppFirstTitle'" -ForegroundColor Green

    if ($oppTotal -gt 0 -and $oppTotal -lt $baselineTotal) {
        Write-Host "  CONCLUSION: filter[external_opportunity.id] WORKS — a known-good, ID-based CRM filter." -ForegroundColor Green
        Write-Host "              This is a reliable fallback even if Test 2's field-name guesses never pan out," -ForegroundColor Green
        Write-Host "              PROVIDED you have some other way to resolve customer/project name -> opportunity id." -ForegroundColor Green
        $opportunityFilterWorks = $true
    }
    else {
        Write-Host "  CONCLUSION: unexpected result — expected a small non-zero count well under baseline." -ForegroundColor Yellow
        $opportunityFilterWorks = $false
    }
}

# ── Summary ────────────────────────────────────────────────────────────────
Write-Section "SUMMARY"
Write-Host "  Test 0 (filter mechanism sanity check)         : $(if ($mechanismWorks) { 'PASS' } else { 'SUSPECT/FAIL' })"
Write-Host "  Test 3 (filter[external_opportunity.id] known) : $(if ($opportunityFilterWorks) { 'PASS' } else { 'SUSPECT/FAIL' })"
Write-Host ""
$summary | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "Interpreting results:" -ForegroundColor Cyan
Write-Host "  - If Test 0 FAILED, ignore Test 2 entirely — fix -KnownParticipantEmail first."
Write-Host "  - 'POSSIBLE MATCH' rows are your best candidates for a real crm_field.field_name."
Write-Host "  - 'ZERO RESULTS' means the API accepted the field name but nothing matched that value"
Write-Host "    (try a different -TestValue/-KnownCompanyId, or the field name may still be wrong)."
Write-Host "  - 'ERROR' rows often include a response body naming the actual valid field — read the Note column."
Write-Host "  - If Test 3 PASSED but Test 2 has no matches, the practical path forward is resolving"
Write-Host "    customer/project name -> CRM id some other way (e.g. a small lookup file you maintain),"
Write-Host "    then filtering by filter[external_opportunity.id] rather than by CRM field name/value."
Write-Host ""
