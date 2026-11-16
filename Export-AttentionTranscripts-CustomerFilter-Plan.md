# Plan: Add -CustomerName / -OpportunityName Filters + Migrate to /conversations/list

**Status:** Planned. Ready to implement.
**Script:** `Export-AttentionTranscripts.ps1`
**Date drafted:** 2026-08-12
**Last updated:** 2026-08-12 (live test results incorporated — open question resolved)

---

## Background

`Export-AttentionTranscripts.ps1` currently fetches all conversations and derives
customer and project names client-side via heuristics. There is no way to filter
by customer or opportunity/project name in a single run. This plan adds:

- `-CustomerName` — substring filter on derived `customer` field (client-side only)
- `-OpportunityName` — substring filter on derived `project` field; server-side
  pre-filter via HubSpot `deals` entity (confirmed working)
- Migration from the deprecated `/conversations` endpoint to `/conversations/list`

### Prior research (see Attention-Customer-Project-Filtering-Plan.md)

- `filter[crm_field.entity_code]` with `companies` + various field names returned
  zero results in live testing (2026-07-23).
- `filter[external_opportunity.id]` works reliably server-side.
- Customer/project names are derived client-side only — no server-side name filter
  is documented. The CRM field filter is the closest available approximation.
- The CRM filter is a three-param combo (`entity_code` + `field_name` + `values`).
  Only ONE set can be sent per request. When both `-CustomerName` and
  `-OpportunityName` are provided, the opportunity filter takes server-side
  priority (more specific). `-CustomerName` is client-side only in that case,
  but AND logic is still enforced.

### Live test results (2026-08-12) — open question resolved

Run: `Test-AttentionCrmFilters.ps1 -EntityCode "deals" -FieldNameCandidates @("dealname","name","deal_name","title") -TestValue "BUNN_Gateway Emulator"`

| FieldName | Result |
|---|---|
| `dealname` (text) | **POSSIBLE MATCH — 9 results**, first: "Bunn Weekly Status Call" |
| `name` (text) | ZERO RESULTS |
| `deal_name` (text) | ZERO RESULTS |
| `title` (text) | ZERO RESULTS |

**Conclusions:**
- `entity_code=deals` + `field_name=dealname` **works server-side** — confirmed.
- Default `OpportunityCrmEntity="deals"` and `OpportunityCrmField="dealname"` are correct.
- The `dealname` value must be the **full `externalOpportunity.title`** string for
  exact server-side match (e.g. `"BUNN_Gateway Emulator"`, not `"Gateway Emulator"`).
- `entity_code=companies` remains confirmed non-functional — **`-CustomerName` has
  no viable server-side path**. The `CustomerCrmEntity`/`CustomerCrmField` params
  are retained as overridable power-user options but the companies filter will NOT
  be injected by default. `-CustomerName` is client-side only.

### Critical implication: -OpportunityName server-side vs client-side value mismatch

The server-side `dealname` filter matches against the **full deal title**
(e.g. `"BUNN_Gateway Emulator"`), which is the unsplit `externalOpportunity.title`.

The client-side `-OpportunityName` filter matches against the **derived `project`
field** (right side of the split, e.g. `"Gateway Emulator"`).

These are different strings. The practical guidance for users:
- Pass the **full deal title** to get server-side pre-filtering benefit
  (e.g. `-OpportunityName "BUNN_Gateway Emulator"`)
- Pass a **partial project name** for client-side-only substring match
  (e.g. `-OpportunityName "Gateway"` — works, but fetches all pages first)
- `-NoCrmFilter` makes the distinction irrelevant — always client-side substring

### API endpoint change

`/conversations` is deprecated. The replacement is `/conversations/list` — identical
filter parameters and response schema, but uses a batch query path that pre-fetches
edges in a single pass, making it significantly faster for large result sets.

### HubSpot CRM integration

Tenant has HubSpot connected to Attention. Confirmed working CRM entity:
- Deals → `entity_code=deals`, `field_name=dealname` ✓ (confirmed 2026-08-12)
- Companies → `entity_code=companies`, `field_name=name` ✗ (confirmed zero results 2026-07-23)

---

## Changes (single file: Export-AttentionTranscripts.ps1)

### Change 1 — Migrate endpoint (line ~366)

**From:**
```powershell
$uri = "$BaseUrl/conversations?$($queryParams -join '&')"
```
**To:**
```powershell
$uri = "$BaseUrl/conversations/list?$($queryParams -join '&')"
```

---

### Change 2 — Add seven new parameters (after line 84)

```powershell
# Optional: filter by customer/company name (case-insensitive substring match).
# Applied client-side only against the derived customer field from Get-CustomerProjectInfo.
# NOTE: The HubSpot companies CRM filter is confirmed non-functional for this tenant —
# no server-side pre-filter is injected for -CustomerName. If both -CustomerName and
# -OpportunityName are provided, the opportunity filter handles server-side narrowing
# and client-side AND logic is enforced for both.
[string]$CustomerName          = "",

# CRM entity type for a customer server-side pre-filter. Retained as a power-user
# override but NOT injected by default (companies filter confirmed non-functional).
# Salesforce equivalent: "account".
[string]$CustomerCrmEntity     = "companies",

# CRM field name for the customer entity. Retained as override only.
# Salesforce equivalent: "Name".
[string]$CustomerCrmField      = "name",

# Optional: filter by opportunity/deal name (case-insensitive substring match).
# Applied client-side against the derived project field from Get-CustomerProjectInfo.
# When -NoCrmFilter is NOT set, also injects a server-side CRM filter using the
# deals entity (entity_code=deals, field_name=dealname — confirmed working).
# NOTE: Server-side filter requires the FULL deal title for exact match
# (e.g. "BUNN_Gateway Emulator"). Partial values still work client-side.
# Takes server-side priority over -CustomerName when both are provided.
[string]$OpportunityName       = "",

# CRM entity type for the opportunity server-side pre-filter.
# Defaults to HubSpot "deals" (confirmed working). Salesforce equivalent: "opportunity".
[string]$OpportunityCrmEntity  = "deals",

# CRM field name for the opportunity entity.
# Defaults to HubSpot "dealname" (confirmed working). Salesforce equivalent: "Name".
[string]$OpportunityCrmField   = "dealname",

# When set, skips all server-side CRM filters even if -CustomerName or -OpportunityName
# are provided. Use when conversations are not linked to CRM records, or when the
# value passed is a partial name rather than a full deal title.
# Client-side substring filtering still applies in all cases.
[switch]$NoCrmFilter
```

---

### Change 3 — Announce block (inside the header Write-Host block, ~line 330)

```powershell
if ($CustomerName -ne "") {
    Write-Host "  Customer Name       : $CustomerName (substring, client-side only)"
}
if ($OpportunityName -ne "") {
    Write-Host "  Opportunity Name    : $OpportunityName"
    if (-not $NoCrmFilter) {
        Write-Host "  CRM Pre-filter      : entity=$OpportunityCrmEntity  field=$OpportunityCrmField  (server-side, exact match)"
        Write-Host "                        NOTE: server-side matches full deal title only (e.g. 'BUNN_Gateway Emulator')"
    } else {
        Write-Host "  CRM Pre-filter      : disabled (-NoCrmFilter) — client-side substring only"
    }
}
if ($CustomerName -ne "" -and $OpportunityName -ne "") {
    Write-Host "  Filter logic        : AND (both must match)"
}
```

---

### Change 4 — Inject CRM query params (inside the `do` loop, after OpportunityIds block, ~line 364)

Only the opportunity (deals) filter is injected — companies filter is not used
(confirmed non-functional). Opportunity filter takes priority when both names
are provided.

```powershell
if (-not $NoCrmFilter -and $OpportunityName -ne "") {
    $queryParams.Add("filter[crm_field.entity_code]=$([System.Uri]::EscapeDataString($OpportunityCrmEntity))")
    $queryParams.Add("filter[crm_field.field_name]=$([System.Uri]::EscapeDataString($OpportunityCrmField))")
    $queryParams.Add("filter[crm_field.values]=$([System.Uri]::EscapeDataString($OpportunityName))")
}
```

> Server-side `dealname` filter uses exact match on the full deal title.
> Client-side pass below provides forgiving substring fallback on the derived
> project field and enforces AND logic with -CustomerName.

---

### Change 5 — Client-side post-filter in the write loop (after `$info` is resolved, ~line 454)

AND logic: a conversation must pass ALL active filters to be written.

```powershell
if ($CustomerName -ne "" -and $info.customer -notlike "*$CustomerName*") { continue }
if ($OpportunityName -ne "" -and $info.project -notlike "*$OpportunityName*") { continue }
```

---

### Change 6 — Summary section (after the `$bySource` breakdown, ~line 576)

```powershell
if ($CustomerName -ne "" -or $OpportunityName -ne "") {
    Write-Host "  Filter results:"
    if ($CustomerName -ne "") {
        Write-Host "    Customer filter    : '$CustomerName' (client-side substring)"
    }
    if ($OpportunityName -ne "") {
        Write-Host "    Opportunity filter : '$OpportunityName' (client-side substring)"
        if (-not $NoCrmFilter) {
            Write-Host "    CRM pre-filter     : entity=$OpportunityCrmEntity  field=$OpportunityCrmField  (server-side)"
        }
    }
    if ($CustomerName -ne "" -and $OpportunityName -ne "") {
        Write-Host "    Logic              : AND"
    }
    Write-Host "    Matched files      : $savedCount"
}
```

---

## Server-side filter logic (revised)

| Filters provided | Server-side CRM filter sent | Client-side filter applied |
|---|---|---|
| `-CustomerName` only | **none** (companies confirmed non-functional) | customer substring |
| `-OpportunityName` only | deals `dealname` filter (exact match on full title) | project substring |
| `-CustomerName` + `-OpportunityName` | deals `dealname` filter | customer AND project substring |
| either + `-NoCrmFilter` | none | substring(es) only |
| neither | none | none |

---

## Behavior matrix

| Invocation | Server-side | Client-side |
|---|---|---|
| `-CustomerName "Bunn"` | none | customer substring |
| `-OpportunityName "BUNN_Gateway Emulator"` | deals filter, exact title match → narrows pages | project substring |
| `-OpportunityName "Gateway"` | none (partial — auto-skips server filter) | project substring |
| `-CustomerName "Bunn" -OpportunityName "BUNN_Gateway Emulator"` | deals filter | customer AND project substring |
| `-CustomerName "Bunn" -NoCrmFilter` | none | customer substring |
| `-OpportunityName "Gateway" -NoCrmFilter` | none | project substring |
| neither set | none | none |

> **Tip:** When using `-OpportunityName` for server-side benefit, pass the full
> HubSpot deal title (e.g. `"BUNN_Gateway Emulator"`). For exploratory/partial
> matching, any substring works client-side — add `-NoCrmFilter` to make intent
> explicit and avoid a no-op server-side filter call.

---

## Usage examples (post-implementation)

```powershell
# Filter by customer name only (client-side substring)
.\Export-AttentionTranscripts.ps1 -CustomerName "Bunn"

# Filter by full deal title — server-side pre-filter narrows pages fetched
.\Export-AttentionTranscripts.ps1 -OpportunityName "BUNN_Gateway Emulator"

# Filter by partial project name — client-side only
.\Export-AttentionTranscripts.ps1 -OpportunityName "Gateway" -NoCrmFilter

# Filter by both (AND logic) — deals server-side + both client-side
.\Export-AttentionTranscripts.ps1 -CustomerName "Bunn" -OpportunityName "BUNN_Gateway Emulator"

# Combine with date filter
.\Export-AttentionTranscripts.ps1 -CustomerName "Bunn" -BeginDate "2025-01-01"
.\Export-AttentionTranscripts.ps1 -CustomerName "Bunn" -OpportunityName "BUNN_Gateway Emulator" -BeginDate "2025-01-01"

# Override HubSpot defaults with Salesforce entity names
.\Export-AttentionTranscripts.ps1 -OpportunityName "Gateway" -OpportunityCrmEntity "opportunity" -OpportunityCrmField "Name"
```

---

## Files to change

- `Export-AttentionTranscripts.ps1` — only file modified
- `Export-AttentionTranscripts-CustomerFilter-Plan.md` — this file (reference only)
