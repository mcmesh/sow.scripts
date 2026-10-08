# Plan: Export-HubSpotEmails.ps1

## Overview

Mirrors the structure of `Export-AttentionTranscripts.ps1` exactly —
auth → paginate → write JSON files → write manifest → update bookmark → zip → summary.

Auth uses a HubSpot Private App token (`HUBSPOT_API_KEY` env var) sent as a `Bearer` token.

**API endpoint:** `POST https://api.hubapi.com/crm/v3/objects/emails/search`
**Required scopes:** `crm.objects.emails.read`, `sales-email-read`

---

## Parameters

| Param | Default | Notes |
|---|---|---|
| `-BeginDate` | `""` | ISO date (`YYYY-MM-DD`); overrides bookmark when supplied |
| `-EndDate` | `""` | ISO date; optional upper bound |
| `-OutputDir` | `C:\Temp\HubSpotEmails` | Where JSON files land |
| `-ZipOutput` | `C:\Temp\HubSpotEmails.zip` | Final archive |
| `-PageSize` | `50` | Max 100 per HubSpot docs |
| `-MaxRetries` | `5` | Exponential back-off cap |
| `-SearchQuery` | `""` | Optional subject substring filter |

---

## Script Structure (7 sections)

### Section 0 — Config + Validate Prerequisites

- Read `$env:HUBSPOT_API_KEY`; fail fast with clear error if missing.
- Validate `-BeginDate` / `-EndDate` format (must be `YYYY-MM-DD`) if supplied.
- Build `Authorization: Bearer $ApiKey` header.

### Helper: `Invoke-WithRetry`

Identical to template: exponential back-off (2^n seconds), fail fast on 401.

### Helper: `ConvertTo-FlatEmailBody`

HubSpot returns `hs_email_html` (HTML) and `hs_email_text` (plain text).
Prefer plain text; if absent, strip HTML tags from `hs_email_html` as fallback.
Returns a flat string.

### Section 1 — Bookmark Resolution

- Read `_lastRun.json` from `$OutputDir`.
- If `-BeginDate` is not explicitly set, use `lastCreatedAt` from bookmark as the lower bound.
- Convert effective date to Unix milliseconds for HubSpot filter format.

### Section 2 — Pagination Loop (Fetch)

Use `POST /crm/v3/objects/emails/search` with:
- `properties` array: `hs_email_subject`, `hs_email_html`, `hs_email_text`,
  `hs_email_from_email`, `hs_email_to_email`, `hs_email_cc`, `hs_email_bcc`,
  `hs_email_direction`, `hs_createdate`, `hs_timestamp`
- `associations`: `contacts`, `deals`
- `filterGroups` on `hs_createdate` (GTE lower bound, LTE upper bound if set)
- Optional `CONTAINS_TOKEN` filter on `hs_email_subject` when `-SearchQuery` is supplied
- Paginate via `paging.next.after` cursor (HubSpot cursor-based, not page-number)
- Deduplication by email `id` using a `HashSet`

> **10,000-result cap:** The HubSpot CRM Search API caps results at 10,000 per query.
> If the total exceeds this, the script automatically chunks by calendar month,
> runs a separate query per chunk, and union-deduplicates — same pattern as
> `SOW_docx_Exporter.ps1`.

### Section 3 — File Writing

One JSON file per email. Filename pattern:

```
{YYYY-MM-DD}_{sanitized-subject}_{id}.json
```

Record shape saved to disk:

```json
{
  "id": "...",
  "subject": "...",
  "body": "...",
  "from": "...",
  "to": "...",
  "cc": "...",
  "bcc": "...",
  "direction": "INBOUND|OUTBOUND",
  "createdAt": "...",
  "associations": {
    "contactIds": ["..."],
    "dealIds":    ["..."]
  }
}
```

### Section 4 — Manifest

Write `_manifest.json` and `_manifest.csv` with columns:
`id`, `filename`, `subject`, `from`, `direction`, `createdAt`

### Section 5 — Bookmark + Zip

- Advance `_lastRun.json` bookmark only on a clean run (zero fetch errors).
- Zip `$OutputDir\*` to `$ZipOutput`.

### Section 6 — Summary

Print:
- API total fetched / files saved / errors
- Direction breakdown (inbound vs outbound counts)
- Zip path

---

## Key Differences from Attention Script

| | `Export-AttentionTranscripts.ps1` | `Export-HubSpotEmails.ps1` |
|---|---|---|
| Auth header | `Bearer $ApiKey` | `Bearer $ApiKey` (identical) |
| HTTP method | `GET` | `POST` (Search API) |
| Pagination | Page-number (`page=N`) | Cursor (`paging.next.after`) |
| Date filtering | Query param `fromDateTime` | `filterGroups` on `hs_createdate` |
| Body parsing | `ConvertTo-FlatTranscript` | `ConvertTo-FlatEmailBody` |
| Customer metadata | 3-source heuristic | Not needed — raw fields only |
| Result cap | None apparent | 10,000/query — auto-chunk by month |

---

## Files to Create

| File | Purpose |
|---|---|
| `Export-HubSpotEmails.ps1` | The exporter script |
| `Export-HubSpotEmails-Plan.md` | This document |
