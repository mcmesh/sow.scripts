# Plan: Customer/Project-Aware Attention Transcript Export

**Status as of 2026-07-23 (Thu).** Investigation and API verification complete.
Next step is implementation (not yet started).

## Goal

Build a new script (working name: `Export-AttentionTranscripts.ps1`) that behaves
like `Export-AttentionConversations.ps1` but outputs slim, transcript-only files
tagged with best-guess **customer** and **project** metadata, so a downstream
consumer (human or script) can sort/move the files into per-customer/per-project
folders. This new script does **not** do the sorting/moving itself.

The script must also be safely re-runnable over time on a rolling basis — e.g.
weekly/monthly — pulling only conversations created since the last run, rather
than re-fetching the full history every time.

## Research summary (see full detail in conversation history if needed)

Read `Export-AttentionConversations.ps1`, inspected real sample data in
`C:\Temp\AttentionConversations` (799 files), and reviewed the Attention.tech API
docs (`docs.attention.com/api-reference/conversation/*`).

**Data shape reality** (confirmed against 799 real exported conversations):
- No `customer` or `project` field exists anywhere in the API payload.
- `externalOpportunity.title` (free-text CRM deal name) present on 47.4% of
  conversations; 95% of those follow a `Customer_ProjectName` convention
  (e.g. `BUNN_Gateway Emulator`, `RenewAire_Connected_ERV`). ~5% don't
  (human-entered CRM data, e.g. `Schamberger Andreas, FG-7-I`).
- `externalAccounts[].id` populated on 61.5%, but it's a bare CRM company ID
  with no name attached anywhere in the payload — confirmed even with
  `withCrmRecords=true` (see Verified Item 1 below).
- Attendee email domain (excluding `meshsystems.com`) is the most complete
  signal: 48.4% of all conversations have exactly one external domain; only
  3.6% mix multiple customer domains. 48% have zero external attendees
  (internal-only syncs — no customer signal possible for these).

## Verified items (via `Test-AttentionCrmFilters.ps1`, run against live tenant)

This diagnostic script is checked in and re-runnable — see its header comment
for usage. Results from the live run on 2026-07-23:

1. **`withCrmRecords=true` does NOT expose human-readable CRM names.**
   Confirmed live: `linkedCrmRecords`, `externalAccounts`, `externalContacts`
   all still show bare IDs only (e.g. `{"code":"companies","id":"18717127232"}`)
   even with `withCrmRecords=true` on the working `/conversations/list` route.
   `extractedIntelligence`/`confirmedExtractedIntelligence` are AI-derived
   summary fields (e.g. "Customer Objective", "Funding Status"), not raw CRM
   record data. The only human-readable signal remains
   `externalOpportunity.title`.

2. **`filter[crm_field.entity_code/field_name/values]` does NOT narrow results**
   for `entity_code=companies`. All 20 tested combinations (10 field-name
   guesses: `id`, `name`, `company_name`, `companyName`, `account_name`,
   `accountName`, `dealname`, `deal_name`, `title`, `domain` — each tried
   against both free-text and known-real-ID values) returned valid-syntax
   but zero matches, with no errors. Likely cause: this workspace's CRM sync
   doesn't populate queryable `org_crm_records` field data for `companies`
   (not just a wrong-field-name-guess problem), since a real ID-based filter
   (`filter[external_opportunity.id]`) works cleanly against the same data.
   **Not yet tried:** `entity_code=deals` with field names like `dealname`/
   `name` — worth a quick try before fully closing this door, but not
   blocking, since a working alternative (below) already exists.

3. **`filter[external_opportunity.id]` DOES work** — confirmed live,
   returned exactly the expected 9 conversations for a known deal ID. This is
   a reliable, documented, ID-based server-side filter. Useful when the
   consumer already knows a specific opportunity ID, but doesn't solve
   name-based lookup (no documented endpoint resolves a customer/deal *name*
   to an ID).

## Decided approach

No server-side "filter by customer/project name" exists. Customer/project
metadata must be derived **client-side**, after fetching, using a
priority-ordered heuristic per conversation:

1. **Primary**: split `externalOpportunity.title` on its delimiter (`_` or
   ` - `) → left = customer, right = project. (~47% coverage, ~95% accurate
   within that.)
2. **Fallback**: external attendee domain (excluding `meshsystems.com`) →
   friendly name (e.g. `bunn.com` → `Bunn`), optionally overridable via a
   small user-maintained domain→name lookup file. (~48% coverage, catches
   many conversations with no opportunity link.)
3. **Last resort**: raw conversation `title` string, flagged low-confidence,
   for manual sorting.
4. Tag each record with `customerSource`/`projectSource`
   (`opportunity`/`domain`/`title`/`none`) so the consumer knows how much to
   trust it.
5. ~48% of conversations have zero external attendees (internal-only) — these
   will land in an "unknown" bucket regardless of method. This is a hard
   floor, not a bug to fix.

## Script design (not yet implemented)

- Reuse `Invoke-WithRetry` and `ConvertTo-FlatTranscript` from
  `Export-AttentionConversations.ps1` — same auth/pagination/retry scaffolding.
- Per-conversation output slimmed to: `id`, `title`, `createdAt`, `transcript`
  (flattened), plus derived `customer`, `project`, `customerSource`,
  `projectSource`, `externalDomains`. Drop `attendees` bulk, CRM intelligence
  blobs, Zoom/import metadata, scorecards, etc.
- Emit `_manifest.json`/`.csv`: one row per conversation with `id`, `filename`,
  `customer`, `project`, `customerSource`, `projectSource`, `createdAt`,
  `externalDomains` — this is what a downstream script/human uses to sort/move
  files. **This script does not move or organize files itself.**
- New optional params:
  - `-BeginDate` → same mechanism as `Export-AttentionConversations.ps1`
    (passthrough to the documented `fromDateTime` filter). Excludes anything
    created before this date — the primary way to run "from a start date
    forward" and skip older files. Already a proven-working filter (used by
    the existing script today), so no further verification needed.
  - `-TitleContains` → passthrough to documented `filter[title]`
  - `-ParticipantEmails` → passthrough to documented `filter[participants.email]`
  - `-OpportunityIds` → passthrough to confirmed-working
    `filter[external_opportunity.id]`, for targeted re-fetch of a known deal
  - `-DomainMap` → path to a JSON file mapping domain → friendly customer name
    (overrides the default "strip TLD" guess)
- **Incremental/repeatable runs**: after a successful run, write the max
  `createdAt` seen to a small state file (e.g. `_lastRun.json` in
  `-OutputDir`). On the next run, if `-BeginDate` isn't explicitly passed, it
  defaults to that bookmarked date automatically — so the script can simply be
  re-run periodically (manually or via a scheduled task) and will naturally
  fetch only conversations created since the previous run, without the caller
  having to track or pass a date each time. `-BeginDate` still overrides the
  bookmark when explicitly supplied (e.g. for a one-off backfill).

## Next steps

1. (Optional, low priority) Quick probe of `entity_code=deals` field-name
   guesses using `Test-AttentionCrmFilters.ps1` — only worth doing if curious;
   doesn't block implementation since the heuristic approach + confirmed
   `filter[external_opportunity.id]` are sufficient to proceed.
2. Implement `Export-AttentionTranscripts.ps1` per the design above.
3. Test against the live tenant, spot-check the customer/project heuristic
   against a sample of real conversations (compare to the 799-file stats
   already gathered).
