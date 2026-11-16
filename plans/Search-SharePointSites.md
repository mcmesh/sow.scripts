# Plan: Search-SharePointSites.ps1

## Goal

Create a new PowerShell script that searches for SharePoint sites by display name,
returning only sites whose name **contains** a caller-supplied filter string.

---

## Status

- [x] Script created
- [x] Manually tested against live tenant

---

## Source / Template

`SOW_Exporter.ps1` — reuse the OAuth2 client-credentials auth block and the
`$headers` / error-handling pattern verbatim.

---

## New Script: `Search-SharePointSites.ps1`

### Parameters

| Parameter       | Type   | Required | Default | Notes                                          |
|-----------------|--------|----------|---------|------------------------------------------------|
| `-NameFilter`   | string | yes      | —       | Case-insensitive substring to match site names |
| `-MaxResults`   | int    | no       | 200     | Upper bound on sites returned                  |

### Credentials / Config

Follow the env-var pattern (no hardcoded secrets):

```
$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf"   # ok to keep (non-secret)
$ClientId     = "5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b"   # ok to keep (non-secret)
$ClientSecret = $env:GRAPH_CLIENT_SECRET                  # MUST come from env var
```

Fail fast with a clear message if `GRAPH_CLIENT_SECRET` is not set (same guard
as `SOW_Exporter.ps1` lines 14-17).

### Graph API Approach

Use the **Graph Search API** (`POST /v1.0/search/query`) with `entityTypes: ["site"]`
— the same endpoint already proven in `SOW_Exporter.ps1`, just targeting a different
entity type.

```json
{
  "requests": [
    {
      "entityTypes": ["site"],
      "query": { "queryString": "<NameFilter>" },
      "fields": ["displayName", "webUrl", "createdDateTime", "description"],
      "from": 0,
      "size": 25,
      "region": "NAM"
    }
  ]
}
```

> The `queryString` is set to the raw `-NameFilter` value. Graph full-text search
> matches against the site display name (and other indexed text), so partial-name
> queries work naturally. Client-side, results are additionally filtered with
> `-match` on `displayName` to enforce the "contains" guarantee when Graph returns
> loosely related hits.

### Pagination

Reuse the `$from / $PageSize / do..while` loop from `SOW_Exporter.ps1` lines 57-106,
stopping when `$from -ge $total` **or** `$allSites.Count -ge $MaxResults`.

### Output

Pretty-print a numbered table to the console:

```
#   Display Name                     URL
--  -------------------------------- ------------------------------------------
1   Contoso SOW Archive              https://contoso.sharepoint.com/sites/SOW
2   Contoso SOW 2024                 https://contoso.sharepoint.com/sites/SOW24
```

Also return the result objects from the script (assign to a variable with
`$sites = .\Search-SharePointSites.ps1 -NameFilter "SOW"`) by writing them to
the pipeline at the end.

### Script Structure (sections)

1. **Param block** — `-NameFilter`, `-MaxResults`
2. **Config** — tenant/client IDs + env-var secret guard
3. **Auth** — identical to `SOW_Exporter.ps1` section 1
4. **Search loop** — paginate Graph Search API, client-side filter on `displayName`
5. **Output** — formatted table + pipeline objects
6. **Summary** — site count found vs. returned

---

## Known Constraints / Risks

- Graph Search API `site` entity type requires `Sites.Read.All` (application
  permission). Confirm the existing app registration has this scope before running.
- Search API region is hardcoded to `NAM`; make it a parameter if multi-geo support
  is needed later.
- `total` returned by Graph Search may be an estimate; the loop guards against
  over-fetching via `$MaxResults`.

---

## Files Changed

| File                          | Action |
|-------------------------------|--------|
| `Search-SharePointSites.ps1`  | CREATE |
| `plans/Search-SharePointSites.md` | CREATE (this file) |
