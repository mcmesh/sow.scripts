# CLAUDE.md

**Core context**: Standalone PowerShell scripts for exporting M365/SaaS tenant data (Attention.tech conversations, SharePoint SOWs, Teams transcripts, pricing sheets). No build system, tests, or package manager — each `.ps1` runs directly via `pwsh ./script.ps1 [args]`.

## Running scripts

All scripts require PowerShell 7+ (some explicitly via `#Requires -Version 7.0`). Invoke directly:

```powershell
pwsh ./Export-AttentionConversations.ps1 -BeginDate "2025-01-01"
pwsh ./MC-Export-TeamsTranscripts_TEST.ps1 -DownloadTranscripts
```

Scripts fetch credentials from **environment variables only** — do not hardcode secrets:
- `Export-AttentionConversations.ps1`, `MC-Export-TeamsTranscripts_TEST.ps1` use `$env:ATTENTION_API_KEY` / `$env:GRAPH_CLIENT_SECRET`
- Set them before running: `$env:GRAPH_CLIENT_SECRET = "..."`

**Pattern violation risk**: SOW and pricing exporters (`SOW_docx_Exporter.ps1`, `SOW_xlsx_Exporter.ps1`, `Test-Access.ps1`) contain **hardcoded tenant and client IDs in plaintext**. When editing these files, ask before attempting to migrate to env-var pattern; do not perpetuate the bad pattern.

- **Attention.tech scripts** — Bearer token via `ATTENTION_API_KEY` env var, REST calls
  to `https://api.attention.tech/v2`.

- **Microsoft Graph / SharePoint scripts** — Azure AD app-only auth (OAuth2 client
  credentials flow) against tenant `<TENANT_ID>`, then calls
  to `https://graph.microsoft.com/v1.0` (search, drive, onlineMeetings APIs).

## Script quirks

- **`SOW_docx_Exporter.ps1`** — Graph Search API becomes unreliable past ~1000 results per query. Script splits search into yearly `$DateChunks` (2010–2026) and refreshes auth token every 500 downloads. **Known issue**: dead code tail after the chunked version's summary — references undefined `$SearchQuery` variable. Treat top half (chunked logic) as the real implementation.

- **`Export-AttentionConversations.ps1`** — Attention API returns transcript field inconsistently (flat string or word-level array). Helper `ConvertTo-FlatTranscript` always coerces to flat-string shape. Retries transient failures with exponential backoff; fails fast on 401 (auth).

- **`MC-Export-TeamsTranscripts_TEST.ps1`** — Primary path fetches via transcript content endpoint. Falls back to searching organizer's OneDrive `Documents/Recordings` by matching `.vtt` file timestamps when Graph error 3004 (expired resource) is returned or speaker attribution is disabled. Requires `Files.Read.All` permission for fallback.

- **`Test-Access.ps1`** — Diagnostic only; not a complete runnable script (references `$headers` before it's defined). Decodes Graph JWT claims and checks SharePoint data-residency region.

## Architecture pattern

All exporters follow: **authenticate (OAuth2 or Bearer token) → paginate through API with `@odata.nextLink` → write JSON/files to disk → zip output directory → print summary**.

Output defaults (override with script parameters):
- Attention: `C:\Temp\AttentionConversations\` + `C:\Temp\AttentionConversations.zip`
- Teams transcripts: `.\TranscriptExport\`
- SOW docs/spreadsheets: `C:\Temp\SOWFiles_*\` + zip

## Validation

No lint, test, or build commands. Verify changes by running the script against the live API/tenant.

## References

- `script-inventory.md` — Full script inventory, credential handling, detailed known issues
- `versions/` — Historical script revisions (not active)
- `export_test_result` — Sample API response shape for reference
- For Attention.tech API documentation: `https://docs.attention.com/api-reference/conversation/list-conversations-optimized`
