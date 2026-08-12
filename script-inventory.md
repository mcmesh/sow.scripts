### Script inventory

- `Export-AttentionConversations.ps1` — pages through all Attention conversations,
  normalizes the transcript field (the API inconsistently returns either a flat string
  or a word-level array — `ConvertTo-FlatTranscript` always coerces to the flat-string
  shape), writes one JSON file per conversation plus a `_manifest.json`, then zips.
  Retries transient failures with exponential backoff; fails fast on 401.
- `SOW_Exporter.ps1` — searches SharePoint (via Graph `/search/query`) for signed SOW
  files (docx/pdf) and downloads them using the base64 "share URL" driveItem pattern
  (`u!<base64(webUrl)>`), then zips.
- `SOW_docx_Exporter.ps1` — same idea as `SOW_Exporter.ps1` but for `*SOW*` docx/doc
  files. Splits the search into yearly date chunks (`$DateChunks`) because the Graph
  Search API becomes unreliable past ~1000 results in a single query, and refreshes
  the auth token every 500 downloads. **Known issue:** the file has a second, older
  copy of the single-query (non-chunked) script logic appended after the chunked
  version's summary block; that trailing block references `$SearchQuery`, which is
  never defined in this file (only `$BaseQuery` is), so it will error if reached —
  treat the chunked version (top of the file) as the real logic and the tail as dead
  code to clean up if touching this file.
- `SOW_xlsx_Exporter.ps1` — same single-query search/download/zip pattern, targeting
  `*Pricing*` spreadsheets (xlsx/xls).
- `Test-Access.ps1` — ad hoc diagnostic snippet for decoding a Graph token's JWT claims
  (aud/appid/tid) and checking a SharePoint site's data-residency region. Not a
  complete runnable script as-is (references `$headers` before it's ever defined).
- `MC-Export-TeamsTranscripts_TEST.ps1` — exports Teams meeting transcripts for a given
  organizer via `getAllTranscripts`. Primary path hits the transcript content endpoint;
  falls back to searching the organizer's OneDrive `Documents/Recordings` folder by
  matching `.vtt` file timestamps when the primary endpoint returns an expired-resource
  error (Graph error code 3004) or when speaker attribution is disabled.
- `export_test_result` — not a script; a saved snippet of captured PowerShell output
  from a prior `getAllTranscripts` run, kept for reference on the API's response shape
  (including `@odata.nextLink` pagination tokens).
- `versions/` — earlier/alternate revisions of the scripts above, kept for reference.
  Not invoked by anything; treat as historical, not part of the active toolset.

## Credential handling — follow the existing good pattern, not the bad one

`Export-AttentionConversations.ps1` and `MC-Export-TeamsTranscripts_TEST.ps1` read
secrets from environment variables (`ATTENTION_API_KEY`, `GRAPH_CLIENT_SECRET`) and
fail with a clear error if unset — this is the pattern to follow for any new or edited
script. The SOW/pricing exporter scripts and `Test-Access.ps1` instead have an Azure AD
tenant ID, client ID, and **client secret hardcoded in plaintext**. Do not add new
hardcoded secrets, and if you're touching one of these files, prefer migrating it to
the env-var pattern over leaving the literal secret in place.
