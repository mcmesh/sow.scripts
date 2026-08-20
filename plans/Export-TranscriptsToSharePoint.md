# Plan: Export-TranscriptsToSharePoint.md

## Goal

New script `Export-TranscriptsToSharePoint.ps1` that:

1. Reads Attention transcript JSON files from `C:\Temp\AttentionTranscripts`
2. Matches each transcript to a SharePoint site using the `customer` and `project`
   fields in the JSON
3. Creates a formatted `.docx` file (Open XML, no Word required) from the
   transcript content
4. Uploads the `.docx` to the `"Call Transcripts"` folder on the matched site
5. Supports `-WhatIf` to preview all decisions without writing or uploading anything

---

## Status

- [x] Plan written
- [ ] Script created
- [ ] Manually tested — WhatIf (matched)
- [ ] Manually tested — WhatIf (no match)
- [ ] Manually tested — real run (upload)
- [ ] Manually tested — folder missing + `-CreateFolderIfMissing`

---

## Source / Pattern

Auth, site search, and folder check/create logic follows `Search-SharePointSites.ps1`
exactly. New capabilities: transcript parsing, Open XML docx construction, and
Graph file upload.

---

## Parameters

| Parameter              | Type   | Default                        | Notes |
|------------------------|--------|--------------------------------|-------|
| `-TranscriptDir`       | string | `C:\Temp\AttentionTranscripts` | Source folder; files starting with `_` are skipped |
| `-FolderName`          | string | `"Call Transcripts"`           | Target folder name at root of Documents library |
| `-CreateFolderIfMissing` | switch | `$false`                    | Creates "Call Transcripts" if absent; requires `Files.ReadWrite.All` |
| `-WhatIf`              | switch | `$false`                       | Prints plan per transcript; no disk writes, no API mutations |

---

## Transcript JSON shape

Each file in `$TranscriptDir` (except `_manifest.*` / `_lastRun.json`) has:

```json
{
  "id":             "f150fdd8-...",
  "title":          "Fairbanks Morse / Mesh Systems Connect",
  "createdAt":      "2026-08-12T13:58:06.986Z",
  "transcript":     "Steagall, John:\nHey...\nBrad Natelborg:\nGood...",
  "customer":       "Fairbanks Morse",
  "project":        "MeshInsights",
  "customerSource": "domain",
  "projectSource":  "none",
  "externalDomains": ["fmdefense.com"]
}
```

Files are UTF-8 with BOM; read via `[System.IO.File]::ReadAllText()` which
handles BOM automatically.

---

## Site Matching

- Graph Search POST with `queryString = $customer` (paged, same pattern as
  `Search-SharePointSites.ps1`)
- Client-side filter: `displayName` must contain **both** `$customer` **and**
  `$project` (case-insensitive regex escape)
- Site naming convention: `FMR - Fairbanks Morse-MeshInsights-COM`
  (prefix + customer + project + suffix)

### Match outcomes

| Result     | Action |
|------------|--------|
| Exactly 1  | Proceed |
| 0 matches  | `[NO MATCH]` warn + skip |
| 2+ matches | `[AMBIGUOUS]` warn + list all + skip |

---

## Folder Check / Create

Verbatim from `Search-SharePointSites.ps1` Phases 1 & 2:

1. Resolve site ID: `GET /v1.0/sites/{hostname}:{path}`
2. Get Documents drive: `GET /v1.0/sites/{siteId}/drives`
   — prefer `name = "Documents"`, fallback first `documentLibrary`
3. Walk root children with `@odata.nextLink` paging to find `$FolderName`
4. Not found + `-CreateFolderIfMissing` → `POST` to root with `conflictBehavior: fail`
5. Not found + no flag → `[NO FOLDER]` warn + skip

---

## Docx Construction (Open XML, no Word required)

A `.docx` file is a ZIP archive. Built using `System.IO.Compression.ZipArchive`
(built into .NET / PowerShell 7 — no external modules required).

### ZIP entries

| Entry | Purpose |
|---|---|
| `[Content_Types].xml` | Declares part content types |
| `_rels/.rels` | Root relationship → document.xml |
| `word/_rels/document.xml.rels` | Document → styles.xml |
| `word/document.xml` | Body content |
| `word/styles.xml` | Heading1, Subtitle, Normal style definitions |

### Document structure

```
[Heading1]   <title>
[Subtitle]   August 12, 2026 at 1:58 PM UTC
[HRule]      ── horizontal rule paragraph ──
[Bold run]   Speaker Name:
[Normal]     Their spoken text...
[Bold run]   Next Speaker:
[Normal]     Their text...
             (repeats for each turn)
```

### Speaker label heuristic

A line is treated as a speaker label when it matches:

```
^\S[^:\n]{1,50}:\s*$
```

— starts with a non-whitespace char, contains a name-like string (max 50 chars),
ends with `:`, and has no sentence-length text on the same line. This avoids
false positives like *"here's the thing:"* mid-paragraph.

### Filename

`<title>.docx` with invalid filesystem chars (`\ / : * ? " < > |`) replaced
by `_`. Written to `$env:TEMP` before upload.

### Collision handling

If two transcripts share the same `title`, the second upload overwrites the
first in SharePoint (Graph PUT replaces on conflict). This is intentional —
no suffix appended.

---

## Upload

```
PUT /v1.0/drives/{driveId}/items/root:/{FolderName}/{encodedFilename}:/content
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
Body: file bytes (-InFile)
```

| HTTP status | Meaning | Action |
|---|---|---|
| 200/201 | Success | `[UPLOADED]` + item webUrl |
| 409 | Already exists | `[EXISTS]` — not an error |
| 403 | Permissions | `[ERROR]` — remind about `Files.ReadWrite.All` |
| Other | Unexpected | `[ERROR]` + status code |

---

## WhatIf Output (per transcript)

```
[WHATIF] Fairbanks Morse / Mesh Systems Connect
  Source  : 2026-08-12_Fairbanks Morse_..._f150fdd8.json
  Site    : FMR - Fairbanks Morse-MeshInsights-COM
  SiteUrl : https://meshsystemsus.sharepoint.com/sites/FMR-FaibanksMorse-MeshInsights-COM
  Target  : Documents/Call Transcripts/Fairbanks Morse _ Mesh Systems Connect.docx
  Folder  : [EXISTS]  (or [MISSING — would create] / [MISSING — would skip])
```

When no site match: `[WHATIF][NO MATCH] customer='...' project='...'`

---

## Summary Block

Always printed at the end:

```
── Summary ─────────────────────────────────────────────────────
  Transcripts scanned : 5
  Uploaded (or WhatIf): 3
  Already exists      : 1
  No site match       : 1
  Ambiguous match     : 0
  No folder (skipped) : 0
  Errors              : 0
────────────────────────────────────────────────────────────────
```

---

## Required Permissions

| Permission           | Purpose                        | Status          |
|----------------------|--------------------------------|-----------------|
| `Sites.Read.All`     | Site search + drive/folder read | Already granted |
| `Files.ReadWrite.All`| Folder create + file upload     | Already granted |

---

## Usage Examples

```powershell
# Dry run — see what would be uploaded and where
pwsh ./Export-TranscriptsToSharePoint.ps1 -WhatIf

# Real run — skip transcripts whose site has no "Call Transcripts" folder yet
pwsh ./Export-TranscriptsToSharePoint.ps1

# Real run — create missing folders automatically
pwsh ./Export-TranscriptsToSharePoint.ps1 -CreateFolderIfMissing

# Custom source directory
pwsh ./Export-TranscriptsToSharePoint.ps1 -TranscriptDir "D:\MyTranscripts" -WhatIf
```

---

## Files Changed

| File | Action |
|---|---|
| `Export-TranscriptsToSharePoint.ps1` | CREATE |
| `plans/Export-TranscriptsToSharePoint.md` | CREATE (this file) |
