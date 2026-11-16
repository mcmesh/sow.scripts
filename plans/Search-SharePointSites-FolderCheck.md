# Plan: Search-SharePointSites-FolderCheck.md

## Goal

Modify `Search-SharePointSites.ps1` to add two new phases after the existing site
search:

- **Phase 1** — Check whether a named folder exists at the root of the `Documents`
  library on a single specified site
- **Phase 2** — Create that folder at the same location if it doesn't exist
  (opt-in via `-CreateIfMissing` switch)

---

## Status

- [x] Script modified
- [x] Manually tested — Phase 1 (folder found)
- [x] Manually tested — Phase 1 (folder not found)
- [x] Manually tested — Phase 2 (folder created)

---

## Source / Template

`Search-SharePointSites.ps1` — all existing auth, search, and output logic is
**unchanged**. New phases are appended after the existing summary block.

---

## New Parameters

| Parameter         | Type   | Required | Default            | Notes                                                       |
|-------------------|--------|----------|--------------------|-------------------------------------------------------------|
| `-SiteUrl`        | string | no       | —                  | If omitted, script runs existing search-only behavior and exits normally |
| `-FolderName`     | string | no       | `"Call Transcripts"` | Exact display name of the folder to check/create          |
| `-CreateIfMissing` | switch | no      | `$false`           | Must be explicitly passed to trigger Phase 2; safe default |

Existing `-NameFilter` and `-MaxResults` params are **unchanged**.

---

## Phase 1 — Verify Folder Exists

Only runs when `-SiteUrl` is supplied.

### Step 1.1 — Resolve site ID from URL

```
GET /v1.0/sites/{hostname}:{serverRelativePath}
```

Parse the site URL to extract:
- `hostname` — e.g. `meshsystemsus.sharepoint.com`
- `serverRelativePath` — e.g. `/sites/BUD-ABInbev-PerfectDraft-COM`

Fail fast with a clear message on 404 (site not found) or auth error.

### Step 1.2 — Get the Documents drive

```
GET /v1.0/sites/{siteId}/drives
```

Find the drive where `name eq "Documents"` and `driveType eq "documentLibrary"`.
Fallback: use the first drive with `driveType eq "documentLibrary"` if no drive
named `"Documents"` exists. Fail with a clear message if no document library is
found at all.

### Step 1.3 — List root children and match folder by name

```
GET /v1.0/drives/{driveId}/root/children
```

Filter the response client-side for an item that:
- Has a `folder` facet (is a folder, not a file)
- `name` matches `-FolderName` (exact, case-insensitive)

Using root children listing (not `/root/search`) because:
- The folder is known to live at root level (confirmed via SharePoint UI screenshot)
- Exact root listing avoids false positives from full-text search
- Does not require `Files.Read.All` — covered by `Sites.Read.All`

### Phase 1 Output

**Found:**
```
[FOUND] "Call Transcripts" exists in the Documents library.
  Drive : Documents
  URL   : https://meshsystemsus.sharepoint.com/sites/.../Forms/AllItems.aspx
```

**Not found:**
```
[NOT FOUND] No folder named "Call Transcripts" was found at the root of Documents.
```

---

## Phase 2 — Create Folder

Only runs when **both** conditions are true:
- Phase 1 found no matching folder
- `-CreateIfMissing` switch was passed

If Phase 1 found the folder, Phase 2 is silently skipped.
If Phase 1 did not find the folder but `-CreateIfMissing` was not passed:

```
[INFO] Run with -CreateIfMissing to create the folder automatically.
```

### Step 2.1 — POST new folder to root of Documents library

```
POST /v1.0/drives/{driveId}/items/root/children
Content-Type: application/json

{
  "name": "Call Transcripts",
  "folder": {},
  "@microsoft.graph.conflictBehavior": "fail"
}
```

`conflictBehavior: "fail"` — if the folder appeared between Phase 1 and Phase 2,
Graph returns a `409 Conflict`. The script catches this and reports it as "already
exists" rather than a hard error.

### Phase 2 Output

**Created:**
```
[CREATED] Folder "Call Transcripts" successfully created.
  URL : https://meshsystemsus.sharepoint.com/sites/.../Call%20Transcripts
```

**409 Conflict (race condition):**
```
[INFO] Folder "Call Transcripts" already exists (conflict returned by API).
```

**403 Forbidden (permission not granted):**
```
[ERROR] Permission denied. Ensure Files.ReadWrite.All has been granted and admin
        consent has been applied to app registration 5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b.
```

---

## Required Permission Change

| Permission          | Purpose                   | Status                                  |
|---------------------|---------------------------|-----------------------------------------|
| `Sites.Read.All`    | Site search + Phase 1     | Already granted                         |
| `Files.ReadWrite.All` | Phase 2 folder creation | **Must be added before Phase 2 will work** |

### How to add `Files.ReadWrite.All`

1. Azure Portal → **Azure Active Directory** → **App registrations**
2. Open app `5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b`
3. **API permissions** → **Add a permission** → **Microsoft Graph**
4. Choose **Application permissions** → search `Files.ReadWrite.All` → Add
5. Click **Grant admin consent for [tenant]** — required for app-only (client credentials) flow

> Phase 1 (read check) works with existing permissions alone.
> Phase 2 will return a `403` until `Files.ReadWrite.All` is granted and consented.

---

## Script Structure After Modification

| Section | Status |
|---|---|
| 1. Param block | **Modified** — add `-SiteUrl`, `-FolderName`, `-CreateIfMissing` |
| 2. Config | Unchanged |
| 3. Auth | Unchanged |
| 4. Search loop | Unchanged |
| 5. Output table | Unchanged |
| 6. Summary | Unchanged |
| 7. Phase 1 — Resolve site + list root children + match folder | **New** |
| 8. Phase 2 — POST folder to root (guarded by `!$folderFound -and $CreateIfMissing`) | **New** |

---

## Composability

```powershell
# Search only (existing behavior, unchanged)
$sites = ./Search-SharePointSites.ps1 -NameFilter "AB InBev"

# Search + check folder on one site
./Search-SharePointSites.ps1 -NameFilter "AB InBev" -SiteUrl $sites[0].WebUrl

# Search + check + create if missing
./Search-SharePointSites.ps1 -NameFilter "AB InBev" -SiteUrl $sites[0].WebUrl -CreateIfMissing

# Check/create with a custom folder name
./Search-SharePointSites.ps1 -NameFilter "AB InBev" -SiteUrl $sites[0].WebUrl -FolderName "Meeting Notes" -CreateIfMissing

# Pass -SiteUrl directly without running a name search
./Search-SharePointSites.ps1 -NameFilter "dummy" -SiteUrl "https://meshsystemsus.sharepoint.com/sites/BUD-ABInbev-PerfectDraft-COM" -CreateIfMissing
```

---

## Known Constraints / Risks

- Root children listing (`/root/children`) returns up to 200 items per page.
  If the Documents library root has more than 200 items, pagination via
  `@odata.nextLink` is required to guarantee the folder is not missed.
  **Implementation must follow nextLink pages until exhausted.**
- `Files.ReadWrite.All` is a broad permission (grants write to all files/folders
  across all sites the app can access). If the principle of least privilege is a
  concern, `Sites.Selected` is a narrower alternative but requires per-site
  configuration via the SharePoint admin API and is out of scope here.
- `-SiteUrl` is optional so that the script remains fully backward-compatible —
  existing callers using only `-NameFilter` see no behavior change.

---

## Files Changed

| File | Action |
|---|---|
| `Search-SharePointSites.ps1` | MODIFY |
| `plans/Search-SharePointSites-FolderCheck.md` | CREATE (this file) |
