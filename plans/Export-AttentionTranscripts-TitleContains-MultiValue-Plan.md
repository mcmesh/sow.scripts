# Plan: Multi-value `-TitleContains` for Export-AttentionTranscripts.ps1

## Problem

`-TitleContains` accepts a single string and passes it directly to the Attention API's
`filter[title]` parameter as an exact substring match. Because the API has no OR/regex
support, searching for `"AB In"` does not return titles spelled `"ABIn"` (no space),
and vice versa. There is no server-side mechanism to express optional whitespace or
alternative spellings in a single request.

## Root Cause

Line 355 of `Export-AttentionTranscripts.ps1`:

```powershell
$queryParams.Add("filter[title]=$([System.Uri]::EscapeDataString($TitleContains))")
```

A single string is URL-encoded and sent. The API returns only records whose title
contains that exact substring.

## Options Considered

| Option | Description | Verdict |
|--------|-------------|---------|
| A | Accept `[string[]]`, run one paginated fetch per term, deduplicate | **Recommended** |
| B | Auto-generate space/no-space variants via a `-TitleContainsNormalized` switch | Nice-to-have on top of A |
| C | Fetch all, post-filter client-side | Rejected — defeats server-side filtering |

## Chosen Solution: Option A — Multi-value `-TitleContains`

### Usage

```powershell
# Single term (backward-compatible)
.\Export-AttentionTranscripts.ps1 -TitleContains "AB In"

# Multiple terms — results are unioned and deduplicated
.\Export-AttentionTranscripts.ps1 -TitleContains "AB In","ABIn"
```

### Changes Required

1. **`param` block** — change type from `[string]` to `[string[]]` and default from `""`
   to `@()`.

2. **Announce block** — join the array with `", "` for display instead of a plain string
   check.

3. **Pagination loop** — wrap in a `foreach ($term in $TitleContains)` (or run once with
   no title filter when the array is empty). Accumulate into the existing
   `$allConversations` list; skip any `id` already present in `$seenIds` to deduplicate
   across terms.

4. **Summary output** — no structural change needed; existing counters already aggregate
   across all conversations regardless of how they were fetched.

### Deduplication Strategy

`$seenIds` (a `HashSet[string]`) is already used in the file-writing phase to skip
duplicate IDs. The same set will be checked during the fetch phase: if a conversation
returned by the second (or later) title-term query has an ID already in `$seenIds`, it
is skipped before being added to `$allConversations`.

### Backward Compatibility

- Callers passing a single string (`-TitleContains "foo"`) continue to work unchanged;
  PowerShell coerces a scalar to a single-element array.
- Callers passing nothing get the existing behavior (no title filter applied).
