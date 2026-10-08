# SOW Selection Rules

## Customer Match

Match supplied names and aliases against contracting-party text first, then
document title, filename, folder, and general text. A conflicting named party
overrides filename/folder hints. Never infer subsidiaries or aliases.

## Grouping

Group revisions using agreement ID, parties, title, dates, normalized filename,
and scope evidence. Hash-identical files are copies with multiple source paths.
Keep amendments and change orders as separate groups linked to a parent. Do not
merge documents solely because their names are similar.

## Signature Status

| Status | Required evidence |
|---|---|
| Execution evidence detected | Explicit fully executed/completion evidence tied to the agreement; no draft/unsigned contradiction |
| Partially signed | Some completed signer evidence, but execution by all required parties is not established |
| Unsigned | Explicit unsigned status, blank signature evidence, or draft-only document |
| Unverified | Insufficient evidence for another status |

Names, dates, signature images, `final`, and `signed` filenames are supporting
evidence only. Internal approval is not execution. User selection changes the
selection method, never the signature status.

## Automatic Selection

Auto-select only a unique high-confidence candidate with execution evidence and
no material conflict. Modified time is only a tie-breaker for equivalent copies.
Return ambiguity when two executed revisions differ, chronology conflicts, or
evidence is incomplete.

## Coverage

Block by default when a relevant file fails extraction or a group is unresolved.
The user may choose a candidate, skip a group, or exclude a failed file. Any skip
requires incomplete coverage and must appear in Review Notes. Zero included
agreements always fails.
