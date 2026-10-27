# Helper Handoff Contract

The helper writes schema `1.0` JSON to stdout and diagnostics to stderr.

| Exit | JSON status | Meaning |
|---|---|---|
| `0` | `ready` | Inputs are selected and coverage is resolved |
| `2` | `needs_input` | Missing value, ambiguity, overwrite, or coverage decision |
| `1` | `failed` | Fatal validation/extraction failure |

Selected documents include stable `documentId` and `groupId`, source paths,
hash, agreement type/parent, normalized and original dates, signature status,
selection confidence/method/evidence, extraction metadata, and located text
sections. Coverage includes enumerated, selected, excluded, failed, and
unresolved counts.

## Update Mode

When the output file exists and neither `-Force` nor `overwriteOutput` is set,
the helper runs in `update` mode. It never modifies the report; it reads the
report's Sources table and Review Notes, then classifies the current folder
contents against it.

Additional top-level fields (schema stays `1.0`; they are additive):

- `mode`: `create`, `update`, or `replace`.
- Each selected document gains `updateStatus` (`unchanged`, `new`, or
  `supersedes`), `supersedesDocumentId`, and `reportSequence` (the document's
  row number in the existing Sources table). Fields are `null` outside update
  mode. `unchanged` documents have `selectionMethod` `Previously reported` and
  an empty `sections` array because their report text is preserved.
- `update.existingReport`: path, sha256, `Report generated` value, and number
  of Sources rows parsed.
- `update.backupPath`: suggested backup path; the agent creates the copy.
- `update.hasChanges`: true when there are new documents, superseding
  documents, or new source paths for already-reported documents.
- `update.newDocumentIds`, `supersedingDocumentIds`, `unchangedDocumentIds`.
- `update.unchangedWithNewSourcePaths`: reported documents with new
  hash-identical copies. A path counts as known when it appears in backticks
  anywhere in the existing report.
- `update.missingFromFolder`: reported documents no longer found (or whose
  file content changed). They stay in the report and a warning is emitted.
- `update.carriedForward`: group and file IDs the existing report's Review Notes
  records as skipped or excluded. They are not asked about again unless a
  decision overrides them.

A group with a previously reported document keeps that document unless another
candidate in the group has equal or higher signature status (and is at least
`Unverified`). Then the group is unresolved and the user chooses: select the
previous document to keep it, or a new one to supersede it. A new group is
selected normally. Older and newer files are treated the same way.

Chronology is requested in update mode only when a new or superseding document
has a missing or duplicate effective date. The decision must then list every
selected group; prefill previously reported groups from `reportSequence`.

An existing file without a recognizable Sources table returns exit `2` so the
user can pick another path or authorize replacement.

## Decision File

Pass decisions back with `-DecisionFile <path>`:

```json
{
  "schemaVersion": "1.0",
  "allowIncompleteCoverage": true,
  "overwriteOutput": false,
  "groupDecisions": [
    { "groupId": "grp-123", "action": "select", "documentId": "doc-456" },
    { "groupId": "grp-789", "action": "skip", "reason": "Not part of this engagement" }
  ],
  "fileDecisions": [
    { "candidateId": "file-123", "action": "exclude", "reason": "Corrupt duplicate confirmed by user" }
  ],
  "chronology": [
    { "groupId": "grp-123", "order": 1 }
  ]
}
```

Actions are lowercase. A `select` decision requires a document from that group.
Every `skip` or `exclude` requires a reason and incomplete coverage permission.
The helper rejects unknown or stale IDs.
