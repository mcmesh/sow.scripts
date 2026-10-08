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
