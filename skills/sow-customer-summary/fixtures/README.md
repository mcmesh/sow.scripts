# Synthetic SOW Fixtures

`New-SowFixtures.ps1` creates minimal DOCX and PDF inputs in a caller-provided
temporary directory. The generated documents contain fictional organizations
and no production data.

```powershell
pwsh ./New-SowFixtures.ps1 -DestinationPath ./generated-fixtures
```

The fixture sets cover:

- a draft and executed version of one SOW;
- a hash-identical executed copy;
- an executed change order;
- an unrelated customer;
- two competing executed revisions requiring a decision;
- two executed agreements with missing dates requiring an explicit order;
- a valid text PDF used to verify successful Poppler extraction;
- an invalid relevant PDF used to verify extraction-failure handling.

Generated files must not be committed. The validation script creates and
removes its own temporary fixture directory.
