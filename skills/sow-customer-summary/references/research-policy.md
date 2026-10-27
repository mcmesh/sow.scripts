# Approved Customer Research Policy

The user approved these defaults on 2026-10-26, including the decisions in their
attached screenshot. Apply them to customer-research summaries unless the user
requests a different scope; do not ask for the same policy approval on every run.

## Inclusion and Identity

- Include relevant execution-unverified SOWs with explicit signature-status
  caveats. Inclusion does not upgrade execution status or establish delivery.
- Preserve materially different revisions as history. Do not assume the newest
  date, version, or `signed` filename supersedes an earlier agreement. Explain
  differences and uncertainty; do not add overlapping prices as separate revenue.
- Organize Microsoft-funded work under the identified end customer, while
  separately naming Microsoft as the contracting/funding party. Cite evidence for
  each role. If the end customer is not established, flag it rather than guess.
- Exclude supplier-to-Mesh agreements from customer engagement reports. Record
  their paths and exclusion reasons in the batch coverage record.
- Keep distinct legal entities separate, with sourced cross-references. A parent,
  subsidiary, brand, or acquisition is not automatically a legal alias.
- Choose the most common unambiguous customer name, counting distinct documents
  rather than duplicate exports. Retain the legal name and supported variants.
- Search the internet for possible aliases and former names; cite URLs and access
  dates. Use document party/scope evidence to confirm matches. Project names
  (such as Yoda) are search terms, not automatically customer names. If online
  research fails, disclose the limitation and proceed with supported identities.

## Evidence and Coverage

- Inventory every PDF/DOCX and search document text as well as filenames. Track
  byte-identical copies separately from text-equivalent copies; retain paths.
- Recover image-only pages with local OCR or visual inspection, keeping page
  locators and extraction limitations. Do not send private SOWs to public OCR.
  Block the affected customer on unreadable relevant inputs unless the user
  explicitly approves incomplete coverage; continue independent customers.
- Keep scope, deliverables, prices, dates, dependencies, and historical evidence
  sourced. Distinguish proposed work from reported past results and verified
  completion. Leave unsupported fields unspecified.
- Record unresolved customer/contract relationships in a batch exceptions list;
  avoid silently losing them or assigning by filename alone.

## Helper Decisions and Limitations

The helper remains conservative: `Auto` requires unique execution evidence and
selects only one candidate per revision group. These defaults guide the agent;
they are not new helper flags or a change to signature classification.

Record policy-authorized inclusions using `groupDecisions` selections, naming
the policy and preserving signature status in Review Notes. Document skips and
exclusions with reasons; enable incomplete coverage only for explicit exclusions,
not to suppress unreadable evidence. Obtain genuine user decisions for remaining
material ambiguities.

To retain multiple revisions that the helper groups together, run each revision
in an isolated working source folder with its copies and combine the sourced
handoffs, recording original paths and grouping rationale. Do not select just
one revision and claim that the others were reviewed or superseded.

## Existing Summaries

- Find existing reports by customer identity, not just exact filename. Preserve
  existing prose, manual edits, document IDs, and useful research notes.
- Back up each report before any update or authorized rename. Prefer adding only
  new evidence and source paths; avoid duplicate appended summaries.
- Preserve prior revisions and their source rows even when supersession is
  explicitly documented. Do not rewrite history to show only the replacement.
- If multiple reports exist for one customer, use the most complete report as the
  update base; retain other files and add nonredundant material with provenance.
  Do not delete or overwrite conflicting user edits without clarification.
- Emit a batch index mapping canonical customer names, aliases, output paths,
  related legal entities, exclusions, and unresolved coverage. Keep this metadata
  separate from customer deliverable claims.
