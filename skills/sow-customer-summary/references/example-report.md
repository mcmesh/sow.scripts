<!--
STYLE REFERENCE ONLY. Fictional customer, documents, IDs, amounts, and paths.
Copy the structure, tone, citation density, and caveat handling, never the facts.
-->

# Customer Engagement Summary: Northwind Sensors

| Report detail | Value |
|---|---|
| Customer | Northwind Sensors, Inc. ("Northwind") [doc-a1b2c3d4e5f6, page 1] |
| Report generated | 2026-01-15 |
| Agreements with execution evidence summarized | 1 [doc-a1b2c3d4e5f6, page 9] |
| Unverified agreements included | 2 without full execution verification: Firmware Hardening (helper: partially signed) and Certification Support (helper: unverified) [doc-b2c3d4e5f6a1, page 6; doc-c3d4e5f6a1b2, page 4] |
| SOW effective-date range | 2023-04-10 to 2025-06-02; the effective date is not explicitly stated for the later agreements [doc-a1b2c3d4e5f6, page 1; doc-c3d4e5f6a1b2, page 4] |
| Engagement period | Initial work due by 2023-10-31; later scopes documented for 2024-2025; no single continuous service period established [doc-a1b2c3d4e5f6, page 7; doc-b2c3d4e5f6a1, page 3] |
| Coverage | 38 PDF/DOCX files enumerated; 3 agreement packages included; 1 helper group excluded (unrelated customer); 0 unresolved groups or relevant extraction failures. |

## Customer Executive Summary

Northwind is described in the SOW as a manufacturer of industrial environmental sensors sold to facility operators. In 2023 it sought a cloud-connected monitoring platform to replace manual readings and shorten time to market for a new sensor line. [doc-a1b2c3d4e5f6, pages 2-3]

The initial contracted work covered a prototype sensor gateway, cloud ingestion, and an operator dashboard. A later SOW addressed firmware hardening and field-update capability ahead of production. [doc-a1b2c3d4e5f6, pages 4-6; doc-b2c3d4e5f6a1, pages 2-3]

A 2025 change order added regulatory certification support. Together, the documents describe a path from prototype to production readiness and market qualification. They do not prove that every planned deliverable was completed. [doc-c3d4e5f6a1b2, pages 1-2]

## Contracted Work

### Initial Statement of Work: Northwind Gateway Prototype

| SOW detail | Value |
|---|---|
| Effective date | 2023-04-10 [doc-a1b2c3d4e5f6, page 1] |
| Project or engagement | Northwind Gateway Prototype; quotation 2023-0042 [doc-a1b2c3d4e5f6, page 1] |
| Selected source | `/data/SOWFiles/2023-0042 Gateway Prototype SoW Northwind - signed.pdf` |
| Signature status | Execution evidence detected: both parties e-signed and the audit trail states "Agreement completed" [doc-a1b2c3d4e5f6, page 9] |
| Selection confidence / method | High / Auto selected |

Northwind wanted a standardized sensor platform and a faster route to commercialization. The SOW planned two parallel workstreams: select a radio technology for the sensor network, and build a gateway-to-cloud prototype. The radio evaluation compared range, power draw, implementation effort, and unit cost. [doc-a1b2c3d4e5f6, pages 3-4]

The prototype scope included gateway hardware design, power-optimized firmware, cloud ingestion and storage configuration, and a dashboard with telemetry export to Northwind's own cloud subscription. Production hardware was left to later work after technology selection. [doc-a1b2c3d4e5f6, pages 5-6]

The fixed fee was USD 120,000, and the SOW required completion by 2023-10-31. These are contracted terms, not confirmation of actual spend or delivery. [doc-a1b2c3d4e5f6, page 7]

**Key deliverables**

- Radio evaluation report, comparison matrix, and a jointly selected approach. [doc-a1b2c3d4e5f6, page 5]
- Three prototype gateways and design artifacts. [doc-a1b2c3d4e5f6, page 6]
- Operator dashboard, telemetry export, and two months of post-installation support. [doc-a1b2c3d4e5f6, page 6]

**Dependencies, assumptions, and constraints**

- Northwind was to provide sensor documentation, sample devices, and engineering contacts; component availability could affect the schedule. [doc-a1b2c3d4e5f6, page 7]
- Regulatory certification was excluded from this scope. [doc-a1b2c3d4e5f6, page 7]

## Additional Work

### Firmware Hardening and Field Update - March 2024

The SOW cites a successful prototype and proposes bringing the firmware to production quality. Planned work included secure boot, signed over-the-air updates, a fault-injection test plan, and integration testing on pilot hardware. Deliverables included hardened firmware, a test report, and release documentation. [doc-b2c3d4e5f6a1, pages 2-4]

The scope was priced at USD 85,000 over an estimated 16 weeks. Northwind approvals and a named product owner were dependencies, and mobile-app work was excluded. This builds on the prototype's gateway design but is a separate agreement under the same master agreement. [doc-b2c3d4e5f6a1, pages 4-5; doc-a1b2c3d4e5f6, page 8]

### Certification Support Change Order - June 2025

This change order references the Firmware Hardening SOW. It covers test-plan review, documentation support, and coordination with an accredited lab for regulatory certification of the gateway. Mesh was to prepare and advise, and Northwind was to submit applications directly. [doc-c3d4e5f6a1b2, pages 1-2]

The fee was USD 6,500, excluding lab and filing fees. The document describes planned support and does not establish that certification was achieved. [doc-c3d4e5f6a1b2, pages 2-3]

## Sources

| Sequence / source ID | Agreement | Source path | Effective date | Signature status | Confidence / method | Selection evidence / locator |
|---|---|---|---|---|---|---|
| 1 / doc-a1b2c3d4e5f6 | Gateway Prototype SOW | `/data/SOWFiles/2023-0042 Gateway Prototype SoW Northwind - signed.pdf` | 2023-04-10 | Execution evidence detected | High / Auto selected | Parties and quotation, p. 1; signing audit and completion, p. 9 |
| 2 / doc-b2c3d4e5f6a1 | Firmware Hardening and Field Update SOW | `/data/SOWFiles/2024-0117 Firmware Hardening SoW Northwind v2 [Signed].pdf` | Document dated 2024-03-18; effective date not explicitly stated | Partially signed; full execution unverified | Medium / User selected | Northwind signature block, p. 6; version history, p. 2 |
| 3 / doc-c3d4e5f6a1b2 | Certification Support Change Order | `/data/SOWFiles/2025-0231 Certification Support CO Northwind SIGNED.pdf` | Document dated 2025-06-02; effective date not explicitly stated | Unverified | Low / User selected | Customer and parent SOW, p. 1; names and dates, p. 4 |

## Review Notes

- **Additional source paths retained for the selected Gateway Prototype document:** `/data/SOWFiles/2023-0042 Gateway Prototype SoW Northwind - signed (1).pdf`. It shares the selected document's hash in the helper handoff. [doc-a1b2c3d4e5f6, source-path metadata]
- **Execution caveat:** The later two agreements remain execution-unverified. Printed names, dates, and "signed" filenames alone are not conclusive. Inclusion was approved by the user for customer research. [doc-b2c3d4e5f6a1, page 6; doc-c3d4e5f6a1b2, page 4]
- **Version discrepancy:** The Firmware Hardening filename and version history indicate v2, while the body heading says v1. Both identifiers are retained rather than assuming either is authoritative. [doc-b2c3d4e5f6a1, filename and page 2]
- **Requirements conflict:** Section 3 allows a 15-second update window, while the appendix retains a 10-second requirement. This needs clarification rather than silent normalization. [doc-b2c3d4e5f6a1, pages 3, 8]
- **Reported past results versus future commitments:** The overview reports prototype update success in lab conditions, while the appendix says field reliability goals were not yet met. Both statements are preserved, and neither establishes completed production validation. [doc-b2c3d4e5f6a1, pages 2, 8]
- **Excluded match:** Helper group `grp-0123456789ab` was excluded because its contracting customer is a different organization. It contributes no Northwind claims. [excluded source, page 1]
- **Financial interpretation:** Listed prices and fees describe the selected scopes only. Total customer revenue, actual payments, complete contract history, and completion of later deliverables are not specified in the selected SOWs.
