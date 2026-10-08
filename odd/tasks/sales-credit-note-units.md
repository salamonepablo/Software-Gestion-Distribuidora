# Sales credit-note units repair

## Objective
Subtract NC quantities and their recorded monetary values from the sales report. Preserve invoice calculations, commissions and cancelled-budget behavior. User explicitly authorized monetary netting after the first quantity-only manual check.

## Constraints
- Branch: repair; base cf7b361. No commit, merge, push or branch deletion until explicitly authorized. User will approve integration into main after functional acceptance.
- Preserve pre-existing MDB and FormMovimentosCuentaCorriente.frm changes and SPC_LEGACY_CONTEXT.md.
- Full verified backup: sibling SPC-Core-resguardo-20261008-192007.
- Preserve ANSI encoding and CRLF in VB6.
- No writes to working MDB; combined report mutates auxiliary tables and must only be run by user or against a disposable database.
- Delivery strategy: single focused repair, ask-on-risk; forecast below 400 authored lines.

## Evidence
FormListadoVentas.frm Liquidar uses qListadoVentasFac (invoice-only composite join) in L1 and combined mode. NC details store positive quantities. NC money has mixed signs, so sign must come from document type, not amount.
User fixture: September 2026, seller GABRIEL PERALTA, customer 34, all products, Todos. Current 45FKA=2,65=4,75=2,ON75=2,total=10. NC A-741 dated 07/09/2026 subtracts one of each; expected 1,3,1,1,total=6.
User manual check confirmed quantities 1/3/1/1, total 6. Invoice amounts remained 938999.98 in quantity-only implementation. Expanded acceptance: subtract NC A-741 amount 362600.00, expected report amount 576399.98, with proportional recorded detail amounts and applicable IVA; do not reprice products. Verify rounding and signed detail treatment against existing NC storage.

## Tasks
- [x] T1 (implementation verified; commit deferred to T3 pending authorization): Extend implemented negative NC quantities with negative recorded NC monetary values for L1 and combined; preserve filters and invoice arithmetic. Route: delegated writer; preparation and coordinated code/test trigger. Observe failing deterministic fixture then passing repair, or report unavailable runner.
- [x] T2 (verified; native assessment unavailable): Independently verify behavior, filters, joins, ANSI/CRLF and preservation of existing changes. Route: delegated verifier. Native assessment/review if available; record unavailable checks honestly.
- [x] T3 (manual case accepted; commit/integration deferred): User VB6 IDE compile and manual acceptance of September case. Commit/integration pending explicit authorization; do not mark complete before acceptance.

## Verification and progress
Investigation used 32-bit PowerShell DAO 3.6 read-only. Native gentle-ai executable unavailable through PATH. Quantity implementation added SQLVentasConNC plus tools/tests/Test-SalesCreditNoteUnits.ps1. Worker observed RED then GREEN using actual helper through ScriptControl and disposable DAO MDB, with filters/composite joins/encoding checks. User screenshot independently confirms total 6, but monetary extension is now required. No commit made. Commissions and budget cancellation deferred.

## Next step
User screenshot and acceptance confirm September fixture: quantities 1/3/1/1 total6 and amount576399.98. Commit authorization and integration remain pending. Next business unit: clarify which NC qualify as seller commission payments before implementation.

## Manual acceptance
User confirmed report screenshot after running in IDE: 45FKA quantity1 amount93100.00; 65 quantity3 amount279299.99; 75 quantity1 amount112000.00; ON75 quantity1 amount91999.99; total6 and576399.98. This proves the reported UI case, not a standalone EXE compilation or all manual edge cases. No commit or merge performed.

## Monetary extension verification
Worker observed monetary RED (expected0 actual121) then GREEN. Production helper uses -ND.TotalLinea and NC.PorcentajeIVA, without Abs or repricing. Read-only real case net576399.9802 rounds576399.98. Independent verifier reran 32bit PowerShell regression and git diff --check: PASS; ANSI/CRLF and unchanged account-movement form versus backup confirmed. Synthetic fixture also rounds576399.98. Native ASSESS unassessable due intended-untracked declaration, so independent verification performed; no native review closure claimed. VB6 compilation/UI acceptance pending. No commits/staging/push. Form diff +22/-2; tests added. Task work-unit commit closure remains deferred until explicit authorization.


## Final Git delivery evidence
User authorized commit/push/main integration/repair deletion. Work-unit commit: d213902. All three final regressions and independent verification passed. Native review unavailable due retained selection mismatch; no lineage created. Independent verification satisfied fallback. User manual examples accepted; standalone EXE compile and print/export remain pending. Preexisting dirty MDB/account-form and user spreadsheet/context excluded and preserved.
