# Cancelled budget sales exclusion

## Objective
Exclude cancelled budgets from sales units and amounts in active FormListadoVentas.frm Liquidar L2 and Todos. Fix adjacent L2 running-subtotal caption arithmetic so totals equal displayed units. Keep NC sales and commission repairs unchanged.

## Constraints
Branch repair; no commit/merge/push/delete without explicit authorization. No working MDB or saved QueryDef changes. Preserve dirty accountform/MDB/context and existing user changes. ANSI/CRLF. Verified sibling backup exists. Commission manual client acceptance pending; spreadsheet LiqCom_Vend_700_sept-2026.xlsx sent by user, not verified here. Forecast below400 lines, ask-on-risk.

## Evidence
qListadoVentasPres joins headers/details by NroPresu without cancellation filter. Anulado Text:33576 NULL,329 si. NULL is active, so predicate must retain NULL. Four active L2 query branches. Cancelled screenshot match budget212944 (not212954),28September2026 product65 quantity1 amount106400. September seller screenshot product65 allclients read-only units28 amount2949940 ->27 amount2843540; matchingclient2 units212800 ->1 unit106400. Exclude source rows, never append another reversal. L2 caption incorrectly adds running Cant each detail, yielding353 instead28 in allclient baseline; change to add row quantity once. Todos already sums auxiliary product quantities.

## Tasks
- [x] T1 (implementation verified; commit deferred): Null-safe cancellation exclusion in four budget selections and narrow L2 row-quantity caption fix, with actual production query regression RED/GREEN. Delegated worker due code/tests coordinated preparation.
- [x] T2 (three independent regressions passed): Independent regressions including accepted NC sales and commission suites; byte/CRLF/source verification. Native assessment if available. Delegated verifier.
- [ ] T3 (pending user IDE case): User IDE/manual cancelled budget212944 case and totals; commit/integration pending explicit approval.

## Checks
Actual source SQL and aggregation assertions using32bitDAO/ScriptControl disposableMDB. ActiveNULL/si cancellation, duplicate/multiproduct rows, dates/seller/product/client filters, empty/cancelled-only results, L2caption/grid equality and Todos combinedproducts. Include signedNC behavior regression. No agents run actualreport on workingMDB (auxiliary writes). Report unavailable VB6compile/manualUI honestly.

## Verification and next step
Worker observed RED cancellation units expected5actual6 and caption2actual3, then GREEN. Independent verifier reran cancelledbudget, salesNC and commission suites: PASS; gitdiffcheck clean, ANSI CRLF passed. User IDE cancelled budget212944 acceptance pending; no standalone compile observed. Working .ldb appeared during verification, likely active user IDE but origin unverified; retained untouched (64bytes), no deletion. No agents opened working DB during tests; disposable fixtures only. No commit/stage/push. Next userSeptember product65 seller screenshot L2 verifies exclusiondelta1unit106400, alsoTodos. Mixed commission extension separately authorized next.


## Final Git delivery evidence
User authorized commit/push/main integration/repair deletion. Work-unit commit: d213902. All three final regressions and independent verification passed. Native review unavailable due retained selection mismatch; no lineage created. Independent verification satisfied fallback. User manual examples accepted; standalone EXE compile and print/export remain pending. Preexisting dirty MDB/account-form and user spreadsheet/context excluded and preserved.
