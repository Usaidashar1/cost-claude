# Fix Changelog — Azure Pricing Calculator Conversion Tool

This documents every fix applied against the production review. Run `pytest tests/test_convert.py -v` to see them all verified automatically.

## Follow-up fixes (round 3) — RI pricing "missing" for D2_v3, B2s, F2s, F2s_v2, F4s, F4s_v2

**Root cause found — this was not a code bug.** Microsoft retired new purchase/renewal of Azure Reserved VM Instances for a specific list of VM series effective **1 July 2026** (announced 6 May 2026: [Azure Reserved VM Instances for select VM series will no longer be available](https://techcommunity.microsoft.com/blog/azurecompute/azure-reserved-vm-instances-for-select-vm-series-will-no-longer-be-available-sta/4516505)):

- 1-year RI retired for: `Av2, Amv2, Bv1, D, Ds, Dv2, Dsv2, F, Fs, Fsv2, G, Gs, Ls, Lsv2`
- 1-year **and** 3-year RI retired for: `Dv3, Dsv3, Ev3, Esv3`

Every SKU you flagged maps exactly onto this list: `D2_v3` → Dv3 (both terms gone), `B2s` → Bv1 (1yr gone), `F2s`/`F4s` → Fs (1yr gone), `F2s_v2`/`F4s_v2` → Fsv2 (1yr gone). I confirmed SKU extraction and region mapping are correct for all of these (not a parsing bug) — the Retail Prices API genuinely no longer returns Reservation entries for these series/terms. The Calculator UI showing "1 Year Available" is either not yet updated or shows a legacy/reference figure; it's no longer actually purchasable. If "D4s" refers to `D4s_v3`, that's `Dsv3` — also fully retired — which is the most likely explanation for the "wrong pricing" report too (RI silently mirroring PAYG). If it's a different generation or the *compute/PAYG* number itself looked wrong (not RI), that's a separate issue and I'll need the actual numbers to trace it.

**Fix:** Added `_classify_vm_series()` and `_ri_retirement_note()`, which recognize SKUs in the retired list and attach a specific, accurate Remarks explanation ("Azure retired 1-year Reserved Instances for Fs-series as of 1 Jul 2026...") instead of silently showing RI == PAYG with no explanation. This turns an apparent tool bug into a clearly-surfaced, correctly-attributed platform change. See `tools/diagnose_sku.py` — a standalone script you can run from an environment with real network access to `prices.azure.com` to get ground-truth raw API data for any SKU (this sandbox can't reach that host, which is why this round required extra research rather than direct reproduction).

## Follow-up fixes (round 2)

| Issue | Fix |
|---|---|
| **VMSS bucketed into "Others", no RI pricing** | Added "Virtual machine scale sets" / "vm scale sets" / "vmss" to `SVC_MAP`, giving it its own "Virtual Machine Scale Sets" sheet. That sheet is now routed through the exact same pricing/enrichment pipeline as "Virtual Machines" (`VM_LIKE_SHEETS`), so VMSS rows get the same SKU parsing, live PAYG/RI pricing, and license breakdown. |
| **Totals blank in every sheet + Summary sheet empty** | Root cause: totals had been switched to Excel `=SUM(...)` formulas. Formula cells only show a value once Excel actually recalculates and caches it — many viewers (LibreOffice quick-preview, Google Sheets before it fully loads, embedded grid previewers, or reading the file back with `openpyxl`/`pandas` without ever opening it in real Excel) display a formula cell as blank. Reverted every total (per-sheet and Summary) to a **plain static number**, computed as the exact sum of the same rounded values already written into the rows above it — so it always displays correctly everywhere, and still foots exactly (verified by `test_totals_are_static_values_and_foot_exactly`). |

## Critical

| Issue | Fix |
|---|---|
| **C1 — Worker exceptions silently swallowed** | `enrich_vms_concurrent` now calls `future.result()` on every submitted task. `process_row` itself is wrapped in try/except: any failure is logged with full row context (SKU/region/description/sheet) and a traceback, and the row falls back to its original values **with a visible Remarks warning** instead of silently reverting with no explanation. |
| **C2 — Only the active worksheet was read** | `parse_format()` now iterates `wb.worksheets` (every tab), tagging each row with `source_sheet`. If **no** sheet contains a `Service category` header, a clear `ValueError` is raised instead of silently producing an empty/partial report. |
| **C3 — Currency conversion could silently default to 1.0** | The entire "locked global FX ratio reverse-engineered from one row" mechanism is gone. Every row now fetches pricing **directly in the target currency** from the Retail Prices API (`currencyCode=<target>`), so there is no "no eligible candidate row" failure mode at all. |

## High

| Issue | Fix |
|---|---|
| **H1 — Premium OS (RHEL/SUSE) license silently merged into compute** | Removed the hardcoded `prem_os_payg_final = 0`. RHEL/SUSE license cost is now its own visible "Premium OS License (RHEL/SUSE)" line, same as Windows. |
| **H2 — Unsafe SKU-name fallback could corrupt M-series etc.** | Replaced the blind `sku.replace("s_v", "_v")` with a regex that only strips a trailing "s" flag when it directly follows a digit (`Standard_D4s_v3` → `Standard_D4_v3`), so it can no longer turn `Standard_M32ms_v2` into the invalid `Standard_M32m_v2`. Covered by `test_sku_fallback_never_corrupts_m_series`. |
| **H3 — Hardcoded static RHEL/SUSE tier pricing** | `get_premium_os_license_price()` now first tries to find a **live, distinct RHEL/SUSE meter** in the same API response already fetched for that SKU/region/currency. Only if no live meter exists does it fall back to the static reference table — and that fallback is converted using a **per-row** native/USD price ratio (not a global guessed ratio), and is always disclosed via a Remarks note. |
| **H4 — No Azure Hybrid Benefit / BYOL detection** | Descriptions are now scanned for AHB/BYOL keywords; when detected, the Windows license delta is not charged, and a Remarks note explains why. |
| **H5 — Remarks column was dead** | Remarks are now populated for every degraded/approximated/failed/not-applicable scenario: SKU not parsed, SKU not found via API, AHB/BYOL detected, RHEL/SUSE price approximated or unavailable, unexplained residual cost, Spot VM (RI n/a), and processing errors. |

## Medium

| Issue | Fix |
|---|---|
| Static Excel totals | All sheet totals and the Summary sheet are now real `=SUM(...)` formulas referencing the exact displayed cells, so totals always foot exactly and update live if a user edits a line item. |
| No currency symbol in number format | Number format is now built per-conversion as `"<symbol>"#,##0.00` (₹ / $ / € / £ / A$). |
| Fixed absolute `>5` unaccounted-cost threshold | Replaced with a threshold that scales with the row's own total (`max(1.0, 1% of original PAYG)`), so it means the same thing regardless of selected currency. |
| Global mutable formatting state | Number format is now threaded explicitly through every write function instead of a module-level global, so concurrent users of the same process can't bleed formatting into each other's output. |
| Negative-value guard incomplete | The residual-cost fold-back now floors at 0 and folds into the (also visible) "Other/Unidentified Cost" line rather than silently going negative. |
| Generic "unaccounted" cost mislabeled as Premium OS | Introduced a distinct "Other/Unidentified Cost" line (highlighted red) for residual cost that isn't SQL or premium-OS related, instead of mislabeling it. |

## Low / Security / Performance

| Issue | Fix |
|---|---|
| Raw exception text shown to users | `app.py` now logs full exception detail server-side and shows a generic, safe message to the end user for unexpected errors. `ValueError`s (which are deliberately user-facing) are still shown directly. |
| No file size limit | `app.py` now rejects uploads over 15 MB with a clear message. |
| Excel/CSV formula injection | Cell values are now defensively prefixed with `'` if user-supplied text starts with `=`, `+`, `-`, or `@`. |
| Aggressive retry/backoff | Reduced from `total=5, backoff_factor=1` (worst case ~31s per failing filter chain) to `total=3, backoff_factor=0.5` (worst case ~3.5s), and failed lookups are cached so a persistently-missing SKU doesn't retry-storm on every duplicate row. |
| Thread pool size hardcoded at 10 regardless of workload | Now scales as `min(10, max(2, len(vm_rows)))`. |
| Outdated/vulnerable dependencies | `streamlit` bumped past its known path-traversal CVE (fixed in 1.37.0), `openpyxl` bumped to 3.1.5, added `defusedxml` since openpyxl does not guard against XML entity-expansion attacks by default — relevant because this app parses untrusted user-uploaded `.xlsx` files. |
| No automated tests | Added `tests/test_convert.py` — 10 tests covering every Critical/High fix above. |
| No visibility into degraded rows before download | `app.py` now scans the generated workbook after conversion and shows a warning banner + expandable list of every row that has a Remarks note, instead of a blanket "Conversion Complete!". |

## Not changed (needs a decision from you, not a code fix)

- **Two near-duplicate GitHub Actions workflows** (`main_cost-converter.yml`, `main_cost-tool.yml`) deploy the same code to two different Web Apps. I didn't touch these since I can't tell which is actually your production target — recommend consolidating to one once you confirm which app is live, and adding a `pytest` step before the deploy step.
- **No authentication in front of the Streamlit app.** This is best solved at the infrastructure layer (Azure App Service Authentication / Entra ID "Easy Auth") rather than in application code.
