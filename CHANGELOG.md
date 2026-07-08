# Fix Changelog — Azure Pricing Calculator Conversion Tool

This documents every fix applied against the production review. Run `pytest tests/test_convert.py -v` to see them all verified automatically.

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
