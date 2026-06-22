# Port Barracks Per-Action Field Visibility into Central

**Date:** 2026-06-22
**Scope:** Frontend (`src/Index.html`) only — no backend changes
**Constraint:** Keep all of Central's existing actions and submit/finalize logic intact. Only the *which fields appear* behavior is changing.

---

## 1. Background

`New-Central` is the active project. `New Barracks` (at `C:\Users\Shane\Desktop\New Barracks`) is the older reference implementation. The two share the same form scaffolding and `handleActionChange()` pattern, but:

- Barracks shows a much leaner set of fields per action (e.g., its warehouseman `STOCK_IN` shows only the DR selector — no location/site/client/WBS/doc fields).
- Barracks labels the rightmost entry-table column `Inventory On Hand`. Central labels it `Target Status`.
- Central has additional actions (`ISSUE`, `TRANSFER_WH`, `RETURN_WH`) and a Pending Actions queue tab that don't exist in Barracks — these are deliberate upgrades and must remain.

The user wants Central's UI to feel like Barracks (cleaner, minimal fields per action), without losing any of Central's added workflows.

## 2. Goals

1. Each Central action shows only the fields it needs (matching Barracks's equivalent action where one exists).
2. The entry table's rightmost column is labeled **Inventory On Hand**.
3. Central's action set, submit logic, queue workflow, and `Code.js` backend are unchanged.

## 3. Non-Goals

- No backend (`Code.js`, `user.js`, `setup.js`, `appsscript.json`) changes.
- No action codes added or removed.
- No refactor of `handleActionChange()` into a declarative config (option B from brainstorming was rejected in favor of surgical edits).
- The password auto-login improvement is a separate change, not bundled here.

## 4. Changes

### 4.1 Table header rename

**File:** `src/Index.html`, around line 261
**Change:** `<th width="160">Target Status</th>` → `<th width="160">Inventory On Hand</th>`

The cell content is already populated by `updateInventoryBadges()` with the inventory-on-hand badge; no JS change is needed for this rename.

### 4.2 Per-action field visibility rules

All changes live inside `handleActionChange()` in `src/Index.html` (around lines 1200-1308). The function continues to dispatch on `act = document.getElementById('action').value`. For each action, the resulting visible field set must be:

| Central action | Role | Visible fields | Hidden fields |
|---|---|---|---|
| `RECEIVE_DR` | warehouseman | **DR ID dropdown** only | loc, site-name, site-id, client, wbs, doc-grp, po-grp, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button |
| `DR_CREATE` | admin | loc, site-name, site-id, client, wbs, **doc (required)**, PO dropdown + "PO to follow" checkbox, admin paste box | dr-select, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button; 10 rows preloaded |
| `PURCHASE_LOG` | admin | loc, site-name, site-id, client, wbs, **doc (required)**, PO dropdown + "PO to follow" checkbox, admin paste box, **Unit Price + Subtotal columns ON** | dr-select, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button; 10 rows preloaded |
| `USAGE` | team leader | **source-doc** (then auto-fills loc / site-name / site-id / client / wbs and locks them) | doc-grp, po-grp, mrc, dr-select, target-loc, target-site, return-type, Add Another Item button |
| `RETURN_CLIENT` | warehouseman | **source-doc**, **MRC** (then auto-fills loc / site-name / site-id / client / wbs and locks them); **Actual Quantity column ON** | doc-grp, po-grp, dr-select, target-loc, target-site, return-type, Add Another Item button |
| `ISSUE` | warehouseman | loc, site-name, site-id, client, wbs, doc-grp (auto-generate), Add Another Item button | po-grp, mrc, dr-select, target-loc, target-site, source-doc, return-type, price cols |
| `TRANSFER_WH` | warehouseman | loc, site-name (source), site-id, client, wbs, doc-grp (auto-generate), **target-loc**, **target-site**, Add Another Item button | po-grp, mrc, dr-select, source-doc, return-type, price cols |
| `RETURN_WH` | team leader | loc, site-name (source), site-id, client, wbs, doc-grp (auto-generate), Add Another Item button | po-grp, mrc, dr-select, target-loc, target-site, source-doc, return-type, price cols |

### 4.3 Specific edits inside `handleActionChange()`

The following toggles need to be added or adjusted (current Central behavior in parens):

1. **Hide `doc-grp` entirely for `RECEIVE_DR`.** Currently Central only hides the `t-doc` input but leaves the label and auto-generate hint visible. Change to add `d-none` to the whole `#doc-grp` element when `act === 'RECEIVE_DR'`.
2. **Hide `doc-grp` entirely for `RETURN_CLIENT`.** Same reason — should be hidden completely (Barracks hides it for source-doc-based actions).
3. **Hide loc/site/client/site-id/wbs for `USAGE` until source-doc is selected.** Currently this "hide until source doc is picked" behavior runs only for `RETURN_CLIENT`. Extend it to `USAGE` so the form starts as `source-doc only`, then reveals the auto-populated fields.
4. **Show `source-doc-grp` for `USAGE`.** Currently only shown for `RETURN_CLIENT`. Add `USAGE` to the source-doc-based set.
5. **Confirm `Add Another Item` button is hidden** for `RECEIVE_DR`, `RETURN_CLIENT`, `USAGE`, `DR_CREATE`, `PURCHASE_LOG` (matches Barracks). Current Central code hides it for `RETURN_CLIENT` and `RECEIVE_DR` only; extend the condition to also cover `USAGE`, `DR_CREATE`, `PURCHASE_LOG`.
6. **Leave `ISSUE`, `TRANSFER_WH`, `RETURN_WH` field rules essentially as-is**, since they have no Barracks equivalent. Audit to confirm they don't accidentally leave PO/MRC/source-doc/dr-select visible.

### 4.4 Source-doc-driven auto-fill flow (`USAGE`)

When `act === 'USAGE'`, `handleSourceDocChange()` must reveal and auto-fill loc / site-name / site-id / client / wbs (locking site-id and client). Currently this re-reveal block runs only when `act === 'RETURN_CLIENT'` (lines 1404-1421). Extend the conditional to include `USAGE`. The items table population for USAGE continues to be driven by the source doc (no change to `populateSourceDocItems` or `handleSiteChange`'s ISSUE/USAGE/RETURN_WH branch).

## 5. Out-of-Scope Cleanups

The `return-type-grp` element (currently always force-hidden inside `handleActionChange`) and any other vestigial UI elements stay as-is — removing them is a separate refactor.

## 6. Testing / Acceptance

Manual QA per action (push via `clasp push`, open the web app, log in as the appropriate role):

- [ ] `RECEIVE_DR`: only the DR ID dropdown shows above the items table. Selecting a DR auto-fills the items.
- [ ] `DR_CREATE`: admin paste box visible; PO dropdown + "PO to follow" visible; 10 blank rows; Unit Price/Subtotal hidden.
- [ ] `PURCHASE_LOG`: same as DR_CREATE but Unit Price and Subtotal columns visible.
- [ ] `USAGE`: only source-doc selector shows initially; picking a doc reveals loc/site/site-id/client/wbs and populates items.
- [ ] `RETURN_CLIENT`: only source-doc + MRC show initially; picking a doc reveals the auto-populated fields and items; Actual Quantity column visible.
- [ ] `ISSUE`, `TRANSFER_WH`, `RETURN_WH`: behavior unchanged; no stray PO/MRC/source-doc fields.
- [ ] Entry table header reads "Inventory On Hand" for every action.
- [ ] Submit / Confirm Transaction still works end-to-end for each action (backend unchanged, sanity check only).

## 7. Risks

- **Field-hide side effects:** A field that's hidden but still validated server-side could break submission. Mitigation: backend is unchanged, and Central currently already drives validation from values held in those hidden inputs after the auto-fill — confirm during QA that no required-field check fails after a source-doc auto-fill on `USAGE`.
- **Role-specific visibility:** Each action only appears in the dropdown for one role, so the field rules don't need to branch on role except where they already do (e.g., admin always sees the doc input).
