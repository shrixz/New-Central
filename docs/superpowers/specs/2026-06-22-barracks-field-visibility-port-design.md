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

**Guiding principle:** Match Barracks's visible field set exactly for each action that has a Barracks equivalent. Do NOT hide a field that Barracks shows. (User correction: "what fields presents in barracks, that's the only thing appeared on central… in warehouseman, the Receive DR is the only thing they can pick and PO is auto, that should also be applied to central but don't remove the other things.")

| Central action | Role | Visible fields | Hidden fields |
|---|---|---|---|
| `RECEIVE_DR` | warehouseman | **DR ID dropdown**, **PO Number (read-only, auto-populated from DR)** | loc, site-name, site-id, client, wbs, doc-grp, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button |
| `DR_CREATE` | admin | loc, site-name, site-id, client, wbs, **doc (required)**, PO dropdown + "PO to follow" checkbox, admin paste box | dr-select, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button; 10 rows preloaded |
| `PURCHASE_LOG` | admin | loc, site-name, site-id, client, wbs, **doc (required)**, PO dropdown + "PO to follow" checkbox, admin paste box, **Unit Price + Subtotal columns ON** | dr-select, mrc, target-loc, target-site, source-doc, return-type, Add Another Item button; 10 rows preloaded |
| `USAGE` | team leader | **source-doc** (then auto-fills loc / site-name / site-id / client / wbs and locks them) | doc-grp, po-grp, mrc, dr-select, target-loc, target-site, return-type, Add Another Item button |
| `RETURN_CLIENT` | warehouseman | **source-doc**, **MRC**, **doc-grp (label + auto-generate hint for non-admin; required input for admin)** (then auto-fills loc / site-name / site-id / client / wbs and locks them); **Actual Quantity column ON** | po-grp, dr-select, target-loc, target-site, return-type, Add Another Item button |
| `ISSUE` | warehouseman | loc, site-name, site-id, client, wbs, doc-grp (auto-generate), Add Another Item button | po-grp, mrc, dr-select, target-loc, target-site, source-doc, return-type, price cols |
| `TRANSFER_WH` | warehouseman | loc, site-name (source), site-id, client, wbs, doc-grp (auto-generate), **target-loc**, **target-site**, Add Another Item button | po-grp, mrc, dr-select, source-doc, return-type, price cols |
| `RETURN_WH` | team leader | loc, site-name (source), site-id, client, wbs, doc-grp (auto-generate), Add Another Item button | po-grp, mrc, dr-select, target-loc, target-site, source-doc, return-type, price cols |

**Note on `RECEIVE_DR` PO behavior:** Central's existing `handleActionChange()` already shows `po-grp` and swaps the dropdown for the read-only input when `act === 'RECEIVE_DR'`, and `handleDRSelection()` already populates `t-po-readonly` with `items[0].poNumber`. No new code is needed for the PO display on RECEIVE_DR — just don't accidentally hide `po-grp` for this action.

### 4.3 Specific edits inside `handleActionChange()`

The following toggles need to be added or adjusted (current Central behavior in parens):

1. **Hide `doc-grp` entirely for `RECEIVE_DR` and `USAGE` only.** Currently Central only hides the `t-doc` input but leaves the `#doc-grp` container visible. Change to add `d-none` to the whole `#doc-grp` element when `act === 'RECEIVE_DR' || act === 'USAGE'`. (Barracks hides doc-grp for WH STOCK_IN and USAGE but keeps it visible — label + auto-generate hint — for RETURN_CLIENT, so we do the same.)
2. **Keep `doc-grp` visible for `RETURN_CLIENT`.** No change to the doc-grp container for this action. The existing non-admin branch hides the `t-doc` input and shows the `t-doc-auto-text` hint — that is the correct Barracks-matching behavior.
3. **Keep `po-grp` visible for `RECEIVE_DR`** with `po-dd-wrap` hidden and `t-po-readonly` shown. Central's existing `handleActionChange()` already does this; no change needed except making sure we do not regress it when editing.
4. **Hide loc/site/client/site-id/wbs for `USAGE` until source-doc is selected.** Currently this "hide until source doc is picked" behavior runs only for `RETURN_CLIENT`. Extend it to `USAGE` so the form starts as `source-doc only`, then reveals the auto-populated fields.
5. **Show `source-doc-grp` for `USAGE`.** Currently only shown for `RETURN_CLIENT`. Add `USAGE` to the source-doc-based set.
6. **Confirm `Add Another Item` button is hidden** for `RECEIVE_DR`, `RETURN_CLIENT`, `USAGE`, `DR_CREATE`, `PURCHASE_LOG` (matches Barracks). Current Central code hides it for `RETURN_CLIENT` and `RECEIVE_DR` only; extend the condition to also cover `USAGE`, `DR_CREATE`, `PURCHASE_LOG`.
7. **Leave `ISSUE`, `TRANSFER_WH`, `RETURN_WH` field rules essentially as-is**, since they have no Barracks equivalent. Audit to confirm they don't accidentally leave PO/MRC/source-doc/dr-select visible.

### 4.4 Source-doc-driven auto-fill flow (`USAGE`)

When `act === 'USAGE'`, `handleSourceDocChange()` must reveal and auto-fill loc / site-name / site-id / client / wbs (locking site-id and client). Currently this re-reveal block runs only when `act === 'RETURN_CLIENT'` (lines 1404-1421). Extend the conditional to include `USAGE`. The items table population for USAGE continues to be driven by the source doc (no change to `populateSourceDocItems` or `handleSiteChange`'s ISSUE/USAGE/RETURN_WH branch).

## 5. Out-of-Scope Cleanups

The `return-type-grp` element (currently always force-hidden inside `handleActionChange`) and any other vestigial UI elements stay as-is — removing them is a separate refactor.

## 6. Testing / Acceptance

Manual QA per action (push via `clasp push`, open the web app, log in as the appropriate role):

- [ ] `RECEIVE_DR`: DR ID dropdown + PO Number (read-only) are the only fields above the items table. Selecting a DR auto-populates the PO and the items.
- [ ] `DR_CREATE`: admin paste box visible; PO dropdown + "PO to follow" visible; 10 blank rows; Unit Price/Subtotal hidden.
- [ ] `PURCHASE_LOG`: same as DR_CREATE but Unit Price and Subtotal columns visible.
- [ ] `USAGE`: only source-doc selector shows initially; picking a doc reveals loc/site/site-id/client/wbs and populates items.
- [ ] `RETURN_CLIENT`: source-doc + MRC + Doc Number (label + auto-generate hint) show initially; picking a doc reveals the auto-populated fields and items; Actual Quantity column visible.
- [ ] `ISSUE`, `TRANSFER_WH`, `RETURN_WH`: behavior unchanged; no stray PO/MRC/source-doc fields.
- [ ] Entry table header reads "Inventory On Hand" for every action.
- [ ] Submit / Confirm Transaction still works end-to-end for each action (backend unchanged, sanity check only).

## 7. Risks

- **Field-hide side effects:** A field that's hidden but still validated server-side could break submission. Mitigation: backend is unchanged, and Central currently already drives validation from values held in those hidden inputs after the auto-fill — confirm during QA that no required-field check fails after a source-doc auto-fill on `USAGE`.
- **Role-specific visibility:** Each action only appears in the dropdown for one role, so the field rules don't need to branch on role except where they already do (e.g., admin always sees the doc input).
