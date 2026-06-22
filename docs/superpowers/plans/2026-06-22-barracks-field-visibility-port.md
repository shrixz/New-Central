# Barracks Field-Visibility Port Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mirror Barracks's per-action field visibility into Central, plus rename the entry-table column to "Inventory On Hand", without touching Central's actions, queue workflow, or backend.

**Architecture:** All changes live inside `src/Index.html`. We edit one HTML header cell and the `handleActionChange()` / `handleSourceDocChange()` / `handleLocationChange()` / `handleSiteChange()` functions in the inline `<script>` block. No backend (`Code.js`, `user.js`) touched. No new files.

**Tech Stack:** Google Apps Script web app, Bootstrap 5, vanilla JS, SweetAlert2. Deployed via `clasp push`.

## Global Constraints

- All edits limited to `src/Index.html`.
- Do not remove, rename, or change the submit-side behavior of any action (`DR_CREATE`, `PURCHASE_LOG`, `RECEIVE_DR`, `ISSUE`, `TRANSFER_WH`, `RETURN_CLIENT`, `USAGE`, `RETURN_WH`).
- Do not edit `Code.js`, `user.js`, `setup.js`, `appsscript.json`, or `trigger.js`.
- No automated test framework exists in this project — verification is manual via `clasp push` and clicking through each action in the browser.
- Commit after each task. Push (`clasp push`) at user's discretion; the plan calls it out as a verification gate but the user decides when to actually push to Apps Script.

---

### Task 1: Rename entry-table column header

**Files:**
- Modify: `src/Index.html` (one line, header row of the Movement transaction table)

**Interfaces:**
- Consumes: nothing.
- Produces: the entry table now reads "Inventory On Hand" instead of "Target Status". `updateInventoryBadges()` already populates the cell with the correct badge — no JS change needed.

- [ ] **Step 1: Locate the header**

Find this line in `src/Index.html` (near line 261):

```html
                    <th width="160">Target Status</th>
```

- [ ] **Step 2: Make the rename**

Replace it with:

```html
                    <th width="160">Inventory On Hand</th>
```

- [ ] **Step 3: Manual verification**

Either: (a) `clasp push` then open the web app and confirm the rightmost column reads "Inventory On Hand" on every action, or (b) inspect the diff and confirm the rename is the only change in the working tree.

Expected: column header text reads "Inventory On Hand" in the Movement tab table.

- [ ] **Step 4: Commit**

```bash
git add src/Index.html
git commit -m "Rename entry-table column header to 'Inventory On Hand'"
```

---

### Task 2: Hide `doc-grp` entirely for `RECEIVE_DR` only (not `RETURN_CLIENT`)

**Files:**
- Modify: `src/Index.html`, inside `handleActionChange()` (around lines 1260-1268, in the `if (userRole === 'admin') { ... } else { ... }` block).

**Interfaces:**
- Consumes: existing `isReceive` flag computed earlier in the same function.
- Produces: when `act === 'RECEIVE_DR'`, the `#doc-grp` column (Doc Number label + input + auto-generate hint) is fully hidden. For `RETURN_CLIENT` and every other action, the existing behavior is preserved (RETURN_CLIENT keeps the label + auto-generate hint visible for non-admin, matching Barracks). Task 4 will extend this hide to cover `USAGE` once that action becomes source-doc-based.

- [ ] **Step 1: Locate the admin/non-admin doc block**

Find this block in `handleActionChange()`:

```js
    if (userRole === 'admin') {
        document.getElementById('t-doc').classList.remove('d-none');
        document.getElementById('t-doc-auto-text').classList.add('d-none');
    } else {
        document.getElementById('t-doc').classList.add('d-none');
        document.getElementById('t-doc').value = '';
        document.getElementById('t-doc-auto-text').classList.remove('d-none');
    }
```

- [ ] **Step 2: Add the doc-grp hide line just before this block**

Replace the block with:

```js
    // Hide the whole Doc Number group only when the action has no doc concept at all.
    // RECEIVE_DR auto-references the DR. RETURN_CLIENT still shows the label + auto-generate
    // hint (matches Barracks), so do NOT include it here. USAGE will be added in Task 4.
    document.getElementById('doc-grp').classList.toggle('d-none', isReceive);

    if (userRole === 'admin') {
        document.getElementById('t-doc').classList.remove('d-none');
        document.getElementById('t-doc-auto-text').classList.add('d-none');
    } else {
        document.getElementById('t-doc').classList.add('d-none');
        document.getElementById('t-doc').value = '';
        document.getElementById('t-doc-auto-text').classList.remove('d-none');
    }
```

- [ ] **Step 3: Manual verification**

After `clasp push`, log in and confirm:
- As warehouseman, select `RECEIVE_DR`: above the items table you see ONLY the DR ID dropdown and the PO Number (read-only, auto-populated when a DR is picked). No Doc Number label/input/hint, no Warehouse Location, no Site, no Client, no WBS.
- As warehouseman, select `RETURN_CLIENT`: you see the Source Doc dropdown, MRC field, AND the Doc Number group (label + "Document ID will be auto-generated" hint for non-admin; required input for admin). Loc/Site/Client/Site-ID/WBS hidden until a source doc is selected.
- As admin, select `DR_CREATE`: Doc Number field is visible and required as before.
- As admin, select `PURCHASE_LOG`: Doc Number field is visible.
- As team leader, select `USAGE`: Doc Number group still visible (will be cleaned up in Task 4).
- As warehouseman, select `ISSUE` or `TRANSFER_WH`: Doc Number group still visible with auto-generate hint.

- [ ] **Step 4: Commit**

```bash
git add src/Index.html
git commit -m "Hide Doc Number group for RECEIVE_DR (keep visible for RETURN_CLIENT)"
```

---

### Task 3: Hide "Add Another Item" button for `DR_CREATE`, `PURCHASE_LOG`, and `USAGE`

**Files:**
- Modify: `src/Index.html`, inside `handleActionChange()` (the `addRowBtn` block around lines 1274-1281).

**Interfaces:**
- Consumes: `act` value of action dropdown.
- Produces: the `#btn-add-row` button is hidden for `DR_CREATE`, `PURCHASE_LOG`, `RECEIVE_DR`, `RETURN_CLIENT`, and `USAGE`. Visible for `ISSUE`, `TRANSFER_WH`, `RETURN_WH`.

Rationale: `DR_CREATE` / `PURCHASE_LOG` preload 10 rows for bulk encode; `RECEIVE_DR` / `RETURN_CLIENT` / `USAGE` populate from a source document. None of them benefit from the "add another item" affordance.

- [ ] **Step 1: Locate the addRowBtn block**

Find in `handleActionChange()`:

```js
    const addRowBtn = document.getElementById('btn-add-row');
    if (addRowBtn) {
        if (isReturn || isReceive) {
            addRowBtn.style.display = 'none';
        } else {
            addRowBtn.style.display = 'inline-block';
        }
    }
```

- [ ] **Step 2: Extend the hide condition**

Replace the block with:

```js
    const addRowBtn = document.getElementById('btn-add-row');
    if (addRowBtn) {
        const hideAddRow = isReturn || isReceive
            || act === 'DR_CREATE' || act === 'PURCHASE_LOG' || act === 'USAGE';
        addRowBtn.style.display = hideAddRow ? 'none' : 'inline-block';
    }
```

- [ ] **Step 3: Manual verification**

After `clasp push`, switch through every action and verify the "Add Another Item" button:
- `RECEIVE_DR`, `RETURN_CLIENT`, `DR_CREATE`, `PURCHASE_LOG`, `USAGE` → hidden.
- `ISSUE`, `TRANSFER_WH`, `RETURN_WH` → visible.

- [ ] **Step 4: Commit**

```bash
git add src/Index.html
git commit -m "Hide Add Another Item button for source-doc and bulk-encode actions"
```

---

### Task 4: Convert `USAGE` to source-doc-based field flow

**Files:**
- Modify: `src/Index.html`, inside `handleActionChange()`, `handleLocationChange()`, `handleSourceDocChange()`, `handleSiteChange()`.

**Interfaces:**
- Consumes: existing `populateSourceDocs()` and `master.receiptDocs` (already used by `RETURN_CLIENT`).
- Produces: `USAGE` action shows only the Source Doc dropdown initially. Picking a doc reveals and auto-fills `loc / site-name / site-id / client / wbs` (with site-id and client locked), then populates the items table from the doc's remaining items. The submit-side behavior of USAGE (Code.js side) is unchanged — items still write as USAGE records.

This is the biggest task in the plan because it shifts USAGE's data-entry surface from "live inventory at a site" to "items from a source receipt doc", matching Barracks. The submit logic in `Code.js` already accepts items + USAGE action, so no backend change is needed.

- [ ] **Step 1: In `handleActionChange()`, add `USAGE` to the source-doc-grp and hideLocFields logic**

Find:

```js
    const isReturn = act === 'RETURN_CLIENT';
    const isReceive = act === 'RECEIVE_DR';

    if (isReturn) {
        document.getElementById('source-doc-grp').classList.remove('d-none');
        document.getElementById('t-source-doc').disabled = false;
    } else {
        document.getElementById('source-doc-grp').classList.add('d-none');
    }

    // Completely hide mapping fields if we are doing a Return (so only 2 fields show) or Receive process
    const hideLocFields = (isReceive || isReturn);
```

Replace with:

```js
    const isReturn = act === 'RETURN_CLIENT';
    const isReceive = act === 'RECEIVE_DR';
    const isUsage = act === 'USAGE';
    const isSourceDocBased = isReturn || isUsage;

    if (isSourceDocBased) {
        document.getElementById('source-doc-grp').classList.remove('d-none');
        document.getElementById('t-source-doc').disabled = false;
        document.getElementById('t-source-doc').value = '';
    } else {
        document.getElementById('source-doc-grp').classList.add('d-none');
    }

    // Hide mapping fields for Receive (uses DR), Return, and Usage (use source-doc).
    const hideLocFields = (isReceive || isSourceDocBased);
```

- [ ] **Step 2: Extend the Task 2 doc-grp hide to also cover USAGE**

Find the line added in Task 2:

```js
    document.getElementById('doc-grp').classList.toggle('d-none', isReceive);
```

Replace with:

```js
    document.getElementById('doc-grp').classList.toggle('d-none', isReceive || isUsage);
```

This hides the Doc Number group for USAGE (matches Barracks). RETURN_CLIENT is intentionally left visible — Barracks shows the label + auto-generate hint for non-admin there.

- [ ] **Step 3: In `handleLocationChange()`, call `populateSourceDocs()` for `USAGE` too**

Find:

```js
  function handleLocationChange() {
    const act = document.getElementById('action').value;
    if (act !== 'RECEIVE_DR' && act !== 'RETURN_CLIENT') {
        updateSiteDropdown();
    }
    if (act === 'RETURN_CLIENT') {
        populateSourceDocs();
    }
    updateInventoryBadges();
  }
```

Replace with:

```js
  function handleLocationChange() {
    const act = document.getElementById('action').value;
    if (act !== 'RECEIVE_DR' && act !== 'RETURN_CLIENT' && act !== 'USAGE') {
        updateSiteDropdown();
    }
    if (act === 'RETURN_CLIENT' || act === 'USAGE') {
        populateSourceDocs();
    }
    updateInventoryBadges();
  }
```

- [ ] **Step 4: In `populateSourceDocs()`, extend the access filter to cover `USAGE` for team leaders**

Find in `populateSourceDocs()` (around line 609):

```js
    if (act === 'RETURN_CLIENT') {
        if (userRole === 'warehouseman' && locAccess) {
            const locs = locAccess.split(',').map(s => s.trim());
            options = options.filter(d => locs.includes(master.receiptDocs[d].location));
        }
    } else if (loc) {
        options = options.filter(d => master.receiptDocs[d].location === loc);
    }
```

Replace with:

```js
    if (act === 'RETURN_CLIENT') {
        if (userRole === 'warehouseman' && locAccess) {
            const locs = locAccess.split(',').map(s => s.trim());
            options = options.filter(d => locs.includes(master.receiptDocs[d].location));
        }
    } else if (act === 'USAGE') {
        if (userRole === 'team leader' && siteAccess) {
            const sites = siteAccess.split(',').map(s => s.trim());
            options = options.filter(d => sites.includes(master.receiptDocs[d].site));
        }
    } else if (loc) {
        options = options.filter(d => master.receiptDocs[d].location === loc);
    }
```

- [ ] **Step 5: In `handleSourceDocChange()`, extend the reveal/auto-fill block to cover `USAGE`**

Find:

```js
      if (!doc || !master.receiptDocs[doc]) {
          if (act === 'RETURN_CLIENT') {
             document.getElementById('loc-grp').style.display = 'none';
             document.getElementById('site-name-grp').style.display = 'none';
             document.getElementById('client-grp').style.display = 'none';
             document.getElementById('site-id-grp').style.display = 'none';
             document.getElementById('wbs-grp').style.display = 'none';
          }
          return;
      }

      const docData = master.receiptDocs[doc];
      
      if (act === 'RETURN_CLIENT') {
          // Re-reveal locked fields
          document.getElementById('loc-grp').style.display = 'block';
          document.getElementById('site-name-grp').style.display = 'block';
          document.getElementById('client-grp').style.display = 'block';
          document.getElementById('site-id-grp').style.display = 'block';
          document.getElementById('wbs-grp').style.display = 'block';

          document.getElementById('t-location').value = docData.location || '';
          updateSiteDropdown(); 
          document.getElementById('t-site-name').value = docData.site || ''; 
          handleSiteChange(); 
          
          document.getElementById('t-location').disabled = true;
          document.getElementById('t-site-name').disabled = true;
          document.getElementById('t-site-id').disabled = true;
          document.getElementById('t-client').disabled = true;
      }
```

Replace with:

```js
      const isSourceDocBased = (act === 'RETURN_CLIENT' || act === 'USAGE');

      if (!doc || !master.receiptDocs[doc]) {
          if (isSourceDocBased) {
             document.getElementById('loc-grp').style.display = 'none';
             document.getElementById('site-name-grp').style.display = 'none';
             document.getElementById('client-grp').style.display = 'none';
             document.getElementById('site-id-grp').style.display = 'none';
             document.getElementById('wbs-grp').style.display = 'none';
          }
          return;
      }

      const docData = master.receiptDocs[doc];

      if (isSourceDocBased) {
          // Re-reveal locked fields
          document.getElementById('loc-grp').style.display = 'block';
          document.getElementById('site-name-grp').style.display = 'block';
          document.getElementById('client-grp').style.display = 'block';
          document.getElementById('site-id-grp').style.display = 'block';
          document.getElementById('wbs-grp').style.display = 'block';

          document.getElementById('t-location').value = docData.location || '';
          updateSiteDropdown();
          document.getElementById('t-site-name').value = docData.site || '';
          handleSiteChange();

          document.getElementById('t-location').disabled = true;
          document.getElementById('t-site-name').disabled = true;
          document.getElementById('t-site-id').disabled = true;
          document.getElementById('t-client').disabled = true;
      }
```

- [ ] **Step 6: In `handleSiteChange()`, remove `USAGE` from the live-inventory item-pull branch**

Find:

```js
          if (act === 'ISSUE' || act === 'USAGE' || act === 'RETURN_WH') {
             document.getElementById('entry-rows').innerHTML = '';
             let hasItems = false;
             
             master.inventory.forEach(invItem => {
                 let stock = 0;
                 if (act === 'ISSUE') { stock = invItem.balances[curLoc || details.location] || 0; }
                 else { stock = invItem.balances[tSite] || 0; }
                 
                 if (stock > 0) {
                     hasItems = true;
                     addRow(invItem.code, invItem.name, invItem.uom, "");
                 }
             });
             
             if (!hasItems) {
                addRow(); 
                Swal.fire('Notice', 'No items currently in stock for this location/site.', 'info');
             }
          } 
```

Replace with:

```js
          if (act === 'ISSUE' || act === 'RETURN_WH') {
             document.getElementById('entry-rows').innerHTML = '';
             let hasItems = false;

             master.inventory.forEach(invItem => {
                 let stock = 0;
                 if (act === 'ISSUE') { stock = invItem.balances[curLoc || details.location] || 0; }
                 else { stock = invItem.balances[tSite] || 0; }

                 if (stock > 0) {
                     hasItems = true;
                     addRow(invItem.code, invItem.name, invItem.uom, "");
                 }
             });

             if (!hasItems) {
                addRow();
                Swal.fire('Notice', 'No items currently in stock for this location/site.', 'info');
             }
          }
          // NOTE: USAGE handled by handleSourceDocChange (items pulled from source receipt doc).
```

- [ ] **Step 7: In `handleSourceDocChange()`, ensure items populate for both actions**

Find the item-population loop at the bottom of `handleSourceDocChange()`:

```js
      const items = docData.items;
      const curLoc = document.getElementById('t-location').value;
      
      for (let code in items) {
          const invItem = master.inventory.find(i => i.code === code);
          const warehouseStock = invItem ? (invItem.balances[curLoc] || 0) : 0;
          const remainingInDoc = items[code].remaining;
          const availableToReturn = Math.min(remainingInDoc, warehouseStock);
          
          if (availableToReturn > 0) {
              addRow(code, items[code].name, items[code].uom, availableToReturn, availableToReturn, true);
          }
      }
      updateInventoryBadges();
```

Replace with:

```js
      const items = docData.items;
      const tSiteForUsage = document.getElementById('t-site-name').value;
      const curLoc = document.getElementById('t-location').value;

      for (let code in items) {
          const invItem = master.inventory.find(i => i.code === code);
          const remainingInDoc = items[code].remaining;
          let prefill = 0;
          let manual = false;

          if (act === 'RETURN_CLIENT') {
              // Return cap = doc remaining ∩ warehouse stock at the source loc
              const warehouseStock = invItem ? (invItem.balances[curLoc] || 0) : 0;
              prefill = Math.min(remainingInDoc, warehouseStock);
              manual = true;
              if (prefill <= 0) continue;
          } else if (act === 'USAGE') {
              // Usage cap = doc remaining ∩ site stock at the destination site
              const siteStock = invItem ? (invItem.balances[tSiteForUsage] || 0) : 0;
              const available = Math.min(remainingInDoc, siteStock);
              if (available <= 0) continue;
              // Leave quantity blank for the user to fill; pass the doc-derived row in manual mode
              prefill = "";
              manual = true;
              addRow(code, items[code].name, items[code].uom, prefill, 0, manual);
              continue;
          } else {
              continue;
          }

          addRow(code, items[code].name, items[code].uom, prefill, prefill, manual);
      }
      updateInventoryBadges();
```

> If `addRow`'s signature differs from what's used above, search for `function addRow(` in `Index.html` and adapt the call to use the same argument positions used elsewhere in this same file for RETURN_CLIENT. The existing `RETURN_CLIENT` call passes `(code, name, uom, qty, actualQty, manualFlag)` — match that shape exactly.

- [ ] **Step 8: Manual verification**

After `clasp push`, log in as a team leader and:
1. Open the Movement tab — only the Source Doc dropdown is visible above the items table (no Location/Site/Client/WBS/Doc fields, no MRC, no DR ID, no PO).
2. Pick a source doc — Location, Site Name, Site ID, Client, WBS auto-fill and lock; the items list populates with items that have remaining stock at the destination site.
3. Enter a quantity for one item and submit; verify the transaction posts as USAGE in the History tab and inventory updates correctly.
4. Switch role to warehouseman (or use another account) and confirm `ISSUE`, `TRANSFER_WH`, `RETURN_WH` still work the OLD way (live inventory pull on site change) — verify by switching to those actions and confirming items populate from inventory.
5. Confirm `RETURN_CLIENT` still works exactly as before.

- [ ] **Step 9: Commit**

```bash
git add src/Index.html
git commit -m "Convert USAGE to source-doc-based field flow (mirror barracks)"
```

---

### Task 5: Final per-action visibility audit and screenshot pass

**Files:**
- Modify (only if defects found): `src/Index.html`.

**Interfaces:** none.

Goal: walk through every action one more time and confirm the field set matches the table in the spec. Fix any stray visible field.

- [ ] **Step 1: After `clasp push`, log in once per role and switch through every action**

For each action, confirm the visible field set matches the spec's table (`docs/superpowers/specs/2026-06-22-barracks-field-visibility-port-design.md`, section 4.2). Take a screenshot or note the visible fields per action.

Per-action expected fields:

- `RECEIVE_DR` (warehouseman): DR ID + PO Number (read-only, auto-populated after DR selection). Nothing else above the items table.
- `DR_CREATE` (admin): Warehouse Location, Site Name, Site ID, Client, WBS, Doc Number (required), PO dropdown + "PO to follow", admin paste box; 10 rows preloaded.
- `PURCHASE_LOG` (admin): same as DR_CREATE plus Unit Price / Subtotal columns.
- `USAGE` (team leader): Source Doc only initially; after pick, Loc/Site/Site ID/Client/WBS appear locked, items populate.
- `RETURN_CLIENT` (warehouseman): Source Doc + MRC + Doc Number group (label + auto-generate hint for non-admin) initially; after pick, Loc/Site/Site ID/Client/WBS appear locked, Actual Quantity column visible.
- `ISSUE` (warehouseman): Loc, Site Name, Site ID, Client, WBS, Doc Number (auto-generate hint); Add Another Item visible.
- `TRANSFER_WH` (warehouseman): Loc, Site Name (source), Site ID, Client, WBS, Doc Number (auto-generate), Target Loc, Target Site; Add Another Item visible.
- `RETURN_WH` (team leader): Loc, Site Name (source), Site ID, Client, WBS, Doc Number (auto-generate); Add Another Item visible.

- [ ] **Step 2: Fix any defect found**

If any action shows a field that the spec says should be hidden (or hides one it should show), edit `handleActionChange()` in `src/Index.html` to add the missing toggle. Code style: prefer `document.getElementById('<id>').classList.toggle('d-none', <condition>)` to match the function's existing style.

- [ ] **Step 3: Confirm entry-table column reads "Inventory On Hand" for every action**

Visual check across all 8 actions.

- [ ] **Step 4: Commit (if any fixes applied) or note no further changes**

```bash
# Only if any edits were made in Step 2:
git add src/Index.html
git commit -m "Audit fix: tidy per-action field visibility"
```

If no defects were found, no commit needed for this task; mark it done.

---

## Plan Self-Review

**Spec coverage:**
- §4.1 header rename → Task 1 ✓
- §4.2 per-action table → Tasks 2, 3, 4, 5 collectively cover every row ✓
- §4.3.1 hide doc-grp for RECEIVE_DR → Task 2 ✓
- §4.3.2 hide doc-grp for RETURN_CLIENT → Task 2 ✓
- §4.3.3 hide loc fields for USAGE until source-doc selected → Task 4 ✓
- §4.3.4 show source-doc-grp for USAGE → Task 4 ✓
- §4.3.5 hide Add Another Item for USAGE/DR_CREATE/PURCHASE_LOG → Task 3 ✓
- §4.3.6 ISSUE/TRANSFER_WH/RETURN_WH audit → Task 5 ✓
- §4.4 USAGE source-doc auto-fill flow → Task 4 ✓
- §6 Acceptance checklist → Task 5 manual QA ✓

**Placeholder scan:** no "TBD"/"TODO"/"figure out later" in any step. Every code-changing step shows the actual code.

**Type/name consistency:** `isSourceDocBased` is defined in Task 4 Step 1 and re-used in Task 4 Step 2 and Step 5; `populateSourceDocs()` and `master.receiptDocs` referenced are existing functions/data; `addRow` signature `(code, name, uom, qty, actualQty, manual)` matches the existing `RETURN_CLIENT` call site referenced in Task 4 Step 7.
