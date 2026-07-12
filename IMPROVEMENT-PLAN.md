# Smart Org Chart — Code Review & Improvement Plan

Execution plan from a full code review (2026-07-10, v1.2.1). Work through phases in order — Phase 1 items are correctness bugs, later phases are performance, dead code, and polish. Each task lists the file(s), the problem, the fix, and acceptance criteria.

## Ground rules

- **Build/verify:** `gulp build` must pass with no new lint errors after every phase. There is no test suite; verify behavior manually with `gulp serve` (demo mode auto-activates on localhost, so MockGraphService supplies data — use the 1,000-person dataset from the Settings gear for stress tests).
- **Never hand-edit `*.module.scss.ts`** files — they are generated from the sibling `.scss` during build.
- **American English** in all code, comments, strings, and docs (organization, color, center, behavior).
- Keep the existing code style: aligned property assignments, section-banner comments, private `_method` naming.
- Do not bump the version or touch `CHANGELOG.md` except where a task says to; the maintainer will handle release notes.

---

## Phase 1 — Bugs (correctness)

### 1.1 Manager cycle / self-manager causes infinite loop in Expand All
**Files:** `src/services/GraphService.ts` (`_buildMaps`, ~line 517), `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (`_handleToggle` ~1074, `_handleExpandLoaded` ~1103)

Real Azure AD tenants sometimes have a user whose manager is themselves (common for CEOs), or short cycles (A→B→A). There is no guard anywhere:

- In `_buildMaps`, a self-manager puts the user into their own `_childrenMap` entry, so the user renders as their own child.
- In `OrgChart._handleExpandLoaded`, the `while (this._mounted)` BFS loop finds the self/cycle node on the unloaded frontier, injects its reports (which include an ancestor), which creates a new unloaded frontier node — **the loop never terminates and the browser hangs**. The same applies to longer cycles reachable via `_handleToggle` (each manual expand adds another copy, unbounded depth).

**Fix (three layers):**
1. In `GraphService._buildMaps`: after resolving `mgrId`, `if (mgrId === user.id) continue;` (skip self-manager). Apply the same guard in `MockGraphService._init` for consistency.
2. Add a small helper in `OrgChart.tsx` (near the other tree helpers): `collectAncestorIds(root, targetId): Set<string>` that returns the ids on the path from root to the target node. In `_handleToggle` and in `_handleExpandLoaded`'s injection step, filter fetched reports: drop any report whose `id` is the node's own id or an ancestor id, so cycles longer than 1 cannot re-inject an ancestor.
3. Safety valve in `_handleExpandLoaded`: cap the outer `while` loop at a generous iteration count (e.g., 50 levels) so a future regression degrades gracefully instead of hanging.

**Acceptance:** Temporarily add a self-manager row and a 2-cycle to `RAW_BASE` in `MockGraphService.ts`, run `gulp serve`, click Expand All in Top Down layout — it terminates, the cyclic users appear once, no console errors. Remove the temporary rows afterward.

### 1.2 `?socFocus` URL param is written even in the default state, and instances clash
**File:** `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (`componentDidUpdate` ~832–852, `updateUrlFocus` ~719)

On first load in drill mode, `_loadTree` sets `drillPath = [rootUser]`, which triggers the persistence block and writes `focusEmail = root's email` into both localStorage and the page URL. Consequences:

- Every visitor's URL silently gains `?socFocus=ceo@…` just by opening the page — copied/shared links carry a meaningless focus param.
- Two org chart web parts on the same page fight over the single un-scoped URL param, each restoring the other's focus.

**Fix:**
1. Compute a "default position" check: in drill mode, when `drillPath.length <= 1` and the drill root equals the loaded `rootNode` user (and in full-tree mode when `focusedUser === null`), treat `focusEmail` as `null`. Only write a non-null `focusEmail` when the user has actually navigated away from the default.
2. `updateUrlFocus(null)` already removes the param — with (1), the default state now cleans the URL instead of polluting it.
3. Keep localStorage behavior the same (it is already instance-scoped); only the URL logic changes.

**Acceptance:** Load the page fresh (no stored state) → URL has no `socFocus` param. Drill into a person → param appears. Navigate back to root via Home → param is removed.

### 1.3 Zoom "Reset" ignores the admin-configured default zoom
**File:** `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (~line 1981)

The reset button always sets `zoomLevel: 1`, but the admin may have configured `defaultZoom` (e.g., 75%). Reset should return to the configured default: `this.props.defaultZoom > 0 ? this.props.defaultZoom : 1`, and the `disabled` condition should compare against that same value.

**Acceptance:** Set Default Org Chart Zoom to 75% in the property pane, zoom in, click Reset → returns to 75%.

### 1.4 Admin property changes don't take effect until page reload
**File:** `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (`componentDidUpdate` ~805)

`chartLayout`, `filterMembers/filterGuests`, `filterDepartments`, `showStats`, and `zoomLevel` are initialized from props only in the constructor. SPFx property pane edits re-render (not remount) the component, so:

- Changing **Default Org Chart Layout** or **Default Zoom** does nothing visibly.
- Turning **off** a feature flag (`enableLayoutToggle`, `enableUserFilter`, `enableDeptFilter`, `enableStats`) hides the button but leaves any persisted state active — e.g., a user who filtered to one department is stuck with the filter and no UI to clear it.

**Fix:** Extend `componentDidUpdate` prop-change handling:
- `prev.defaultLayout !== this.props.defaultLayout` → if `!this.props.enableLayoutToggle` or no stored layout override, `this._setLayout(this.props.defaultLayout)`.
- `prev.defaultZoom !== this.props.defaultZoom` → `setState({ zoomLevel: defaultZoom > 0 ? defaultZoom : 1 })` and call `_autoFitZoom()` when it becomes 0.
- Feature flag turned off → reset the corresponding state to neutral (`enableUserFilter` → `filterMembers: true, filterGuests: true`; `enableDeptFilter` → `filterDepartments: new Set()`; `enableStats` → `showStats: false`; `enableLayoutToggle` → `chartLayout: props.defaultLayout || 'drill'`).

**Acceptance:** With the web part in edit mode, toggle each of these properties and confirm the chart updates live without a page reload; disabling User Filter while a filter is active un-hides the filtered users.

### 1.5 Filtered-out managers silently orphan their whole subtree
**File:** `src/services/GraphService.ts` (`_loadUsers` ~102, `_applyUserFilters` ~131, `_buildMaps` ~517)

User filters (`excludedPatterns`, `hideNoJobTitle`, `hideDisabledAccounts`, etc.) remove users from the flat list, but `_buildMaps` then builds parent→child maps from the surviving users only. Reports whose manager was filtered out keep a `_pendingManagerIds` entry pointing at a non-existent user, so their entire branch disappears from the org chart (they remain in the directory). Example: hiding accounts without a job title hides a director *and* unlinks their 40-person org.

**Fix:** Bridge the chain in `_buildMaps`. Approach:
1. Run `_applyUserFilters` as today to get the visible set, but keep the pre-filter list accessible inside `_buildMaps` (change `_loadUsers` to pass both, or store the raw list on a private field before filtering).
2. Build a raw manager map (userId → managerId) from **all** users. For each visible user whose resolved manager is not visible, walk up the raw chain (with a visited-set guard) until reaching a visible manager or running out; use that as the effective manager.
3. `getManagerChain` then works unchanged because `_managerMap` now only contains visible-to-visible edges.

**Acceptance:** In demo mode, add an exclusion pattern matching a mid-level manager (e.g., `an.gupta`) in the property pane. Their reports (Benjamin Chang etc.) now appear under the next visible manager (Kevin Park) instead of vanishing from the org chart.

### 1.6 Inconsistent fallback defaults for guest/disabled filters
**File:** `src/webparts/smartOrgChart/SmartOrgChartWebPart.ts` (~lines 106–107)

`onInit` defaults `hideGuestUsers`/`hideDisabledAccounts` to `true`, but `render()` falls back with `|| false`. Harmless today (onInit runs first) but a trap for future refactors. Change the two `render()` fallbacks to `?? true` so both places agree, or simply pass `this.properties.hideGuestUsers` directly since onInit guarantees a value. Match the pattern used for the `enable*` flags.

### 1.7 Presence failure handling marks the cache fresh
**File:** `src/services/GraphService.ts` (`getPresence` ~240–259)

When the `getPresencesByUserId` call throws (Presence.Read.All not granted), the `catch` breaks out of the chunk loop but `_presenceExpiry` is still set, and every subsequent 60-second poll retries the full failing call for all visible users. Track a private `_presenceUnavailable = true` on the first hard failure and return early on subsequent calls (reset the flag if a later call succeeds — e.g., retry at most once every 15 minutes). This stops a permanent 1-request/min failure loop against Graph for tenants that never approve the permission.

**Acceptance:** In a code walkthrough (no live tenant needed): first failure sets the flag; subsequent `getPresence` calls within the back-off window return the empty cache without a network call.

---

## Phase 2 — Performance

### 2.1 Parallelize `_supplementUnlicensedReports`
**File:** `src/services/GraphService.ts` (~line 331)

The SharePoint Search path issues **one awaited Graph request per distinct manager, serially**. In a 2,000-person org with ~200 managers that is 200 sequential round trips (easily 60+ seconds) before first render. Process managers with a concurrency pool of ~8 (slice into chunks and `Promise.all` each chunk, same pattern as `_checkFrontierNodes` in OrgChart.tsx). Keep the per-request `try/catch` so one bad manager doesn't fail the batch. Note: `knownIds`/`newUsers` mutation is safe under `Promise.all` because the callbacks are synchronous after the await, but double-check no duplicate can slip in when two managers share an unlicensed report — dedupe via `knownIds` check immediately before push.

### 2.2 Batch photo setState in EmployeeDirectory
**File:** `src/webparts/smartOrgChart/components/EmployeeDirectory/EmployeeDirectory.tsx` (`_drainPhotoQueue` ~143)

One `setState` per photo means up to 200 re-renders when a page loads. OrgChart already solves this (`_loadPhotos`, OrgChart.tsx:1033) by batching flushes every 10 photos — apply the same pattern here.

### 2.3 Make Expand All injection linear instead of quadratic
**File:** `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (`_handleExpandLoaded` ~1103)

Each `injectChildren` call clones the entire tree, and it's called once per frontier node (plus `expandLoaded` per batch). For a 1,000-person org this is O(n²) clones. Refactor to inject a whole batch in one traversal: build a `Map<string, IOrgNode[]>` of nodeId → children for the batch, then one recursive clone that consults the map (`injectChildrenBatch(root, map)`). Keep `injectChildren` for the single-node `_handleToggle` case or reimplement it via the batch helper.

**Acceptance:** Demo mode, 1,000-person dataset, Expand All completes without noticeable UI freeze; result tree identical to before (spot-check a few branches).

### 2.4 Memoize per-render tree scans
**File:** `src/webparts/smartOrgChart/components/OrgChart/OrgChart.tsx` (`render` ~1702–1707)

`countTreeUsers`, `getUniqueDepts`, and `computeStats` walk the full tree / user list on **every** render (including every mousemove-driven pan state change — note `isDragging` state flips re-render the component). Cache them: compute in a private field keyed by the `rootNode`/`allUsers` reference (e.g., `if (this._statsCacheFor !== rootNode) { recompute; this._statsCacheFor = rootNode; }`). Low effort, removes repeated O(n) scans during interaction.

### 2.5 Debounce presence refresh on directory filter typing
**File:** `src/webparts/smartOrgChart/components/EmployeeDirectory/EmployeeDirectory.tsx` (`componentDidUpdate` ~85)

Every search keystroke triggers `_refreshPresence()`. The TTL cache usually makes it cheap, but new users on the filtered page still cause Graph calls per keystroke. Debounce the presence refresh ~500 ms (store a timeout handle, clear on the next update and on unmount).

---

## Phase 3 — Dead code removal

All confirmed unused via grep; each removal must keep `gulp build` green.

1. **`exportOrgChartToPng` + `html2canvas`** — PNG export was removed from the UI in v1.1.0, but the function and its top-level `import html2canvas from 'html2canvas'` remain in `src/services/PdfExportService.ts:1,143-160`, keeping html2canvas (~48 KB gzipped) in the shipped bundle. Delete the function and the import, remove `"html2canvas"` from `package.json` dependencies, and run `npm install` to update the lockfile.
2. **`exportVCard`** — the vCard button was removed from the profile card in v1.2.1; the function at `PdfExportService.ts:162-180` is unreferenced. Delete it.
3. **`EmployeeDirectory.exportPdf`** (`EmployeeDirectory.tsx:114-117`) and **`exportDirectoryToPdf`** (`PdfExportService.ts:58-90`) — no UI calls either (the directory toolbar only exports CSV). *Decision required:* the cleaner option is to delete both; alternatively add a PDF button next to the CSV export button. Default to **deleting** unless the maintainer says otherwise.
4. **`_directoryRef` / `_orgChartRef`** in `SmartOrgChart.tsx:88-89,277,294` — created and attached but never read. Remove the fields and the `ref=` props.
5. **`chartLayout` prop on `OrgTree`** (`OrgChart.tsx:552,602`) — threaded through recursion but never used for rendering. Remove from the interface and both call sites.
6. **`showEmail`/`showPhone` on `IOrgChartProps`** (`IOrgChartProps.ts:9-10`) — passed from `SmartOrgChart.tsx:300-301` but never read inside OrgChart (cards deliberately don't show email/phone; the profile popup always shows them). Remove from the interface and the parent's JSX.
7. **Always-true conditions** — `count >= 0` in `FilterPanel` (`OrgChart.tsx:410`) and `i < ancestorChain.length` in the ancestor strip (`OrgChart.tsx:2022`). Simplify both (the chevron after each ancestor is intentional — just drop the redundant condition).
8. **`matchesQuery` vs `matchUserQuery`** (`OrgChart.tsx:75-94`) — duplicate logic. Implement `matchesQuery(node, q)` as `matchUserQuery(node.user, q)`.

---

## Phase 4 — Polish, accessibility, consistency

### 4.1 American English spelling sweep
Per project convention, fix British spellings (do not touch `.claude/commands/release-prep.md` — different project scope):
- `CHANGELOG.md:50` "organisation" → "organization"; `CHANGELOG.md:51` "centre" → "center"
- `README.md:32` "organisation" → "organization"
- `OrgChart.tsx:220` comment "Coloured" → "Colored"

### 4.2 Popup keyboard accessibility
**File:** `OrgChart.tsx`

- The layout picker, department filter, and user filter popups close only via backdrop click. Add an Escape-key handler (a single `keydown` listener that closes whichever popup is open — mirror the pattern in `PersonCard`).
- `PersonCard` overlay: add `role="dialog"`, `aria-modal="true"`, and an `aria-label` with the person's name. Move initial focus to the close button on open (a `ref` + `focus()` in an effect) so keyboard/screen-reader users land inside the dialog.

### 4.3 Directory table semantics
**File:** `EmployeeDirectory.tsx` (`_renderListView` ~320)

Add `scope="col"` to the `<th>` elements.

### 4.4 CSV BOM literal
**File:** `PdfExportService.ts:29`

`const CSV_BOM = '﻿'` holds an invisible literal BOM character — editors/formatters can silently strip it. Replace with the explicit escape: `const CSV_BOM = '﻿';`.

### 4.5 Print window robustness
**File:** `PdfExportService.ts` (`openPrintWindow` ~49)

`win.onload` after `document.write` is unreliable in some browsers (the document may already be "loaded", so print never fires). After `win.document.close()`, call print directly with a short fallback:
```ts
win.focus();
const doPrint = () => { try { win.print(); } catch { /* ignore */ } };
if (win.document.readyState === 'complete') setTimeout(doPrint, 100);
else win.onload = doPrint;
```

### 4.6 Restore collapse state after clearing search
**File:** `OrgChart.tsx` (`_onSearchChange` ~927)

Typing a search expands all branches containing matches, but clearing the search leaves everything expanded. Snapshot the pre-search `rootNode` (private field) when a search begins (transition empty → non-empty) and restore it when the query transitions back to empty, unless the tree object changed in between (compare the stored reference against a second field tracking the last non-search root; if the user expanded/collapsed nodes during the search, prefer keeping the current state — simplest correct rule: only restore if `rootNode` was untouched except by `expandToMatches`, which you can track with a boolean flag set in `_onSearchChange` and cleared by any other rootNode mutation).
If that bookkeeping proves too invasive, the fallback is acceptable UX: on clearing the search, run `collapseAll` + re-expand to the initially loaded depth via `expandLoaded`. Choose whichever is simpler to implement cleanly.

### 4.7 MockGraphService id normalization
**File:** `MockGraphService.ts` (`getManagerChain` ~607, `getDirectReports` ~594, `hasDirectReports` ~602)

The base class lowercases `userId` before map lookups; the mock does not. Mock ids are already lowercase emails so it works today, but normalize (`userId.toLowerCase()`) in the mock's map lookups to keep the contract identical — this prevents demo-vs-live behavior drift if a caller ever passes mixed case.

---

## Explicitly out of scope (do not do)

- CSS `zoom` → `transform: scale()` migration (Firefox < 126 compat): `_fixConnectorLines` and auto-fit math depend on `zoom` semantics; SharePoint's supported browsers all handle `zoom`. Not worth the risk.
- Directory column sorting, new export formats, or any new features.
- Version bump / changelog entries / release packaging.
- Reworking the localStorage view persistence (admin `defaultView` being overridden by a user's stored view is intentional).

## Final verification checklist

1. `gulp build` — clean (no new lint warnings).
2. `gulp serve` demo mode:
   - Directory: search, alphabet filter, dept/office filters, pagination, CSV export, card/list toggle.
   - Org chart drill mode: drill in/out, breadcrumbs, Find Me, profile popup (Escape closes), root picker.
   - Top Down layout: Expand All on the 1,000-person dataset (fast, terminates), Collapse All, zoom in/out/reset (respects default zoom), search expands matches and clearing restores.
   - Property pane: change default layout/zoom/feature flags with the page open — all take effect live.
   - URL: no `socFocus` param at default position; appears on focus; removed on return to root.
3. `git diff --stat` review — confirm no `*.module.scss.ts` files were hand-edited and `package.json`/lockfile only lost `html2canvas`.
