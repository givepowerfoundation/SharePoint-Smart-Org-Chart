# Changelog

All notable changes to Smart Org Chart are documented here.

---

## [1.2.1] — 2026-06-16

### Changed
- **Export PDF moved to org chart toolbar** — the PDF export button is now in the org chart toolbar (where the PNG button was), keeping export controls alongside the chart rather than in the header bar.
- **Employee Directory photo sizes increased** — profile photos in the card view are now 80 px (up from 64 px) and in the list view 52 px (up from 40 px). Initials avatars and presence dots scale proportionally.

### Fixed
- **Department filter blanked the entire chart** — when a department filter was applied, the root node (e.g. CEO) was hidden because they were not in the selected department, causing the whole tree to disappear. Ancestor nodes are now kept visible when any of their descendants match the active filter.
- **Connector lines not updated after filtering** — horizontal connector bars between levels retained stale positions after a department, member, or guest filter was changed, or after the zoom level or layout was switched. `_fixConnectorLines` now runs whenever any of these values change.
- **Full-width horizontal connector flash** — the horizontal connector bar between tree levels briefly appeared full-width before JavaScript recalculated its span. The CSS default is now zero-width so there is no visible flash on initial render.

### Removed
- **Header-level PDF download button** — the Download (↓) button that appeared in the header bar next to the Settings gear has been removed. PDF export for the org chart is now available directly in the org chart toolbar.
- **Profile card action buttons simplified** — the Meet (schedule meeting), vCard (download contact), Copy Link, and Call buttons have been removed from the person profile card popup. The remaining actions are Chat in Teams, Email, and Focus.

---

## [1.2.0] — 2026-06-07

### Added
- **Export Employee Directory to CSV** — new Export button in the directory toolbar downloads the current filtered list as a UTF-8 CSV file (Excel-compatible). Exported columns match the user's visible field preferences (name, title, department, office, email, phone).
- **"View from person…" root picker** — a new search field in the org chart toolbar lets any user temporarily re-root the chart at any person in the organization. A cancel button resets back to the admin-configured root.
- **Configurable default org chart zoom** — new admin property pane setting (Default Org Chart Zoom). Options: Auto-fit (default), 50%, 75%, 100%, 125%, 150%. When a fixed zoom is set, the auto-fit behavior is suppressed.
- **Configurable font size** — text scale can now be set both by the admin (Default Font Size, in the property pane Visual Style group) and by each user (Font Size, in the Settings panel). Options range from 75% to 175%.
- **Demo dataset picker in Settings panel** — when Demo Data is enabled, the 150 / 500 / 1,000 person dataset size buttons are now accessible from the Settings panel (gear icon) instead of a banner below the header.

### Changed
- **Demo Data moved to its own property pane group** — the Use Demo Data toggle has moved from the General group to a dedicated Demo group at the bottom of the property pane.
- **Layout names standardized** — the three org chart layouts are now named *Drill-Down*, *Top Down*, and *Left to Right* consistently across the toolbar, layout picker, and property pane.
- **Page size control improved** — the Max employees per page slider in the property pane now includes a companion number input for precise values, in addition to the drag slider.
- **Statistics based on all users** — the stats bar (Total, Members, Guests, Depts) now reflects all visible users in the directory rather than being derived by walking the currently-loaded tree nodes.
- **Zoom level not persisted** — the org chart zoom level is no longer saved to the browser session. Each visit starts at the admin-configured zoom or auto-fit.

### Removed
- **"Avg span" removed from stats bar** — the average reporting span figure has been removed. The stats bar now shows Total, Members, Guests, and Depts.
- **Demo data banner removed** — the banner that appeared below the header in demo mode has been removed; dataset size controls are now in the Settings panel.

---

## [1.1.0] — 2026-06-02

### Changed
- **Zoom minimum lowered to 25%** — the org chart can now be zoomed out to 25% (previously 40%), giving more room to view large hierarchies.
- **Expand All now auto-fits zoom** — after expanding the full organization tree, the view automatically zooms to fit the newly loaded content.
- **Connector lines end at card edges** — the horizontal connector bar in the vertical tree now ends at the center of the leftmost and rightmost cards in each row, rather than extending to the full sub-tree width.
- **Large tree horizontal scrolling fixed** — the full tree is now accessible by scrolling left and right after Expand All, including nodes at the far left and right of wide org charts.

### Removed
- **Export to PNG** — the PNG image export button has been removed from the header bar. PDF export remains available.

---

## [1.0.1] — 2026-05-30

### Added
- **Room account exclusion by pattern** — new *Exclude room accounts by pattern* field in the property pane. Comma-separated patterns (e.g. `conf-, room-, mrm@`) hide matching mailboxes from all views regardless of their account-enabled status, covering room mailboxes that have sign-in enabled as well as disabled ones.
- **Office location on org chart cards** — the office/location field now appears on node cards in the tree and drill-down views when enabled in User Preferences.
- **Configurable data source** — choose between Graph API (live, no indexing delay), SharePoint Search (legacy), or Auto (Graph with SP Search fallback).
- **List view in Employee Directory** — toggle between card grid and compact single-row list view.
- **Admin user filters** — property pane options to hide Azure AD guest accounts, hide disabled accounts, restrict to tenant domain, and exclude accounts by name pattern.

### Changed
- **Consistent accent color** — card borders, avatar backgrounds, job title text, and org chart connector lines now all follow the SharePoint site's color theme instead of using per-department hashed colors.
- **Expand All loads the complete org** — clicking Expand All in tree modes now performs a full BFS load from Microsoft Graph, fetching every level of the hierarchy rather than only expanding the pre-loaded nodes.
- **Horizontal tree layout** — the left-right tree now uses the same recursive layout as the top-down tree: each manager's reports appear in a column to their right, expanding outward as branches are opened.
- **Direct report count** — the badge on the expand button always shows a person's own direct report count; it no longer changes as deeper nodes are expanded.
- **Hide disabled accounts and hide guest accounts** now default to **On** for new web part instances.
- **Profile photo sizes increased** for clearer display at larger card sizes.
- Long job titles now wrap to two lines on org chart cards instead of truncating.

### Fixed
- `hideDisabledAccounts` and `hideGuestUsers` filters were not being applied when the Graph API data source was active.
- Horizontal and vertical tree layouts now maintain strict level alignment — all nodes at the same org depth appear in the same row/column.

### Removed
- Multi-column wrapping for managers with 8+ direct reports has been removed. All direct reports now appear in a single row (vertical) or column (horizontal), consistent with a traditional org chart layout.

---

## [1.0.0] — 2025-12-01

### Added
- Initial release.
- Employee Directory with A–Z filter, real-time search, paginated card grid, three card sizes, and per-user field visibility toggles.
- Org Chart with drill-down, vertical tree, and horizontal tree layouts.
- Person profile card with presence status, manager chain, and action buttons (Chat, Email, Call, Focus).
- Org chart toolbar: search, Find Me, layout picker, stats bar, department filter, user type filter, zoom controls.
- Four themes: Modern, Minimal, Corporate, Dark.
- PDF and PNG export.
- Demo mode with 150 / 500 / 1,000 person mock datasets.
- User preferences panel persisted to `localStorage`.
- Graph API and SharePoint Search data sources with automatic fallback.
- CI release workflow producing `smart-org-chart.sppkg`.
