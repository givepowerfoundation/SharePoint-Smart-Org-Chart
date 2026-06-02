# Changelog

All notable changes to Smart Org Chart are documented here.

---

## [1.1.0] — 2026-06-02

### Changed
- **Zoom minimum lowered to 25%** — the org chart can now be zoomed out to 25% (previously 40%), giving more room to view large hierarchies.
- **Expand All now auto-fits zoom** — after expanding the full organisation tree, the view automatically zooms to fit the newly loaded content.
- **Connector lines end at card edges** — the horizontal connector bar in the vertical tree now ends at the centre of the leftmost and rightmost cards in each row, rather than extending to the full sub-tree width.
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
