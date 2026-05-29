# Smart Org Chart

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![SPFx](https://img.shields.io/badge/SPFx-1.18.2-green.svg)](https://aka.ms/spfx)
[![Node](https://img.shields.io/badge/Node-18_LTS-brightgreen.svg)](https://nodejs.org)

A SharePoint Framework (SPFx) web part that provides a searchable **Employee Directory** and an interactive **Org Chart**, both powered by Microsoft Graph.

![Smart Org Chart overview](docs/screenshots/01-overview.png)

---

## Features

| Feature | Detail |
|---|---|
| **Employee Directory** | Responsive card grid — photo, name, job title, department, office, email, phone |
| **Real-time search** | Filters by name, title, email, or department as you type |
| **Alphabet filter** | A–Z filter bar, switchable between first name or last name |
| **Org Chart** | Hierarchical tree built from the Azure AD `manager` field |
| **Three chart layouts** | Drill-down, vertical tree, and horizontal tree |
| **Lazy expand** | Child nodes are fetched on demand when you expand them |
| **Wide-org wrapping** | When a manager has 8+ direct reports the children automatically wrap into 2 rows/columns, keeping the chart readable without zooming |
| **Department & user filters** | Narrow the chart to specific departments or hide guests/members |
| **User account filters** | Admin-controlled: hide disabled accounts, hide guests, restrict to tenant domain, exclude by name pattern |
| **Statistics panel** | Headcount by department, overlaid on the chart |
| **User preferences** | Per-user settings saved to `localStorage` (card size, sort order, visible fields, compact mode, etc.) |
| **Photo support** | Profile photos from Graph with base64 caching; initials avatar fallback |
| **Export** | Download the current view as a PDF or PNG |
| **Demo mode** | Built-in mock data (150 / 500 / 1,000 people) for testing without Graph permissions |
| **Themes** | Modern, Minimal, Corporate, and Dark |
| **Configurable data source** | Graph API (live, no indexing delay), SharePoint Search, or Auto (Graph with SP Search fallback) |

---

## Screenshots

| Employee Directory | Org Chart — Drill-down |
|---|---|
| ![Directory](docs/screenshots/03-directory-overview.png) | ![Org Chart](docs/screenshots/06-orgchart-drill.png) |

| Vertical tree | Dark theme |
|---|---|
| ![Vertical](docs/screenshots/08-orgchart-vertical.png) | ![Dark](docs/screenshots/15-theme-dark.png) |

See the full **[User Guide](docs/USER_GUIDE.md)** for all features and screenshots.

---

## Prerequisites

| Requirement | Detail |
|---|---|
| **Node.js** | 18 LTS (SPFx 1.18 is not compatible with Node 20+) |
| **SharePoint** | Online (Microsoft 365) |
| **Permissions to deploy** | SharePoint site owner or higher |
| **Permissions to approve Graph API** | Global Administrator or SharePoint Administrator |

---

## Getting Started

### 1. Clone and install

```bash
git clone https://github.com/sregan1/SharePoint-Smart-Org-Chart.git
cd SharePointSmartOrgChart
npm install
```

### 2. Configure the dev workbench URL

Open `config/serve.json` and replace the placeholder with your SharePoint site URL:

```json
{
  "initialPage": "https://YOURTENANT.sharepoint.com/sites/YOURSITE/_layouts/workbench.aspx"
}
```

### 3. Start the development server

```bash
gulp serve
```

This opens the SharePoint workbench at the URL above. Add **Smart Org Chart** to the page and configure it via the property pane.

> **Tip:** If you just want to see the UI without a SharePoint tenant, use demo mode — see [Demo Mode](#demo-mode) below.

---

## Building and Deploying

### Package for production

```bash
npm run package
```

Produces `sharepoint/solution/smart-org-chart.sppkg`.

### Deploy to SharePoint

1. Go to **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**.
2. Upload `smart-org-chart.sppkg`.
3. Check **Make this solution available to all sites** for tenant-wide deployment.
4. Click **Deploy**.

### Approve Microsoft Graph permissions

The web part requests two delegated Graph permissions. A Global or SharePoint Administrator must approve them:

1. In SharePoint Admin Center go to **Advanced** → **API access**.
2. Approve:
   - `Microsoft Graph — User.Read.All`
   - `Microsoft Graph — User.ReadBasic.All`

> Without these approvals the web part loads but cannot retrieve user data.

### Add the web part to a page

1. Edit any SharePoint page.
2. Click **+** and search for **Smart Org Chart**.
3. Open the property pane and set at minimum:
   - **Top-Level User** — UPN or email of your CEO / org root (e.g. `ceo@company.com`).
4. Publish the page.

---

## Demo Mode

To try the web part without a SharePoint tenant:

```bash
npm run demo:screenshots
```

This builds a standalone webpack bundle, starts a local HTTP server, and uses Puppeteer to capture screenshots of all views. Output goes to `docs/screenshots/`.

You can also run the demo server manually and open it in a browser (see `demo/` folder).

---

## Configuration

### Property Pane (admin, per web part)

**General**

| Setting | Description |
|---|---|
| Use demo data | Renders with built-in mock people instead of Graph |
| Default view | Which view opens first — Directory or Org Chart |

**Branding**

| Setting | Description |
|---|---|
| App title | Text shown in the header alongside the current view name |
| Logo URL | Full URL to a PNG/SVG/JPG (e.g. `https://contoso.sharepoint.com/sites/mysite/SiteAssets/logo.png`) |

**Visual Style**

| Setting | Description |
|---|---|
| Chart theme | Modern, Minimal, Corporate, or Dark |
| Default layout | Drill-Down, Vertical, or Horizontal |

**Data Source**

| Option | Description |
|---|---|
| Auto (default) | Tries Graph API first; falls back to SharePoint Search if Graph is unavailable |
| Graph API | Always reads live Azure AD data — no indexing delay; recommended |
| SharePoint Search | Legacy behavior using the People Search index |

> Graph API is strongly recommended. New users and manager changes appear immediately without waiting for the SharePoint search index to update.

**User Filters**

| Setting | Default | Description |
|---|---|---|
| Exclude accounts | _(empty)_ | Comma-separated words or patterns. Any user whose name, email, or UPN contains one of these is hidden everywhere (e.g. `conf-room, noreply, service`). |
| Only show tenant users | Off | Hides accounts whose email domain does not match your tenant (removes gmail.com, hotmail.com, etc.). |
| Hide Azure AD guest accounts | On | Hides guest (B2B) accounts from all views. |
| Hide disabled accounts | On | Hides accounts with blocked sign-in (former employees, service accounts). |

**Org Chart**

| Setting | Description |
|---|---|
| Top-Level User | UPN or email of the person at the root of the chart (e.g. `ceo@company.com`). Required. |
| Levels to load below root | How many hierarchy levels to fetch on initial load (1–8). |

**Org Chart Features** — show or hide individual toolbar buttons.

**Directory**

| Setting | Description |
|---|---|
| Max employees per page | How many cards to show per page (10–200). |

### Settings Panel (per user, via ⚙ icon)

| Setting | Description |
|---|---|
| Alphabet filter field | Sort/filter by first name or last name |
| Card size | Small / Medium / Large |
| Show email | Toggle email addresses on directory and org chart cards |
| Show phone | Toggle phone numbers on directory and org chart cards |
| Show department | Toggle department badge on cards |
| Show office location | Toggle office location on directory and org chart cards |
| Manager levels shown | Levels of the reporting chain to show above a focused person |
| Compact cards | Smaller cards; useful for seeing more of the org chart at once |

User preferences are saved to `localStorage` and persist across page reloads.

---

## Project Structure

```
SharePointSmartOrgChart/
├── config/                          # SPFx build configuration
├── demo/                            # Standalone demo tool (screenshots)
│   ├── index.tsx                    # Demo entry point
│   ├── webpack.demo.config.js       # Standalone webpack config
│   └── take-screenshots.js          # Puppeteer screenshot script
├── docs/
│   ├── USER_GUIDE.md                # End-user documentation
│   └── screenshots/                 # Auto-generated screenshots
└── src/
    ├── services/
    │   ├── GraphService.ts          # Microsoft Graph calls + photo caching + tree builder
    │   └── MockGraphService.ts      # Demo data service
    └── webparts/smartOrgChart/
        ├── SmartOrgChartWebPart.ts  # Web part entry + property pane
        └── components/
            ├── SmartOrgChart.tsx    # Root component — header, view switcher
            ├── EmployeeDirectory/   # Directory view
            ├── OrgChart/            # Chart view (all three layouts)
            └── SettingsPanel/       # User preferences panel
```

---

## Technology Stack

| Technology | Version |
|---|---|
| SharePoint Framework (SPFx) | 1.18.2 |
| React | 17.0.1 |
| TypeScript | 4.7.4 |
| Fluent UI React | 8.x |
| Microsoft Graph API | v1.0 |

---

## Troubleshooting

**"Failed to load employees"** — Graph API permissions have not been approved. See [Approve Microsoft Graph permissions](#approve-microsoft-graph-permissions).

**Org Chart shows "User not found"** — Check that Top-Level User contains a valid UPN or email (e.g. `john.doe@company.com`), not a display name.

**New users or manager changes not appearing** — If you are using SharePoint Search as the data source, the search index may not have updated yet (indexing can take hours). Switch the Data Source setting to **Graph API** or **Auto** for real-time data.

**Disabled / former employees still showing** — Open the property pane → User Filters and enable **Hide disabled accounts**. This requires the Graph API data source; SharePoint Search does not expose account status.

**Photos not loading** — User photos require the `User.Read.All` scope. Verify that profile photos are set in Microsoft 365.

**Build errors after `npm install`** — Ensure you are using Node.js 18. SPFx 1.18 is not compatible with Node 20+.

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

## License

[MIT](LICENSE) © 2026 Sean Regan
