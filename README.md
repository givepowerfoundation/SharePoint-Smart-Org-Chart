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
| **Department & user filters** | Narrow the chart to specific departments or hide guests/members |
| **Statistics panel** | Headcount by department, overlaid on the chart |
| **User preferences** | Per-user settings saved to `localStorage` (card size, sort order, visible fields, etc.) |
| **Photo support** | Profile photos from Graph with base64 caching; initials avatar fallback |
| **Export** | Download the current view as a PDF or PNG |
| **Demo mode** | Built-in mock data (150 / 500 / 1,000 people) for testing without Graph permissions |
| **Themes** | Modern, Minimal, Corporate, and Dark |

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

| Setting | Description |
|---|---|
| Default view | Which view opens first — Directory or Org Chart |
| Top-Level User | Root of the org chart (UPN or email) |
| Company name | Displayed in the header |
| Company logo URL | Full URL to a PNG/SVG/JPG (e.g. `https://contoso.sharepoint.com/sites/mysite/SiteAssets/logo.png`) |
| Levels to load | How many levels below root to fetch (1–8) |
| Manager levels shown | How many levels above a focused user to show (0–5) |
| Card size | Small / Medium / Large |
| Directory page size | How many people per page (10–200) |
| Use demo data | Renders with built-in mock people instead of Graph |

### Settings Panel (per user, via ⚙ icon)

| Setting | Description |
|---|---|
| Alphabet filter field | Sort/filter by first name or last name |
| Show email | Toggle email addresses on cards |
| Show phone | Toggle phone numbers on cards |
| Show department | Toggle department on cards |
| Show office location | Toggle office location on cards |

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

**Photos not loading** — User photos require the `User.Read.All` scope. Verify that profile photos are set in Microsoft 365.

**Build errors after `npm install`** — Ensure you are using Node.js 18. SPFx 1.18 is not compatible with Node 20+.

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

## License

[MIT](LICENSE) © 2026 Sean Regan
