# Smart Org Chart — User Guide
**Version 1.3.0**

Smart Org Chart is a SharePoint web part that gives your organization two complementary views of its people data: a searchable **Employee Directory** and an interactive **Org Chart**, both powered by Microsoft Graph.

---

## Contents

1. [Getting Started](#1-getting-started)
2. [Header Bar](#2-header-bar)
3. [Employee Directory](#3-employee-directory)
4. [Org Chart](#4-org-chart)
   - [Drill-Down Layout](#drill-down-layout)
   - [Tree Layouts (Vertical & Horizontal)](#tree-layouts-vertical--horizontal)
   - [Toolbar](#org-chart-toolbar)
5. [Person Profile Card](#5-person-profile-card)
6. [User Preferences](#6-user-preferences)
7. [Exporting](#7-exporting)
8. [Admin Configuration](#8-admin-configuration)
9. [Themes](#9-themes)
10. [Tips](#10-tips)

---

## 1. Getting Started

Add the **Smart Org Chart** web part to any modern SharePoint page. On first use, an admin must open the property pane (pencil icon → edit web part) and set the **Top-Level User** field to the UPN or email of the person who should appear at the root of the org chart (typically the CEO or department head).

![Overview of the Smart Org Chart web part showing the header bar and a vertical org chart tree](screenshots/01-overview.png)

---

## 2. Header Bar

The header bar is always visible at the top of the web part.

![Header bar with logo, view title, and action buttons](screenshots/02-header-bar.png)

| Element | Description |
|---|---|
| **Logo** | Company logo configured by the admin (optional) |
| **App / View title** | Company name (if set) and the name of the current view |
| **View toggle** | Switches between Employee Directory and Org Chart |
| **Settings gear** | Opens the [User Preferences](#6-user-preferences) panel |

---

## 3. Employee Directory

The Employee Directory displays your organization's people as a paginated grid of cards.

![Employee Directory showing the A–Z filter bar, search box, and employee cards](screenshots/03-directory-overview.png)

### A–Z filter

Click any letter in the alphabet bar to filter the list to employees whose first name (or last name, depending on your [preference](#6-user-preferences)) starts with that letter. Click the active letter again, or click **All**, to clear the filter.

![Alphabet filter bar with the letter S selected](screenshots/04-directory-alpha-filter.png)

### Search

Type in the search box above the alphabet bar to filter by name, job title, department, or email. Results update as you type and can be combined with the A–Z filter.

### Employee cards

Each card shows the employee's profile photo (or their initials if no photo is available), name, job title, and any additional fields you have enabled in [User Preferences](#6-user-preferences): email address, phone number, department, and office location.

![Close-up of an employee card with photo, name, title, email and department](screenshots/05-directory-card.png)

Click any card to open the [Person Profile Card](#5-person-profile-card).

### List view

Toggle between the card grid and a compact list view using the view buttons above the alphabet bar.

![Employee Directory in list view showing a compact single-row layout per person](screenshots/17-directory-list.png)

### Card sizes

Three card sizes are available in User Preferences:

| Size | Best for |
|---|---|
| **Small** | Browsing a large number of people at once |
| **Medium** | Balanced view (default) |
| **Large** | Accessibility or touch-friendly use |

### Pagination

Navigation arrows and a page indicator appear at the bottom of the directory when there are more employees than the configured page size (set by the admin in the property pane).

---

## 4. Org Chart

The Org Chart displays your organization's reporting hierarchy. Three layouts are available, switchable from the toolbar.

### Drill-Down Layout

The default layout. Shows one level of the hierarchy at a time, starting from the root person.

![Drill-down org chart showing the CEO card and their direct reports in a grid](screenshots/06-orgchart-drill.png)

**Navigating:**

- **Click a person's card** (or their expand button) to drill into their direct reports.
- The **breadcrumb bar** at the top shows your current path. Click any name in the breadcrumb to jump back up.
- Click the **house icon** at the left of the breadcrumb to return to the root.

![Drill-down breadcrumb showing CEO → VP Engineering → Director Backend](screenshots/07-orgchart-drill-breadcrumb.png)

Each card shows the number of direct reports. Cards with no reports open the [Person Profile Card](#5-person-profile-card) when clicked.

### Tree Layouts (Vertical & Horizontal)

The vertical and horizontal tree layouts display the full loaded hierarchy simultaneously.

![Vertical tree org chart showing the full hierarchy from CEO downwards](screenshots/08-orgchart-vertical.png)

![Horizontal tree org chart showing the full hierarchy spreading left-to-right](screenshots/18-orgchart-horizontal.png)

**Expanding and collapsing:**

- Click the **chevron button** at the bottom of a card to expand or collapse that branch.
- Use **Expand All** / **Collapse All** in the toolbar to expand or collapse every node at once. **Expand All** fetches the complete organization, loading any levels that have not yet been retrieved from Microsoft Graph.

**Panning:** Click and drag anywhere on the chart canvas to scroll.

**Zooming:** Use the **+** / **−** buttons in the toolbar (bottom-right) or the reset button to return to the default zoom level (100%, or whatever fixed zoom your admin has configured).

### Org Chart Toolbar

The toolbar appears above the chart and contains the following controls (some may be hidden by the admin):

![Org chart toolbar showing search, Find Me, layout picker, stats, filters and zoom](screenshots/09-orgchart-toolbar.png)

| Control | Description |
|---|---|
| **Search box** | Search all people in the org. A live dropdown shows up to 8 matching people; click a result to focus the chart on that person. In tree mode, matching cards are highlighted in the chart. |
| **Expand All / Collapse All** | Available in tree modes only. Expand All loads and shows the entire organization. |
| **Find Me** | Centers the org chart on your own profile. |
| **View from person…** | Type a name to temporarily re-root the chart at any person. The chart reloads from that person downward. Click the × to return to the default root. |
| **View** | Opens the layout picker to switch between Drill-Down, Top Down, and Left to Right. |
| **Stats bar** (chart icon) | Toggles a summary bar showing total people, members, guests, and departments. |
| **Department filter** (tools icon) | Shows checkboxes to show only selected departments. A badge on the button shows how many filters are active. |
| **Country filter** (globe icon) | Shows checkboxes to show only selected countries, each with its color swatch. Hidden if no one in the chart has a country set. |
| **Employee type filter** (funnel icon) | Shows checkboxes to show only selected employee types (for example Employee, Contractor, PTE). People with no employee type set are grouped under **Not set**. Hidden if no one in the chart has an employee type set. |
| **Export PDF** (PDF icon) | Downloads the currently visible org chart as a PDF. |
| **Export CSV** (Excel icon) | Downloads the currently visible org chart as a CSV spreadsheet. |
| **Zoom controls** | Visible in tree modes only. Adjust the scale from 25% to 150%. |

#### Stats bar

![Stats bar showing 312 people, 290 members, 8 guests, 6 departments](screenshots/10-orgchart-stats.png)

#### Department filter

![Department filter dropdown showing checkboxes for Engineering, HR, Finance, Sales](screenshots/11-orgchart-deptfilter.png)

#### Country and employee type filters

The country and employee type filters work exactly like the department filter: check one
or more values to narrow the chart, and ancestors of matching people stay visible so you
can still see the reporting chain above them. All three can be combined — for example
Engineering **and** Kenya **and** Contractor. With nothing checked, everyone is shown.

Both read attributes from Azure AD (`country` and `employeeType`) and are only available
when the web part is configured to use the Graph API data source. If your organization
doesn't fill these fields in, the buttons don't appear at all.

---

## 5. Person Profile Card

Click any employee card (in either the Directory or Org Chart) to open their full profile.

![Person profile card showing photo, name, title, department badge, email, phone, office, manager chain and action buttons](screenshots/12-person-card.png)

The profile card shows:

- **Profile photo** and **presence status** (Available, Busy, Away, etc.) pulled from Microsoft Teams
- **Name, job title, and department**
- **Email, phone numbers, and office location**
- **Disabled / Guest badges** where applicable
- **Reports to** — the person's manager chain up to 8 levels, shown as clickable chips. Click any chip to jump to that manager.
- **Action buttons:** Chat in Teams, Email, and Focus (re-centers the org chart on this person)

Press **Escape** or click the overlay to close the card.

---

## 6. User Preferences

Click the **gear icon** in the header bar to open the Preferences panel. All settings here are personal — they are saved to your browser and do not affect other users.

![User Preferences panel open showing all settings](screenshots/13-preferences-panel.png)

### Directory

| Setting | Options | Description |
|---|---|---|
| **Alphabet Filter By** | First Name, Last Name | Which part of the name the A–Z bar filters on |
| **Card Size** | Small, Medium, Large | Controls the size of employee cards |
| **Font Size** | 75% – 175% | Scales all text in the web part. Useful for accessibility or high-density displays. |

### Cards & Fields

These toggles control what appears on employee cards in **both** the Employee Directory and the Org Chart node cards.

| Setting | Default | Description |
|---|---|---|
| **Email address** | On | Show the email address on each card |
| **Phone number** | On | Show the phone number on each card |
| **Department** | On | Show the department badge on each card |
| **Office location** | On | Show the office location on each card |

### Org Chart

| Setting | Default | Description |
|---|---|---|
| **Manager levels shown** | 1 | When you focus the chart on a specific person, how many levels above them to show in the ancestor breadcrumb trail (0–5) |
| **Compact cards** | Off | Reduces card height; useful for seeing more of the chart at once |

Click **Save** to apply changes. **Cancel** discards unsaved changes.

> Your preferences (including the last active view, chart layout, and filters) are remembered across page reloads.

---

## 7. Exporting

### Export Org Chart to PDF

Click the **PDF** button in the org chart toolbar to export the currently visible chart as a PDF.

- **Drill-Down mode:** exports the current person's header and their direct reports.
- **Tree modes (Top Down / Left to Right):** exports all currently loaded and expanded nodes.

> To include deeper levels in the export, click **Expand All** in the toolbar before exporting.

### Export Directory to CSV

Click the **Export** button in the Employee Directory toolbar to download the current filtered list as a CSV file that opens directly in Excel. The exported columns match your visible field preferences (name, title, department, office, email, phone).

### Export Org Chart to CSV

Click the **Excel** icon button in the org chart toolbar to download the currently visible org chart tree as a CSV file.

---

## 8. Admin Configuration

Open the SharePoint property pane (**Edit** → click the web part → pencil icon on the side) to configure the web part for your site. These settings apply to everyone viewing the page, and most take effect immediately as you change them — no page reload required.

![SharePoint property pane showing all configuration groups](screenshots/14-property-pane.png)

### General

| Setting | Description |
|---|---|
| **Default View** | Which view opens when a user first loads the page (Employee Directory or Org Chart). Users can switch views and their choice is remembered. |

### Branding

| Setting | Description |
|---|---|
| **App Title** | Text shown in the header bar alongside the current view name (e.g. "Contoso"). |
| **Logo URL** | Full URL to a PNG, SVG, or JPG logo file. Open the image in your browser and copy the address bar URL. SharePoint "Copy link" sharing URLs will not work. |
| **Country pill colors** | Optional. Sets a specific color per country for the country pill on employee cards, one `Country=#hex` mapping per line (for example `Kenya=#ffe08a`). Country names must match the value in Azure AD. Any country you don't list gets an automatic color, so this is only needed where you want a particular color. The pill is filled with your color and labelled in black, so pick light or mid-tone colors — a very dark color makes the label hard to read. |

### Page placement

The web part can be added to an ordinary page section or to a **full-width column** section.
Full width gives the org chart the entire browser width, which helps with large
organizations. Add a full-width column section to the page first, then add Smart Org Chart
to it. Not every page template offers a full-width column — it is generally available on
Communication site pages.

### Visual Style

| Setting | Options | Description |
|---|---|---|
| **Chart Theme** | Modern, Minimal, Corporate, Dark | Color scheme for employee cards and the org chart |
| **Default Font Size** | 75% – 175% | Starting text scale for all users. Users can override this in their own Preferences. |
| **Default Org Chart Layout** | Drill-Down, Top Down, Left to Right | The layout shown when a user opens the chart for the first time. Their choice is remembered after that. |

### Data Source

Controls where user and org data is loaded from.

| Option | Description |
|---|---|
| **Auto** (default) | Tries Microsoft Graph first; falls back to SharePoint Search if Graph is unavailable. Recommended for most tenants. |
| **Graph API** | Always reads directly from Azure Active Directory. New users and manager changes appear immediately — no indexing delay. |
| **SharePoint Search** | Uses the SharePoint People Search index. Changes can take hours to appear. Provided for backwards compatibility. |

> **Graph API is strongly recommended.** It is real-time and ensures that new starters, leavers, and reporting-line changes are reflected immediately.

### User Filters

These filters are applied globally — hidden users do not appear in the directory, the org chart, stats, or search results.

| Setting | Default | Description |
|---|---|---|
| **Exclude accounts** | _(empty)_ | Comma-separated words or patterns (case-insensitive). Any user whose display name, email, or UPN contains one of these is hidden. Useful for removing shared mailboxes (`noreply`), service accounts (`svc-`), or room mailboxes (`conf-`). |
| **Only show tenant users** | Off | When on, hides accounts whose email domain does not match your organization's domain. Removes external gmail.com, hotmail.com, and other personal email accounts. |
| **Hide Azure AD guest accounts** | **On** | Hides B2B guest accounts (users from other organizations invited to your tenant). |
| **Hide disabled accounts** | **On** | Hides accounts with blocked sign-in — typically former employees or deactivated service accounts. Requires the Graph API data source. |
| **Hide accounts without a job title** | Off | Hides any account that has no job title set in Azure AD. Useful for removing shared mailboxes and system accounts that don't have profile data. |
| **Hide accounts without a department** | Off | Hides any account that has no department set in Azure AD. |

> The **Hide disabled accounts** and **Hide guest accounts** filters are **on by default** for new web part instances.

> If a hidden account had direct reports, those reports are automatically re-linked to the next visible manager above them — hiding an account never orphans part of the org chart.

### Org Chart

| Setting | Description |
|---|---|
| **Top-Level User** | UPN or email of the person at the root of the org chart (e.g. `ceo@company.com`). Required. |
| **Levels to load below root** | How many hierarchy levels to fetch on initial load (1–8). Higher values load more data upfront; lower values are faster but require more click-through to explore deep branches. |
| **Default Org Chart Zoom** | Starting zoom level for the org chart: Auto-fit (default), 50%, 75%, 100%, 125%, or 150%. When set to a fixed value, the chart opens at that zoom rather than fitting to the visible area. |

### Org Chart Features

These toggles show or hide individual toolbar buttons and features. Hide controls that are not relevant to your users to simplify the interface.

| Setting | Default |
|---|---|
| Find Me button | On |
| Layout toggle | On |
| Org stats bar | On |
| Department filter | On |
| Country filter | On |
| Employee type filter | On |

### Directory

| Setting | Description |
|---|---|
| **Max employees per page** | How many employee cards to show per page (10–200). |

### Demo

| Setting | Description |
|---|---|
| **Use Demo Data** | When on, replaces live Microsoft 365 data with sample employees. Useful for testing or demonstrations without real user data. The dataset size (150 / 500 / 1,000 people) can be changed from the Settings panel. If **Top-Level User** is empty, turning this on fills it in with the sample CEO automatically. |

---

## 9. Themes

The **Chart Theme** property controls the visual appearance of employee cards and the org chart. All four themes support the same set of features. Card accent colors and connector lines follow your SharePoint site's color theme automatically.

| Theme | Description |
|---|---|
| **Modern** (default) | White cards with accent colors matching your SharePoint site theme |
| **Minimal** | Clean flat design with a subtle left-border accent; low visual noise |
| **Corporate** | Unified blue palette, consistent with a formal enterprise look |
| **Dark** | Dark navy background, ideal for digital signage or low-light environments |

![Side-by-side comparison of Modern, Minimal, Corporate and Dark themes](screenshots/15-themes.png)

---

## 10. Tips

- **Find someone fast:** Use the search box in the org chart toolbar — it searches across the entire organization, not just the loaded levels.
- **Explore a sub-tree:** Use **View from person…** in the org chart toolbar to temporarily re-root the chart at any person. Great for exploring a specific department or team without changing the admin configuration.
- **See the whole org:** Click **Expand All** in the tree toolbar. The web part will load every level from Microsoft Graph and expand them automatically.
- **Navigate deep hierarchies:** In drill-down mode, click through cards to explore; use the breadcrumb to jump back up without reloading the tree.
- **Compare departments:** Use the department filter to isolate one or more teams in the org chart. Ancestor nodes remain visible so you can see the full reporting chain above any matching person.
- **Presence badges:** Presence status (green = Available, red = Busy, yellow = Away) updates every 60 seconds from Microsoft Teams.
- **Export the directory:** Click **Export** in the directory toolbar to download the current filtered list as a CSV file for use in Excel.
- **Export the org chart:** Use the **PDF** button in the org chart toolbar to save the currently visible chart. For a complete export, click **Expand All** first.
- **Hide room mailboxes:** If meeting rooms appear in the directory, add their naming pattern (e.g. `conf-`) to **Exclude accounts** in the property pane.
