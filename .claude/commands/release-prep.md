You are preparing the **Smart Org Chart** SPFx web part for a public release. Work through each step below in order. Stop and report any blocking issues before continuing.

**Requested version:** $ARGUMENTS
(If blank, read `config/package-solution.json` for the current version, then ask the user what the new version should be before proceeding with any changes.)

---

## Step 1 — TypeScript check

Run `.\node_modules\.bin\tsc.cmd --noEmit`. If there are errors, list them and stop. Do not continue until the user confirms they are resolved.

---

## Step 2 — Determine the current feature set

Read these files:
- `src/webparts/smartOrgChart/SmartOrgChartWebPart.ts` — the property pane `getPropertyPaneConfiguration()` method lists every admin setting and feature toggle
- `src/webparts/smartOrgChart/components/SmartOrgChart.tsx` — the `currentView` state and header bar reveal the available views and user-facing features

From these, produce a definitive list of all views, admin settings, and user-facing features. Do not assume — derive it from the code.

---

## Step 3 — Update `docs/USER_GUIDE.md`

Read `docs/USER_GUIDE.md` in full. Then update it:
- Change the version number in the heading to the new version.
- Update the Table of Contents to match the actual sections you will write.
- Remove any sections that describe features no longer in the app (check against the feature list from Step 2).
- Add or update sections for features that are present in the app but missing or out of date in the guide. Use clear, user-facing language — write for a non-technical SharePoint admin audience.
- Keep accurate existing content unchanged.
- Do not remove the screenshots section; update screenshot filenames only if they change.

Write the complete updated file.

---

## Step 4 — Update `README.md`

Read `README.md` in full. Then update it:
- Update any version badge or version reference to the new version.
- Update the features table to exactly match the current feature set (no removed features, no missing new ones).
- Keep the prerequisites, installation, permissions, troubleshooting, and tech stack sections accurate.

Write the complete updated file.

---

## Step 5 — Update `CHANGELOG.md`

Run `git log --oneline $(git describe --tags --abbrev=0)..HEAD` to get commits since the last tag. If that fails (no prior tags), run `git log --oneline -20` instead.

Read `CHANGELOG.md` in full, then prepend a new entry for the new version following the existing format:
- Use today's date.
- Group changes under Added / Changed / Fixed / Removed headings as appropriate.
- Base the entries on the git log and your knowledge of changes made in this session.

Write the complete updated file.

---

## Step 6 — Security check

Before bumping versions, verify the repo is safe for public release:

1. Check `config/serve.json` — `initialPage` must be the placeholder `https://YOURTENANT.sharepoint.com/sites/YOURSITE/_layouts/workbench.aspx`, not a real tenant URL. If it contains a real URL, replace it with the placeholder.
2. Check `config/package-solution.json` — the `developer.name` field must not contain an internal organisation name. If it does, clear it to `""`.
3. Grep `src/` and `config/` for any hardcoded tenant domains (e.g. `.sharepoint.com`, `.onmicrosoft.com`) or personal email addresses. Report any findings.

Fix any issues found before continuing.

---

## Step 7 — Bump version numbers

Update exactly these two fields:
- `package.json` → `"version"`: set to the new version (e.g. `"1.2.0"`)
- `config/package-solution.json` → `"solution"."version"`: set to four-part form (e.g. `"1.2.0.0"`)

---

## Step 8 — Regenerate UserGuide.docx

Run:
```
node docs/generate-word.js
```

Report whether it succeeded or failed. If the script does not exist, note "N/A — no generate-word.js" in the checklist and continue to Step 9.

---

## Step 9 — Full release build

Run:
```
npm run package
```

This bundles in production mode and produces `sharepoint/solution/smart-org-chart.sppkg`. This may take a minute or two — wait for it to complete. Report success or any build errors. If errors occur, stop and report them — do not print the checklist until they are resolved.

---

## Step 10 — Release readiness checklist

Replace X.X.X below with the actual version number, and update the ✅/❌ status of each automated item based on what actually happened above. Then print it:

```
Release checklist for vX.X.X
─────────────────────────────────────────────────────────
Automated (completed by this command):
  ✅ TypeScript compiles cleanly
  ✅ docs/USER_GUIDE.md updated
  ✅ README.md updated
  ✅ CHANGELOG.md updated
  ✅ Security check passed (no tenant URLs or org names exposed)
  ✅ Version bumped in package.json and config/package-solution.json
  ✅ UserGuide.docx regenerated         (or ❌ N/A — no generate-word.js)
  ✅ Release build complete — smart-org-chart.sppkg ready

Still needed (manual steps):
  [ ] Review all doc changes:     git diff
  [ ] Commit everything:          git add -A && git commit -m "vX.X.X Release"
  [ ] Tag the release:            git tag vX.X.X && git push origin vX.X.X
  [ ] Push to main (triggers CI): git push origin main
  [ ] Verify GitHub Release:      https://github.com/sregan1/SharePoint-Smart-Org-Chart/releases
  [ ] Upload .sppkg to SharePoint app catalog if distributing directly
```
