# VériRapport

**Automated Power BI Report Testing & Validation**

Playwright automation for Power BI reports. It exports the data behind every
visual for regression testing, and checks that a report migrated from Power BI
**Import mode** to **Direct Lake** still shows the same data.

> **New to the app?** See **[HOW-TO-USE.md](HOW-TO-USE.md)** for a first-time walkthrough.

---

## 1. First-Time Setup Guide

### Option 1 — One file, double-click (new machine, easiest)

Share **`Installer.cmd`** (email it, Teams/Slack it, drop it on a shared drive).
It is a single standalone file; the rest of the repo doesn't need to be there
yet. On the new machine, double-click it.

A console window explains what it is about to do, then:
- clones/updates the repo to `%USERPROFILE%\verirapport`
- installs Node/Git via `winget` if missing
- runs `npm install` for the root project and the desktop app, and installs the
  Playwright browser
- prompts once for `E2E_USERNAME` / `E2E_PASSWORD` (a dedicated test account, not
  your personal MFA-protected login) and saves them to
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env`
- creates a desktop shortcut to the app

It does **not** sign in or run a comparison automatically — those open a real
browser and take several minutes, so they are left for you to start. The window
stays open at the end and tells you which shortcut to double-click next.

If something fails partway, fix the reported issue and double-click
`Installer.cmd` again — steps already completed are skipped.

### Option 2 — One bootstrap command (PowerShell)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u='https://raw.githubusercontent.com/chakshugoel1/VeriRapport/main/bootstrap-install.ps1'; $p=Join-Path $env:TEMP 'bootstrap-install.ps1'; iwr -UseBasicParsing $u -OutFile $p; powershell -NoProfile -ExecutionPolicy Bypass -File $p"
```

If your organisation blocks `Invoke-WebRequest`, copy `bootstrap-install.ps1`
(or `Installer.cmd`) to the new machine and run it directly.

> **First-run security prompt:** because these files are downloaded from the
> internet, Windows SmartScreen may show "Windows protected your PC". Click
> **More info → Run anyway**. This is normal for unsigned scripts shared outside
> the Microsoft Store.

### Option 3 — Run the installer from an existing clone

```powershell
.\install.ps1
```

If script execution is blocked:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

What it does, in order:
1. **Prerequisites** — Node/Git via `winget` if missing, `npm install` for the
   root project and `app-desktop`, Playwright browser install.
2. **Secrets** — prompts for `E2E_USERNAME` / `E2E_PASSWORD` and saves them to
   `%USERPROFILE%\Power_BI_report_validation_credentials\.env`. Also creates a
   `pbi-service-principal.json` template there (optional — only needed for the
   DAX-based `main:run` flow; the parity flow works without it).
3. **Auth session** — reports whether a saved sign-in session exists. Does **not**
   open a browser to sign in unless you pass `-RunAuthSetup`.
4. **Parity smoke check** — reports readiness. Does **not** run unless you pass
   `-RunParity`.
5. **Desktop shortcut** — creates `VeriRapport.lnk` pointing at
   `app-desktop\launch-desktop.bat`. Asks where to put it (default: your Desktop).

Useful installer flags:

```powershell
.\install.ps1 -Mode verify                # check only, no installs/changes
.\install.ps1 -Mode fast                  # skips npm install / browser download
.\install.ps1 -Mode full -NonInteractive  # no prompts (unattended)
.\install.ps1 -RunAuthSetup -RunParity    # also sign in and run parity (no desktop app needed)
```

npm shortcuts (from an already-cloned repo):

```powershell
npm run install:setup
npm run install:verify
npm run install:fast
```

---

### Manual setup (if you're not using the installer)

Do these steps once per machine. Credentials (Step 4) are not needed again per run.

#### Prerequisites

- **Node.js** LTS (v20+) — [nodejs.org](https://nodejs.org)
- **Git**
- Access to the Power BI workspace(s) you want to test, with a test account

#### Step 1 — Clone the repo

```powershell
git clone https://github.com/chakshugoel1/VeriRapport.git
cd VeriRapport
```

#### Step 2 — Install dependencies

```powershell
npm install
```

#### Step 3 — Install the Playwright browser

```powershell
npm run install-browsers
```

Downloads Chromium into Playwright's local cache (once per machine).

#### Step 4 — Add your credentials

Credentials live **outside** the project folder so they are never committed or
synced:

**`%USERPROFILE%\Power_BI_report_validation_credentials\.env`**
```
E2E_USERNAME=your.name@example.com
E2E_PASSWORD=your-password
```
Use a dedicated test account, not your personal MFA-protected one. Create the
folder and file yourself if you are skipping the installer.

**`%USERPROFILE%\Power_BI_report_validation_credentials\pbi-service-principal.json`**
(optional — only for the DAX-based `main:run` flow; the parity flow falls back to
your signed-in user token)
```json
{
  "tenantId": "<tenant id>",
  "clientId": "<client id>",
  "clientSecret": "<client secret>"
}
```

> Neither file lives in the repo, and both are covered by `.gitignore`.

**Credentials are never typed into the desktop app.** The `.env` file is the only
source for them.

#### Step 5 — Sign in once

```powershell
npm run test:setup
```

This opens a browser at `https://app.powerbi.com/home`, signs in with the
credentials from Step 4, and saves the session so later runs don't sign in again.
Do this before `parity` or `main:run` from the command line.

Report IDs for a parity run are entered in the desktop app (or configured in
`tests/helpers/comparison-config.helpers.ts`) — see section 2.

---

## 2. How it works

### What it does

VériRapport opens a real Power BI report in a Playwright-controlled browser,
pulls the data out of every visual with Power BI's `exportData()` API (not
scraped text), and writes it to Excel for comparison. Two flows:

| Flow | Command | What it checks |
|---|---|---|
| **Report Parity** | `npm run parity` | Does a report migrated from **Import mode** to **Direct Lake** still show the same data as the original? |
| **Main Run** | `npm run main:run` | Does one report still show the same data as a saved baseline? (needs a report configured — see below) |

### How the report is embedded

Power BI's embed SDK expects the parent page to be on the `app.powerbi.com`
origin. `harness/harness.html` is served locally by Playwright at an
`app.powerbi.com` URL (`page.route()`), so the embedded report iframe is
same-origin and `postMessage` calls work. Details: `tests/helpers/harness.helpers.ts`.

### Report Parity — how the comparison works

1. The **source** (Import mode) and **target** (Direct Lake) report identities are
   applied one after the other to the same embed mechanism.
2. Every visual on every page being compared is exported to `expected.xlsx`
   (source) and `actual.xlsx` (target).
3. `excel-compare.helpers.ts` regroups each sheet into its individual visuals and
   compares **visual to visual**, so a visual that moved or a column that was
   renamed is not treated as a data difference.
4. `parity-summary.xlsx` is generated with these sheets:
   - **Parity Summary** — overall verdict and counts
   - **Page Comparison** — one row per report page
   - **Visual Comparison** — one row per visual, with an exact status
     (identical / header text differs / duplicated / values differ / missing / ambiguous)
   - **Differences** — rows that exist on one side only
   - **Changed Values** — which cell changed, from what to what

The Playwright test fails only on a real automation error (sign-in failed, a
report wouldn't embed, a file wasn't written). Finding that the two reports'
*data* differs is a normal result: it appears in the console, in the summary
workbook and as an annotation in the HTML report. Set `FAIL_ON_PARITY_DIFF=1` to
make a data mismatch fail the test (e.g. as a CI gate).

Full details on pairs, pages and reading results: **`REPORT-PARITY-README.md`**.
How the pieces fit together: **`architecture.md`**.

### Desktop app

`app-desktop/` is an Electron app that runs the same scripts without a terminal.
Open it with the shortcut the installer created, or:

```powershell
cd app-desktop
npm install
npm start
```

What you see:
- **Reports to compare** — enter Tenant ID / Group ID / Report ID / Dataset ID for
  the Source (Expected, Import) and Target (Actual, Direct Lake) report, plus a
  **Pair name** (it names the results folder). You can paste a Power BI report link
  above each side to fill the report, workspace and tenant IDs; the Dataset ID is
  not in the link and must be typed once. **Use Source tenant** copies the tenant
  across. **Save** / **Delete** keep named pairs (IDs and options only, never
  credentials). An optional **Case-insensitive compare** checkbox treats labels
  that differ only by case or separators as equal.
- **Run** — **▶ Run comparison** exports both reports and compares them. If there
  is no saved sign-in session it signs in first automatically. **■ Stop** ends the
  whole run including the browser. **Sign in again** re-does sign-in on demand.
  Progress and elapsed time show while it runs; the raw log is under
  **Technical details**.
- **Results** — a verdict (PASS / FAIL / NOT COMPARABLE), a per-page breakdown, and
  buttons to open the summary and the results folder. **Recent runs** lists the
  last runs.
- **Filters (optional)** — a collapsed section. **Load filters** reads the
  reports' slicers (it only reads); values you pick are applied to both reports on
  your next run. It needs a saved sign-in session.

Values entered in the app override the pair in
`tests/helpers/comparison-config.helpers.ts` for that run only; nothing in the
repo is edited. See **`app-desktop/README.md`**.

### Command reference

| Command | Purpose |
|---|---|
| `npm run test:setup` | Sign in and save the session (run first, and again if it expires) |
| `npm run parity` | Export source + target and compare |
| `npm run parity:source` | Export only the source (refresh the baseline) |
| `npm run parity:target` | Export only the target and compare with the existing baseline |
| `npm run main:run` | Export one report and compare with the saved baseline |
| `npm run compare:excel` | Standalone Excel-to-Excel comparison (no browser; set `COMPARE_FILE_A` / `COMPARE_FILE_B`) |
| `npm run discover:slicers` | List a report's filters/slicers — see the header comment in `tests/specs/discover-slicers.spec.ts` for the `DISCOVER_*` variables |
| `npm run show:report` | Open the last Playwright HTML report |
| `npm run test:unit` | Browser-free unit tests |
| `npm run typecheck` | `tsc --noEmit` for the whole project |
| `npm run install-browsers` | Install the Playwright Chromium browser |

With several pairs in `comparison-config.helpers.ts`, pick one by name:
```powershell
$env:PAIR='My Report'; npm run parity
```
(the name must match the pair's `name` exactly, ignoring case; if `PAIR` is not
set, the first pair is used).

### Settings (environment variables)

Set these in your shell, or in the `.env` file next to the credentials.

| Variable | Used for |
|---|---|
| `E2E_USERNAME`, `E2E_PASSWORD` | Test-account sign-in (`.env` only) |
| `PBI_REPORT_URL` | Report for `main:run` (or set `PBI_GROUP_ID` + `PBI_REPORT_ID`) |
| `PBI_TENANT_ID`, `PBI_GROUP_ID`, `PBI_REPORT_ID`, `PBI_DATASET_ID` | Alternative to the URL for `main:run` |
| `PBI_RLS_ROLE` | RLS role for `main:run`, only if the dataset requires one (default: none) |
| `PBI_WORKSPACE_URL` | Optional sign-in landing page; default is `https://app.powerbi.com/home` |
| `PAIR` | Choose a configured pair for parity |
| `FAIL_ON_PARITY_DIFF=1` | Fail the parity test on a data mismatch |
| `COMPARE_FILE_A`, `COMPARE_FILE_B` | Files for `compare:excel` |
| `DISCOVER_*` | Options for `discover:slicers` |

(`PBI_UI_RUNTIME_*`, `PBI_SLICER_*` and `PBI_COMPARE_*` are set by the desktop app
for its own runs.)

### Where results land

- `playwright-report-parity/<pair-name>/` — `expected.xlsx`, `actual.xlsx`,
  `parity-summary.xlsx`, plus `parity-result.json` and `discovered-slicers.json`
- `playwright-report-main/` — Main Run output (`main-run.xlsx`,
  `expected-values.xlsx`, `execution-summary.xlsx`, `main-run.json`)
- `playwright-report/` — Playwright HTML report
- `test-results/` — traces, screenshots and videos for failed runs

All of these are local, git-ignored output folders — safe to delete; they are
recreated on the next run.

### Troubleshooting

- **`npm run test:setup` fails immediately** — check that
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env` exists with
  `E2E_USERNAME` and `E2E_PASSWORD` set.
- **Desktop app: Run comparison opened a sign-in browser first** — expected: no
  saved session was found, so it signed in before comparing.
- **Load filters says you are not signed in** — click **Run comparison** or
  **Sign in again** once first.
- **A run fails with HTTP 401/403 fetching report metadata** — if the error
  mentions the service principal, check `pbi-service-principal.json`; that file is
  optional for the parity flow, so a missing one is not itself the problem.
- **A command-line parity run stops before opening a browser** — the pair in
  `comparison-config.helpers.ts` has blank IDs. Fill them in, or use the desktop app.
- **Chromium not found** — run `npm run install-browsers` again.
