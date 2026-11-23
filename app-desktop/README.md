# VériRapport

A small Electron app that wraps the project's existing Playwright scripts
(`npm run test:setup`, `npm run parity`, `npm run discover:slicers`) behind a
GUI — for running things without a terminal. It's a wrapper, not a separate
implementation: everything it does ultimately runs the same scripts you could
run from PowerShell yourself.

## What changed in the friendlier UI

- **Saved pairs and last form** - the form is restored on launch; "Save" keeps a named pair (IDs, pages, options) you can reload from the dropdown. Stored in the app's own data folder (`ui-state.json`), IDs only, never credentials.
- **Paste a Power BI link** into the box above each side to fill report, workspace and tenant IDs.
- **One Run button** - signs in first automatically when there is no saved session. "Sign in again" is separate.
- **Stop** ends the whole run, including the browser it started.
- **Live progress** (page X of N, elapsed time) taken from the run's own log lines; the raw log is under "Show detailed log".
- **Clear errors** - missing or badly formatted IDs are listed next to the Run button and the field is highlighted.
- **After a run** - verdict, per-page results, a "where the time went" breakdown, buttons to open the summary or results folder, and a Recent runs list.
- Filters moved into an optional section below.

## What the UI does

- **Report identities** — enter the Source and Target report's Tenant ID /
  Group ID / Report ID / Dataset ID directly. This overrides
  `tests/helpers/comparison-config.helpers.ts`'s hardcoded `PAIRS` entry for
  that run only, via a temporary runtime-config file
  (`.runtime/parity-runtime-config.json`) — nothing is edited or saved
  permanently in the repo.
- **Discover Filters** — reads report-level ("global") filters instantly. For
  page-level slicers: leave the Pages field blank and it scans pages in
  order, crawling only the first one that actually has slicers (most reports
  repeat the same fields across pages, so one representative page is usually
  enough — this avoids a second full-report embed). Type specific page names
  to crawl those instead.
- **Run Authentication** — runs `npm run test:setup`, opening a real browser
  to log in and save the session.
- **Run Parity** — runs `npm run parity`. Checks for a saved login session
  first; if there isn't one yet, it transparently runs Authentication first
  instead of failing partway through a run with a confusing token error.
- **Run Auth + Parity** — runs both in sequence unconditionally.

## Credentials are never entered here

`E2E_USERNAME` / `E2E_PASSWORD` live **only** in
`%USERPROFILE%\Power_BI_report_validation_credentials\.env` — set once via
the root installer's prompt (`.\install.ps1`), or by editing that file
directly. There is no login field anywhere in this UI, by design.

The Power BI **report identity** fields (Tenant/Group/Report/Dataset ID) are
a separate, unrelated concept — those *are* entered in the UI, since they
describe *what to validate*, not *who's logged in*.

## Run (dev)

```powershell
cd app-desktop
npm install
npm start
```

## Launching without a terminal

The root installer (`install.ps1`) creates a desktop shortcut
(`VériRapport.lnk`, pointing at `launch-desktop.bat`) as its
last step — that's the normal way most people will open this app.

You can also do it manually:

- Double-click [launch-desktop.bat](launch-desktop.bat) directly.
- Or recreate a Desktop shortcut yourself, independent of the installer:
  ```powershell
  cd app-desktop
  npm run make-shortcut
  ```
  This runs `CreateDesktopShortcut.ps1` and creates a *second*, differently
  named shortcut (`VériRapport.lnk`) always on the Desktop — useful
  if you want to recreate just the shortcut without rerunning the whole
  installer, but note it's a distinct mechanism from the one the installer
  uses, not the same shortcut recreated.

## How the report-identity override works under the hood

Runtime overrides are opt-in and only take effect when both of these are set
(which this app sets automatically for you when you click Run):
- `PBI_UI_RUNTIME_OVERRIDE=1`
- `PBI_UI_RUNTIME_CONFIG_PATH` pointing at the temporary JSON config

With those unset — e.g. if you close this app and run `npm run parity`
directly from a terminal — `comparison-config.helpers.ts` falls back to
whatever is hardcoded in its `PAIRS` array, exactly as if this app didn't
exist. Nothing about the CLI flow changes because this app is installed.
