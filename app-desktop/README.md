# VériRapport

A small Electron app that wraps the project's existing Playwright scripts
(`npm run test:setup`, `npm run parity`, `npm run discover:slicers`) behind a
GUI — for running things without a terminal. It's a wrapper, not a separate
implementation: everything it does ultimately runs the same scripts you could
run from PowerShell yourself.

## What the UI does

- **Saved pairs** - the last form is restored on launch. **Save** keeps a named pair (IDs, pages, options) you can reload from the *Saved pairs* dropdown; **Delete** removes it. Stored in the app's own data folder (`ui-state.json`), IDs only, never credentials.
- **Reports to compare** - enter the Source and Target report's Tenant ID / Group ID / Report ID / Dataset ID, or paste a Power BI report link into the box above each side to fill report, workspace and tenant IDs. **Use Source tenant** copies the tenant across. These values override the default pair in `tests/helpers/comparison-config.helpers.ts` for that run only, via a temporary runtime-config file (`.runtime/parity-runtime-config.json`) - nothing in the repo is edited.
- **▶ Run comparison** - runs `npm run parity`. If there is no saved sign-in session it runs `npm run test:setup` first, automatically. **Sign in again** runs sign-in on its own.
- **■ Stop** - ends the whole run, including the browser it started.
- **Live progress** (page X of N, elapsed time) taken from the run's own log lines; the raw log is under **Technical details**.
- **Clear errors** - missing or badly formatted IDs are listed next to the Run button and the field is highlighted.
- **After a run** - verdict (PASS / FAIL / NOT COMPARABLE), per-page results, buttons to open the summary or results folder, and a **Recent runs** list.
- **Filters (optional)** - a collapsed section. **Load filters** reads report-level filters and page slicers (read-only, needs a saved sign-in); pick pages and slicer values to apply on both reports in the next run.

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
(`VeriRapport.lnk`, pointing at `launch-desktop.bat`) as its
last step — that's the normal way most people will open this app.

You can also do it manually:

- Double-click [launch-desktop.bat](launch-desktop.bat) directly.
- Or recreate a Desktop shortcut yourself, independent of the installer:
  ```powershell
  cd app-desktop
  npm run make-shortcut
  ```
  This runs `CreateDesktopShortcut.ps1`, which (re)creates `VeriRapport.lnk`
  on your Desktop - useful to recreate just the shortcut without rerunning the
  whole installer.

## How the report-identity override works under the hood

Runtime overrides are opt-in and only take effect when both of these are set
(which this app sets automatically for you when you click Run):
- `PBI_UI_RUNTIME_OVERRIDE=1`
- `PBI_UI_RUNTIME_CONFIG_PATH` pointing at the temporary JSON config

With those unset — e.g. if you close this app and run `npm run parity`
directly from a terminal — `comparison-config.helpers.ts` falls back to
the pair(s) in its `PAIRS` array (blank by default - fill in the IDs first), exactly as if this app didn't
exist. Nothing about the CLI flow changes because this app is installed.
