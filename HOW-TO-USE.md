# How to use VériRapport

A first-time walkthrough of the desktop app. The usual order is:
**Enter both reports → Run comparison → Review the result.** Filters are optional.

## Before you start

- You have the IDs for both reports you are comparing — the Import-mode
  **source** and the Direct-Lake **target**: Tenant, Group/Workspace, Report and
  Dataset ID. (Pasting a Power BI report link fills the first three.)
- A **test account** is set in the `.env` file in
  `%USERPROFILE%\Power_BI_report_validation_credentials` (the installer asks for
  it once, or edit the file directly). Credentials are never typed into the app.

## The flow

```text
┌──────────────────────────────────────────────────────────────────┐
│                 START — open the VériRapport app                 │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 1 · Reports to compare                              (Card 1)│
│                                                                  │
│   Source · Expected (Import)      Target · Actual (Direct Lake)  │
│     • paste report link (optional)  • paste report link          │
│     • Tenant ID                     • Tenant ID                  │
│     • Group / Workspace ID          • Group / Workspace ID       │
│     • Report ID                     • Report ID                  │
│     • Dataset ID                    • Dataset ID                 │
│                                                                  │
│   • Pair name  (e.g. Sales report)                               │
│   • Case-insensitive compare  (optional)                         │
│   • Save → keeps this pair in the "Saved pairs" list             │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ OPTIONAL · Filters                         (collapsed section)   │
│                                                                  │
│   Click "Load filters" to read the reports' pages and slicers    │
│   (read-only). Pick pages and slicer values to apply on both     │
│   reports. Your picks are used on the next run.                  │
│   Needs a saved sign-in — run a comparison or "Sign in again"    │
│   once first.                                                    │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 2 · Run comparison                                  (Card 2)│
│                                                                  │
│   Click "▶ Run comparison".                                      │
│   No saved sign-in yet? It signs in first, automatically.        │
│   Source (Import)        →  expected.xlsx                        │
│   Target (Direct Lake)   →  actual.xlsx                          │
│   Progress and elapsed time show while it runs.                  │
│   "■ Stop" ends the run and closes its browser.                  │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 3 · Review the result                                       │
│                                                                  │
│   • Verdict: PASS / FAIL / NOT COMPARABLE                        │
│   • Plain-English summary + per-page breakdown                   │
│   • Open the summary (parity-summary.xlsx) or the results folder │
│   • "Recent runs" keeps the last runs for one-click access       │
└──────────────────────────────────────────────────────────────────┘
```

## Step notes

**1 · Reports to compare.** Paste a Power BI report link into the box above each
side to fill the Report, Workspace and Tenant IDs. The Dataset ID is not in the
link, so type it once — **Save** keeps it with the pair. **Use Source tenant**
copies the tenant ID to the target. Give the pair a name: it names the results
folder. Turn on *Case-insensitive compare* only if labels that differ just by
case or separators should count as equal.

**Filters (optional).** Open the *Filters (optional)* section and click **Load
filters**. It reads the reports without changing anything and lists pages and
slicer values. Pick what you want; it is applied to both reports on the next run.
Loading filters cannot sign you in, so run a comparison or click **Sign in again**
once first.

**2 · Run comparison.** Click **▶ Run comparison**. If there is no saved sign-in, a
browser opens and signs in with the test account first. Both reports are then
exported and compared. The raw log is under **Technical details**. Use **Sign in
again** if your saved session has expired.

**3 · Review.** The verdict and per-page results appear in the app. Open
`parity-summary.xlsx` for the full detail (Parity Summary, Page Comparison,
Visual Comparison, Differences, Changed Values), or open the results folder for
`expected.xlsx` and `actual.xlsx`.
