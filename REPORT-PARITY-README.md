# Report Parity (Migration) Validator

Validates that a Power BI report migrated from **Import mode** → **Direct Lake mode**
still shows **identical data**.

```
source (Import mode)  → exported → expected.xlsx   ← EXPECTED values
target (Direct Lake)  → exported → actual.xlsx     ← ACTUAL values
                        compare  → parity-summary.xlsx
```

---

## 1. Configure your report pair

Two ways to do this — pick whichever fits how you're running things.

### Option A — Desktop app (no code editing)

Open the desktop app (`app-desktop/`, or the shortcut the installer created)
and fill in the Source and Target report's Tenant ID / Group ID / Report ID /
Dataset ID directly in the UI (or paste a report link). This overrides the pair
in `comparison-config.helpers.ts` for that run only — nothing to edit, nothing to save permanently. See the main
`README.md`'s "Desktop app" section for the full flow.

### Option B — Edit the code (a permanent, version-controlled default)

Open **`tests/helpers/comparison-config.helpers.ts`** and fill in both the
**source** and **target** GUIDs in the `PAIRS` array:

```ts
export const PAIRS: ReportPair[] = [
  {
    name: 'Power BI Report',
    source: {   // Import mode — EXPECTED
      tenantId:  '...',
      groupId:   '...',
      reportId:  '...',
      datasetId: '...',
    },
    target: {   // Direct Lake — ACTUAL
      tenantId:  '...',
      groupId:   '...',
      reportId:  '...',
      datasetId: '...',
    },
  },
];
```

**Where to find the GUIDs:** open the report in Power BI. The URL is
`app.powerbi.com/groups/<groupId>/reports/<reportId>/...`. The `datasetId` is on
the report's *Settings → Semantic model* page (or via the workspace's dataset list).

The shipped file has all IDs blank. If you leave any field blank (or as the literal `REPLACE_WITH_...` placeholder
text), the test **fails immediately with a clear message** before opening a
browser — it will not silently compare the wrong thing.

### Optional settings (code only)

```ts
pages: ['Overview', 'Details'],      // only validate these tabs (default: ALL tabs)
pageMap: { 'Old Name': 'New Name' },  // only if a tab was RENAMED during migration
                                      // (source name → target name)
// inside source / target: rlsRole: 'RoleName' — only if the dataset requires an RLS role
```

By default all pages are auto-discovered and matched by name — no config needed
if the tab names are unchanged.

---

## 2. Run it

```powershell
npm run parity          # export source + target, then compare   ← normal use
npm run parity:source   # export ONLY the source → refresh the expected baseline
npm run parity:target   # export ONLY the target → compare against existing baseline
```

`parity:target` is the fast loop: once you have a baseline, re-validate the
migrated report without re-exporting the Import-mode report each time. If no
baseline exists yet, it fails with an instruction to run `parity:source` first.

Multiple pairs configured? Select one:

```powershell
$env:PAIR='Power BI Report'; npm run parity    # PowerShell (must match the pair's name)
```

---

## 3. Output

Everything lands in **`playwright-report-parity/<PairName>/`**:

| File | Contents |
|---|---|
| `expected.xlsx` | Source (Import mode) data — one sheet per report page |
| `actual.xlsx` | Target (Direct Lake) data — one sheet per report page |
| `parity-summary.xlsx` | **The deliverable.** Verdict + per-page diffs + sample differing rows |
| `parity-result.json` | Machine-readable result (read by the desktop app) |
| `discovered-slicers.json` | Filters found by Load filters / `discover:slicers` |

`parity-summary.xlsx` has these sheets, in reading order:

- **At a glance** — the verdict, whether the same filters were used on both
  reports, report version / data-refresh facts, a count of visuals in five
  plain groups, and a page-by-page table.
- **What differs** — real differences only, with the three biggest changes.
- **All visuals** — one row per visual, worst first. The groups are:
  ❌ real difference, 🔍 check by eye, ℹ️ differs as expected (version / refresh
  date / last closed period), ⚪ nothing to compare (empty on both sides or
  removed on purpose), ✅ match. The **Why it came back empty** column says why
  a visual exported nothing (for example "Returned 0 rows" or an export error).
- **Changed values** — every changed cell, biggest change first (with % change).
- **Filters Applied** — filters and slicers found, cleared, kept and applied.
- **Run details** — report IDs, files, page and mapping counts.
- **Differences** — side-by-side boxes for each differing visual.

Visuals that export no data are no longer counted as "identical": they are
listed as "nothing to compare" with the reason. A visual that returns exactly
30,000 rows hit Power BI's export limit and is flagged on *At a glance*.

**The test does NOT fail (does not go red) just because the data differs.**
A data mismatch is a normal, expected finding — it's recorded as a
non-failing annotation on the test result and fully detailed in
`parity-summary.xlsx`. The test only fails on a genuine automation problem
(login failed, a report wouldn't embed, a file didn't get written). If you
want a real data mismatch to fail the test outright (e.g. wiring this into a
CI gate), set `FAIL_ON_PARITY_DIFF=1` before running.

---

## Comparing with a KPI mapping file (redesigned reports)

Use this when the new report was **redesigned** (for example by ReportIQ), so visual types, titles and layout changed. A KPI mapping Excel file tells VériRapport which old visual became which new one.

- **Desktop app:** in step 1, click **Browse…** next to *Mapping file (optional)* and pick the Excel file. Click **Clear** to go back to the normal comparison.
- **Config / CLI:** set `mappingFile: 'C:/path/to/mapping.xlsx'` on the pair in `tests/helpers/comparison-config.helpers.ts`.

What the file needs (column names may vary: they are found by meaning, ignoring case, spaces and symbols): old page, old page ID, **old visual ID**, **match type**, new page, new page ID, new visual ID, old/new fields, fields dropped, reason. Extra columns are ignored. If your file uses other header names, list them in a settings JSON (`{"columns": {"oldVisualId": ["My ID header"]}}`) and point `mappingSettingsFile` (or `PBI_MAPPING_SETTINGS`) at it.

How each match type is handled:

| Match type | What happens |
|---|---|
| Same KPI | Values compared directly. Field names that differ only in spelling (`Project Number` vs `PROJECT_NUMBER`) count as the same. A card whose label was renamed (`M_Last Available Period_EC` vs `Last closed period`) is paired through the mapping's field lists, or by position, so its value is still compared. |
| Same table, columns changed | Only the columns both reports have are compared. Removed columns are listed. |
| Merged into a multi-value card | The old card's value is looked up inside the new card: by field name, then through the mapping file's old/new field lists, then by value. Only if all three fail is the card **Not verified** (listed for a look, not a failure). |
| No same KPI / dropped | Listed as **Removed on purpose**, with the reason. Not a failure. |

The *All visuals* sheet gets three extra columns: Old visual ID, New visual ID, Why changed. Filters written against the old report are translated to the new field and page names where the mapping shows the rename.

Notes: each export now also writes a small `*.visual-index.json` file next to `expected.xlsx` / `actual.xlsx` (which visual ID is which block). Mapping mode needs it, so re-export the source once after updating. Without a mapping file, nothing changes.

## Saved filters in the reports (no-filter baseline)

By default, before exporting each page, VériRapport **clears what the report author left applied** on both reports: report-level filters ("Filters on all pages"), page-level filters ("Filters on this page") and every slicer's selection.

- **No filters picked in the app:** both reports are compared on their full data.
- **Filters picked in the app:** only those are applied, after clearing.
- **Not cleared:** visual-level filters (they are usually part of a KPI's definition, for example a card that only counts open projects), and **row-level security (RLS)**, which always applies to the signed-in user and cannot be removed by the tool.
- The summary's *At a glance* notes and the *Filters Applied* sheet list what was cleared or kept on each side. If something could not be cleared on one side, the run is marked *not comparable* with the reason.
- To export the reports exactly as saved (the old behaviour), tick **Keep the filters and slicer selections saved in the reports** in the app, or set `PBI_KEEP_PREAPPLIED_FILTERS=1`.

## Notes

- **Comparison is order-independent.** Rows and visuals are matched as a
  multiset, so a different row order or visual order will *not* be reported as a
  difference — only genuinely different data is.
- **Only data-bearing visuals are compared.** Images, shapes, text boxes,
  buttons, and slicers are skipped (they carry no tabular data).
- **Switching pairs?** Delete the old `playwright-report-parity/<Pair>/` folder,
  or the stale `expected.xlsx` will be reused.
- **The existing main-run flow is unchanged.** `npm run main:run` behaves
  independently — the parity feature only activates on the `parity*` scripts.
- Auth is shared with the existing setup. If the token can't be found, run
  `npm run test:setup` first (or use the desktop app, which signs in for you
  automatically if no session is found when you click Run comparison).
