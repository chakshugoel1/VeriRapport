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

`parity-summary.xlsx` has five sheets:

- **Parity Summary** — PASS/FAIL verdict, which reports were compared, counts
- **Page Comparison** — per page: rows in source, rows in target, rows only on
  one side, and status (Identical / Different / Missing / Only in target)
- **Visual Comparison** — one row per Power BI visual, with an exact status
  (identical / header text differs / duplicated / values differ / missing /
  ambiguous)
- **Differences** — rows that exist on one side only
- **Changed Values** — which cell changed, from what to what (with % change)

**The test does NOT fail (does not go red) just because the data differs.**
A data mismatch is a normal, expected finding — it's recorded as a
non-failing annotation on the test result and fully detailed in
`parity-summary.xlsx`. The test only fails on a genuine automation problem
(login failed, a report wouldn't embed, a file didn't get written). If you
want a real data mismatch to fail the test outright (e.g. wiring this into a
CI gate), set `FAIL_ON_PARITY_DIFF=1` before running.

---

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
