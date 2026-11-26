const FIELD_PLACEHOLDERS = {
  tenantId: 'Enter tenant ID',
  groupId: 'Enter group/workspace ID',
  reportId: 'Enter report ID',
  datasetId: 'Enter dataset ID',
};

const SOURCE_KEYS = [
  ['tenantId', 'Tenant ID'],
  ['groupId', 'Workspace ID'],
  ['reportId', 'Report ID'],
  ['datasetId', 'Dataset ID'],
];

const TARGET_KEYS = [
  ['tenantId', 'Tenant ID'],
  ['groupId', 'Workspace ID'],
  ['reportId', 'Report ID'],
  ['datasetId', 'Dataset ID'],
];

function createRows(containerId, prefix, items) {
  const container = document.getElementById(containerId);
  items.forEach(([key, label]) => {
    const row = document.createElement('div');
    row.className = 'row';

    const fieldWrap = document.createElement('div');
    fieldWrap.className = 'field-wrap';

    const lbl = document.createElement('label');
    lbl.textContent = label;
    lbl.setAttribute('for', `${prefix}_${key}`);

    const input = document.createElement('input');
    input.id = `${prefix}_${key}`;
    input.placeholder = FIELD_PLACEHOLDERS[key] || '';

    fieldWrap.appendChild(lbl);

    row.appendChild(fieldWrap);
    row.appendChild(input);
    container.appendChild(row);
  });
}

const GUID_RE = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
const GUID_KEYS = ['tenantId', 'groupId', 'reportId', 'datasetId'];

createRows('source-fields', 'source', SOURCE_KEYS);
createRows('target-fields', 'target', TARGET_KEYS);

// Inline format check: a pasted ID with a stray space/char is the most common
// reason a run fails minutes in, so flag it the moment the field loses focus.
function checkGuidField(input) {
  const v = input.value.trim();
  const bad = v !== '' && !GUID_RE.test(v);
  input.classList.toggle('invalid', bad);
  let hint = input.parentElement.querySelector(`.hint[data-for="${input.id}"]`);
  if (bad && !hint) {
    hint = document.createElement('div');
    hint.className = 'hint';
    hint.dataset.for = input.id;
    hint.textContent = 'This does not look like an ID (expected xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).';
    input.after(hint);
  } else if (!bad && hint) {
    hint.remove();
  }
  return !bad;
}
['source', 'target'].forEach((side) => GUID_KEYS.forEach((k) => {
  const input = document.getElementById(`${side}_${k}`);
  input.addEventListener('blur', () => { input.value = input.value.trim(); checkGuidField(input); });
  input.addEventListener('input', () => { if (input.classList.contains('invalid')) checkGuidField(input); scheduleSaveForm(); });
}));

const logEl = document.getElementById('log');
const buttons = [
  document.getElementById('btnSetup'),
  document.getElementById('btnParity'),
  document.getElementById('btnDiscover'),
];

function appendLog(line) {
  try { updateProgressFromLog(String(line)); } catch { /* progress is cosmetic */ }
  logEl.textContent += `${line}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

// Discover Filters can't complete without a saved login session: it runs
// with --no-deps (skips the interactive 'setup' auth step, deliberately, so
// discovery doesn't force a fresh sign-in every time), and cross-report
// discovery specifically runs its browser windows off-screen for speed - so
// if a Microsoft sign-in page ever appeared, there would be no way to see or
// interact with it. Gate the button on a known-good session instead of
// letting that run start and fail confusingly.
let _isBusy = false;
let _hasSession = false;

function updateButtonStates() {
  buttons.forEach(b => { b.disabled = _isBusy; });
  const discoverBtn = document.getElementById('btnDiscover');
  discoverBtn.disabled = _isBusy || !_hasSession;
  const note = document.getElementById('discoverAuthNote');
  if (note) note.style.display = (!_isBusy && !_hasSession) ? 'block' : 'none';
}

function setBusy(isBusy) {
  _isBusy = isBusy;
  updateButtonStates();
  document.getElementById('btnStop').style.display = isBusy ? '' : 'none';
  document.getElementById('btnStop').disabled = false;
  document.getElementById('progressWrap').classList.toggle('show', isBusy);
  document.getElementById('idleNote').style.display = isBusy ? 'none' : '';
  if (isBusy) { resetProgress(); startTimer(); } else { stopTimer(); }
}

async function refreshAuthGate() {
  try {
    const status = await window.verirapportDesktop.getAuthStatus();
    _hasSession = !!(status && status.hasSession);
  } catch {
    _hasSession = false;
  }
  updateButtonStates();
}
refreshAuthGate();

function readIdentity(prefix) {
  const obj = {};
  ['tenantId', 'groupId', 'reportId', 'datasetId'].forEach((k) => {
    obj[k] = document.getElementById(`${prefix}_${k}`).value.trim();
  });
  return obj;
}

function getPagesInScope() {
  return [];
}

function buildConfig() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  const pairName = document.getElementById('pairName').value.trim() || 'PowerBI-Report';
  const pages = getSelectedRunPages();
  const slicerScenario = buildScenarioFromPicks(pages);
  const lenientTextCompare = document.getElementById('lenientTextCompare')?.checked === true;

  return {
    pairName,
    source,
    target,
    pages: pages.length ? pages : undefined,
    slicerScenario,
    lenientTextCompare,
    mappingFile: (document.getElementById('mappingFile')?.value || '').trim() || undefined,
    keepSavedFilters: document.getElementById('keepSavedFilters')?.checked === true,
  };
}

const FIELD_LABELS = { tenantId: 'Tenant ID', groupId: 'Workspace (Group) ID', reportId: 'Report ID', datasetId: 'Dataset ID' };

/** Returns human-readable problems; each carries the input to focus. */
function validateForParity(cfg) {
  const problems = [];
  for (const side of ['source', 'target']) {
    const sideName = side === 'source' ? 'Source' : 'Target';
    for (const k of GUID_KEYS) {
      const input = document.getElementById(`${side}_${k}`);
      if (!cfg[side][k]) problems.push({ input, text: `${sideName}: ${FIELD_LABELS[k]} is empty` });
      else if (!GUID_RE.test(cfg[side][k])) problems.push({ input, text: `${sideName}: ${FIELD_LABELS[k]} is not in the expected ID format` });
    }
  }
  if (cfg.source.reportId && cfg.source.reportId === cfg.target.reportId && cfg.source.groupId === cfg.target.groupId) {
    problems.push({ input: document.getElementById('target_reportId'), text: 'Source and Target are the same report — nothing to compare' });
  }
  return problems;
}

function showErrors(problems) {
  const el = document.getElementById('errorBanner');
  if (!problems || problems.length === 0) { el.classList.remove('show'); el.innerHTML = ''; return; }
  el.innerHTML = `<strong>Fix ${problems.length === 1 ? 'this' : 'these'} before running:</strong><ul>${problems.map(p => `<li>${p.text}</li>`).join('')}</ul>`;
  el.classList.add('show');
  problems.forEach(p => p.input && p.input.classList.add('invalid'));
  const first = problems.find(p => p.input);
  if (first) first.input.focus();
}
function showErrorText(text) {
  const el = document.getElementById('errorBanner');
  el.textContent = text;
  el.classList.add('show');
}

const IDENTITY_REQUIRED_KEYS = ['tenantId', 'groupId', 'reportId', 'datasetId'];

/**
 * Discovery only ever reads ONE report identity, but it must now be supplied
 * explicitly from the UI. We accept either a fully-populated Source or a
 * fully-populated Target identity; we do NOT fall back to the hardcoded pair.
 */
function resolveDiscoverIdentity() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  const sourceTouched = IDENTITY_REQUIRED_KEYS.some((k) => source[k]);
  const targetTouched = IDENTITY_REQUIRED_KEYS.some((k) => target[k]);

  if (!sourceTouched && !targetTouched) {
    return {
      identity: null,
      side: null,
      missing: IDENTITY_REQUIRED_KEYS.map((k) => `source.${k}`),
      note: 'fill in all Source report fields (or fill in all Target report fields instead)',
    };
  }

  const side = sourceTouched ? 'source' : 'target';
  const candidate = sourceTouched ? source : target;
  const missing = IDENTITY_REQUIRED_KEYS.filter((k) => !candidate[k]).map((k) => `${side}.${k}`);
  return {
    identity: missing.length === 0 ? candidate : null,
    side,
    missing,
    note: side === 'source'
      ? 'fill in all Source report fields before discovering'
      : 'fill in all Target report fields before discovering',
  };
}

/** True when Source AND Target are both fully filled in - the trigger for
 *  cross-report discovery instead of the single-side flow. */
function bothIdentitiesFilled() {
  const source = readIdentity('source');
  const target = readIdentity('target');
  return IDENTITY_REQUIRED_KEYS.every((k) => source[k]) && IDENTITY_REQUIRED_KEYS.every((k) => target[k]);
}

// Mirrors tests/helpers/report-export.helpers.ts's pageNameKey exactly (same
// reasoning as main.js's copy: this is plain browser JS, no TypeScript
// compilation step, so it can't import the .ts version directly).
function pageNameKeyJs(name) {
  return String(name || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
}

function dedupePageNames(names) {
  const seen = new Set();
  const out = [];
  for (const n of names) {
    const key = pageNameKeyJs(n);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(n);
  }
  return out;
}

window.verirapportDesktop.onLog((line) => appendLog(line));
window.verirapportDesktop.onState((state) => {
  // Display only - button enablement is driven by onBusy below, which spans
  // the WHOLE flow (e.g. all of Auth+Parity), not just whichever individual
  // step this text currently names. See beginFlow/endFlow in main.js.
  void state;
});
window.verirapportDesktop.onBusy((isBusy) => setBusy(isBusy));

document.getElementById('btnSetup').addEventListener('click', async () => {
  appendLog('--- Running setup only ---');
  await window.verirapportDesktop.runSetup();
  await refreshAuthGate();
});

// ── Run: one button does the right thing ─────────────────────────────────────

document.getElementById('btnSetup').addEventListener('click', async () => {
  showErrors(null);
  appendLog('--- Signing in (a browser window will open) ---');
  const res = await window.verirapportDesktop.runSetup();
  if (res && !res.ok) showErrorText(res.error);
  else { toast('✓ Sign-in complete — session saved'); appendLog('--- Sign-in complete ---'); }
  await refreshAuthGate();
});

document.getElementById('btnStop').addEventListener('click', async () => {
  document.getElementById('btnStop').disabled = true;
  setProgressText('Stopping…');
  await window.verirapportDesktop.cancelRun();
});

document.getElementById('btnParity').addEventListener('click', async () => {
  showErrors(null);
  const cfg = buildConfig();
  const problems = validateForParity(cfg);
  if (problems.length > 0) { showErrors(problems); return; }

  // No saved session => a plain parity run would fail minutes in, deep inside
  // token acquisition. Sign in first, automatically, instead.
  const authStatus = await window.verirapportDesktop.getAuthStatus();
  const filterCount = Object.values(cfg.slicerScenario.pages).reduce((n, sels) => n + sels.length, 0);
  appendLog(authStatus.hasSession ? '--- Starting comparison ---' : '--- No saved login session - signing in first, then comparing ---');
  if (filterCount > 0) appendLog(`--- ${filterCount} filter selection(s) applied ---`);
  if (cfg.lenientTextCompare) appendLog('--- Case-insensitive compare enabled ---');
  appendLog(cfg.keepSavedFilters
    ? '--- Keeping the filters saved in the reports ---'
    : '--- Clearing the filters saved in the reports before export ---');
  if (cfg.mappingFile) appendLog(`--- Mapping-file mode: ${cfg.mappingFile} ---`);
  document.getElementById('parityResults').innerHTML = '';
  resetProgress();

  await saveFormNow();
  const res = authStatus.hasSession
    ? await window.verirapportDesktop.runParity(cfg)
    : await window.verirapportDesktop.runSetupAndParity(cfg);
  if (res.ok) {
    renderParityResults(res.result);
    await rememberRun(cfg, res.result);
    toast(res.result ? 'Comparison finished' : 'Run finished');
  } else {
    showErrorText(friendlyError(res.error));
  }
  await refreshAuthGate();
});

/** Turns raw failure text into something actionable. */
function friendlyError(msg) {
  const m = String(msg || 'The run failed.');
  if (/stopped by you/i.test(m)) return 'Run stopped. Nothing was compared.';
  if (/npm/i.test(m) && /find|found/i.test(m)) return m;
  if (/exit code/i.test(m)) return `${m}. Open "Technical details" below to see the cause — a common one is an expired sign-in, which "Sign in again" fixes.`;
  return m;
}

// ── Progress, elapsed time and Stop ──────────────────────────────────────────
//
// Derived from the log lines the run already prints ("[export:source] Report
// has 8 page(s)" and one '→ page "X"' line per page) - no change to the run.

const prog = { total: { source: 0, target: 0 }, done: { source: 0, target: 0 }, side: '', phase: 'Starting…', startedAt: 0, timer: null };

function resetProgress() {
  prog.total = { source: 0, target: 0 }; prog.done = { source: 0, target: 0 };
  prog.side = ''; prog.phase = 'Starting…';
  setProgressText('Starting…');
  setProgressFraction(null);
}
function setProgressText(t) { document.getElementById('progressText').innerHTML = `<strong>${t}</strong>`; }
function setProgressFraction(f) {
  const bar = document.getElementById('progress');
  if (f === null) { bar.classList.add('indeterminate'); bar.firstElementChild.style.width = ''; return; }
  bar.classList.remove('indeterminate');
  bar.firstElementChild.style.width = `${Math.max(2, Math.min(100, Math.round(f * 100)))}%`;
}
function fmtClock(ms) {
  const s = Math.floor(ms / 1000);
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
}
function updateProgressFromLog(line) {
  let m = /\[export:(source|target)\] Report has (\d+) page/i.exec(line);
  if (m) { prog.side = m[1].toLowerCase(); prog.total[prog.side] = Number(m[2]); }
  m = /\[export:(source|target)\]\s+→ page "(.+)"/i.exec(line);
  if (m) {
    prog.side = m[1].toLowerCase();
    prog.done[prog.side]++;
    prog.phase = `${prog.side === 'source' ? 'Source' : 'Target'} report — page ${prog.done[prog.side]} of ${prog.total[prog.side] || '?'}: ${m[2]}`;
  } else if (/Wrote \d+ sheet/i.test(line) && prog.side) {
    prog.phase = `${prog.side === 'source' ? 'Source' : 'Target'} report exported`;
  } else if (/compar/i.test(line) && /excel|workbook|summary/i.test(line)) {
    prog.phase = 'Comparing the two workbooks…';
  } else if (/test:setup|signing in/i.test(line)) {
    prog.phase = 'Signing in…';
  } else if (!m && !/^\[export:/.test(line)) {
    return;
  }
  const known = prog.total.source + prog.total.target;
  const expected = (prog.total.source && prog.total.target) ? known : Math.max(prog.total.source, prog.total.target) * 2;
  setProgressText(prog.phase);
  setProgressFraction(expected > 0 ? (prog.done.source + prog.done.target) / expected : null);
}
function startTimer() {
  prog.startedAt = Date.now();
  clearInterval(prog.timer);
  prog.timer = setInterval(() => { document.getElementById('elapsed').textContent = fmtClock(Date.now() - prog.startedAt); }, 1000);
  document.getElementById('elapsed').textContent = '0:00';
}
function stopTimer() { clearInterval(prog.timer); prog.timer = null; }

// ── Toast ────────────────────────────────────────────────────────────────────
let _toastTimer = null;
function toast(text) {
  const el = document.getElementById('toast');
  el.textContent = text; el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}

// ── Saved pairs, last form, recent runs ──────────────────────────────────────
//
// Everything here lives in the app's own data folder (see state-save in
// main.js). Report IDs and run history only - never credentials.

let _state = { pairs: {}, lastForm: null, recent: [] };

function readForm() {
  return {
    pairName: document.getElementById('pairName').value.trim(),
    lenientTextCompare: document.getElementById('lenientTextCompare').checked,
    mappingFile: (document.getElementById('mappingFile')?.value || '').trim(),
    keepSavedFilters: document.getElementById('keepSavedFilters').checked,
    source: { ...readIdentity('source') },
    target: { ...readIdentity('target') },
  };
}
function writeForm(f) {
  if (!f) return;
  document.getElementById('pairName').value = f.pairName || '';
  document.getElementById('lenientTextCompare').checked = !!f.lenientTextCompare;
  setMappingFile(f.mappingFile || '');
  document.getElementById('keepSavedFilters').checked = !!f.keepSavedFilters;
  for (const side of ['source', 'target']) {
    ['tenantId', 'groupId', 'reportId', 'datasetId'].forEach((k) => {
      const input = document.getElementById(`${side}_${k}`);
      input.value = (f[side] && f[side][k]) || '';
      if (GUID_KEYS.includes(k)) checkGuidField(input);
    });
  }
}

// Optional KPI mapping file. Only the path is kept; the run reads the file and
// checks it before any browser opens (a bad file stops the run with a clear message).
function setMappingFile(p) {
  const input = document.getElementById('mappingFile');
  const note = document.getElementById('mappingNote');
  if (!input) return;
  input.value = p || '';
  if (note) {
    note.style.display = p ? '' : 'none';
    note.textContent = p
      ? 'Mapping mode: visuals are paired using this file instead of by title. Pages and visuals marked "removed on purpose" are listed, not failed.'
      : '';
  }
}
document.getElementById('btnPickMapping')?.addEventListener('click', async () => {
  const res = await window.verirapportDesktop.pickMappingFile();
  if (res && res.ok) { setMappingFile(res.path); scheduleSaveForm(); }
});
document.getElementById('btnClearMapping')?.addEventListener('click', () => { setMappingFile(''); scheduleSaveForm(); });

let _saveTimer = null;
function scheduleSaveForm() {
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(saveFormNow, 600);
}
async function saveFormNow() {
  clearTimeout(_saveTimer);
  _state.lastForm = readForm();
  try { await window.verirapportDesktop.saveState(_state); } catch { /* non-fatal */ }
}
['pairName', 'lenientTextCompare', 'keepSavedFilters'].forEach((id) => {
  const el = document.getElementById(id);
  el.addEventListener('input', scheduleSaveForm);
  el.addEventListener('change', scheduleSaveForm);
});

function renderPairSelect(selected) {
  const sel = document.getElementById('pairSelect');
  sel.innerHTML = '<option value="">— new pair —</option>';
  for (const name of Object.keys(_state.pairs).sort((a, b) => a.localeCompare(b))) {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    sel.appendChild(opt);
  }
  sel.value = selected && _state.pairs[selected] ? selected : '';
  document.getElementById('btnDeletePair').disabled = !sel.value;
}

document.getElementById('pairSelect').addEventListener('change', () => {
  const name = document.getElementById('pairSelect').value;
  document.getElementById('btnDeletePair').disabled = !name;
  if (!name) return;
  writeForm(_state.pairs[name]);
  showErrors(null);
  scheduleSaveForm();
  toast(`Loaded "${name}"`);
});

document.getElementById('btnSavePair').addEventListener('click', async () => {
  const f = readForm();
  if (!f.pairName) {
    showErrorText('Give the pair a name first (the "Pair name" field) — that is the name it is saved under.');
    document.getElementById('pairName').focus();
    return;
  }
  showErrors(null);
  _state.pairs[f.pairName] = f;
  await saveFormNow();
  renderPairSelect(f.pairName);
  toast(`Saved "${f.pairName}"`);
});

document.getElementById('btnDeletePair').addEventListener('click', async () => {
  const name = document.getElementById('pairSelect').value;
  if (!name) return;
  delete _state.pairs[name];
  await saveFormNow();
  renderPairSelect('');
  toast(`Deleted "${name}" (the form keeps its values)`);
});

function relTime(ts) {
  const d = new Date(ts);
  const mins = Math.round((Date.now() - ts) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${mins} min ago`;
  if (mins < 60 * 24) return `${Math.round(mins / 60)} h ago`;
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function renderRecent() {
  const card = document.getElementById('recentCard');
  const host = document.getElementById('recentRuns');
  host.innerHTML = '';
  const items = _state.recent || [];
  card.style.display = items.length ? '' : 'none';
  for (const r of items) {
    const row = document.createElement('div');
    row.className = 'recent-item';
    const when = document.createElement('span'); when.className = 'when'; when.textContent = relTime(r.ts);
    const what = document.createElement('span'); what.className = 'what';
    what.textContent = `${r.pairName} · ${r.durationMs ? fmtClock(r.durationMs) : ''}`.replace(/ · $/, '');
    const badge = document.createElement('span');
    const v = VERDICT_BADGE[r.verdict] || VERDICT_BADGE.fail;
    badge.className = v.cls; badge.textContent = v.text;
    row.append(when, what, badge);
    for (const [label, p, folder] of [['Open summary', r.summaryFile, false], ['Open folder', r.folder, true]]) {
      if (!p) continue;
      const b = document.createElement('button');
      b.className = 'ghost'; b.type = 'button'; b.textContent = label;
      b.addEventListener('click', async () => {
        const res = folder ? await window.verirapportDesktop.openOutputFolder(p) : await window.verirapportDesktop.openOutputFile(p);
        if (!res.ok) toast(`Could not open: ${res.error}`);
      });
      row.appendChild(b);
    }
    host.appendChild(row);
  }
}

async function rememberRun(cfg, result) {
  const sc = result && (result.scenarios || [])[0];
  if (!sc) return;
  _state.recent = [{
    ts: Date.now(), pairName: cfg.pairName, verdict: sc.ui.verdict,
    summaryFile: sc.summaryFile, folder: result.folder, durationMs: result.durationMs,
  }, ...(_state.recent || [])].slice(0, 8);
  await saveFormNow();
  renderRecent();
}

// ── Paste a Power BI link to fill the IDs ────────────────────────────────────
//   https://app.powerbi.com/groups/<workspace>/reports/<report>/ReportSection?ctid=<tenant>

function fillFromLink(side, text) {
  const t = String(text || '');
  const g = /\/groups\/([0-9a-f-]{36})/i.exec(t);
  const r = /\/reports\/([0-9a-f-]{36})/i.exec(t);
  const c = /[?&]ctid=([0-9a-f-]{36})/i.exec(t);
  if (!r) { toast('That does not look like a Power BI report link'); return; }
  const set = (k, v) => { if (!v) return; const i = document.getElementById(`${side}_${k}`); i.value = v; checkGuidField(i); };
  set('reportId', r[1]); set('groupId', g && g[1]); set('tenantId', c && c[1]);
  const need = [!g && 'workspace', !c && 'tenant'].filter(Boolean);
  toast(`Filled report${g ? ', workspace' : ''}${c ? ', tenant' : ''}. Dataset ID is not in the link${need.length ? ` (and no ${need.join('/')})` : ''} — type it once; Save keeps it.`);
  scheduleSaveForm();
}
for (const side of ['source', 'target']) {
  const box = document.getElementById(`${side}_link`);
  box.addEventListener('paste', (e) => {
    const text = (e.clipboardData || window.clipboardData).getData('text');
    e.preventDefault();
    fillFromLink(side, text);
    box.value = '';
  });
  box.addEventListener('change', () => { if (box.value.trim()) { fillFromLink(side, box.value); box.value = ''; } });
}
document.getElementById('btnCopyTenant').addEventListener('click', () => {
  const v = document.getElementById('source_tenantId').value.trim();
  if (!v) { toast('Fill the Source tenant first'); return; }
  const i = document.getElementById('target_tenantId'); i.value = v; checkGuidField(i); scheduleSaveForm();
});

// ── Startup: restore the last form, saved pairs and recent runs ──────────────
(async () => {
  try {
    const st = await window.verirapportDesktop.loadState();
    _state = { pairs: {}, lastForm: null, recent: [], ...(st || {}) };
  } catch { /* first run */ }
  writeForm(_state.lastForm);
  const last = _state.lastForm && _state.lastForm.pairName;
  renderPairSelect(last && _state.pairs[last] ? last : '');
  renderRecent();
})();

// ── Filter discovery + picking ───────────────────────────────────────────────
//
// _filterPicks: Map<fieldKey, {
//   title, isHierarchy, targetLabel, targets, options,
//   byPage: { [pageName]: { visualName } },   // which pages this field was actually found on
//   value: string,                            // currently chosen value ('' = none picked)
// }>
window._filterPicks = new Map();
let _lastGlobalFilters = [];
let _discoveredPages = [];
let _selectedRunPages = new Set();

function fieldKey(title, targetLabel) {
  return `${title}::${targetLabel}`;
}

// Groups raw per-page discovery results into one entry per unique field,
// tracking exactly which pages it was actually found on (this is what lets
// the UI correctly disable a hierarchy field for pages where it doesn't
// have its own slicer visual, while leaving flat fields enabled everywhere).
function groupDiscoveredFields(byPage) {
  const groups = new Map();
  for (const [pageName, slicers] of Object.entries(byPage)) {
    for (const s of slicers) {
      if (s.errorMessage) continue; // not pickable if we couldn't read it
      const key = fieldKey(s.title, s.targetLabel);
      let g = groups.get(key);
      if (!g) {
        g = {
          title: s.title, isHierarchy: s.isHierarchy, targetLabel: s.targetLabel,
          targets: s.targets, options: s.options || [], byPage: {},
        };
        groups.set(key, g);
      }
      g.byPage[pageName] = { visualName: s.name };
      if (s.options && s.options.length > g.options.length) g.options = s.options;
    }
  }
  return groups;
}

function getOrCreatePick(picksMap, key, group) {
  let pick = picksMap.get(key);
  if (!pick) {
    pick = { ...group, value: '' };
    picksMap.set(key, pick);
  } else {
    // Refresh discovery data (byPage/options/targets) on re-discovery, keep the user's picks.
    pick.isHierarchy = group.isHierarchy;
    pick.targetLabel = group.targetLabel;
    pick.targets = group.targets;
    pick.byPage = group.byPage;
    pick.sideTargets = group.sideTargets;
    if (group.options.length > pick.options.length) pick.options = group.options;
  }
  return pick;
}

function getPageUniverse() {
  if (Array.isArray(_discoveredPages) && _discoveredPages.length > 0) {
    return _discoveredPages;
  }
  return getPagesInScope();
}

function getSelectedRunPages() {
  const universe = getPageUniverse();
  if (universe.length === 0) return [];

  const selected = universe.filter(p => _selectedRunPages.has(p));
  // Safety: if selection hasn't been initialized yet, default to all pages.
  return selected.length > 0 ? selected : universe;
}

function getDiscoverPagesCsv() {
  // Discovery scope = the pages picked in the shared page selector.
  const selected = getPageUniverse().filter(p => _selectedRunPages.has(p));
  return selected.length > 0 ? selected.join(',') : '';
}

function buildScenarioFromPicks(pageScope) {
  // Fixed name (not a timestamp) — this is what makes each run overwrite the
  // same expected.xlsx/actual.xlsx/parity-summary.xlsx instead of creating a
  // new playwright-report-parity/<pair name>/ui-<timestamp>/ folder every time.
  const scenario = { name: 'ui-filters', pages: {} };
  for (const pageName of pageScope) scenario.pages[pageName] = [];
  if (pageScope.length === 0) return scenario;

  const addPicksFromMap = (picksMap, side) => {
    for (const pick of picksMap.values()) {
      if (!pick.value) continue;
      for (const pageName of pageScope) {
        if (pick.sideTargets && !side) {
          // Shared picker for a cross-report "identical" match: one VALUE
          // picked once, but each side's underlying table.column binding can
          // differ (e.g. a Direct Lake migration renamed the table) even
          // though the on-screen slicer TITLE is identical - matching by
          // title never guaranteed matching bindings. Emitting the source
          // side's targets for both sides is exactly what produced the
          // FILTER FIELD MISMATCH abort (and, before that check existed, a
          // silent one-sided-only filter). Each side must use its OWN
          // discovered targets.
          for (const s of ['source', 'target']) {
            const sd = pick.sideTargets[s];
            if (!sd) continue;
            const sel = { title: pick.title, values: [pick.value], targets: sd.targets, side: s };
            if (pick.isHierarchy) {
              const onThisPage = sd.byPage[pageName];
              if (!onThisPage) continue;
              sel.isHierarchy = true;
              sel.visualName = onThisPage.visualName;
            }
            scenario.pages[pageName].push(sel);
          }
          continue;
        }
        const sel = { title: pick.title, values: [pick.value], targets: pick.targets };
        if (side) sel.side = side; // undefined = applies to both sides, as always
        if (pick.isHierarchy) {
          const onThisPage = pick.byPage[pageName];
          if (!onThisPage) continue; // hierarchy can only be set where its visual exists
          sel.isHierarchy = true;
          sel.visualName = onThisPage.visualName;
        }
        scenario.pages[pageName].push(sel);
      }
    }
  };

  if (window._discoverMode === 'split') {
    // Cross-report discovery found the two reports' filters do NOT match -
    // each side's picks were made against that side's own fields, so they
    // must only ever apply on that side (see selectionsForSide in
    // slicer-config.helpers.ts, which is what actually enforces this at
    // apply time).
    addPicksFromMap(window._filterPicksSource || new Map(), 'source');
    addPicksFromMap(window._filterPicksTarget || new Map(), 'target');
  } else {
    addPicksFromMap(window._filterPicks, undefined);
  }
  return scenario;
}

function renderPageScopeSelector(pageNames) {
  const host = document.getElementById('pageScopeResult');
  host.innerHTML = '';

  if (!Array.isArray(pageNames) || pageNames.length === 0) {
    host.innerHTML = '<div class="placeholder">Click Load filters once to load report pages. Selected pages here will be used for all slicers.</div>';
    return;
  }

  const validSelection = pageNames.filter(p => _selectedRunPages.has(p));
  _selectedRunPages = validSelection.length > 0 ? new Set(validSelection) : new Set(pageNames);

  const wrap = document.createElement('div');
  wrap.className = 'discover-page';
  const h3 = document.createElement('h3');
  h3.textContent = 'Selected pages (applies to all slicers below)';
  wrap.appendChild(h3);

  const pagesWrap = document.createElement('div');
  pagesWrap.style.marginTop = '6px';
  pagesWrap.style.display = 'flex';
  pagesWrap.style.flexWrap = 'wrap';
  pagesWrap.style.gap = '10px';

  const allLabel = document.createElement('label');
  allLabel.style.fontSize = '12px';
  allLabel.style.display = 'flex';
  allLabel.style.alignItems = 'center';
  allLabel.style.gap = '4px';
  allLabel.style.fontWeight = '600';

  const allCb = document.createElement('input');
  allCb.type = 'checkbox';

  const syncAllCheckboxState = () => {
    const selected = pageNames.filter(p => _selectedRunPages.has(p)).length;
    allCb.checked = selected === pageNames.length;
    allCb.indeterminate = selected > 0 && selected < pageNames.length;
  };

  const syncPageCheckboxes = () => {
    for (const cb of pagesWrap.querySelectorAll('input[data-page]')) {
      const page = cb.getAttribute('data-page');
      cb.checked = page ? _selectedRunPages.has(page) : false;
    }
  };

  allCb.addEventListener('change', () => {
    _selectedRunPages = allCb.checked ? new Set(pageNames) : new Set();
    syncAllCheckboxState();
    syncPageCheckboxes();
  });

  allLabel.appendChild(allCb);
  allLabel.appendChild(document.createTextNode('All pages'));
  pagesWrap.appendChild(allLabel);

  for (const pageName of pageNames) {
    const label = document.createElement('label');
    label.style.fontSize = '12px';
    label.style.display = 'flex';
    label.style.alignItems = 'center';
    label.style.gap = '4px';

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.setAttribute('data-page', pageName);
    cb.checked = _selectedRunPages.has(pageName);
    cb.addEventListener('change', () => {
      if (cb.checked) _selectedRunPages.add(pageName);
      else _selectedRunPages.delete(pageName);
      syncAllCheckboxState();
    });

    label.appendChild(cb);
    label.appendChild(document.createTextNode(pageName));
    pagesWrap.appendChild(label);
  }

  syncAllCheckboxState();
  wrap.appendChild(pagesWrap);
  host.appendChild(wrap);
}

function renderGlobalFilters(globalFilters) {
  const el = document.getElementById('globalFiltersResult');
  el.innerHTML = '';
  if (!globalFilters || globalFilters.length === 0) {
    el.innerHTML = '<div class="note">No report-level filters found (or not checked in this discovery mode) — this report likely repeats fields as per-page slicers instead. Use page discovery below.</div>';
    return;
  }
  const pageDiv = document.createElement('div');
  pageDiv.className = 'discover-page';
  const h3 = document.createElement('h3');
  h3.textContent = `Global filters (${globalFilters.length}) — already applied to every page`;
  pageDiv.appendChild(h3);
  for (const f of globalFilters) {
    const row = document.createElement('div');
    row.className = 'discover-slicer';
    const t = f.target || {};
    const label = t.table && t.column ? `${t.table}.${t.column}` : JSON.stringify(t);
    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = label;
    const meta = document.createElement('span');
    meta.className = 'meta';
    meta.textContent = f.values ? f.values.join(', ') : '(hierarchy selection)';
    row.appendChild(name);
    row.appendChild(meta);
    pageDiv.appendChild(row);
  }
  el.appendChild(pageDiv);
}

// ── Parity results (in-page summary of parity-result.json) ──────────────────
//
// A more detailed, always-visible alternative to opening parity-summary.xlsx
// by hand - NOT a replacement for it. Full per-visual detail (sample rows,
// etc.) stays exclusive to the xlsx; this view is page-level, plus a link to
// open each of the three generated files for anyone who wants to go deeper.

const VERDICT_BADGE = {
  pass:          { cls: 'badge lg',        text: '✅ PASS' },
  fail:          { cls: 'badge lg fail',   text: '❌ FAIL' },
  not_comparable: { cls: 'badge lg review', text: '⚠️ NOT COMPARABLE' },
};

const PAGE_STATUS_BADGE = {
  'identical':       { cls: 'badge',        text: 'Identical' },
  'header-only':     { cls: 'badge',        text: 'Labels differ only' },
  'different':       { cls: 'badge fail',   text: 'Different' },
  'only-in-source':  { cls: 'badge review', text: 'Missing in Target' },
  'only-in-target':  { cls: 'badge review', text: 'Only in Target' },
};

async function openOutputFile(filePath, label) {
  const res = await window.verirapportDesktop.openOutputFile(filePath);
  if (!res.ok) appendLog(`--- Could not open ${label}: ${res.error} ---`);
}

function renderParityResults(result) {
  const el = document.getElementById('parityResults');
  el.innerHTML = '';

  if (!result) {
    el.innerHTML = '<div class="note">No parity summary to show yet — either this was a source-only run (nothing was compared), or the run did not get far enough to produce one. Check the log above.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = 'var(--space-3)';
  heading.textContent = 'Parity results';
  el.appendChild(heading);

  for (const sc of result.scenarios || []) {
    const pageDiv = document.createElement('div');
    pageDiv.className = 'discover-page';

    const h3 = document.createElement('h3');
    h3.textContent = sc.name && sc.name !== '(unfiltered)' ? `Scenario: ${sc.name}` : 'Result';
    pageDiv.appendChild(h3);

    const verdictRow = document.createElement('div');
    verdictRow.className = 'parity-verdict';
    const verdict = VERDICT_BADGE[sc.ui.verdict] || VERDICT_BADGE.fail;
    const verdictBadge = document.createElement('span');
    verdictBadge.className = verdict.cls;
    verdictBadge.textContent = verdict.text;
    verdictRow.appendChild(verdictBadge);
    if (sc.ui.differingSheets > 0) {
      const note = document.createElement('span');
      note.className = 'note';
      note.textContent = `${sc.ui.differingSheets} page(s) with real differences`;
      verdictRow.appendChild(note);
    }
    pageDiv.appendChild(verdictRow);

    for (const line of sc.ui.narrative || []) {
      const p = document.createElement('div');
      p.className = 'note';
      p.style.marginBottom = '4px';
      p.textContent = line;
      pageDiv.appendChild(p);
    }

    for (const page of sc.ui.pages || []) {
      const row = document.createElement('div');
      row.className = 'discover-slicer';
      const name = document.createElement('span');
      name.className = 'name';
      name.textContent = page.name;
      const meta = document.createElement('span');
      meta.className = 'meta';
      const statusBadge = PAGE_STATUS_BADGE[page.status] || PAGE_STATUS_BADGE.different;
      const badgeEl = document.createElement('span');
      badgeEl.className = statusBadge.cls;
      badgeEl.textContent = statusBadge.text;
      meta.appendChild(document.createTextNode(`${page.rowsExpected} / ${page.rowsActual} rows `));
      meta.appendChild(badgeEl);
      row.appendChild(name);
      row.appendChild(meta);
      pageDiv.appendChild(row);
    }

    if ((sc.ui.pagesOnlyInSource || []).length || (sc.ui.pagesOnlyInTarget || []).length) {
      const p = document.createElement('div');
      p.className = 'note';
      p.style.marginTop = '6px';
      const parts = [];
      if (sc.ui.pagesOnlyInSource.length) parts.push(`only in Source: ${sc.ui.pagesOnlyInSource.join(', ')}`);
      if (sc.ui.pagesOnlyInTarget.length) parts.push(`only in Target: ${sc.ui.pagesOnlyInTarget.join(', ')}`);
      p.textContent = `Pages ${parts.join('; ')}`;
      pageDiv.appendChild(p);
    }

    const links = document.createElement('div');
    links.className = 'parity-file-links';
    const fileButtons = [
      ['Open Parity Summary', sc.summaryFile],
      ['Open Source file', sc.expectedFile],
      ['Open Target file', sc.actualFile],
    ];
    const folderBtn = document.createElement('button');
    folderBtn.className = 'ghost'; folderBtn.type = 'button'; folderBtn.textContent = 'Open results folder';
    folderBtn.addEventListener('click', async () => {
      const r = await window.verirapportDesktop.openOutputFolder(result.folder);
      if (!r.ok) appendLog(`--- Could not open folder: ${r.error} ---`);
    });
    for (const [label, filePath] of fileButtons) {
      const btn = document.createElement('button');
      btn.className = label === 'Open Parity Summary' ? '' : 'ghost';
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => openOutputFile(filePath, label));
      links.appendChild(btn);
    }
    links.appendChild(folderBtn);
    pageDiv.appendChild(links);

    el.appendChild(pageDiv);
  }
}

function renderFieldPicker(key, pick) {
  const wrap = document.createElement('div');
  wrap.className = 'discover-page';

  const h3 = document.createElement('h3');
  h3.textContent = pick.title;
  if (pick.isHierarchy) {
    const badge = document.createElement('span');
    badge.className = 'badge hierarchy';
    badge.textContent = 'hierarchy — top level only';
    h3.appendChild(badge);
  }
  wrap.appendChild(h3);

  const metaLine = document.createElement('div');
  metaLine.className = 'note';
  metaLine.textContent = pick.targetLabel;
  wrap.appendChild(metaLine);

  // Value picker
  const select = document.createElement('select');
  select.style.marginTop = '6px';
  select.style.width = '100%';
  select.style.padding = '6px';
  select.style.fontSize = '12px';
  const noneOpt = document.createElement('option');
  noneOpt.value = '';
  noneOpt.textContent = pick.options.length > 0 ? '— select a value —' : '(no values available)';
  select.appendChild(noneOpt);
  for (const v of pick.options) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    if (v === pick.value) opt.selected = true;
    select.appendChild(opt);
  }
  select.addEventListener('change', () => { pick.value = select.value; });
  wrap.appendChild(select);

  const note = document.createElement('div');
  note.className = 'note';
  note.style.marginTop = '6px';
  note.textContent = 'This value applies to the selected pages above.';
  wrap.appendChild(note);

  return wrap;
}

function renderDiscoverResults(data) {
  window._discoverMode = 'shared';
  const globalFilters = data.globalFilters || [];
  const byPage = data.pages || {};
  const allPages = Array.isArray(data.allPages) ? data.allPages : [];
  _lastGlobalFilters = globalFilters;
  _discoveredPages = allPages.length > 0 ? allPages : Object.keys(byPage);

  renderGlobalFilters(globalFilters);
  renderPageScopeSelector(getPageUniverse());

  const groups = groupDiscoveredFields(byPage);

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  if (groups.size === 0) {
    container.innerHTML = '<div class="note">No pages were crawled for slicer visuals yet. Tick one or more pages and click Load filters again to build pickable filters.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = '4px';
  heading.textContent = 'Pick values to apply on your selected pages:';
  container.appendChild(heading);

  for (const [key, group] of groups) {
    const pick = getOrCreatePick(window._filterPicks, key, group);
    container.appendChild(renderFieldPicker(key, pick));
  }
}

/**
 * Cross-report mode: source and target discovered the SAME filter titles
 * (matchDetails.identical), so the user only picks a value once - one
 * shared control per field, not two columns. BUT "same title" does not mean
 * "same underlying table.column": a Direct Lake migration can rename the
 * table backing a field while the on-screen slicer title stays identical
 * (observed in production - a field titled the same on both reports bound
 * to 'Germany - Extract'[Recruitment Area] on the source and to
 * 'V_FACT_APPLICATION_GERMANY'[RECRUITMENT_AREA] on the target). Reusing
 * just one side's targets for both reports is exactly what produced a
 * FILTER FIELD MISMATCH abort (previously, before that check existed, a
 * silently one-sided filter). So each merged pick keeps BOTH sides'
 * targets/visualNames (sideTargets), and buildScenarioFromPicks emits two
 * side-tagged selections from the one shared value - see there.
 */
function renderMatchedDiscoverResults(result) {
  window._discoverMode = 'shared';
  _lastGlobalFilters = [];
  const sourcePage = result.sourcePage;
  const targetPage = result.targetPage;

  renderGlobalFilters(null);
  renderPageScopeSelector(getPageUniverse());

  const sourceGroups = groupDiscoveredFields(sourcePage ? { [sourcePage]: result.sourceFields || [] } : {});
  const targetGroups = groupDiscoveredFields(targetPage ? { [targetPage]: result.targetFields || [] } : {});

  // Pair up by normalized title - matchDetails.identical already guarantees
  // every source field has exactly one title+hierarchy-ness counterpart in
  // target and vice versa, AND that no title resolves to more than one
  // distinct field binding on either side (see duplicateTitleConflicts in
  // matchDiscoveredFields, cross-report-match.helpers.ts) - so this pairing
  // should never need to choose between more than one candidate. The
  // grouping-by-title below still defends against it happening anyway: it
  // is exactly what produced a real bug previously (two "Bewerbungseingang
  // Cluster" slicers on different fields both getting silently paired to
  // the SAME target field). If a title ever has more than one distinct
  // variant on either side here, skip it rather than guess.
  const sourceGroupsByTitleKey = new Map();
  for (const g of sourceGroups.values()) {
    const tk = pageNameKeyJs(g.title);
    const list = sourceGroupsByTitleKey.get(tk);
    if (list) list.push(g); else sourceGroupsByTitleKey.set(tk, [g]);
  }
  const targetGroupsByTitleKey = new Map();
  for (const g of targetGroups.values()) {
    const tk = pageNameKeyJs(g.title);
    const list = targetGroupsByTitleKey.get(tk);
    if (list) list.push(g); else targetGroupsByTitleKey.set(tk, [g]);
  }

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  const merged = [];
  for (const [titleKey, srcGroupsForTitle] of sourceGroupsByTitleKey) {
    const tgtGroupsForTitle = targetGroupsByTitleKey.get(titleKey);
    if (!tgtGroupsForTitle) continue; // should not happen when result.identical is true
    if (srcGroupsForTitle.length > 1 || tgtGroupsForTitle.length > 1) {
      console.warn(`[verirapport] Shared picker: "${srcGroupsForTitle[0].title}" has more than one distinct field binding on at least one side - skipping rather than guessing which pairs with which. This should not happen when identical=true.`);
      continue;
    }
    const sourceGroup = srcGroupsForTitle[0];
    const targetGroup = tgtGroupsForTitle[0];
    merged.push({ key: fieldKey(sourceGroup.title, sourceGroup.targetLabel), sourceGroup, targetGroup });
  }

  if (merged.length === 0) {
    container.innerHTML = '<div class="note">No pages were crawled for slicer visuals yet. Tick one or more pages and click Load filters again to build pickable filters.</div>';
    return;
  }

  const heading = document.createElement('div');
  heading.className = 'section-title';
  heading.style.marginTop = '4px';
  heading.textContent = 'Pick values to apply on your selected pages (applied separately to each report\'s own fields):';
  container.appendChild(heading);

  for (const { key, sourceGroup, targetGroup } of merged) {
    const sameLabel = sourceGroup.targetLabel === targetGroup.targetLabel;
    const group = {
      title: sourceGroup.title,
      isHierarchy: sourceGroup.isHierarchy,
      targetLabel: sameLabel
        ? sourceGroup.targetLabel
        : `${sourceGroup.targetLabel}  (target report: ${targetGroup.targetLabel})`,
      options: sourceGroup.options.length >= targetGroup.options.length ? sourceGroup.options : targetGroup.options,
      sideTargets: {
        source: { targets: sourceGroup.targets, byPage: sourceGroup.byPage },
        target: { targets: targetGroup.targets, byPage: targetGroup.byPage },
      },
    };
    const pick = getOrCreatePick(window._filterPicks, key, group);
    container.appendChild(renderFieldPicker(key, pick));
  }
}

/**
 * Cross-report mode: source and target discovered DIFFERENT filters (or a
 * flat-vs-hierarchy mismatch on a same-named field), so a single shared
 * picker would be wrong - render two independent columns instead, with
 * their own pick state, and tag each pick with which side it came from so
 * buildScenarioFromPicks only applies it there.
 */
function renderSplitDiscoverResults(result) {
  window._discoverMode = 'split';
  if (!window._filterPicksSource) window._filterPicksSource = new Map();
  if (!window._filterPicksTarget) window._filterPicksTarget = new Map();

  const container = document.getElementById('discoverResults');
  container.innerHTML = '';

  const d = result.matchDetails || {};
  const reasons = [];
  if (d.onlyInSource && d.onlyInSource.length) reasons.push(`only in Source: ${d.onlyInSource.join(', ')}`);
  if (d.onlyInTarget && d.onlyInTarget.length) reasons.push(`only in Target: ${d.onlyInTarget.join(', ')}`);
  if (d.hierarchyMismatch && d.hierarchyMismatch.length) reasons.push(`type differs (flat vs hierarchy): ${d.hierarchyMismatch.join(', ')}`);
  if (d.duplicateTitleConflicts && d.duplicateTitleConflicts.length) reasons.push(`same title, different field on at least one side: ${d.duplicateTitleConflicts.join(', ')}`);

  const banner = document.createElement('div');
  banner.className = 'mismatch-banner';
  banner.textContent = 'Filters are not identical between the two reports' +
    (reasons.length ? ` (${reasons.join('; ')})` : ' (no common filters were found)') +
    ' — select values separately for each report below.';
  container.appendChild(banner);

  const columns = document.createElement('div');
  columns.className = 'split-columns';

  const buildColumn = (label, pageName, fields, picksMap) => {
    const col = document.createElement('div');
    col.className = 'split-column';
    const h = document.createElement('div');
    h.className = 'section-title';
    h.textContent = pageName ? `${label} filters (page: ${pageName})` : `${label} filters`;
    col.appendChild(h);

    const groups = groupDiscoveredFields(pageName ? { [pageName]: (fields || []) } : {});
    if (groups.size === 0) {
      const p = document.createElement('div');
      p.className = 'note';
      p.textContent = 'No slicers found.';
      col.appendChild(p);
      return col;
    }
    for (const [key, group] of groups) {
      const pick = getOrCreatePick(picksMap, key, group);
      col.appendChild(renderFieldPicker(key, pick));
    }
    return col;
  };

  columns.appendChild(buildColumn('Source', result.sourcePage, result.sourceFields, window._filterPicksSource));
  columns.appendChild(buildColumn('Target', result.targetPage, result.targetFields, window._filterPicksTarget));
  container.appendChild(columns);
}

document.getElementById('btnDiscover').addEventListener('click', async () => {
  // Defense in depth: the button is already disabled without a session, but
  // re-check at click time in case the UI state is ever stale - starting
  // this run with no session is a guaranteed, confusing failure (see
  // updateButtonStates's comment for why).
  await refreshAuthGate();
  if (!_hasSession) {
    showErrorText('No saved login session yet. Click "Sign in again" (or Run comparison) first.');
    return;
  }

  const pairName = document.getElementById('pairName').value.trim() || 'PowerBI-Report';
  const skipGlobalCheck = document.getElementById('discoverSkipGlobalCheck')?.checked === true;

  if (bothIdentitiesFilled()) {
    const source = readIdentity('source');
    const target = readIdentity('target');
    appendLog('--- Discovering filters on BOTH reports (comparing source vs target) ---');
    if (skipGlobalCheck) {
      appendLog('--- Skipping report-level filter check (uncheck the box above to re-enable) ---');
    }
    document.getElementById('globalFiltersResult').textContent = '';
    document.getElementById('pageScopeResult').textContent = 'Loading page list…';
    document.getElementById('discoverResults').textContent = 'Discovering source, then target… two fast first-match scans, one after the other.';

    const res = await window.verirapportDesktop.discoverCrossReport(pairName, source, target, skipGlobalCheck);
    if (!res.ok) {
      document.getElementById('pageScopeResult').textContent = '';
      document.getElementById('discoverResults').textContent = `Discovery failed: ${res.error}`;
      return;
    }

    const result = res.result;
    _discoveredPages = dedupePageNames([...(result.sourceAllPages || []), ...(result.targetAllPages || [])]);
    renderPageScopeSelector(getPageUniverse());

    if (result.identical) {
      appendLog(`--- Filters match between reports (page "${result.sourcePage}") — one shared picker ---`);
      renderMatchedDiscoverResults(result);
    } else {
      appendLog('--- Filters do NOT match between reports — select values separately below ---');
      renderSplitDiscoverResults(result);
    }
    return;
  }

  const pagesCsv = getDiscoverPagesCsv();

  const { identity, side, missing, note } = resolveDiscoverIdentity();
  if (missing.length > 0) {
    showErrorText(`Load filters needs a complete report: ${note || 'fill in the IDs above'}.`);
    return;
  }

  appendLog(pagesCsv
    ? `--- Discovering filters: ${skipGlobalCheck ? 'pages' : 'global check + pages'} [${pagesCsv}] (applies nothing) ---`
    : `--- Discovering filters: ${skipGlobalCheck ? '(global check skipped)' : 'global check only'} (applies nothing) ---`);
  appendLog(`--- Using ${side === 'target' ? 'Target' : 'Source'} report entered above ---`);
  document.getElementById('globalFiltersResult').textContent = skipGlobalCheck ? '' : 'Checking…';
  document.getElementById('pageScopeResult').textContent = 'Loading page list…';
  document.getElementById('discoverResults').textContent = pagesCsv ? 'Crawling pages… this can take a while on a page-heavy report.' : '';
  const res = await window.verirapportDesktop.discoverSlicers(pairName, pagesCsv, identity, side, skipGlobalCheck);
  if (!res.ok) {
    document.getElementById('globalFiltersResult').textContent = '';
    document.getElementById('discoverResults').textContent = `Discovery failed: ${res.error}`;
    return;
  }
  renderDiscoverResults(res.result);
});