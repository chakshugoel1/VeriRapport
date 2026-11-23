/**
 * timing.helpers.ts
 *
 * Lightweight phase timer used to find where a parity run actually spends its
 * time. It only measures and logs — it never changes what the run does.
 *
 *   await timed('source', 'embed', () => loadHarnessPage(page, token));
 *   ...
 *   console.log(formatTimingSummary());
 *   writeTimingReport('playwright-report-parity/<pair>/timing.json');
 */

import * as fs   from 'fs';
import * as path from 'path';

export interface TimingEntry {
  /** Grouping, e.g. "source", "target", "compare". */
  phase: string;
  /** What was timed, e.g. "embed", "page:CXO:export". */
  label: string;
  ms:    number;
}

const entries: TimingEntry[] = [];
const runStart = Date.now();

export function recordTiming(phase: string, label: string, ms: number): void {
  entries.push({ phase, label, ms });
}

/** Runs `fn`, records how long it took (even if it throws), returns its result. */
export async function timed<T>(phase: string, label: string, fn: () => Promise<T>): Promise<T> {
  const t0 = Date.now();
  try {
    return await fn();
  } finally {
    recordTiming(phase, label, Date.now() - t0);
  }
}

export function getTimings(): TimingEntry[] {
  return entries.slice();
}

export function resetTimings(): void {
  entries.length = 0;
}

const fmt = (ms: number): string => `${(ms / 1000).toFixed(1)}s`;

/** Sum of ms per "kind" (the label with any ":<page>:" part removed). */
function kindOf(label: string): string {
  const parts = label.split(':');
  return parts.length >= 3 ? parts[parts.length - 1] : label;
}

export function formatTimingSummary(): string {
  const lines: string[] = [];
  lines.push('\n──────────── TIMING SUMMARY ────────────');
  const phases = [...new Set(entries.map(e => e.phase))];
  for (const phase of phases) {
    const inPhase = entries.filter(e => e.phase === phase);
    const total = inPhase.reduce((s, e) => s + e.ms, 0);
    lines.push(`  ${phase}: ${fmt(total)} (sum of measured steps)`);
    const byKind = new Map<string, { ms: number; n: number }>();
    for (const e of inPhase) {
      const k = kindOf(e.label);
      const cur = byKind.get(k) ?? { ms: 0, n: 0 };
      cur.ms += e.ms; cur.n += 1;
      byKind.set(k, cur);
    }
    for (const [k, v] of [...byKind.entries()].sort((a, b) => b[1].ms - a[1].ms)) {
      lines.push(`      ${k.padEnd(14)} ${fmt(v.ms).padStart(9)}  x${v.n}`);
    }
  }
  lines.push(`  wall clock since start: ${fmt(Date.now() - runStart)}`);
  lines.push('────────────────────────────────────────');
  return lines.join('\n');
}

export function writeTimingReport(file: string): void {
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(
      file,
      JSON.stringify({ wallClockMs: Date.now() - runStart, entries }, null, 2),
      'utf8',
    );
  } catch (e) {
    console.warn(`[timing] Could not write ${file} (non-fatal): ${(e as Error).message}`);
  }
}
