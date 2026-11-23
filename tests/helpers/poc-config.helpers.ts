import * as path from 'path';
import * as dotenv from 'dotenv';

export interface PocConfig {
  tenantId: string;
  groupId: string;
  reportId: string;
  workspaceUrl: string;
  reportUrl: string;
  datasetId: string;
  rlsRole: string;
}

// No tenant/workspace/report/dataset IDs are baked in: they come from the app's
// fields (applyReportIdentity) or from env vars / the .env file.
const DEFAULT_RLS_ROLE = 'DynamicRoles';

const SECRETS_DIR = path.join(process.env['USERPROFILE'] ?? process.env['HOME'] ?? '', 'Power_BI_report_validation_credentials');
const ENV_FILE = path.join(SECRETS_DIR, '.env');

dotenv.config({ path: ENV_FILE, override: false });

function parseGroupAndReportIds(reportUrl: string): { groupId?: string; reportId?: string } {
  const match = reportUrl.match(/\/groups\/([^\/]+)\/reports\/([^\/?]+)/i);
  if (!match) return {};
  return { groupId: match[1], reportId: match[2] };
}

// Neutral Power BI landing page used for sign-in. A workspace page is only used
// when the user explicitly sets PBI_WORKSPACE_URL.
const SIGN_IN_URL = 'https://app.powerbi.com/home';

function buildReportUrl(groupId: string, reportId: string): string {
  return `https://app.powerbi.com/groups/${groupId}/reports/${reportId}?experience=power-bi`;
}

function firstNonEmpty(...values: Array<string | undefined>): string | undefined {
  for (const value of values) {
    const trimmed = (value ?? '').trim();
    if (trimmed) return trimmed;
  }
  return undefined;
}

export function getPocConfig(): PocConfig {
  const envReportUrl = (process.env['PBI_REPORT_URL'] ?? '').trim();
  const idsFromReport = parseGroupAndReportIds(envReportUrl);

  const tenantId = firstNonEmpty(
    process.env['PowerBITenantId'],
    process.env['PBI_TENANT_ID'],
  ) ?? '';
  const groupId = firstNonEmpty(
    process.env['PowerBIWorkspaceId'],
    process.env['PBI_GROUP_ID'],
    idsFromReport.groupId,
  ) ?? '';
  const reportId = firstNonEmpty(
    process.env['PowerBIReportId'],
    process.env['PBI_REPORT_ID'],
    idsFromReport.reportId,
  ) ?? '';

  const workspaceUrl = (process.env['PBI_WORKSPACE_URL'] ?? '').trim() || SIGN_IN_URL;

  const reportUrl = (envReportUrl || (groupId && reportId ? buildReportUrl(groupId, reportId) : '')).trim();

  const datasetId = firstNonEmpty(
    process.env['PBI_DATASET_ID'],
    process.env['PowerBIDatasetId'],
  ) ?? '';

  const rlsRole = firstNonEmpty(
    process.env['PBI_RLS_ROLE'],
    process.env['PowerBIRLSRole'],
    DEFAULT_RLS_ROLE,
  ) as string;

  return {
    tenantId,
    groupId,
    reportId,
    workspaceUrl,
    reportUrl,
    datasetId,
    rlsRole,
  };
}

export function getConfiguredTargetPages<T extends { name: string; displayName: string }>(
  defaultPages: T[],
): T[] {
  const raw = (process.env['PBI_TARGET_PAGES_JSON'] ?? '').trim();
  if (!raw) return defaultPages;

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return defaultPages;

    const pages = parsed
      .filter((p: any) => p && typeof p.name === 'string' && typeof p.displayName === 'string')
      .map((p: any) => ({ name: p.name, displayName: p.displayName })) as T[];

    return pages.length > 0 ? pages : defaultPages;
  } catch {
    return defaultPages;
  }
}
