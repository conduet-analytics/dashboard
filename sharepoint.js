/**
 * sharepoint.js
 * -------------
 * Browser-side data layer — reads the SharePoint Excel file directly
 * via Microsoft Graph API using the user's MSAL token.
 *
 * OPTIMIZED: fetches ALL rows once, caches in memory.
 * Period switching and refreshes use the cache (no re-fetch).
 */

'use strict';

// ── SharePoint / Graph config ────────────────────────────────────────────────
const SP_CONFIG = {
  graphBase:  'https://graph.microsoft.com/v1.0',
  driveId:    'b!d-WPImoa-0mrT0yO4zMkGQ4ENzHxU5xEmMMRHEAXItThtQ_sSWUrTK3KCi7xIKJt',
  itemId:     '01WXH6UQLLYEJT652WHVB3YHHWKHKE7JN5',
  sheetName:  'Main_Raw Data',
  batchSize:  5000,
  concurrency: 4,
};

const MONTH_ORDER = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December'
];

// ── Column name → snake_case field name map ──────────────────────────────────
const FIELD_MAP = {
  'Location':                           'location',
  'Agent Name':                         'agent_name',
  'Line Manager':                       'line_manager',
  'Week End':                           'week_end',
  'Cal Month':                          'cal_month',
  'Year':                               'year',
  'Emp Status':                         'emp_status',
  'Messaging Handled':                  'messaging_handled',
  'Messaging AHT (sec)':               'messaging_aht_sec',
  'Messaging within SLA':              'messaging_within_sla',
  'Email Handled':                      'email_handled',
  'Email AHT (sec)':                   'email_aht_sec',
  'Email Contacts w/ Handle Time':     'email_contacts_handle_time',
  'Email SLA < 12 hrs':               'email_sla_12hrs',
  'Productive Hours':                   'productive_hours',
  'Overall CSAT Score':                'overall_csat_score',
  'Overall Survey Count without blank': 'overall_survey_count',
  'Messaging CSAT Score':              'messaging_csat_score',
  'Messaging Survey Count without blank': 'messaging_survey_count',
  'Email CSAT Score':                  'email_csat_score',
  'Email Survey Count without blank':  'email_survey_count',
  'Total QA Score':                    'qa_score',
  'Total QA Reviews':                  'qa_reviews',
  'Login Hours':                       'login_hours',
  'Scheduled Hours':                   'scheduled_hours',
  'Rostered less Leave':               'rostered_less_leave',
  'OOA':                               'ooa',
};

// ── In-memory cache ─────────────────────────────────────────────────────────
let _cache = null; // { allRows: [], periods: [], fetchedAt: Date }

// ── Graph API fetch helper ───────────────────────────────────────────────────
async function graphFetch(path, token, sessionId) {
  const headers = { Authorization: `Bearer ${token}` };
  if (sessionId) headers['workbook-session-id'] = sessionId;

  const res = await fetch(`${SP_CONFIG.graphBase}${path}`, { headers });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API ${res.status}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

// ── Convert Excel date serial → YYYY-MM-DD ───────────────────────────────────
function excelDateToISO(serial) {
  if (!serial) return null;
  if (typeof serial === 'string' && serial.includes('-')) return serial.split('T')[0];
  const d = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  return d.toISOString().split('T')[0];
}

// ── Map a raw Excel row array to a named-field object ────────────────────────
function mapRow(colIdx, rawRow) {
  const obj = {};
  for (const [excelName, fieldName] of Object.entries(FIELD_MAP)) {
    const i = colIdx[excelName];
    let val = i !== undefined ? rawRow[i] : null;
    if (val === '' || val === undefined) val = null;
    if (fieldName === 'week_end' && val !== null) val = excelDateToISO(val);
    obj[fieldName] = val;
  }
  return obj;
}

// ── Load ALL rows from SharePoint Excel (single fetch, cached) ──────────────
async function loadAllRows(token) {
  const { graphBase, driveId, itemId, sheetName, batchSize, concurrency } = SP_CONFIG;
  const wbBase = `/drives/${driveId}/items/${itemId}/workbook`;
  const wsPath = `${wbBase}/worksheets('${encodeURIComponent(sheetName)}')`;

  // Create a non-persistent session
  const sessionRes = await fetch(`${graphBase}${wbBase}/createSession`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ persistChanges: false }),
  });
  if (!sessionRes.ok) {
    const errText = await sessionRes.text();
    throw new Error(`Failed to create workbook session: ${sessionRes.status} — ${errText.slice(0, 200)}`);
  }
  const session = await sessionRes.json();
  const sessionId = session.id;

  try {
    // Get total row count + headers in parallel
    const [dimRes, hdrRes] = await Promise.all([
      graphFetch(`${wsPath}/usedRange(valuesOnly=true)?$select=rowCount`, token, sessionId),
      graphFetch(`${wsPath}/range(address='A1:AZ1')?$select=values`, token, sessionId),
    ]);

    const totalRows = dimRes.rowCount;
    const headers = hdrRes.values[0];
    const colIdx = {};
    headers.forEach((h, i) => { if (h) colIdx[h] = i; });

    // Build batch ranges
    const batchDefs = [];
    for (let start = 2; start <= totalRows; start += batchSize) {
      const end = Math.min(start + batchSize - 1, totalRows);
      batchDefs.push({ start, end });
    }

    // Fetch ALL rows in parallel batches
    const allRows = [];
    for (let i = 0; i < batchDefs.length; i += concurrency) {
      const chunk = batchDefs.slice(i, i + concurrency);
      const results = await Promise.all(
        chunk.map(({ start, end }) =>
          graphFetch(
            `${wsPath}/range(address='A${start}:AZ${end}')?$select=values`,
            token, sessionId
          )
        )
      );
      for (const batchRes of results) {
        for (const rawRow of batchRes.values) {
          const mapped = mapRow(colIdx, rawRow);
          if (mapped.year && mapped.cal_month) {
            allRows.push(mapped);
          }
        }
      }
    }

    return allRows;

  } finally {
    fetch(`${graphBase}${wbBase}/closeSession`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'workbook-session-id': sessionId,
        'Content-Type': 'application/json',
      },
    }).catch(() => {});
  }
}

// ── Ensure cache is populated ────────────────────────────────────────────────
async function ensureCache(forceRefresh) {
  if (_cache && !forceRefresh) return _cache;

  const token = await getToken();
  if (!token) throw new Error('No access token — please sign in');

  const allRows = await loadAllRows(token);

  // Derive periods from data
  const seen = new Set();
  const periods = [];
  for (const r of allRows) {
    const key = `${r.year}-${r.cal_month}`;
    if (!seen.has(key)) {
      seen.add(key);
      periods.push({ year: Number(r.year), month: String(r.cal_month) });
    }
  }
  periods.sort((a, b) =>
    a.year !== b.year
      ? a.year - b.year
      : MONTH_ORDER.indexOf(a.month) - MONTH_ORDER.indexOf(b.month)
  );

  _cache = { allRows, periods, fetchedAt: new Date() };
  return _cache;
}

// ── Public API (used by app.js) ──────────────────────────────────────────────

/**
 * Fetch rows for the given filters. Uses cache — no re-fetch unless forceRefresh.
 */
async function fetchRows(filters, forceRefresh) {
  const cache = await ensureCache(forceRefresh);
  let rows = cache.allRows;

  // Filter by year + months (Jan → selected month)
  if (filters.year && filters.month) {
    const upTo = MONTH_ORDER.indexOf(filters.month);
    const monthsSet = new Set(MONTH_ORDER.slice(0, upTo + 1));
    rows = rows.filter(r =>
      Number(r.year) === Number(filters.year) && monthsSet.has(r.cal_month)
    );
  } else if (filters.year) {
    rows = rows.filter(r => Number(r.year) === Number(filters.year));
  }

  if (filters.location) rows = rows.filter(r => r.location === filters.location);
  if (filters.empStatus) rows = rows.filter(r => r.emp_status === filters.empStatus);

  return rows;
}

/**
 * Get available periods. Uses cache — no separate API call.
 */
async function fetchAvailablePeriods() {
  const cache = await ensureCache(false);
  return cache.periods;
}

/**
 * Force a fresh fetch from SharePoint (called by Refresh button).
 */
async function forceRefreshData() {
  _cache = null;
}
