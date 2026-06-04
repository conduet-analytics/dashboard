/**
 * sharepoint.js
 * -------------
 * Browser-side data layer — reads the SharePoint Excel file directly
 * via Microsoft Graph API using the user's MSAL token.
 *
 * This replaces the Node.js backend dataSource.js for GitHub Pages deployment.
 *
 * To switch to a different data source later (Azure SQL, Dataverse, etc.),
 * replace the loadFromSharePoint() function — the rest of the app doesn't change.
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

// ── Load raw rows from SharePoint Excel via Graph API ────────────────────────
/**
 * @param {Object} filters - { year, month, location, empStatus }
 * @param {string} token   - Microsoft Graph access token from MSAL
 * @returns {Promise<Array>} Array of row objects with standardized field names
 */
async function loadFromSharePoint(filters, token) {
  const { graphBase, driveId, itemId, sheetName, batchSize, concurrency } = SP_CONFIG;
  const wbBase = `/drives/${driveId}/items/${itemId}/workbook`;
  const wsPath = `${wbBase}/worksheets('${encodeURIComponent(sheetName)}')`;

  // Create a non-persistent session for better performance
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

    const calMonthIdx = colIdx['Cal Month'];
    const yearIdx = colIdx['Year'];

    // Build month filter set: Jan → selected month
    const monthsSet = new Set();
    if (filters.year && filters.month) {
      const upTo = MONTH_ORDER.indexOf(filters.month);
      MONTH_ORDER.slice(0, upTo + 1).forEach(m => monthsSet.add(`${filters.year}-${m}`));
    } else if (filters.year) {
      MONTH_ORDER.forEach(m => monthsSet.add(`${filters.year}-${m}`));
    }

    // Build batch ranges
    const batchDefs = [];
    for (let start = 2; start <= totalRows; start += batchSize) {
      const end = Math.min(start + batchSize - 1, totalRows);
      batchDefs.push({ start, end });
    }

    // Fetch in parallel batches
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
          const rowMonth = rawRow[calMonthIdx];
          const rowYear = rawRow[yearIdx];
          if (!monthsSet.size || monthsSet.has(`${rowYear}-${rowMonth}`)) {
            const mapped = mapRow(colIdx, rawRow);
            if (filters.location && mapped.location !== filters.location) continue;
            if (filters.empStatus && mapped.emp_status !== filters.empStatus) continue;
            allRows.push(mapped);
          }
        }
      }
    }

    return allRows;

  } finally {
    // Always close the session
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

// ── Public API (used by app.js) ──────────────────────────────────────────────

/**
 * Fetch rows from SharePoint for the given filters.
 * @param {Object} filters - { year, month }
 * @returns {Promise<Array>} Filtered row objects
 */
async function fetchRows(filters) {
  const token = await getToken();
  if (!token) throw new Error('No access token — please sign in');
  return loadFromSharePoint(filters, token);
}

/**
 * Get available periods (year/month) from the data.
 * Reads the Year and Cal Month columns to find all distinct periods.
 */
async function fetchAvailablePeriods() {
  const token = await getToken();
  if (!token) throw new Error('No access token — please sign in');

  const { graphBase, driveId, itemId, sheetName } = SP_CONFIG;
  const wsPath = `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')`;

  // Get headers to find Year and Cal Month column indices
  const hdrRes = await graphFetch(`${wsPath}/range(address='A1:AZ1')?$select=values`, token);
  const headers = hdrRes.values[0];
  const yearCol = headers.indexOf('Year');
  const monthCol = headers.indexOf('Cal Month');

  if (yearCol === -1 || monthCol === -1) {
    throw new Error('Could not find Year or Cal Month columns in the spreadsheet');
  }

  // Convert to column letters (A, B, ... Z, AA, etc.)
  const colLetter = (idx) => {
    let s = '';
    while (idx >= 0) {
      s = String.fromCharCode(65 + (idx % 26)) + s;
      idx = Math.floor(idx / 26) - 1;
    }
    return s;
  };

  const yearLetter = colLetter(yearCol);
  const monthLetter = colLetter(monthCol);

  // Get used range row count
  const dimRes = await graphFetch(`${wsPath}/usedRange(valuesOnly=true)?$select=rowCount`, token);
  const totalRows = dimRes.rowCount;

  // Fetch year and month columns only (lighter than full sheet)
  const seen = new Set();
  const periods = [];
  const batchSize = 10000;

  for (let start = 2; start <= totalRows; start += batchSize) {
    const end = Math.min(start + batchSize - 1, totalRows);
    // Fetch both columns in one range request
    const rangeRes = await graphFetch(
      `${wsPath}/range(address='${yearLetter}${start}:${monthLetter}${end}')?$select=values`,
      token
    );
    for (const row of rangeRes.values) {
      // The range is from yearCol to monthCol, so indices shift
      const y = yearCol <= monthCol ? row[0] : row[row.length - 1];
      const m = yearCol <= monthCol ? row[monthCol - yearCol] : row[0];
      if (y && m) {
        const key = `${y}-${m}`;
        if (!seen.has(key)) {
          seen.add(key);
          periods.push({ year: Number(y), month: String(m) });
        }
      }
    }
  }

  return periods.sort((a, b) =>
    a.year !== b.year
      ? a.year - b.year
      : MONTH_ORDER.indexOf(a.month) - MONTH_ORDER.indexOf(b.month)
  );
}
