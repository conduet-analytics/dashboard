/**
 * app.js
 * ------
 * Frontend logic. Calls API endpoints, renders results.
 * NO calculations here — only display logic.
 *
 * API_BASE switches automatically:
 *   - When backend is running locally → calls http://localhost:3000/api
 *   - When deployed on Azure → set BACKEND_URL in config
 *   - When hosted on GitHub Pages (no backend) → calls /static-api (pre-built JSON)
 *
 * Auth: token from auth.js is passed in every API request header.
 * Backend uses it to call Microsoft Graph on behalf of the user.
 */

'use strict';

// ── API configuration ─────────────────────────────────────────────────────────
// Change BACKEND_URL to your Azure App Service URL when deployed
const BACKEND_URL = null; // e.g. 'https://kpi-api.azurewebsites.net'

const API_BASE = BACKEND_URL
  ? `${BACKEND_URL}/api`
  : (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')
    ? 'http://localhost:3000/api'
    : null; // GitHub Pages — use static JSON fallback

const USE_STATIC = !API_BASE;

// ── Utility: build API URL ────────────────────────────────────────────────────
function apiUrl(endpoint, params = {}) {
  if (!USE_STATIC) {
    const qs = new URLSearchParams(params).toString();
    return `${API_BASE}/${endpoint}${qs ? '?' + qs : ''}`;
  }
  // Static JSON fallback mapping (GitHub Pages, no backend)
  const staticMap = {
    'periods':         'periods.json',
    'kpis':            'kpis.json',
    'trends':          'trends.json',
    'team-summary':    params.view === 'monthly'
                         ? 'team-summary-monthly.json'
                         : 'team-summary-weekly.json',
  };
  // Use relative path for GitHub Pages — /static-api/filename.json
  const repoBase = window.location.pathname.replace(/\/[^/]*$/, ''); // e.g. /dashboard
  return `${repoBase}/static-api/${staticMap[endpoint] || endpoint + '.json'}`;
}

// ── Fetch wrapper — includes auth token for live backend ─────────────────────
async function apiFetch(endpoint, params = {}) {
  const url = apiUrl(endpoint, params);
  const headers = { 'Content-Type': 'application/json' };

  // Attach user's token when calling live backend
  if (!USE_STATIC && typeof getToken === 'function') {
    try {
      const token = await getToken();
      if (token) headers['Authorization'] = `Bearer ${token}`;
    } catch (e) {
      console.warn('Could not get token for API call:', e.message);
    }
  }

  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`API error ${res.status} for ${url}`);
  return res.json();
}

// ── Formatters ────────────────────────────────────────────────────────────────
const fmt = {
  dec:  v => v == null ? '—' : Number(v).toFixed(2),
  pct:  v => v == null ? '—' : Number(v).toFixed(2) + '%',
  num:  v => v == null ? '—' : Number(v).toLocaleString(),
};

function trend(curr, prev, higherBetter) {
  if (curr == null || prev == null) return '';
  const d = curr - prev;
  const p = Math.abs(d / (prev || 1));
  if (p < 0.005) return '<span class="fl">▬</span>';
  const good = higherBetter ? d > 0 : d < 0;
  return good ? '<span class="up">▲</span>' : '<span class="dn">▼</span>';
}

// ── KPI metric definitions (display only, no formulas) ───────────────────────
const METRICS = [
  { key: 'csat',           label: 'CSAT',                  fmt: fmt.dec, higher: true  },
  { key: 'messaging_csat', label: 'Messaging CSAT',         fmt: fmt.dec, higher: true  },
  { key: 'email_csat',     label: 'Email CSAT',             fmt: fmt.dec, higher: true  },
  { key: 'aht_messaging',  label: 'AHT Messaging (mins)',   fmt: fmt.dec, higher: false },
  { key: 'aht_email',      label: 'AHT Email (mins)',       fmt: fmt.dec, higher: false },
  { key: 'quality',        label: 'Quality',                fmt: fmt.pct, higher: true  },
  { key: 'cph',            label: 'CPH',                    fmt: fmt.dec, higher: true  },
  { key: 'messaging_sla',  label: 'Messaging SLA',          fmt: fmt.pct, higher: true  },
  { key: 'email_sla',      label: 'Email SLA',              fmt: fmt.pct, higher: true  },
  { key: 'attendance',     label: 'Attendance %',           fmt: fmt.pct, higher: true  },
  { key: 'adherence',      label: 'Adherence %',            fmt: fmt.pct, higher: true  },
  { key: 'total_contacts', label: 'Total Contacts Handled', fmt: fmt.num, higher: null  },
];

// ─────────────────────────────────────────────────────────────────────────────
// RENDER: Snapshot Cards
// ─────────────────────────────────────────────────────────────────────────────
function renderSnapCards(currentKpis, prevKpis) {
  // Accepts either { conduet, hrd } directly or wraps it
  const current  = currentKpis?.conduet ? currentKpis : { conduet: currentKpis, hrd: null };
  const previous = prevKpis?.conduet    ? prevKpis    : { conduet: prevKpis,    hrd: null };
  const c = current?.conduet?.total;
  const p = previous?.conduet?.total;
  const h = current?.hrd?.total;

  const badgeCls = (curr, prev, hb) => {
    if (!prev || curr == null) return '';
    const d = (curr - prev) / (prev || 1) * 100;
    const cls = hb ? (d > 0.5 ? 'b-up' : d < -0.5 ? 'b-dn' : 'b-fl')
                   : (d < -0.5 ? 'b-up' : d > 0.5 ? 'b-dn' : 'b-fl');
    const icon = { 'b-up': '▲', 'b-dn': '▼', 'b-fl': '▬' }[cls];
    return `<span class="badge ${cls}">${icon} ${(d > 0 ? '+' : '')}${d.toFixed(1)}%</span>`;
  };

  const cards = [
    { lbl: 'Conduet HC',   val: c?.hc ?? '—',
      sub: `T:${current?.conduet?.tenured?.hc ?? '—'} · NH:${current?.conduet?.new_hires?.hc ?? '—'}`,
      badge: badgeCls(c?.hc, p?.hc, true), cls: '' },
    { lbl: 'HRD HC',       val: h?.hc ?? '—', sub: 'Total only',
      badge: '', cls: 'r' },
    { lbl: 'CSAT',         val: fmt.dec(c?.csat),
      sub: `T:${fmt.dec(current?.conduet?.tenured?.csat)} · NH:${fmt.dec(current?.conduet?.new_hires?.csat)}`,
      badge: badgeCls(c?.csat, p?.csat, true), cls: 't' },
    { lbl: 'Msg CSAT',     val: fmt.dec(c?.messaging_csat),
      sub: `HRD: ${fmt.dec(h?.messaging_csat)}`,
      badge: badgeCls(c?.messaging_csat, p?.messaging_csat, true), cls: 't' },
    { lbl: 'Quality',      val: fmt.pct(c?.quality),
      sub: `T:${fmt.pct(current?.conduet?.tenured?.quality)} · NH:${fmt.pct(current?.conduet?.new_hires?.quality)}`,
      badge: badgeCls(c?.quality, p?.quality, true), cls: 'g' },
    { lbl: 'Msg AHT',      val: (c?.aht_messaging ?? '—') + (c?.aht_messaging ? 'm' : ''),
      sub: `T:${fmt.dec(current?.conduet?.tenured?.aht_messaging)} · NH:${fmt.dec(current?.conduet?.new_hires?.aht_messaging)}`,
      badge: badgeCls(c?.aht_messaging, p?.aht_messaging, false), cls: 'o' },
    { lbl: 'Msg SLA',      val: fmt.pct(c?.messaging_sla),
      sub: `HRD: ${fmt.pct(h?.messaging_sla)}`,
      badge: badgeCls(c?.messaging_sla, p?.messaging_sla, true), cls: 'b' },
    { lbl: 'Contacts',     val: fmt.num((c?.total_contacts || 0) + (h?.total_contacts || 0)),
      sub: `C:${fmt.num(c?.total_contacts)} · HRD:${fmt.num(h?.total_contacts)}`,
      badge: '', cls: 'r' },
  ];

  document.getElementById('snap-cards').innerHTML = cards.map(card =>
    `<div class="card ${card.cls}">
      <div class="card-lbl">${card.lbl}</div>
      <div class="card-val">${card.val}</div>
      <div class="card-sub">${card.sub}</div>
      ${card.badge}
    </div>`
  ).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER: Charts
// ─────────────────────────────────────────────────────────────────────────────
let charts = {};
function renderCharts(trends) {
  const labels = trends.map(t => t.label);
  const get = (grp, subgrp, key) => trends.map(t => t.kpis?.[grp]?.[subgrp]?.[key] ?? null);

  const opts = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'bottom', labels: { font: { size: 9.5 }, boxWidth: 9, padding: 7 } } },
    scales: {
      x: { ticks: { font: { size: 9.5 } }, grid: { display: false } },
      y: { ticks: { font: { size: 9.5 } }, grid: { color: '#f0f0f0' } }
    }
  };

  const line = (id, datasets, yExtra) => {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(document.getElementById(id), {
      type: 'line',
      data: { labels, datasets },
      options: { ...opts, scales: { ...opts.scales, y: { ...opts.scales.y, ...(yExtra || {}) } } }
    });
  };
  const bar = (id, datasets) => {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(document.getElementById(id), {
      type: 'bar', data: { labels, datasets }, options: opts
    });
  };
  const ds = (lbl, data, color, dash) => ({
    label: lbl, data, borderColor: color,
    backgroundColor: dash ? 'transparent' : color + '22',
    tension: .35, pointRadius: 3, borderWidth: 2, fill: !dash,
    ...(dash ? { borderDash: dash } : {})
  });
  const dsb = (lbl, data, color, stack) => ({
    label: lbl, data, backgroundColor: color, borderRadius: 4, stack
  });

  line('c-csat', [
    ds('Conduet', get('conduet', 'total', 'csat'), '#5b2c8d'),
    ds('C-Tenured', get('conduet', 'tenured', 'csat'), '#8e44ad', [4, 3]),
    ds('C-NH', get('conduet', 'new_hires', 'csat'), '#e07b00', [2, 3]),
    ds('HRD', get('hrd', 'total', 'csat'), '#c0392b', [6, 3]),
  ]);
  line('c-mcsat', [
    ds('Conduet', get('conduet', 'total', 'messaging_csat'), '#5b2c8d'),
    ds('C-Tenured', get('conduet', 'tenured', 'messaging_csat'), '#8e44ad', [4, 3]),
    ds('C-NH', get('conduet', 'new_hires', 'messaging_csat'), '#e07b00', [2, 3]),
    ds('HRD', get('hrd', 'total', 'messaging_csat'), '#c0392b', [6, 3]),
  ]);
  line('c-qa', [
    ds('Conduet', get('conduet', 'total', 'quality'), '#1b7f3a'),
    ds('C-Tenured', get('conduet', 'tenured', 'quality'), '#27ae60', [4, 3]),
    ds('C-NH', get('conduet', 'new_hires', 'quality'), '#e07b00', [2, 3]),
    ds('HRD', get('hrd', 'total', 'quality'), '#c0392b', [6, 3]),
  ], { min: 78, max: 95 });
  line('c-aht', [
    ds('Conduet', get('conduet', 'total', 'aht_messaging'), '#1565c0'),
    ds('HRD', get('hrd', 'total', 'aht_messaging'), '#c0392b', [6, 3]),
  ], { min: 9, max: 13 });
  line('c-sla', [
    ds('Conduet', get('conduet', 'total', 'messaging_sla'), '#0d7f7f'),
    ds('HRD', get('hrd', 'total', 'messaging_sla'), '#c0392b', [6, 3]),
  ]);
  bar('c-vol', [
    dsb('Conduet-Tenured', get('conduet', 'tenured', 'total_contacts'), 'rgba(91,44,141,.75)', 'v'),
    dsb('Conduet-NH', get('conduet', 'new_hires', 'total_contacts'), 'rgba(224,123,0,.75)', 'v'),
    dsb('HRD', get('hrd', 'total', 'total_contacts'), 'rgba(192,57,43,.75)', 'v'),
  ]);
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER: KPI Table
// ─────────────────────────────────────────────────────────────────────────────
function renderTable(tableId, periods) {
  if (!periods || !periods.length) {
    document.getElementById(tableId).innerHTML =
      '<tr><td colspan="10" style="padding:20px;text-align:center;color:#888">No data for selected period.</td></tr>';
    return;
  }

  const labels = periods.map(p => p.label);
  let h = `<thead><tr><th>KPIs</th>${labels.map(l => `<th>${l}</th>`).join('')}</tr></thead><tbody>`;

  // HC section
  h += `<tr class="r-sec"><td>TOTAL ACTIVE HC</td>${periods.map(p => {
    const ct = p.kpis?.conduet?.total?.hc || 0;
    const ht = p.kpis?.hrd?.total?.hc || 0;
    return `<td>${ct + ht}</td>`;
  }).join('')}</tr>`;
  h += `<tr class="r-c"><td>Conduet</td>${periods.map(p => `<td>${p.kpis?.conduet?.total?.hc ?? '—'}</td>`).join('')}</tr>`;
  h += `<tr class="r-t"><td>Tenured</td>${periods.map(p => `<td>${p.kpis?.conduet?.tenured?.hc ?? '—'}</td>`).join('')}</tr>`;
  h += `<tr class="r-n"><td>New Hires</td>${periods.map(p => `<td>${p.kpis?.conduet?.new_hires?.hc ?? '—'}</td>`).join('')}</tr>`;
  h += `<tr class="r-h"><td>HRD</td>${periods.map(p => `<td>${p.kpis?.hrd?.total?.hc ?? '—'}</td>`).join('')}</tr>`;

  // KPI rows
  for (const m of METRICS) {
    h += `<tr class="r-sec"><td>${m.label}</td>`;
    periods.forEach((p, i) => {
      const v    = p.kpis?.conduet?.total?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.total?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''}${m.fmt(v)}</td>`;
    });
    h += `</tr>`;

    // Conduet
    h += `<tr class="r-c"><td>Conduet</td>${periods.map(p => `<td>${m.fmt(p.kpis?.conduet?.total?.[m.key])}</td>`).join('')}</tr>`;
    // Tenured
    h += `<tr class="r-t"><td>Tenured</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.conduet?.tenured?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.tenured?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''}${m.fmt(v)}</td>`;
    });
    h += `</tr>`;
    // New Hires
    h += `<tr class="r-n"><td>New Hires</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.conduet?.new_hires?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.new_hires?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''}${m.fmt(v)}</td>`;
    });
    h += `</tr>`;
    // HRD
    h += `<tr class="r-h"><td>HRD</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.hrd?.total?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.hrd?.total?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''}${m.fmt(v)}</td>`;
    });
    h += `</tr>`;
  }

  h += `</tbody>`;
  document.getElementById(tableId).innerHTML = h;
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN: load all dashboard data
// ─────────────────────────────────────────────────────────────────────────────
let selectedYear, selectedMonth;

async function loadDashboard() {
  try {
    showStatus('Loading KPI data...');

    const year  = selectedYear  || new Date().getFullYear();
    const month = selectedMonth || null;

    // Fire API calls in parallel
    const [trendsData, weeklySummary, monthlySummary] = await Promise.all([
      apiFetch('trends',       { year, ...(month ? { month } : {}) }),
      apiFetch('team-summary', { year, ...(month ? { month } : {}), view: 'weekly' }),
      apiFetch('team-summary', { year, ...(month ? { month } : {}), view: 'monthly' }),
    ]);

    // Derive snapshot from latest weekly period (avoids needing separate /api/kpis)
    const latestPeriod  = weeklySummary.periods?.[weeklySummary.periods.length - 1];
    const prevPeriod    = weeklySummary.periods?.[weeklySummary.periods.length - 2];

    // Render
    renderSnapCards(latestPeriod?.kpis, prevPeriod?.kpis);
    renderCharts(trendsData.trends);
    renderTable('tbl-weekly',  weeklySummary.periods);
    renderTable('tbl-monthly', monthlySummary.periods);

    // Update period selector from available data
    document.getElementById('last-refresh-time').textContent = new Date().toLocaleString();
    hideStatus();

  } catch (e) {
    showError(e.message);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Period selector
// ─────────────────────────────────────────────────────────────────────────────
async function initPeriodSelector() {
  try {
    const { periods } = await apiFetch('periods');
    const yearSel  = document.getElementById('sel-year');
    const monthSel = document.getElementById('sel-month');

    // Get unique years
    const years = [...new Set(periods.map(p => p.year))].sort((a, b) => b - a);
    yearSel.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');

    const populateMonths = (year) => {
      const months = periods.filter(p => p.year === parseInt(year)).map(p => p.month);
      monthSel.innerHTML = months.map(m => `<option value="${m}">${m}</option>`).join('');
      monthSel.value = months[months.length - 1]; // default to latest month
      selectedMonth = monthSel.value;
    };

    selectedYear = years[0];
    populateMonths(selectedYear);

    yearSel.addEventListener('change', () => {
      selectedYear = parseInt(yearSel.value);
      populateMonths(selectedYear);
      loadDashboard();
    });

    monthSel.addEventListener('change', () => {
      selectedMonth = monthSel.value;
      loadDashboard();
    });

    document.getElementById('btn-load')?.addEventListener('click', loadDashboard);

  } catch (e) {
    console.error('Period selector failed:', e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Status helpers
// ─────────────────────────────────────────────────────────────────────────────
function showStatus(msg) {
  document.getElementById('status-bar').textContent = msg;
  document.getElementById('status-bar').style.display = 'block';
}
function hideStatus() {
  document.getElementById('status-bar').style.display = 'none';
}
function showError(msg) {
  const bar = document.getElementById('status-bar');
  bar.textContent = '❌ ' + msg;
  bar.style.background = '#fdecea';
  bar.style.color = '#c0392b';
  bar.style.display = 'block';
}

// ─────────────────────────────────────────────────────────────────────────────
// Tab switching
// ─────────────────────────────────────────────────────────────────────────────
window.showTab = function (id, btn) {
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
  document.getElementById('pane-' + id).classList.add('active');
  btn.classList.add('active');
};

// ─────────────────────────────────────────────────────────────────────────────
// Bootstrap
// Called by auth.js after successful sign-in (or if already signed in)
// ─────────────────────────────────────────────────────────────────────────────
async function onSignedIn(account) {
  // Show API mode badge
  const badge = document.getElementById('api-badge');
  if (badge) {
    badge.textContent = USE_STATIC ? '📁 Static Data' : '🟢 Live API';
    badge.title = USE_STATIC
      ? 'Using pre-built JSON (deploy backend for live data)'
      : 'Connected to live backend → SharePoint';
  }
  await initPeriodSelector();
  await loadDashboard();
}

// Static mode: show login screen, then load static data after sign-in click
// The login is cosmetic in static mode (no real backend auth needed)
// but provides consistent UX and prepares for live backend later
if (USE_STATIC) {
  // Override signIn for static mode — just show the dashboard
  window.signIn = async function () {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById('page-app').classList.add('active');
    const badge = document.getElementById('api-badge');
    if (badge) { badge.textContent = '📁 Static Data'; badge.title = 'Using pre-built data — no backend yet'; }
    await initPeriodSelector();
    await loadDashboard();
  };

  document.addEventListener('DOMContentLoaded', async () => {
    // Stay on login screen — wait for user to click Sign In
    // (login screen is already the default active page in index.html)
  });
}
