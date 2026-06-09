/**
 * app.js
 * ------
 * Frontend logic — renders cards, charts, and exec-level KPI tables
 * with inline SVG sparkline trends.
 */

'use strict';

const MONTH_ORDER_APP = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December'
];

// ── Formatters ───────────────────────────────────────────────────────────────
const fmt = {
  dec: v => v == null ? '—' : Number(v).toFixed(2),
  pct: v => v == null ? '—' : Number(v).toFixed(2) + '%',
  num: v => v == null ? '—' : Number(v).toLocaleString(),
};

function trend(curr, prev, higherBetter) {
  if (curr == null || prev == null) return '';
  const d = curr - prev;
  const p = Math.abs(d / (prev || 1));
  if (p < 0.005) return '<span class="fl">▬</span>';
  const good = higherBetter ? d > 0 : d < 0;
  return good ? '<span class="up">▲</span>' : '<span class="dn">▼</span>';
}

// ── SVG Sparkline generator ─────────────────────────────────────────────────
function sparkline(values, higherBetter) {
  const nums = values.filter(v => v != null && !isNaN(v));
  if (nums.length < 2) return '';

  const w = 70, h = 24, pad = 2;
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  const range = max - min || 1;

  const points = nums.map((v, i) => {
    const x = pad + (i / (nums.length - 1)) * (w - pad * 2);
    const y = pad + (1 - (v - min) / range) * (h - pad * 2);
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  });

  // Color based on trend direction
  const first = nums[0], last = nums[nums.length - 1];
  const trending = last > first ? 'up' : last < first ? 'down' : 'flat';
  let color;
  if (higherBetter === null) color = '#5b2c8d';
  else if (higherBetter) color = trending === 'up' ? '#1b7f3a' : trending === 'down' ? '#c0392b' : '#b45309';
  else color = trending === 'down' ? '#1b7f3a' : trending === 'up' ? '#c0392b' : '#b45309';

  // Dot on last point
  const lastPt = points[points.length - 1].split(',');

  return `<div class="sparkline-cell">
    <svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
      <polyline points="${points.join(' ')}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="${lastPt[0]}" cy="${lastPt[1]}" r="2.5" fill="${color}"/>
    </svg>
  </div>`;
}

// ── KPI metric definitions ───────────────────────────────────────────────────
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
// SECTION SWITCHING
// ─────────────────────────────────────────────────────────────────────────────
window.showSection = function (id, btn) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.sec-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('sec-' + id).classList.add('active');
  btn.classList.add('active');
};

// ─────────────────────────────────────────────────────────────────────────────
// RENDER: Snapshot Cards
// ─────────────────────────────────────────────────────────────────────────────
function renderSnapCards(currentKpis, prevKpis) {
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
    plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, boxWidth: 10, padding: 8 } } },
    scales: {
      x: { ticks: { font: { size: 10 } }, grid: { display: false } },
      y: { ticks: { font: { size: 10 } }, grid: { color: '#f0f0f0' } }
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
// RENDER: KPI Table (exec style with sparklines + metric spacing)
// ─────────────────────────────────────────────────────────────────────────────
function renderTable(tableId, periods) {
  if (!periods || !periods.length) {
    document.getElementById(tableId).innerHTML =
      '<tr><td colspan="10" style="padding:24px;text-align:center;color:#888">No data for selected period.</td></tr>';
    return;
  }

  const labels = periods.map(p => p.label);
  let h = `<thead><tr><th>KPIs</th>${labels.map(l => `<th>${l}</th>`).join('')}<th>Trend</th></tr></thead><tbody>`;

  // HC section
  h += `<tr class="r-sec"><td>TOTAL ACTIVE HC</td>${periods.map(p => {
    const ct = p.kpis?.conduet?.total?.hc || 0;
    const ht = p.kpis?.hrd?.total?.hc || 0;
    return `<td>${ct + ht}</td>`;
  }).join('')}<td>${sparkline(periods.map(p => (p.kpis?.conduet?.total?.hc || 0) + (p.kpis?.hrd?.total?.hc || 0)), true)}</td></tr>`;
  h += `<tr class="r-c"><td>Conduet</td>${periods.map(p => `<td>${p.kpis?.conduet?.total?.hc ?? '—'}</td>`).join('')}<td>${sparkline(periods.map(p => p.kpis?.conduet?.total?.hc), true)}</td></tr>`;
  h += `<tr class="r-t"><td>Tenured</td>${periods.map(p => `<td>${p.kpis?.conduet?.tenured?.hc ?? '—'}</td>`).join('')}<td></td></tr>`;
  h += `<tr class="r-n"><td>New Hires</td>${periods.map(p => `<td>${p.kpis?.conduet?.new_hires?.hc ?? '—'}</td>`).join('')}<td></td></tr>`;
  h += `<tr class="r-h"><td>HRD</td>${periods.map(p => `<td>${p.kpis?.hrd?.total?.hc ?? '—'}</td>`).join('')}<td>${sparkline(periods.map(p => p.kpis?.hrd?.total?.hc), true)}</td></tr>`;

  // Spacer
  h += `<tr class="r-spacer"><td colspan="${labels.length + 2}"></td></tr>`;

  for (let mi = 0; mi < METRICS.length; mi++) {
    const m = METRICS[mi];

    // Metric header row with sparkline for conduet total
    const conduetVals = periods.map(p => p.kpis?.conduet?.total?.[m.key]);
    h += `<tr class="r-sec"><td>${m.label}</td>`;
    periods.forEach((p, i) => {
      const v    = p.kpis?.conduet?.total?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.total?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''} ${m.fmt(v)}</td>`;
    });
    h += `<td>${sparkline(conduetVals, m.higher)}</td></tr>`;

    // Conduet row
    h += `<tr class="r-c"><td>Conduet</td>${periods.map(p => `<td>${m.fmt(p.kpis?.conduet?.total?.[m.key])}</td>`).join('')}<td></td></tr>`;

    // Tenured row
    h += `<tr class="r-t"><td>Tenured</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.conduet?.tenured?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.tenured?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''} ${m.fmt(v)}</td>`;
    });
    h += `<td></td></tr>`;

    // New Hires row
    h += `<tr class="r-n"><td>New Hires</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.conduet?.new_hires?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.conduet?.new_hires?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''} ${m.fmt(v)}</td>`;
    });
    h += `<td></td></tr>`;

    // HRD row
    const hrdVals = periods.map(p => p.kpis?.hrd?.total?.[m.key]);
    h += `<tr class="r-h"><td>HRD</td>`;
    periods.forEach((p, i) => {
      const v = p.kpis?.hrd?.total?.[m.key];
      const prev = i > 0 ? periods[i - 1].kpis?.hrd?.total?.[m.key] : null;
      h += `<td>${m.higher !== null ? trend(v, prev, m.higher) : ''} ${m.fmt(v)}</td>`;
    });
    h += `<td>${sparkline(hrdVals, m.higher)}</td></tr>`;

    // Spacer between metric groups
    if (mi < METRICS.length - 1) {
      h += `<tr class="r-spacer"><td colspan="${labels.length + 2}"></td></tr>`;
    }
  }

  h += `</tbody>`;
  document.getElementById(tableId).innerHTML = h;
}

// ─────────────────────────────────────────────────────────────────────────────
// DATA: Fetch from SharePoint + compute KPIs
// ─────────────────────────────────────────────────────────────────────────────

function weekLabel(we) {
  const d = new Date(we);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
}

// Always return last 6 weeks from the data (regardless of filter)
function buildWeeklySummary(rows) {
  const weeks = [...new Set(rows.map(r => r.week_end))].sort();
  const last6 = weeks.slice(-6);
  return last6.map(we => ({
    label: weekLabel(we),
    key:   we,
    kpis:  computeGroupedKPIs(rows.filter(r => r.week_end === we)),
  }));
}

// Always return last 6 months from the data (regardless of filter)
function buildMonthlySummary(rows) {
  const byMonth = {};
  for (const r of rows) {
    if (!r.cal_month) continue;
    if (!byMonth[r.cal_month]) byMonth[r.cal_month] = [];
    byMonth[r.cal_month].push(r);
  }
  const allMonths = MONTH_ORDER_APP
    .filter(m => byMonth[m] && byMonth[m].length)
    .map(m => {
      const year = byMonth[m][0]?.year;
      return {
        label: `${m.slice(0, 3)} '${String(year).slice(-2)}`,
        key:   `${year}-${m}`,
        month: m,
        kpis:  computeGroupedKPIs(byMonth[m]),
      };
    });
  return allMonths.slice(-6);
}

function buildTrends(rows) {
  const byMonth = {};
  for (const r of rows) {
    if (!r.cal_month) continue;
    if (!byMonth[r.cal_month]) byMonth[r.cal_month] = [];
    byMonth[r.cal_month].push(r);
  }
  return MONTH_ORDER_APP
    .filter(m => byMonth[m] && byMonth[m].length)
    .map(m => {
      const year = byMonth[m][0]?.year;
      return {
        label: `${m.slice(0, 3)} '${String(year).slice(-2)}`,
        year, month: m,
        kpis: computeGroupedKPIs(byMonth[m]),
      };
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN: load all dashboard data
// ─────────────────────────────────────────────────────────────────────────────
let selectedYear, selectedMonth;

let _forceNext = false;

async function loadDashboard() {
  try {
    showStatus(_forceNext ? 'Refreshing from SharePoint...' : 'Loading...');

    const year  = selectedYear  || new Date().getFullYear();
    const month = selectedMonth || null;

    if (_forceNext) { await forceRefreshData(); _forceNext = false; }

    const rows = await fetchRows({
      year,
      month: month || MONTH_ORDER_APP[new Date().getMonth()],
    });

    showStatus('Computing KPIs...');

    const weeklySummary  = buildWeeklySummary(rows);
    const monthlySummary = buildMonthlySummary(rows);
    const trends         = buildTrends(rows);

    const latestPeriod = weeklySummary[weeklySummary.length - 1];
    const prevPeriod   = weeklySummary[weeklySummary.length - 2];

    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById('page-app').classList.add('active');

    renderSnapCards(latestPeriod?.kpis, prevPeriod?.kpis);
    renderCharts(trends);
    renderTable('tbl-weekly',  weeklySummary);
    renderTable('tbl-monthly', monthlySummary);

    document.getElementById('last-refresh-time').textContent = new Date().toLocaleString();
    hideStatus();

  } catch (e) {
    showError(e.message);
    console.error('Dashboard load error:', e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Period selector
// ─────────────────────────────────────────────────────────────────────────────
async function initPeriodSelector() {
  try {
    showStatus('Loading available periods...');
    const periods = await fetchAvailablePeriods();
    const yearSel  = document.getElementById('sel-year');
    const monthSel = document.getElementById('sel-month');

    const years = [...new Set(periods.map(p => p.year))].sort((a, b) => b - a);
    yearSel.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');

    const populateMonths = (year) => {
      const months = periods.filter(p => p.year === parseInt(year)).map(p => p.month);
      monthSel.innerHTML = months.map(m => `<option value="${m}">${m}</option>`).join('');
      monthSel.value = months[months.length - 1];
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

  } catch (e) {
    console.error('Period selector failed:', e);
    showError('Failed to load periods: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Status helpers
// ─────────────────────────────────────────────────────────────────────────────
function showStatus(msg) {
  const bar = document.getElementById('status-bar');
  bar.textContent = msg;
  bar.style.background = '#e8f4fd';
  bar.style.color = '#1565c0';
  bar.style.display = 'block';
}
function hideStatus() {
  document.getElementById('status-bar').style.display = 'none';
}
function showError(msg) {
  const bar = document.getElementById('status-bar');
  bar.textContent = 'Error: ' + msg;
  bar.style.background = '#fdecea';
  bar.style.color = '#c0392b';
  bar.style.display = 'block';
}

// ─────────────────────────────────────────────────────────────────────────────
// Tab switching (weekly/monthly within KPI Tables)
// ─────────────────────────────────────────────────────────────────────────────
window.showTab = function (id, btn) {
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
  document.getElementById('pane-' + id).classList.add('active');
  btn.classList.add('active');
};

// ─────────────────────────────────────────────────────────────────────────────
// Bootstrap — called by auth.js after successful sign-in
// ─────────────────────────────────────────────────────────────────────────────
async function onSignedIn(account) {
  const badge = document.getElementById('api-badge');
  if (badge) {
    badge.textContent = '🟢 Live SharePoint';
    badge.title = 'Connected to SharePoint via Microsoft Graph';
  }
  await initPeriodSelector();
  await loadDashboard();
}
