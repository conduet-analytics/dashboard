/**
 * calculator.js (browser version)
 * --------------------------------
 * Pure KPI calculation functions.
 * NO I/O, NO network — only math.
 * All formulas match the official KPI Formula Sheet.
 */

'use strict';

const _sum = (rows, field) =>
  rows.reduce((s, r) => s + (Number(r[field]) || 0), 0);

const _headcount = (rows) =>
  new Set(rows.map(r => r.agent_name)).size;

// ── Individual KPI formulas ──────────────────────────────────────────────────

const _csat = (rows) => {
  const num = rows.reduce((s, r) => s + (Number(r.overall_csat_score) || 0) * (Number(r.overall_survey_count) || 0), 0);
  const den = _sum(rows, 'overall_survey_count');
  return den ? +(num / den).toFixed(2) : null;
};

const _messagingCsat = (rows) => {
  const num = rows.reduce((s, r) => s + (Number(r.messaging_csat_score) || 0) * (Number(r.messaging_survey_count) || 0), 0);
  const den = _sum(rows, 'messaging_survey_count');
  return den ? +(num / den).toFixed(2) : null;
};

const _emailCsat = (rows) => {
  const num = rows.reduce((s, r) => s + (Number(r.email_csat_score) || 0) * (Number(r.email_survey_count) || 0), 0);
  const den = _sum(rows, 'email_survey_count');
  return den ? +(num / den).toFixed(2) : null;
};

const _messagingAHT = (rows) => {
  const handled = _sum(rows, 'messaging_handled');
  const totalSec = _sum(rows, 'messaging_aht_sec');
  return handled ? +(totalSec / handled / 60).toFixed(2) : null;
};

const _emailAHT = (rows) => {
  const contacts = _sum(rows, 'email_contacts_handle_time');
  const totalSec = _sum(rows, 'email_aht_sec');
  return contacts ? +(totalSec / contacts / 60).toFixed(2) : null;
};

const _quality = (rows) => {
  const score = _sum(rows, 'qa_score');
  const reviews = _sum(rows, 'qa_reviews');
  return reviews ? +((score / reviews)).toFixed(2) : null;
};

const _cph = (rows) => {
  const contacts = _sum(rows, 'messaging_handled') + _sum(rows, 'email_handled');
  const hours = _sum(rows, 'productive_hours');
  return hours ? +(contacts / hours).toFixed(2) : null;
};

const _messagingSLA = (rows) => {
  const within = _sum(rows, 'messaging_within_sla');
  const handled = _sum(rows, 'messaging_handled');
  return handled ? +(within / handled * 100).toFixed(2) : null;
};

const _emailSLA = (rows) => {
  const within = _sum(rows, 'email_sla_12hrs');
  const handled = _sum(rows, 'email_handled');
  return handled ? +(within / handled * 100).toFixed(2) : null;
};

const _attendance = (rows) => {
  const login = _sum(rows, 'login_hours');
  const scheduled = _sum(rows, 'scheduled_hours');
  return scheduled ? +(login / scheduled * 100).toFixed(2) : null;
};

const _adherence = (rows) => {
  const rostered = _sum(rows, 'rostered_less_leave');
  const ooa = _sum(rows, 'ooa');
  return rostered ? +((rostered - ooa) / rostered * 100).toFixed(2) : null;
};

const _totalContacts = (rows) =>
  _sum(rows, 'messaging_handled') + _sum(rows, 'email_handled');

// ── Aggregate: compute all KPIs for a set of rows ───────────────────────────

function computeKPIs(rows) {
  if (!rows || !rows.length) return null;
  return {
    hc:              _headcount(rows),
    csat:            _csat(rows),
    messaging_csat:  _messagingCsat(rows),
    email_csat:      _emailCsat(rows),
    aht_messaging:   _messagingAHT(rows),
    aht_email:       _emailAHT(rows),
    quality:         _quality(rows),
    cph:             _cph(rows),
    messaging_sla:   _messagingSLA(rows),
    email_sla:       _emailSLA(rows),
    attendance:      _attendance(rows),
    adherence:       _adherence(rows),
    total_contacts:  _totalContacts(rows),
  };
}

function computeGroupedKPIs(rows) {
  const conduetRows = rows.filter(r => r.location === 'conduet');
  const hrdRows     = rows.filter(r => ['Hard Rock', 'Las Vegas'].includes(r.location));

  return {
    conduet: {
      total:     computeKPIs(conduetRows),
      tenured:   computeKPIs(conduetRows.filter(r => r.emp_status === 'Tenured')),
      new_hires: computeKPIs(conduetRows.filter(r => r.emp_status === 'New Hires')),
    },
    hrd: {
      total: computeKPIs(hrdRows),
    },
  };
}
