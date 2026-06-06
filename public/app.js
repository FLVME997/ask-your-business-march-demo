let DATA;
let SOURCE_EVIDENCE = {};
let selected = 'ALL';
let currentId = null;
let REVIEW_UI = { q: '', status: 'needs_review', severity: '', area: '', scrollTop: 0, scrollLeft: 0 };
const LS_KEY = 'ayb_owner_accountant_portal_v08';
const APP_VERSION = '1.0.0';

const MONTH_ASSETS = {
  '2020-03': { original_file_name: '31 03 2020.xlsx', review_workbook: '/review-workbooks/University_March_2020_Import_Review_Workbook.xlsx', pdf_report: '/reports/University_March_2020_Import_Review_Report.pdf' },
  '2020-04': { original_file_name: '30.04.2020..xlsx', review_workbook: '/review-workbooks/University_April_2020_Import_Review_Workbook.xlsx', pdf_report: '/reports/University_April_2020_Import_Review_Report.pdf' },
  '2020-05': { original_file_name: '29.05.2020..xlsx', review_workbook: '/review-workbooks/University_May_2020_Import_Review_Workbook.xlsx', pdf_report: '/reports/University_May_2020_Import_Review_Report.pdf' }
};

let STATE = blankState();

const fmt = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const money = n => `${fmt.format(Number(n || 0))} RSD`;
const esc = x => String(x ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
const now = () => new Date().toISOString();

const STATUS_META = {
  needs_review: { label: 'Needs review', tone: 'warn' },
  in_review: { label: 'In review', tone: 'warn' },
  owner_reviewed: { label: 'Owner reviewed', tone: 'info' },
  management_ready: { label: 'Management-ready', tone: 'purple' },
  needs_owner_clarification: { label: 'Needs owner clarification', tone: 'warn' },
  needs_accountant_review: { label: 'Needs accountant review', tone: 'warn' },
  accountant_certified: { label: 'Accountant-certified', tone: 'good' },
  rejected: { label: 'Rejected / excluded', tone: 'bad' },
  escalated: { label: 'Escalated', tone: 'bad' }
};

function blankState() {
  return {
    role: 'owner',
    decisions: {},
    txOverrides: {},
    controlOverrides: {},
    manualAdjustments: {},
    mappingRules: [],
    categorySuggestions: [],
    reportGroupSuggestions: [],
    batchStatuses: {},
    clarificationTasks: {},
    auditEvents: []
  };
}
function loadState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) STATE = { ...blankState(), ...JSON.parse(raw) };
    STATE.categorySuggestions = STATE.categorySuggestions || [];
    STATE.reportGroupSuggestions = STATE.reportGroupSuggestions || [];
  } catch (e) { console.warn(e); }
}
function saveState() { localStorage.setItem(LS_KEY, JSON.stringify(STATE)); }

function cls(s) {
  s = String(s || '').toLowerCase();
  if (s.includes('accountant-certified') || s.includes('cert') || s.includes('pass') || s === 'inflow' || s.includes('ready')) return 'good';
  if (s.includes('reject') || s.includes('fail') || s.includes('high') || s.includes('escal')) return 'bad';
  if (s.includes('review') || s.includes('open') || s.includes('warn') || s.includes('medium') || s === 'outflow' || s.includes('pending') || s.includes('clarification')) return 'warn';
  if (s.includes('edited') || s.includes('map') || s.includes('management')) return 'purple';
  return 'info';
}
function pill(t, k) { return `<span class="pill ${k || cls(t)}">${esc(t)}</span>`; }
function table(h, rows) { return `<div class="table-wrap"><table><thead><tr>${h.map(x => `<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`; }
function actionBtn(label, js, klass = '') { return `<button class="btn ${klass}" onclick="${js}">${esc(label)}</button>`; }

function role() { return STATE.role || 'owner'; }
function setRole(r) {
  STATE.role = r;
  saveState();
  renderAll();
  setActiveSection(r === 'owner' ? 'owner' : 'accountant');
}
function allowedTabs() {
  if (role() === 'owner') return ['owner', 'certified', 'shared', 'assistant', 'reports'];
  return ['accountant', 'certified', 'review', 'transactions', 'mapping', 'validation', 'certification', 'shared', 'reports'];
}
function activeSection() { return document.querySelector('.section.active')?.id || (role() === 'owner' ? 'owner' : 'accountant'); }
function setActiveSection(id) {
  const allowed = allowedTabs();
  if (!allowed.includes(id)) id = role() === 'owner' ? 'owner' : 'accountant';
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x.dataset.section === id));
  document.querySelectorAll('.section').forEach(x => x.classList.toggle('active', x.id === id));
}

function month() { return DATA.months.find(m => m.period === selected); }
function label() { return selected === 'ALL' ? 'All months' : month().label; }
function scopeMonths() { return selected === 'ALL' ? DATA.months : [month()]; }
function allItems() { return DATA.months.flatMap(m => m.manual_review_queue.map(r => ({ ...r, month_label: m.label }))); }
function allTx() { return DATA.months.flatMap(m => m.transactions.map(t => ({ ...t, month_label: m.label }))); }
function itemById(id) { return allItems().find(i => i.review_item_id === id); }
function txById(id) { return allTx().find(t => t.transaction_id === id); }
function valById(id) { return DATA.months.flatMap(m => m.validation_results).find(v => v.validation_id === id); }
function scopeItems() { return selected === 'ALL' ? allItems() : allItems().filter(i => i.period === selected); }

function itemStatus(i) {
  const s = STATE.decisions[i.review_item_id]?.status || String(i.status || 'open').toLowerCase();
  return s === 'open' ? 'needs_review' : s;
}
function statusLabel(s) { return STATUS_META[s]?.label || String(s || 'Unknown'); }
function isDoneForQueue(i) { return ['management_ready', 'accountant_certified', 'rejected'].includes(itemStatus(i)); }
function blocksManagement(i) { return ['needs_review', 'in_review', 'needs_owner_clarification', 'escalated'].includes(itemStatus(i)); }
function blocksFormal(i) { return !['accountant_certified', 'rejected'].includes(itemStatus(i)); }
function openOwnerClarification(i) { return itemStatus(i) === 'needs_owner_clarification'; }
function needsAccountantWork(i) { return ['needs_review', 'in_review', 'owner_reviewed', 'needs_accountant_review', 'needs_owner_clarification', 'management_ready', 'escalated'].includes(itemStatus(i)); }

function generalIssueType(i) {
  const raw = `${i.issue_type || ''} ${i.description || ''} ${i.details || ''}`.toLowerCase();
  if (raw.includes('formula') || raw.includes('validation') || raw.includes('monthly') || raw.includes('tie-out') || raw.includes('tie_out') || String(i.related_object_id || '').startsWith('VAL')) return 'Validation / control';
  if (raw.includes('mapping') || raw.includes('new_category') || raw.includes('new category') || raw.includes('category') || raw.includes('confidence')) return 'Mapping / category';
  if (raw.includes('transaction') || raw.includes('large') || raw.includes('amount') || raw.includes('eur') || raw.includes('non_operating') || raw.includes('non-operating')) return 'Transaction review';
  if (raw.includes('sheet') || raw.includes('template') || raw.includes('naming') || raw.includes('source') || raw.includes('file')) return 'Source structure';
  return 'Other';
}
function generalOptions(v = '') { return ['Validation / control', 'Mapping / category', 'Transaction review', 'Source structure', 'Other'].map(o => `<option value="${esc(o)}" ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join(''); }


const DEFAULT_MAPPED_CATEGORIES = [
  'Tuition instalment revenue',
  'Academic renewal revenue',
  'Exam fees',
  'Library/copying revenue',
  'Cafe/bar revenue',
  'ATM rental income',
  'Cash surplus / difference',
  'Currency conversion / FX movement',
  'Salaries and wages',
  'Payroll contributions',
  'Bank fees / financial services',
  'Internet and phone',
  'Utilities / communal services',
  'Fuel / electricity / utilities',
  'Maintenance services',
  'Office supplies',
  'Cafe/bar inventory',
  'External services',
  'Legal / professional services',
  'Legal dispute / settlement',
  'Owner / related-party payment',
  'Taxes / government fees',
  'Transport costs',
  'Taxi / postage / courier',
  'University membership fee',
  'Marketing / bookkeeping / professional services',
  'Expert witness / professional services',
  'Unmapped / needs accountant review',
  'Manual control adjustment',
  'Financial result adjustment'
];
const DEFAULT_REPORT_GROUPS = [
  'Operating revenue',
  'Other operating revenue',
  'Operating expense',
  'Financial income',
  'Financial expense',
  'Financial activity',
  'Non-operating income',
  'Non-operating expense',
  'Owner / related-party movement',
  'Manual adjustment',
  'Unmapped / needs accountant review'
];
function normalizeSuggestion(v) {
  return String(v || '').trim().replace(/\s+/g, ' ');
}
function uniqueSorted(values) {
  const seen = new Set();
  return values.map(normalizeSuggestion).filter(Boolean).filter(v => {
    const key = v.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
}
function readSuggestionValue(v, type = 'category') {
  if (typeof v === 'string') return v;
  if (!v) return '';
  return type === 'report_group' ? (v.group || v.name || v.report_group || '') : (v.category || v.name || v.suggested_category || '');
}
function categorySuggestions() {
  const vals = [...DEFAULT_MAPPED_CATEGORIES];
  (DATA?.months || []).forEach(m => {
    (m.mapping_suggestions || []).forEach(x => vals.push(x.suggested_category));
    (m.transactions || []).forEach(t => vals.push(t.suggested_category));
  });
  Object.values(STATE.txOverrides || {}).forEach(t => vals.push(t.suggested_category));
  Object.values(STATE.manualAdjustments || {}).forEach(t => vals.push(t.suggested_category));
  Object.values(STATE.controlOverrides || {}).forEach(c => vals.push(c.affected_report_line));
  (STATE.mappingRules || []).forEach(r => vals.push(r.suggested_category));
  (STATE.categorySuggestions || []).forEach(v => vals.push(readSuggestionValue(v, 'category')));
  return uniqueSorted(vals);
}
function reportGroupSuggestions() {
  const vals = [...DEFAULT_REPORT_GROUPS];
  (DATA?.months || []).forEach(m => {
    (m.mapping_suggestions || []).forEach(x => vals.push(x.report_group));
    (m.transactions || []).forEach(t => vals.push(t.report_group));
  });
  Object.values(STATE.txOverrides || {}).forEach(t => vals.push(t.report_group));
  Object.values(STATE.manualAdjustments || {}).forEach(t => vals.push(t.report_group));
  (STATE.mappingRules || []).forEach(r => vals.push(r.report_group));
  (STATE.reportGroupSuggestions || []).forEach(v => vals.push(readSuggestionValue(v, 'report_group')));
  return uniqueSorted(vals);
}
function addLocalSuggestion(value, type = 'category', meta = {}) {
  const clean = normalizeSuggestion(value);
  if (!clean) return false;
  const all = type === 'report_group' ? reportGroupSuggestions() : categorySuggestions();
  if (all.some(x => x.toLowerCase() === clean.toLowerCase())) return false;
  if (type === 'report_group') {
    STATE.reportGroupSuggestions = STATE.reportGroupSuggestions || [];
    STATE.reportGroupSuggestions.push({ group: clean, created_at: now(), ...meta });
  } else {
    STATE.categorySuggestions = STATE.categorySuggestions || [];
    STATE.categorySuggestions.push({ category: clean, created_at: now(), ...meta });
  }
  audit(type === 'report_group' ? 'report_group_suggestion_added' : 'mapped_category_suggestion_added', { value: clean, ...meta });
  return true;
}
function suggestionDatalists() {
  return `<datalist id="mappedCategorySuggestions">${categorySuggestions().map(v => `<option value="${esc(v)}"></option>`).join('')}</datalist><datalist id="reportGroupSuggestions">${reportGroupSuggestions().map(v => `<option value="${esc(v)}"></option>`).join('')}</datalist>`;
}
function categoryInput(id, label, value, placeholder = 'Select existing or type a new category') {
  return `<div class="field"><label>${esc(label)}</label><input id="${esc(id)}" list="mappedCategorySuggestions" value="${esc(value || '')}" placeholder="${esc(placeholder)}"><div class="mini">Choose an existing suggestion or type a new category. New saved categories appear in this dropdown for future review items.</div></div>`;
}
function reportGroupInput(id, label, value, placeholder = 'Select existing or type a new report group') {
  return `<div class="field"><label>${esc(label)}</label><input id="${esc(id)}" list="reportGroupSuggestions" value="${esc(value || '')}" placeholder="${esc(placeholder)}"><div class="mini">Report group suggestions are learned from previous mappings and manual entries.</div></div>`;
}
function categorySuggestionRows() {
  const localSet = new Set((STATE.categorySuggestions || []).map(x => normalizeSuggestion(readSuggestionValue(x, 'category')).toLowerCase()));
  return categorySuggestions().map(c => ({ category: c, source: localSet.has(c.toLowerCase()) ? 'Added locally during review' : 'Template/import/default suggestion' }));
}

function stats(period = selected) {
  const items = period === 'ALL' ? allItems() : allItems().filter(i => i.period === period);
  const out = { total: items.length, management_blocking: 0, formal_blocking: 0, owner_clarifications: 0, accountant_pending: 0 };
  Object.keys(STATUS_META).forEach(k => out[k] = 0);
  items.forEach(i => {
    const s = itemStatus(i);
    out[s] = (out[s] || 0) + 1;
    if (blocksManagement(i)) out.management_blocking += 1;
    if (blocksFormal(i)) out.formal_blocking += 1;
    if (openOwnerClarification(i)) out.owner_clarifications += 1;
    if (needsAccountantWork(i)) out.accountant_pending += 1;
  });
  out.resolved_for_management = out.total - out.management_blocking;
  out.resolved_for_formal = out.total - out.formal_blocking;
  return out;
}
function baseQuality(period = selected) {
  if (period === 'ALL') return Number(DATA.portfolio_summary.average_data_quality_score || 0);
  return Number(DATA.months.find(m => m.period === period).summary.data_quality_score || 0);
}
function localQuality(period = selected) {
  const st = stats(period), base = baseQuality(period);
  return st.total ? Math.min(100, base + (100 - base) * (st.resolved_for_management / st.total)) : base;
}
function managementStatus(period = selected) {
  const st = stats(period), bs = batchStatus(period);
  if (period !== 'ALL' && bs.accountant_certified) return { label: 'Accountant-certified', tone: 'good', detail: 'Ready for formal reports and owner dashboards.' };
  if (period !== 'ALL' && bs.management_ready) return { label: 'Management-ready', tone: 'purple', detail: 'Usable for owner management dashboards with accountant-certification limitation.' };
  if (st.management_blocking === 0) return { label: 'Ready for management review', tone: 'purple', detail: 'No management blockers remain, but period has not been marked management-ready yet.' };
  return { label: 'Needs review before management use', tone: 'warn', detail: `${st.management_blocking} item(s) still block management-ready status.` };
}
function formalStatus(period = selected) {
  const st = stats(period), bs = batchStatus(period);
  if (period !== 'ALL' && bs.accountant_certified) return { label: 'Accountant-certified', tone: 'good', detail: 'Formal accountant certification recorded.' };
  if (st.formal_blocking === 0) return { label: 'Ready for accountant certification', tone: 'purple', detail: 'All non-rejected items are accountant-certified; batch certification can be recorded.' };
  return { label: 'Accountant certification pending', tone: 'warn', detail: `${st.formal_blocking} item(s) still need accountant-level certification or exclusion.` };
}
function batchStatus(period) { return STATE.batchStatuses[period] || {}; }

function mappingFor(t) {
  return [...(STATE.mappingRules || [])].reverse().find(r => (!r.period || r.period === 'ALL' || r.period === t.period) && r.direction === t.direction && String(r.raw_category) === String(t.raw_category));
}
function effTx(t) {
  let out = { ...t };
  const rule = mappingFor(out);
  if (rule && !STATE.txOverrides[out.transaction_id]?.suggested_category) {
    out.suggested_category = rule.suggested_category;
    out.report_group = rule.report_group || out.report_group;
    out.mapping_rule_applied = rule.mapping_rule_id;
  }
  if (STATE.txOverrides[out.transaction_id]) out = { ...out, ...STATE.txOverrides[out.transaction_id] };
  return out;
}
function manualAdjustmentRows(period = selected) {
  return Object.values(STATE.manualAdjustments || {}).filter(a => period === 'ALL' || a.period === period).filter(a => {
    const d = STATE.decisions[a.review_item_id];
    return d && !['rejected', 'escalated'].includes(d.status) && !a.excluded;
  });
}
function displayTxRows(period = selected) {
  const months = period === 'ALL' ? DATA.months : [DATA.months.find(m => m.period === period)];
  return [...months.flatMap(m => m.transactions.map(effTx)), ...manualAdjustmentRows(period)];
}
function adjSummary(period = selected) {
  const months = period === 'ALL' ? DATA.months : [DATA.months.find(m => m.period === period)];
  const tx = displayTxRows(period).filter(t => !t.excluded);
  const inflows = tx.filter(t => t.direction === 'inflow').reduce((a, t) => a + Number(t.amount_rsd_equivalent || 0), 0);
  const outflows = tx.filter(t => t.direction === 'outflow').reduce((a, t) => a + Number(t.amount_rsd_equivalent || 0), 0);
  const opening = period === 'ALL' ? DATA.portfolio_summary.opening_position : months[0].summary.opening_position;
  return { opening, inflows, outflows, net: inflows - outflows, closing: Number(opening || 0) + inflows - outflows, rows: tx.length };
}

function linkedReviewItemsForTx(txId) {
  return allItems().filter(i => String(i.related_object_id || '') === String(txId || ''));
}
function dataReadinessForTx(t) {
  const linked = linkedReviewItemsForTx(t.transaction_id);
  const statuses = linked.map(itemStatus);
  const batch = batchStatus(t.period);
  if (t.excluded || statuses.includes('rejected')) return { key: 'excluded', label: 'Rejected / excluded', tone: 'bad', management: false, accounting: false, detail: 'Excluded during review.' };
  if (statuses.some(s => ['needs_review', 'in_review', 'needs_owner_clarification', 'escalated'].includes(s))) return { key: 'not_ready', label: 'Needs review', tone: 'warn', management: false, accounting: false, detail: 'Has unresolved review item(s).' };
  if (batch.accountant_certified || (linked.length && statuses.every(s => s === 'accountant_certified'))) return { key: 'accounting_certified', label: 'Accountant-certified', tone: 'good', management: true, accounting: true, detail: 'Formal accounting review recorded.' };
  if (statuses.includes('accountant_certified')) return { key: 'partly_certified', label: 'Partly accountant-certified', tone: 'purple', management: true, accounting: false, detail: 'Some linked items are accountant-certified, but period certification is not complete.' };
  if (statuses.some(s => ['management_ready', 'owner_reviewed', 'needs_accountant_review'].includes(s)) || batch.management_ready) return { key: 'management_ready', label: 'Management-ready', tone: 'purple', management: true, accounting: false, detail: 'Usable for owner/business analysis with accountant-certification limitation.' };
  if (!linked.length) return { key: 'auto_prepared', label: 'System-prepared', tone: 'info', management: true, accounting: false, detail: 'No exception was raised for this row. Accounting certification still depends on period status.' };
  return { key: 'not_ready', label: 'Needs review', tone: 'warn', management: false, accounting: false, detail: 'Review status derived from linked issue(s).' };
}
function certifiedRows(period = selected) {
  return displayTxRows(period).map(t => {
    const e = effTx(t);
    return { ...e, readiness: dataReadinessForTx(e), linked_review_items: linkedReviewItemsForTx(e.transaction_id) };
  });
}
function sourceMonthsFor(period = selected) { return period === 'ALL' ? DATA.months : [monthByPeriod(period)]; }
function certifiedSummary(period = selected) {
  const rows = certifiedRows(period);
  const sourceMonths = sourceMonthsFor(period).filter(Boolean);
  const original = {
    opening: period === 'ALL' ? DATA.portfolio_summary.opening_position : sourceMonths[0]?.summary?.opening_position,
    inflows: sourceMonths.reduce((a,m)=>a+Number(m.summary.total_inflows||0),0),
    outflows: sourceMonths.reduce((a,m)=>a+Number(m.summary.total_outflows||0),0),
    net: sourceMonths.reduce((a,m)=>a+Number(m.summary.net_movement||0),0),
    closing: period === 'ALL' ? DATA.portfolio_summary.closing_position : sourceMonths[0]?.summary?.closing_position,
    rows: sourceMonths.reduce((a,m)=>a+Number(m.summary.transactions_extracted||0),0)
  };
  const managementRows = rows.filter(r => r.readiness.management && !r.excluded);
  const accountingRows = rows.filter(r => r.readiness.accounting && !r.excluded);
  const rejectedRows = rows.filter(r => r.readiness.key === 'excluded' || r.excluded);
  const pendingRows = rows.filter(r => !r.readiness.management && r.readiness.key !== 'excluded');
  const sumRows = rs => {
    const inflows = rs.filter(t => t.direction === 'inflow').reduce((a,t)=>a+Number(t.amount_rsd_equivalent||0),0);
    const outflows = rs.filter(t => t.direction === 'outflow').reduce((a,t)=>a+Number(t.amount_rsd_equivalent||0),0);
    const opening = Number(original.opening || 0);
    return { opening, inflows, outflows, net: inflows - outflows, closing: opening + inflows - outflows, rows: rs.length };
  };
  return { original, management: sumRows(managementRows), accounting: sumRows(accountingRows), rejectedRows, pendingRows, allRows: rows, managementRows, accountingRows, adjustments: manualAdjustmentRows(period) };
}
function categoryBreakdownRows(rows, limit = 12) {
  const groups = {};
  rows.filter(r => !r.excluded).forEach(t => {
    const key = t.suggested_category || 'Unmapped';
    groups[key] = groups[key] || { category: key, inflows: 0, outflows: 0, net: 0, rows: 0 };
    const amt = Number(t.amount_rsd_equivalent || 0);
    if (t.direction === 'inflow') groups[key].inflows += amt; else groups[key].outflows += amt;
    groups[key].net += t.direction === 'inflow' ? amt : -amt;
    groups[key].rows += 1;
  });
  return Object.values(groups).sort((a,b)=>Math.abs(b.net)-Math.abs(a.net)).slice(0, limit);
}
function beforeAfterRows(sum) {
  const diff = (a,b) => Number(b||0) - Number(a||0);
  return [
    ['Rows', sum.original.rows, sum.management.rows, diff(sum.original.rows, sum.management.rows)],
    ['Inflows', money(sum.original.inflows), money(sum.management.inflows), money(diff(sum.original.inflows, sum.management.inflows))],
    ['Outflows', money(sum.original.outflows), money(sum.management.outflows), money(diff(sum.original.outflows, sum.management.outflows))],
    ['Net movement', money(sum.original.net), money(sum.management.net), money(diff(sum.original.net, sum.management.net))],
    ['Closing cash', money(sum.original.closing), money(sum.management.closing), money(diff(sum.original.closing, sum.management.closing))]
  ].map(r => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td><td>${esc(r[2])}</td><td>${esc(r[3])}</td></tr>`);
}
function certificationLimitations(period = selected) {
  const st = stats(period), mg = managementStatus(period), fm = formalStatus(period), sum = certifiedSummary(period);
  const limits = [];
  if (st.management_blocking > 0) limits.push(`${st.management_blocking} item(s) still block management-ready status.`);
  if (st.formal_blocking > 0) limits.push(`${st.formal_blocking} item(s) still need accountant-level certification or exclusion.`);
  if (sum.pendingRows.length) limits.push(`${sum.pendingRows.length} transaction/control-linked row(s) remain outside management analysis.`);
  if (sum.rejectedRows.length) limits.push(`${sum.rejectedRows.length} row(s) have been rejected/excluded from the certified view.`);
  if (sum.adjustments.length) limits.push(`${sum.adjustments.length} manual adjustment(s) were created during review.`);
  if (!limits.length && fm.tone === 'good') limits.push('No open review limitations for the selected scope based on the current browser-local state.');
  else if (!limits.length) limits.push('No management blockers, but formal accountant period certification may still be pending.');
  return { mg, fm, limits };
}

function taskFor(id) { return STATE.clarificationTasks[id] || null; }
function openTasks(period = selected) {
  return Object.values(STATE.clarificationTasks || {}).filter(t => (period === 'ALL' || t.period === period) && t.status !== 'closed');
}
function audit(action, payload = {}) { STATE.auditEvents.push({ at: now(), action, ...payload }); }


function simpleMonthStatus(m) {
  const mg = managementStatus(m.period);
  const fm = formalStatus(m.period);
  const ts = openTasks(m.period).length;
  let owner = 'No owner action';
  let tone = 'good';
  if (ts > 0) { owner = `${ts} answer needed`; tone = 'warn'; }
  else if (mg.tone === 'warn') { owner = 'Accountant team working'; tone = 'info'; }
  return { mg, fm, owner, tone };
}
function monthTrendSentence() {
  const rows = DATA.months.map(m => ({ label: m.label, net: adjSummary(m.period).net, closing: adjSummary(m.period).closing }));
  const best = [...rows].sort((a,b)=>b.net-a.net)[0];
  const worst = [...rows].sort((a,b)=>a.net-b.net)[0];
  const last = rows[rows.length - 1];
  return `Across the imported months, ${best.label} had the strongest cash movement (${money(best.net)}), while ${worst.label} had the weakest (${money(worst.net)}). Latest closing cash shown is ${money(last.closing)}.`;
}
function groupOpenIssues(period = selected) {
  const out = {};
  (period === 'ALL' ? allItems() : allItems().filter(i => i.period === period)).forEach(i => {
    if (!needsAccountantWork(i)) return;
    const g = generalIssueType(i);
    out[g] = out[g] || { group: g, count: 0, high: 0, medium: 0, low: 0, value: 0 };
    out[g].count += 1;
    const sev = String(i.severity || '').toLowerCase();
    if (sev.includes('high')) out[g].high += 1;
    else if (sev.includes('medium')) out[g].medium += 1;
    else out[g].low += 1;
    const tx = txById(i.related_object_id);
    if (tx) out[g].value += Number(effTx(tx).amount_rsd_equivalent || 0);
  });
  return Object.values(out).sort((a,b)=>b.high-a.high || b.count-a.count);
}
function topPriorityItems(limit = 8, period = selected) {
  const rank = i => {
    const sev = String(i.severity || '').toLowerCase();
    const tx = txById(i.related_object_id);
    return (sev.includes('high') ? 1000000000 : sev.includes('medium') ? 500000000 : 0) + Math.abs(Number(tx?.amount_rsd_equivalent || 0));
  };
  return (period === 'ALL' ? allItems() : allItems().filter(i => i.period === period))
    .filter(needsAccountantWork)
    .sort((a,b)=>rank(b)-rank(a))
    .slice(0, limit);
}
function templateAutomationRows() {
  return DATA.months.map(m => {
    const s = m.summary;
    const reuse = Number(s.template_reuse_rate_percent || 0);
    let msg = reuse >= 95 ? 'High automation - review exceptions only' : reuse > 0 ? 'Template reused - review new patterns' : 'Initial setup month - build template';
    return { month: m.label, reuse, catReuse: Number(s.category_reuse_rate_percent || 0), extracted: s.transactions_extracted, review: s.manual_review_items, quality: s.data_quality_score, msg };
  });
}
function canBulkManagementReady(i) {
  if (!needsAccountantWork(i)) return false;
  if (String(i.severity || '').toLowerCase().includes('high')) return false;
  if (generalIssueType(i) === 'Validation / control') return false;
  if (itemStatus(i) === 'needs_owner_clarification') return false;
  return true;
}
function bulkMarkLowRiskManagementReady() {
  if (role() !== 'accountant') return alert('Switch to Accountant Workbench first.');
  const candidates = scopeItems().filter(canBulkManagementReady);
  if (!candidates.length) return alert('No low-risk non-control review items match the current month scope.');
  if (!confirm(`Mark ${candidates.length} low-risk exception item(s) as management-ready with accountant certification still pending?`)) return;
  candidates.forEach(i => {
    STATE.decisions[i.review_item_id] = {
      review_item_id: i.review_item_id,
      period: i.period,
      status: 'management_ready',
      resolution_type: 'Bulk management-ready exception clearance',
      role: role(),
      review_area: generalIssueType(i),
      reviewer: 'Accountant Operations',
      note: 'Bulk-cleared as low-risk for management analysis. Accountant certification still pending.',
      decided_at: now(),
      before: { review_item: i, transaction: txById(i.related_object_id) || null },
      after: { review_item_status: 'management_ready' },
      source_reference: i.source_reference,
      related_object_id: i.related_object_id,
      management_analysis_status: 'usable_with_limitations',
      accountant_review_status: 'pending_or_reviewed',
      formal_reporting_status: 'not_certified'
    };
    audit('bulk_management_ready_item', { review_item_id: i.review_item_id, period: i.period });
  });
  saveState(); renderAll(); setActiveSection('accountant');
  alert(`${candidates.length} low-risk item(s) marked management-ready.`);
}
function sendVisibleOwnerQuestions() {
  if (role() !== 'accountant') return alert('Switch to Accountant Workbench first.');
  const candidates = scopeItems().filter(i => needsAccountantWork(i) && !taskFor(i.review_item_id) && (generalIssueType(i) === 'Transaction review' || String(i.suggested_action || '').toLowerCase().includes('confirm'))).slice(0, 10);
  if (!candidates.length) return alert('No obvious owner-context questions found in this scope.');
  if (!confirm(`Create owner clarification tasks for ${candidates.length} item(s)?`)) return;
  candidates.forEach(i => {
    STATE.clarificationTasks[i.review_item_id] = { task_id: `TASK-${i.review_item_id}`, period: i.period, review_item_id: i.review_item_id, status: 'open', question: i.suggested_action || i.description, asked_by: 'Accountant Operations', asked_at: now(), owner_answer: '' };
    STATE.decisions[i.review_item_id] = {
      review_item_id: i.review_item_id,
      period: i.period,
      status: 'needs_owner_clarification',
      resolution_type: 'Owner clarification requested',
      role: role(),
      review_area: generalIssueType(i),
      reviewer: 'Accountant Operations',
      note: 'Sent to owner for business-context clarification.',
      owner_question: i.suggested_action || i.description,
      decided_at: now(),
      before: { review_item: i, transaction: txById(i.related_object_id) || null },
      after: { review_item_status: 'needs_owner_clarification' },
      source_reference: i.source_reference,
      related_object_id: i.related_object_id,
      management_analysis_status: 'not_ready',
      accountant_review_status: 'needs_owner_clarification',
      formal_reporting_status: 'not_certified'
    };
    audit('bulk_owner_question_created', { review_item_id: i.review_item_id, period: i.period });
  });
  saveState(); renderAll(); setActiveSection('accountant');
  alert(`${candidates.length} owner clarification task(s) created.`);
}

function renderAll() {
  document.getElementById('badge').textContent = `${DATA.company_code} • v${APP_VERSION}`;
  renderRoleBar();
  renderTabs();
  renderOwnerPortal();
  renderAccountantWorkbench();
  renderShared();
  renderCertifiedData();
  renderReview();
  renderTx();
  renderMapping();
  renderValidation();
  renderCert();
  renderAssistant();
  renderReports();
  if (!allowedTabs().includes(activeSection())) setActiveSection(role() === 'owner' ? 'owner' : 'accountant');
}
function renderRoleBar() {
  document.getElementById('roleLabel').textContent = role() === 'owner' ? 'Owner Portal' : 'Accountant Workbench';
  document.getElementById('roleExplain').textContent = role() === 'owner'
    ? 'Owner sees clean business visibility, tasks to answer, management-ready warnings, and assistant answers with limitations.'
    : 'Accountant team handles intake alignment, mapping, validation, owner questions, and certification controls.';
  document.querySelectorAll('.rolebtn').forEach(b => {
    b.classList.toggle('active', b.dataset.role === role());
    b.onclick = () => setRole(b.dataset.role);
  });
}
function renderTabs() {
  const allowed = allowedTabs();
  document.querySelectorAll('.tab').forEach(b => {
    b.style.display = allowed.includes(b.dataset.section) ? '' : 'none';
  });
}

function kpi(label, value, sub, klass = '') {
  return `<div class="card"><div class="kpi-label">${esc(label)}</div><div class="kpi-value ${klass}">${value}</div><div class="kpi-sub">${esc(sub || '')}</div></div>`;
}
function statusBox(title, status, body = '') {
  return `<div class="status-box ${status.tone}"><div class="mini">${esc(title)}</div><strong>${esc(status.label)}</strong><div class="muted">${esc(body || status.detail || '')}</div></div>`;
}

function renderOwnerPortal() {
  const a = adjSummary();
  const totalTasks = openTasks().length;
  const mgAll = managementStatus();
  const fmAll = formalStatus();
  const currentCompany = 'University Demo';
  const cards = [
    kpi('Current cash shown', money(a.closing), 'Latest imported management view', a.closing >= 0 ? 'pos' : 'neg'),
    kpi('3-month cash movement', money(a.net), monthTrendSentence(), a.net >= 0 ? 'pos' : 'neg'),
    kpi('Questions for you', totalTasks, 'Only owner-context questions are shown here', totalTasks ? 'warnText' : 'pos'),
    kpi('Accountant status', fmAll.label, 'Formal certification handled by accountant team', fmAll.tone === 'good' ? 'pos' : 'warnText')
  ].join('');
  const ownerTasks = openTasks().slice(0, 6).map(t => {
    const i = itemById(t.review_item_id);
    return `<div class="soft-item clickable-row" data-open-item="${esc(t.review_item_id)}"><strong>${esc(t.period)} • ${esc(i?.source_reference || t.review_item_id)}</strong><div class="muted">${esc(t.question || i?.suggested_action || i?.description || '')}</div>${t.owner_answer ? `<div class="mini">Current answer: ${esc(t.owner_answer)}</div>` : ''}</div>`;
  }).join('');
  const monthRows = DATA.months.map(m => {
    const st = simpleMonthStatus(m), aa = adjSummary(m.period);
    return `<tr><td>${esc(m.label)}</td><td>${pill(st.mg.label, st.mg.tone)}</td><td>${pill(st.fm.label, st.fm.tone)}</td><td>${pill(st.owner, st.tone)}</td><td>${money(aa.net)}</td><td>${money(aa.closing)}</td></tr>`;
  });
  const headline = totalTasks
    ? `<div class="owner-action-banner"><strong>${totalTasks} question(s) need your answer.</strong><div>Your accountant team only sends questions when they need business context. Answering these helps them finish certification faster.</div><button class="btn primary" onclick="setActiveSection('shared')">Answer questions</button></div>`
    : `<div class="ready-banner"><strong>No owner action needed right now.</strong><div>Your accountant team can continue working in the background. You can use management-ready information with the limitations shown.</div></div>`;
  document.getElementById('owner').innerHTML = `
    <div class="owner-hero card">
      <div>
        <div class="portal-tag">Owner portal • simplified business view</div>
        <h2>${esc(currentCompany)}</h2>
        <p class="muted">This view hides accounting complexity. It shows only what you need to know: business status, cash movement, questions for you, and whether accountant certification is still pending.</p>
      </div>
      <div class="owner-status-stack">
        ${pill(mgAll.label, mgAll.tone)} ${pill(fmAll.label, fmAll.tone)} ${pill(totalTasks ? 'Owner action needed' : 'No owner action', totalTasks ? 'warn' : 'good')}
      </div>
    </div>
    ${headline}
    <div class="grid cards">${cards}</div>
    <div class="grid two">
      <div class="card"><h2>Plain-English status</h2>
        <div class="status-box ${mgAll.tone}"><strong>Business view</strong><div>${esc(mgAll.detail)}</div></div>
        <div class="status-box ${fmAll.tone}"><strong>Accounting view</strong><div>${esc(fmAll.detail)}</div></div>
        <div class="notice"><strong>Important:</strong> management-ready data is useful for business decisions and visibility. Accountant-certified data is needed for formal accounting, tax, or statutory reporting.</div>
      </div>
      <div class="card"><h2>What you need to do</h2>${ownerTasks || '<div class="notice oknotice">Nothing right now. Your accountant team is handling the review work.</div>'}</div>
    </div>
    <div class="card"><h2>Company timeline</h2><p class="muted">Simple monthly view. Detailed validation and mapping tasks stay in the accountant workbench.</p>${table(['Month','Business view','Accounting status','Owner action','Cash movement','Closing cash'], monthRows)}</div>
    <div class="card"><h2>Business summary</h2><div class="answer">${esc(monthTrendSentence())}\n\nThe owner portal will eventually show all companies here. For now, this demo uses one real company dataset while the accountant workbench proves the data-alignment workflow.</div></div>
  `;
  document.querySelectorAll('[data-open-item]').forEach(r => r.addEventListener('click', () => openModal(r.dataset.openItem)));
}

function renderAccountantWorkbench() {
  const st = stats();
  const a = adjSummary();
  const tasks = openTasks();
  const pending = scopeItems().filter(needsAccountantWork);
  const issueGroups = groupOpenIssues();
  const priority = topPriorityItems(8);
  const automationRows = templateAutomationRows().map(r => `<tr><td>${esc(r.month)}</td><td>${r.reuse.toFixed(1)}%</td><td>${r.catReuse.toFixed(1)}%</td><td>${r.extracted}</td><td>${r.review}</td><td>${Number(r.quality).toFixed(1)}/100</td><td>${esc(r.msg)}</td></tr>`);
  const groupRows = issueGroups.map(g => `<tr><td>${pill(g.group, 'purple')}</td><td>${g.count}</td><td>${g.high}</td><td>${g.medium}</td><td>${g.low}</td><td>${money(g.value)}</td><td><button class="btn" onclick="REVIEW_UI.area='${esc(g.group)}'; REVIEW_UI.status=''; setActiveSection('review'); renderReview();">Open group</button></td></tr>`);
  const priorityRows = priority.map(i => {
    const tx = txById(i.related_object_id), et = tx ? effTx(tx) : null;
    return `<tr class="clickable-row" data-open-item="${esc(i.review_item_id)}"><td>${esc(i.period)}</td><td>${pill(i.severity)}</td><td>${pill(generalIssueType(i),'purple')}</td><td>${pill(statusLabel(itemStatus(i)), STATUS_META[itemStatus(i)]?.tone)}</td><td>${et ? money(et.amount_rsd_equivalent) : ''}</td><td>${esc(i.description)}</td><td>${esc(i.source_reference || '')}</td></tr>`;
  });
  const monthRows = DATA.months.map(m => {
    const s = stats(m.period), mg = managementStatus(m.period), fm = formalStatus(m.period), bs = batchStatus(m.period);
    return `<tr><td>${esc(m.label)}</td><td>${s.total}</td><td>${s.management_blocking}</td><td>${s.formal_blocking}</td><td>${openTasks(m.period).length}</td><td>${pill(mg.label, mg.tone)}</td><td>${pill(fm.label, fm.tone)}</td><td>${bs.management_ready ? esc(bs.management_ready_at || '') : ''}</td></tr>`;
  });
  const cards = [
    kpi('Open accountant items', pending.length, 'Exception-only queue, not row-by-row review', pending.length ? 'warnText' : 'pos'),
    kpi('Owner questions open', tasks.length, 'Business-context items sent to owner', tasks.length ? 'warnText' : 'pos'),
    kpi('Low-risk bulk candidates', scopeItems().filter(canBulkManagementReady).length, 'Can be moved to management-ready in bulk', 'purpleText'),
    kpi('Adjusted closing cash', money(a.closing), 'After local review edits and adjustments', a.closing >= 0 ? 'pos' : 'neg')
  ].join('');
  document.getElementById('accountant').innerHTML = `
    <div class="accountant-hero card">
      <div>
        <div class="portal-tag">Accountant operations center • one-company workflow first</div>
        <h2>Make data ready with exception-only review</h2>
        <p class="muted">This side is designed for your accountant team: apply templates, review only exceptions, ask the owner only for business context, mark management-ready, and later accountant-certify the period.</p>
      </div>
      <div class="portal-actions">
        <button class="btn primary" onclick="setActiveSection('review')">Open exception queue</button>
        <button class="btn purpleBtn" onclick="bulkMarkLowRiskManagementReady()">Bulk management-ready low-risk</button>
        <button class="btn warn" onclick="sendVisibleOwnerQuestions()">Create owner questions</button>
        <button class="btn good" onclick="exportAccountantPack()">Export accountant pack</button>
      </div>
    </div>
    <div class="grid cards">${cards}</div>
    <div class="grid two">
      <div class="card"><h2>Template automation for similar files</h2><p class="muted">March created the initial template. April and May show the automation effect when similar monthly workbooks are imported.</p>${table(['Month','Template reuse','Category reuse','Rows','Review items','Quality','Automation note'], automationRows)}</div>
      <div class="card"><h2>Recommended workflow</h2><div class="timeline">
        <div class="step"><div class="num">1</div><div><strong>Apply company template</strong><div class="muted">Known sheets, categories and validation checks are reused.</div></div></div>
        <div class="step"><div class="num">2</div><div><strong>Review exceptions only</strong><div class="muted">Low-risk items can move to management-ready; high-risk/control issues stay for accountant review.</div></div></div>
        <div class="step"><div class="num">3</div><div><strong>Ask owner only when needed</strong><div class="muted">Owner only receives simple business-context questions.</div></div></div>
        <div class="step"><div class="num">4</div><div><strong>Certify period</strong><div class="muted">Management-ready first, accountant-certified after formal review.</div></div></div>
      </div></div>
    </div>
    <div class="card"><h2>Exception groups</h2>${groupRows.length ? table(['Group','Open items','High','Medium','Low/info','Linked value','Action'], groupRows) : '<div class="notice oknotice">No open exception groups in this scope.</div>'}</div>
    <div class="card"><h2>Priority queue</h2>${priorityRows.length ? table(['Period','Severity','Area','Status','Amount','Issue','Source'], priorityRows) : '<div class="notice oknotice">No priority items remain.</div>'}</div>
    <div class="card"><h2>Month readiness board</h2>${table(['Month','Queue','Mgmt blockers','Formal blockers','Owner tasks','Management','Accounting','Mgmt-ready at'], monthRows)}</div>
  `;
  document.querySelectorAll('[data-open-item]').forEach(r => r.addEventListener('click', () => openModal(r.dataset.openItem)));
}

function renderShared() {
  const tasks = openTasks();
  const rows = tasks.map(t => {
    const i = itemById(t.review_item_id);
    return `<tr class="clickable-row" data-open-item="${esc(t.review_item_id)}"><td>${esc(t.period)}</td><td>${esc(t.review_item_id)}</td><td>${pill(t.status)}</td><td>${esc(t.question || '')}</td><td>${esc(t.owner_answer || '')}</td><td>${esc(t.asked_at || '')}</td><td>${esc(i?.source_reference || '')}</td></tr>`;
  });
  const decisionRows = Object.values(STATE.decisions).slice(-15).reverse().map(d => `<tr><td>${esc(d.period)}</td><td>${esc(d.review_item_id)}</td><td>${pill(statusLabel(d.status))}</td><td>${esc(d.reviewer)}</td><td>${esc(d.note || '')}</td><td>${esc(d.decided_at)}</td></tr>`);
  document.getElementById('shared').innerHTML = `
    <div class="card"><h2>Shared owner-accountant workspace</h2><p class="muted">This is where accountant questions and owner answers stay attached to the same review item, so both sides remain aligned.</p>${rows.length ? table(['Period','Review item','Status','Question to owner','Owner answer','Asked at','Source'], rows) : '<div class="notice oknotice">No open owner-accountant tasks.</div>'}</div>
    <div class="card"><h2>Recent review decisions</h2>${decisionRows.length ? table(['Period','Review item','Decision','Reviewer / actor','Note','Timestamp'], decisionRows) : '<div class="notice">No local decisions yet.</div>'}</div>
  `;
  document.querySelectorAll('[data-open-item]').forEach(r => r.addEventListener('click', () => openModal(r.dataset.openItem)));
}


function renderCertifiedData() {
  const sum = certifiedSummary();
  const lim = certificationLimitations();
  const scopeLabel = label();
  const catRows = categoryBreakdownRows(sum.managementRows).map(g => `<tr><td>${esc(g.category)}</td><td>${g.rows}</td><td>${money(g.inflows)}</td><td>${money(g.outflows)}</td><td>${money(g.net)}</td></tr>`);
  const adjRows = sum.adjustments.map(a => `<tr><td>${esc(a.period)}</td><td>${esc(a.transaction_id)}</td><td>${esc(a.date || '')}</td><td>${pill(a.direction || '')}</td><td>${money(a.amount_rsd_equivalent)}</td><td>${esc(a.suggested_category || '')}</td><td>${esc(a.description || '')}</td><td>${esc(a.review_item_id || '')}</td></tr>`);
  const rejRows = sum.rejectedRows.slice(0, 100).map(t => `<tr><td>${esc(t.period)}</td><td>${esc(t.transaction_id)}</td><td>${esc(t.date || '')}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category || '')}</td><td>${esc(t.suggested_category || '')}</td><td>${esc(t.description || '')}</td><td>${t.linked_review_items.map(x=>esc(x.review_item_id)).join('<br>')}</td></tr>`);
  const pendingRows = sum.pendingRows.slice(0, 100).map(t => `<tr class="clickable-row" data-open-tx="${esc(t.linked_review_items[0]?.review_item_id || '')}"><td>${esc(t.period)}</td><td>${esc(t.transaction_id)}</td><td>${pill(t.readiness.label, t.readiness.tone)}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category || '')}</td><td>${esc(t.suggested_category || '')}</td><td>${esc(t.description || '')}</td><td>${t.linked_review_items.map(x=>esc(x.review_item_id)).join('<br>')}</td></tr>`);
  const certifiedRowsPreview = sum.managementRows.slice(0, 120).map(t => `<tr><td>${esc(t.period)}</td><td>${esc(t.transaction_id)}</td><td>${esc(t.date || '')}</td><td>${pill(t.readiness.label, t.readiness.tone)}</td><td>${pill(t.direction || '')}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category || '')}</td><td>${esc(t.suggested_category || '')}</td><td>${esc(t.report_group || '')}</td><td>${esc(t.source_reference || '')}</td></tr>`);
  document.getElementById('certified').innerHTML = `
    <div class="card certified-hero">
      <div>
        <div class="portal-tag">Certified data view • ${esc(scopeLabel)}</div>
        <h2>What data is now usable?</h2>
        <p class="muted">This tab separates the original extracted data from the current management-ready and accountant-certified layers. It uses your local review decisions, edits, manual adjustments, and rejected rows.</p>
      </div>
      <div class="owner-status-stack">${pill(lim.mg.label, lim.mg.tone)} ${pill(lim.fm.label, lim.fm.tone)}</div>
    </div>
    <div class="grid cards">
      ${kpi('Management-ready rows', sum.management.rows, 'Rows usable for business/owner analysis', sum.management.rows ? 'pos' : 'warnText')}
      ${kpi('Accounting-certified rows', sum.accounting.rows, 'Rows formally certified in current local state', sum.accounting.rows ? 'pos' : 'warnText')}
      ${kpi('Rejected / excluded rows', sum.rejectedRows.length, 'Excluded from the certified view', sum.rejectedRows.length ? 'warnText' : 'pos')}
      ${kpi('Manual adjustments', sum.adjustments.length, 'Created from review/control decisions', sum.adjustments.length ? 'purpleText' : 'pos')}
    </div>
    <div class="grid two">
      <div class="card"><h2>Before vs after totals</h2><p class="muted">Original source totals compared with the current management-ready layer after local review edits, exclusions, and manual adjustments.</p>${table(['Metric','Original extracted','Management-ready current','Change'], beforeAfterRows(sum))}</div>
      <div class="card"><h2>Remaining limitations</h2><p class="muted">These limitations are what the owner chatbot and dashboards should disclose.</p>${lim.limits.map(x=>`<div class="issue">${esc(x)}</div>`).join('')}</div>
    </div>
    <div class="card"><h2>Management-ready category breakdown</h2><p class="muted">This is the first analysis layer: grouped by mapped/management category, using only rows currently accepted for management use.</p>${catRows.length ? table(['Mapped category','Rows','Inflows','Outflows','Net'], catRows) : '<div class="notice">No management-ready rows yet.</div>'}</div>
    <div class="grid two">
      <div class="card"><h2>Manual adjustments created during review</h2>${adjRows.length ? table(['Period','ID','Date','Direction','Amount','Category','Description','Linked review item'], adjRows) : '<div class="notice oknotice">No manual adjustments have been created yet.</div>'}</div>
      <div class="card"><h2>Rejected / excluded rows</h2>${rejRows.length ? table(['Period','ID','Date','Amount','Raw category','Mapped category','Description','Review item'], rejRows) : '<div class="notice oknotice">No rejected rows in this browser-local review state.</div>'}</div>
    </div>
    <div class="card"><h2>Rows still not ready for management analysis</h2><p class="muted">Click a row to open its linked review item and resolve it.</p>${pendingRows.length ? table(['Period','ID','Readiness','Amount','Raw category','Mapped category','Description','Review item'], pendingRows) : '<div class="notice oknotice">No pending management-blocking rows in this scope.</div>'}</div>
    <div class="card"><h2>Management-ready transaction preview</h2><p class="muted">First 120 rows in the current management-ready layer. Full export is available through Export certified package.</p>${certifiedRowsPreview.length ? table(['Period','ID','Date','Readiness','Direction','Amount','Raw category','Mapped category','Report group','Source'], certifiedRowsPreview) : '<div class="notice">No management-ready rows yet.</div>'}</div>
    <div class="toolbar"><button class="btn primary" onclick="exportPackage()">Export certified package</button><button class="btn good" onclick="exportManagementAnalysisPack()">Export management analysis pack</button><button class="btn" onclick="exportAccountantPack()">Export accountant review pack</button></div>
  `;
  document.querySelectorAll('[data-open-tx]').forEach(r => { if (r.dataset.openTx) r.addEventListener('click', () => openModal(r.dataset.openTx)); });
}

function captureReviewUI() {
  const q = document.getElementById('rq'), rs = document.getElementById('rs'), rv = document.getElementById('rv'), ra = document.getElementById('ra'), wrap = document.querySelector('#reviewTable .table-wrap');
  if (q) REVIEW_UI.q = q.value || '';
  if (rs) REVIEW_UI.status = rs.value || '';
  if (rv) REVIEW_UI.severity = rv.value || '';
  if (ra) REVIEW_UI.area = ra.value || '';
  if (wrap) { REVIEW_UI.scrollTop = wrap.scrollTop || 0; REVIEW_UI.scrollLeft = wrap.scrollLeft || 0; }
}
function itemMatchesReviewUI(i) {
  const q = String(REVIEW_UI.q || '').toLowerCase(), sf = REVIEW_UI.status || '', vf = REVIEW_UI.severity || '', af = REVIEW_UI.area || '';
  if (sf && itemStatus(i) !== sf) return false;
  if (vf && i.severity !== vf) return false;
  if (af && generalIssueType(i) !== af) return false;
  const hay = JSON.stringify({ i, d: STATE.decisions[i.review_item_id] || {}, c: STATE.controlOverrides[i.review_item_id] || {}, t: taskFor(i.review_item_id) || {} }).toLowerCase();
  if (q && !hay.includes(q)) return false;
  return true;
}
function restoreReviewScroll() { requestAnimationFrame(() => { const wrap = document.querySelector('#reviewTable .table-wrap'); if (wrap) { wrap.scrollTop = REVIEW_UI.scrollTop || 0; wrap.scrollLeft = REVIEW_UI.scrollLeft || 0; wrap.onscroll = () => { REVIEW_UI.scrollTop = wrap.scrollTop || 0; REVIEW_UI.scrollLeft = wrap.scrollLeft || 0; }; } }); }
function nextMatchingReviewItem(excludeId) { return scopeItems().find(x => x.review_item_id !== excludeId && itemMatchesReviewUI(x) && !['accountant_certified', 'rejected', 'management_ready'].includes(itemStatus(x))) || scopeItems().find(x => x.review_item_id !== excludeId && needsAccountantWork(x)); }

function renderReview() {
  const items = scopeItems();
  const areas = [...new Set(items.map(generalIssueType))].sort();
  const sevs = [...new Set(items.map(i => i.severity))].sort();
  const statusOpts = [['', 'All statuses'], ...Object.keys(STATUS_META).map(k => [k, STATUS_META[k].label])].map(([v, l]) => `<option value="${esc(v)}" ${REVIEW_UI.status === v ? 'selected' : ''}>${esc(l)}</option>`).join('');
  const sevOpts = `<option value="" ${REVIEW_UI.severity === '' ? 'selected' : ''}>All severities</option>` + sevs.map(x => `<option value="${esc(x)}" ${REVIEW_UI.severity === x ? 'selected' : ''}>${esc(x)}</option>`).join('');
  const areaOpts = `<option value="" ${REVIEW_UI.area === '' ? 'selected' : ''}>All review areas</option>` + areas.map(x => `<option value="${esc(x)}" ${REVIEW_UI.area === x ? 'selected' : ''}>${esc(x)}</option>`).join('');
  document.getElementById('review').innerHTML = `<div class="card"><h2>Manual Review Queue</h2><p class="muted">Accountant/team workspace for issue review. Use management-ready when the owner dashboard can use the data with limitations. Use accountant-certified only when professional/accountant review is complete.</p><div class="toolbar"><input id="rq" placeholder="Search issue, source, category..." value="${esc(REVIEW_UI.q)}" oninput="renderReviewTable()"><select id="rs" onchange="renderReviewTable()">${statusOpts}</select><select id="rv" onchange="renderReviewTable()">${sevOpts}</select><select id="ra" onchange="renderReviewTable()">${areaOpts}</select></div><div class="toolbar"><button class="btn primary" onclick="exportPackage()">Export certified package</button><button class="btn good" onclick="exportAccountantPack()">Export accountant review pack</button><button class="btn" onclick="exportState()">Export decisions only</button><button class="btn bad" onclick="resetProgress()">Reset local progress</button></div><div id="reviewTable"></div></div>`;
  renderReviewTable();
}
function renderReviewTable() {
  captureReviewUI();
  const rows = scopeItems().filter(itemMatchesReviewUI).map(i => {
    const tx = txById(i.related_object_id), et = tx ? effTx(tx) : null, d = STATE.decisions[i.review_item_id], task = taskFor(i.review_item_id);
    return `<tr class="clickable-row" data-review-id="${esc(i.review_item_id)}"><td>${esc(i.period)}</td><td>${esc(i.review_item_id)}</td><td>${pill(i.severity)}</td><td>${pill(statusLabel(itemStatus(i)), STATUS_META[itemStatus(i)]?.tone)}</td><td>${pill(generalIssueType(i), 'purple')}</td><td class="mini">${esc(i.issue_type)}</td><td>${et ? esc(et.raw_category) : ''}</td><td>${et ? money(et.amount_rsd_equivalent) : ''}</td><td>${task ? pill(task.status) : ''}</td><td>${esc(i.source_reference)}</td><td>${esc(i.description)}</td><td>${d ? esc(d.resolution_type) : 'Click to review'}</td></tr>`;
  });
  document.getElementById('reviewTable').innerHTML = table(['Period', 'ID', 'Severity', 'Status', 'Review area', 'Raw type', 'Raw category', 'Amount', 'Owner task', 'Source', 'Description', 'Decision'], rows.length ? rows : [`<tr><td colspan="12"><div class="notice oknotice">No items match the current filters.</div></td></tr>`]);
  document.querySelectorAll('[data-review-id]').forEach(r => r.addEventListener('click', () => openModal(r.dataset.reviewId)));
  restoreReviewScroll();
}

function renderTx() {
  document.getElementById('transactions').innerHTML = `<div class="card"><h2>Transactions after local edits</h2><p class="muted">Includes original normalized transactions, local transaction edits, mapping-rule effects, and manual adjustments from validation/control reviews.</p><div class="toolbar"><input id="tq" placeholder="Search transactions..." oninput="renderTxTable()"><select id="tdir" onchange="renderTxTable()"><option value="">All directions</option><option value="inflow">Inflows</option><option value="outflow">Outflows</option></select><select id="tex" onchange="renderTxTable()"><option value="active">Active only</option><option value="all">Show rejected too</option></select></div><div id="txTable"></div></div>`;
  renderTxTable();
}
function renderTxTable() {
  const q = String(document.getElementById('tq')?.value || '').toLowerCase(), dir = document.getElementById('tdir')?.value || '', all = document.getElementById('tex')?.value === 'all';
  const rows = displayTxRows(selected).filter(t => all || !t.excluded).filter(t => !dir || t.direction === dir).filter(t => !q || JSON.stringify(t).toLowerCase().includes(q)).slice(0, 900).map(t => `<tr><td>${t.period}</td><td>${esc(t.transaction_id)}</td><td>${esc(t.date)}</td><td>${pill(t.direction)}</td><td>${esc(t.payment_method)}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category)}</td><td>${esc(t.suggested_category)}</td><td>${esc(t.report_group)}</td><td>${esc(t.description)}</td><td>${esc(t.source_reference)}</td><td>${pill(t.excluded ? 'Rejected' : (t.review_status || 'staged'))}</td></tr>`);
  document.getElementById('txTable').innerHTML = table(['Period', 'ID', 'Date', 'Direction', 'Method', 'Amount', 'Raw category', 'Mapped category', 'Report group', 'Description', 'Source', 'Status'], rows);
}
function renderMapping() {
  const rules = (STATE.mappingRules || []).map(r => `<tr><td>${esc(r.period || 'ALL')}</td><td>${pill(r.direction)}</td><td>${esc(r.raw_category)}</td><td>${esc(r.suggested_category)}</td><td>${esc(r.report_group)}</td><td>${esc(r.created_from_review_item_id)}</td><td>${esc(r.reviewer)}</td><td>${esc(r.created_at)}</td></tr>`);
  const base = scopeMonths().flatMap(m => m.mapping_suggestions).map(x => `<tr><td>${x.period}</td><td>${pill(x.direction)}</td><td>${esc(x.raw_category)}</td><td>${esc(x.suggested_category)}</td><td>${esc(x.report_group)}</td><td>${x.rows_affected}</td><td>${money(x.total_rsd_equivalent)}</td><td>${pill(x.confidence)}</td></tr>`);
  const suggestionRows = categorySuggestionRows().map(x => `<tr><td>${esc(x.category)}</td><td>${pill(x.source, x.source.includes('local') ? 'purple' : 'info')}</td></tr>`);
  document.getElementById('mapping').innerHTML = `<div class="card"><h2>Mapped category dropdown suggestions</h2><p class="muted">These are the categories shown in the mapped-category dropdown inside the review popup. The list combines default categories, previous import mappings, normalized transactions, approved mapping rules, and new categories typed during review.</p>${table(['Mapped category suggestion', 'Source'], suggestionRows)}</div><div class="card"><h2>Approved mapping rules created during review</h2>${rules.length ? table(['Period', 'Direction', 'Raw category', 'Mapped category', 'Report group', 'From item', 'Reviewer', 'Created'], rules) : '<div class="notice">No local mapping rules created yet.</div>'}</div><div class="card"><h2>Original mapping suggestions</h2>${table(['Period', 'Direction', 'Raw category', 'Suggested category', 'Report group', 'Rows', 'Total', 'Confidence'], base)}</div>`;
}
function renderValidation() {
  const rows = scopeMonths().flatMap(m => m.validation_results).map(v => {
    const linked = allItems().find(i => i.related_object_id === v.validation_id);
    const c = linked ? STATE.controlOverrides[linked.review_item_id] : null;
    return `<tr><td>${v.period}</td><td>${esc(v.validation_id || '')}</td><td>${esc(v.check_name)}</td><td>${pill(v.status)}</td><td>${esc(v.severity)}</td><td>${esc(v.difference)}</td><td>${c ? pill('Control reviewed', 'purple') : ''}</td><td>${esc(c?.trusted_source || '')}</td><td>${esc(c?.accounting_treatment || '')}</td><td>${esc(v.explanation)}</td></tr>`;
  });
  document.getElementById('validation').innerHTML = `<div class="card"><h2>Validation / control results</h2><p class="muted">These can be resolved in the Review Center even when no single transaction is linked. Use this area for formula, monthly summary, tie-out, and source-control issues.</p>${table(['Period', 'ID', 'Check', 'Status', 'Severity', 'Original difference', 'Local status', 'Trusted source', 'Treatment', 'Explanation'], rows)}</div>`;
}
function renderCert() {
  const rows = scopeMonths().map(m => {
    const s = stats(m.period), mg = managementStatus(m.period), fm = formalStatus(m.period), bs = batchStatus(m.period);
    const mgDisabled = s.management_blocking > 0 || selected === 'ALL';
    const certDisabled = s.formal_blocking > 0 || selected === 'ALL';
    return `<tr><td>${m.label}</td><td>${pill(mg.label, mg.tone)}</td><td>${pill(fm.label, fm.tone)}</td><td>${s.management_blocking}</td><td>${s.formal_blocking}</td><td>${localQuality(m.period).toFixed(1)}/100</td><td>${bs.management_ready ? esc(bs.management_ready_at || '') : ''}</td><td>${bs.accountant_certified ? esc(bs.accountant_certified_at || '') : ''}</td><td>${selected === 'ALL' ? '<span class="mini">Choose one month</span>' : `<button class="btn purpleBtn" ${mgDisabled ? 'disabled' : ''} onclick="markMonthManagementReady()">Mark management-ready</button> <button class="btn good" ${certDisabled ? 'disabled' : ''} onclick="certifyMonthAccountant()">Accountant certify</button>`}</td></tr>`;
  });
  document.getElementById('certification').innerHTML = `<div class="card"><h2>Batch status and certification</h2><p class="muted">Management-ready means the owner can use dashboards with warnings. Accountant-certified means professional/accounting review is complete for that period.</p>${table(['Month', 'Management status', 'Accountant status', 'Mgmt blockers', 'Formal blockers', 'Quality', 'Management-ready at', 'Certified at', 'Action'], rows)}</div>`;
}

function renderAssistant() {
  const qs = role() === 'owner'
    ? ['Can I use this data for business decisions?', 'What does my accountant need from me?', 'How did May perform?']
    : ['What should the accountant team review next?', 'Which months are management-ready?', 'What blocks accountant certification?'];
  document.getElementById('assistant').innerHTML = `<div class="card"><h2>Assistant Demo</h2><p class="muted">Current version is rule-based and uses the loaded/imported data plus your local review state. Later this becomes the database-backed AI assistant.</p><div class="toolbar">${qs.map(q => `<button class="btn" onclick="ask('${esc(q)}')">${esc(q)}</button>`).join('')}</div><div class="toolbar"><input id="customQ" placeholder="Type a question..."><button class="btn primary" onclick="ask(document.getElementById('customQ').value)">Ask</button></div><div id="answer" class="answer">Select a question.</div></div>`;
}
function ask(q) {
  q = String(q || '').toLowerCase();
  const st = stats(), a = adjSummary(), mg = managementStatus(), fm = formalStatus(), tasks = openTasks();
  let ans = '';
  if (q.includes('accountant need') || q.includes('owner') || q.includes('clarification')) {
    ans = `Open owner clarification tasks: ${tasks.length}\n\n${tasks.slice(0, 8).map(t => `- ${t.period} ${t.review_item_id}: ${t.question}`).join('\n') || 'No owner clarification tasks are currently open.'}\n\nLimitation: this is browser-local demo state.`;
  } else if (q.includes('business') || q.includes('management')) {
    ans = `Management status: ${mg.label}\n${mg.detail}\n\nAdjusted closing cash: ${money(a.closing)}\nManagement blockers: ${st.management_blocking}\n\nAccountant certification status: ${fm.label}\n${fm.detail}\n\nUse this data for management dashboards only when labelled management-ready. Formal accounting use requires accountant certification.`;
  } else if (q.includes('may')) {
    const aa = adjSummary('2020-05'), sm = managementStatus('2020-05'), sf = formalStatus('2020-05');
    ans = `May 2020 based on current local review state:\nNet cash movement: ${money(aa.net)}\nClosing position: ${money(aa.closing)}\nManagement status: ${sm.label}\nAccountant status: ${sf.label}\n\nThe assistant should label this as management-ready or accountant-certified depending on the status recorded in the certification workflow.`;
  } else if (q.includes('review next') || q.includes('blocks')) {
    const items = scopeItems().filter(needsAccountantWork).slice(0, 10);
    ans = `Accountant/team queue for ${label()}: ${items.length} item(s) shown below.\n\n${items.map(i => `- ${i.review_item_id}: ${statusLabel(itemStatus(i))} | ${generalIssueType(i)} | ${i.description}`).join('\n') || 'No pending accountant work in this scope.'}`;
  } else if (q.includes('month')) {
    ans = DATA.months.map(m => `${m.label}: ${managementStatus(m.period).label}; ${formalStatus(m.period).label}`).join('\n');
  } else {
    ans = `Current scope: ${label()}\nAdjusted inflows: ${money(a.inflows)}\nAdjusted outflows: ${money(a.outflows)}\nAdjusted net movement: ${money(a.net)}\nAdjusted closing: ${money(a.closing)}\n\nManagement status: ${mg.label}\nAccountant status: ${fm.label}\n\nLimitations: this answer uses local browser review state, not PostgreSQL yet.`;
  }
  document.getElementById('answer').textContent = ans;
}
function renderReports() {
  const links = DATA.months.map(m => `<div class="step"><div class="num">PDF</div><div><strong>${m.label} report</strong><div class="muted">Human review report generated from the import package.</div></div><a class="btn primary" href="${m.summary.pdf_report}" target="_blank">Open PDF</a></div><div class="step"><div class="num">XLSX</div><div><strong>${m.label} generated review workbook</strong><div class="muted">Structured workbook with transactions, mappings, validation results, and review queue.</div></div><a class="btn" href="${reviewWorkbookPath(m.period)}" target="_blank">Open workbook</a></div>`).join('');
  document.getElementById('reports').innerHTML = `<div class="card"><h2>Reports and source evidence files</h2><p class="muted">These generated files support manual checking. The production version will also open the private original Excel workbook from secure storage.</p><div class="timeline">${links}</div></div>`;
}


function monthByPeriod(period) { return DATA.months.find(m => m.period === period); }
function periodSlug(period) {
  if (period === '2020-03') return 'March';
  if (period === '2020-04') return 'April';
  if (period === '2020-05') return 'May';
  return period || 'Period';
}
function reviewWorkbookPath(period) {
  const m = periodSlug(period);
  return `/review-workbooks/University_${m}_2020_Import_Review_Workbook.xlsx`;
}
function reportPdfPath(period) { return monthByPeriod(period)?.summary?.pdf_report || '#'; }
function openLink(path) { window.open(path, '_blank', 'noopener'); }
function sourceRowNumberFromRef(ref) {
  const r = String(ref || '');
  const m1 = r.match(/row\s*(\d+)/i);
  if (m1) return Number(m1[1]);
  const m2 = r.match(/[A-Z]+(\d+)\s*:/i) || r.match(/[A-Z]+(\d+)/i);
  if (m2) return Number(m2[1]);
  return null;
}
function sourceSheetFromRef(ref) {
  const r = String(ref || '');
  if (!r) return '';
  return r.split('!')[0].replace(/,$/, '').trim();
}
function monthTransactions(period) { return (monthByPeriod(period)?.transactions || []); }
function evidenceTransactionRows(i, tx) {
  const period = i.period;
  const out = [];
  if (tx) out.push(tx);
  const ref = i.source_reference || tx?.source_reference || '';
  const sheet = tx?.source_sheet || sourceSheetFromRef(ref);
  const row = Number(tx?.source_row || sourceRowNumberFromRef(ref) || 0);
  if (sheet && row) {
    monthTransactions(period).forEach(t => {
      if (String(t.source_sheet || '') === String(sheet) && Math.abs(Number(t.source_row || 0) - row) <= 3) out.push(t);
    });
  }
  const text = `${i.description || ''} ${i.details || ''} ${i.suggested_action || ''}`;
  const re = /(\d{2}\.\d{2}\.\d{4}\.?)[^\d]{0,40}row\s*(\d+)/gi;
  let m;
  while ((m = re.exec(text))) {
    const mentionedSheet = m[1];
    const mentionedRow = Number(m[2]);
    monthTransactions(period).forEach(t => {
      if (String(t.source_sheet || '').includes(mentionedSheet) && Math.abs(Number(t.source_row || 0) - mentionedRow) <= 3) out.push(t);
    });
  }
  const seen = new Set();
  return out.filter(t => {
    const k = t.transaction_id || `${t.source_reference}-${t.source_row}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  }).sort((a,b) => String(a.source_sheet).localeCompare(String(b.source_sheet)) || Number(a.source_row||0)-Number(b.source_row||0)).slice(0, 18);
}
function linkedValidationRows(i) {
  const period = i.period;
  const rows = [];
  const direct = valById(i.related_object_id);
  if (direct) rows.push(direct);
  const txt = `${i.related_object_id || ''} ${i.description || ''} ${i.details || ''}`;
  (monthByPeriod(period)?.validation_results || []).forEach(v => {
    if (txt.includes(v.validation_id) || txt.toLowerCase().includes(String(v.check_name || '').toLowerCase())) rows.push(v);
  });
  const seen = new Set();
  return rows.filter(v => { const k = v.validation_id || v.check_name; if (seen.has(k)) return false; seen.add(k); return true; }).slice(0, 10);
}
function monthlyStatusEvidence(i) {
  const m = monthByPeriod(i.period);
  const all = m?.monthly_status_summary || [];
  if (!all.length) return [];
  const row = sourceRowNumberFromRef(i.source_reference);
  if (row) return all.filter(x => Number(x.source_row || x.source_row_number || 0) === row).slice(0, 8);
  const needle = `${i.description || ''} ${i.details || ''}`.toLowerCase();
  const matched = all.filter(x => needle.includes(String(x.label || x.monthly_label || '').toLowerCase())).slice(0, 8);
  return matched.length ? matched : all.filter(x => Number(x.total_rsd_equivalent || x.amount_rsd_equivalent || 0) !== 0).slice(0, 8);
}
function objectRows(obj, preferred = []) {
  if (!obj) return '';
  const keys = preferred.concat(Object.keys(obj).filter(k => !preferred.includes(k))).filter(k => obj[k] !== undefined && obj[k] !== null && String(obj[k]) !== '');
  return keys.map(k => `<tr><td class="mini">${esc(k)}</td><td>${esc(obj[k])}</td></tr>`).join('');
}
function sourceEvidencePanel(i, tx) {
  const effective = tx ? effTx(tx) : null;
  const txRows = evidenceTransactionRows(i, tx);
  const valRows = linkedValidationRows(i);
  const monthlyRows = monthlyStatusEvidence(i);
  const m = monthByPeriod(i.period);
  const dailyRows = (m?.daily_balances || []).filter(d => {
    const sheet = tx?.source_sheet || sourceSheetFromRef(i.source_reference || '');
    return sheet ? String(d.source_sheet || '').includes(sheet) : false;
  }).slice(0, 4);
  const txTableRows = txRows.map(t => `<tr><td>${esc(t.source_sheet)}</td><td>${esc(t.source_row)}</td><td>${esc(t.transaction_id)}</td><td>${esc(t.direction)}</td><td>${esc(t.payment_method)}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category)}</td><td>${esc(t.suggested_category)}</td><td>${esc(t.description)}</td></tr>`);
  const valTableRows = valRows.map(v => `<tr><td>${esc(v.validation_id || '')}</td><td>${esc(v.check_name)}</td><td>${pill(v.status)}</td><td>${esc(v.severity)}</td><td>${esc(v.expected_value)}</td><td>${esc(v.actual_value)}</td><td>${esc(v.difference)}</td><td>${esc(v.explanation)}</td></tr>`);
  const monRows = monthlyRows.map(r => `<tr><td>${esc(r.source_sheet || 'Mesecni promet')}</td><td>${esc(r.source_row || r.source_row_number || '')}</td><td>${esc(r.side || '')}</td><td>${esc(r.label || r.monthly_label || '')}</td><td>${money(r.total_rsd_equivalent ?? r.amount_rsd_equivalent)}</td><td>${esc(r.cash_rsd ?? '')}</td><td>${esc(r.bank_rsd ?? '')}</td></tr>`);
  const dailyTableRows = dailyRows.map(r => `<tr><td>${esc(r.date)}</td><td>${esc(r.source_sheet)}</td><td>${money(r.opening_total)}</td><td>${money(r.inflows_total)}</td><td>${money(r.outflows_total)}</td><td>${money(r.closing_total)}</td><td>${pill(r.status)}</td></tr>`);
  const normalizedBlock = effective
    ? `<details open><summary>Data our app created from the source row</summary><table><tbody>${objectRows(effective, ['transaction_id','date','direction','payment_method','currency_original','amount_rsd_equivalent','raw_category','suggested_category','report_group','counterparty','counterparty_type','description','source_reference','source_sheet','source_row','confidence','review_status'])}</tbody></table></details>`
    : `<details open><summary>Data our app created for this control issue</summary><table><tbody>${objectRows(valRows[0] || i, ['validation_id','check_name','status','severity','expected_value','actual_value','difference','explanation','review_item_id','issue_type','description','details','suggested_action','source_reference'])}</tbody></table></details>`;
  return `<div class="source-box evidence-box"><h3>Source evidence viewer</h3>
    <p class="muted">Use this section to compare the original source reference against the normalized data our app created. The demo uses anonymized/generated review workbooks and PDF reports; the production app will open the private original Excel file from secure storage.</p>
    <div class="toolbar">
      <button class="btn primary" onclick="openLink('${esc(reportPdfPath(i.period))}')">Open PDF review report</button>
      <button class="btn" onclick="openLink('${esc(reviewWorkbookPath(i.period))}')">Open generated review workbook</button>
      <span class="mini">Exact source: ${esc(i.source_reference || effective?.source_reference || 'Workbook/monthly control issue')}</span>
    </div>
    ${normalizedBlock}
    ${txTableRows.length ? `<details open><summary>Original Excel row context / nearby extracted rows</summary>${table(['Sheet','Row','Transaction ID','Direction','Method','Amount','Raw category','Mapped category','Description'], txTableRows)}</details>` : '<div class="notice">No direct transaction row was linked. Use the validation/control evidence below and the workbook/report buttons above.</div>'}
    ${valTableRows.length ? `<details open><summary>Linked validation/control evidence</summary>${table(['ID','Check','Status','Severity','Expected','Actual','Difference','Explanation'], valTableRows)}</details>` : ''}
    ${monRows.length ? `<details><summary>Mesecni promet / monthly status evidence</summary>${table(['Sheet','Row','Side','Label','Total RSD equivalent','Cash RSD','Bank RSD'], monRows)}</details>` : ''}
    ${dailyTableRows.length ? `<details><summary>Daily balance tie-out for the same sheet</summary>${table(['Date','Sheet','Opening','Inflows','Outflows','Closing','Status'], dailyTableRows)}</details>` : ''}
  </div>`;
}


function smallObjRows(obj, preferred = []) {
  if (!obj) return '<div class="mini">No data available.</div>';
  const keys = preferred.length ? preferred.filter(k => obj[k] !== undefined && obj[k] !== null && obj[k] !== '') : Object.keys(obj).slice(0, 16);
  if (!keys.length) return '<div class="mini">No populated fields available.</div>';
  return table(['Field', 'Value'], keys.map(k => `<tr><td class="mini">${esc(k)}</td><td>${esc(obj[k])}</td></tr>`));
}
function evidenceTxTable(rows, title = 'Rows') {
  if (!rows || !rows.length) return `<div class="source-box"><h3>${esc(title)}</h3><div class="mini">No related rows found in the generated import package.</div></div>`;
  const body = rows.map(t => `<tr><td>${esc(t.source_sheet || '')}</td><td>${esc(t.source_row || '')}</td><td>${esc(t.date || '')}</td><td>${pill(t.direction || '')}</td><td>${esc(t.payment_method || '')}</td><td>${money(t.amount_rsd_equivalent)}</td><td>${esc(t.raw_category || '')}</td><td>${esc(t.suggested_category || '')}</td><td>${esc(t.description || '')}</td><td>${esc(t.transaction_id || '')}</td></tr>`);
  return `<div class="source-box"><h3>${esc(title)}</h3>${table(['Sheet','Row','Date','Direction','Method','Amount','Raw category','Mapped category','Description','Transaction ID'], body)}</div>`;
}
function monthlyEvidenceTable(rows) {
  if (!rows || !rows.length) return '<div class="mini">No monthly status rows available for this issue.</div>';
  const body = rows.map(r => {
    const rowNo = r.source_row ?? r.source_row_number ?? '';
    const label = r.label ?? r.monthly_label ?? '';
    const amount = r.total_rsd_equivalent ?? r.amount_rsd_equivalent ?? '';
    return `<tr><td>${esc(rowNo)}</td><td>${esc(r.side || '')}</td><td>${esc(label)}</td><td>${esc(r.cash_rsd ?? '')}</td><td>${esc(r.cash_eur_converted_rsd ?? r.cash_eur ?? '')}</td><td>${esc(r.bank_rsd ?? '')}</td><td>${esc(r.bank_eur_converted_rsd ?? r.bank_eur ?? '')}</td><td>${money(amount || 0)}</td></tr>`;
  });
  return table(['Row','Side','Monthly label','Cash RSD','Cash EUR/RSD','Bank RSD','Bank EUR/RSD','Total RSD eq.'], body);
}
function evidencePanel(i, tx) {
  const ev = SOURCE_EVIDENCE[i.review_item_id] || {};
  const wb = ev.original_workbook || {};
  const parsed = ev.source_reference_parsed || {};
  const linkedTx = ev.linked_transaction || tx || null;
  const linkedVal = ev.linked_validation || valById(i.related_object_id) || null;
  const sourceContexts = ev.source_context?.sheet_contexts || [];
  const sameCat = ev.source_context?.same_category_transactions || [];
  const dailyRows = ev.daily_balance_context || [];
  const monthlyRows = ev.monthly_status_context || [];
  const firstSourceRows = sourceContexts.map((ctx, idx) => evidenceTxTable(ctx.nearby_transactions, `Original Excel row context ${idx + 1}: ${ctx.reference?.sheet || ''}${ctx.reference?.row ? ' around row ' + ctx.reference.row : ''}`)).join('');
  const dailyTable = dailyRows.length ? table(['Date','Sheet','Opening','Inflows','Outflows','Closing','Status'], dailyRows.map(r => `<tr><td>${esc(r.date)}</td><td>${esc(r.source_sheet)}</td><td>${money(r.opening_total)}</td><td>${money(r.inflows_total)}</td><td>${money(r.outflows_total)}</td><td>${money(r.closing_total)}</td><td>${pill(r.status || '')}</td></tr>`)) : '<div class="mini">No daily balance control row found for this source sheet.</div>';
  return `<div class="evidence-panel">
    <h3>Source evidence viewer</h3>
    <div class="notice evidence-note"><strong>Purpose:</strong> compare the original workbook reference, the extracted/normalized row, and the generated review documents before making a decision. In production, this panel will open the secured original Excel file directly. In this Railway demo, the raw workbook is not bundled publicly; the exact workbook/sheet/row reference and extracted source rows are shown instead.</div>
    <div class="evidence-grid">
      <div class="source-box"><h3>1. Original workbook reference</h3>
        <strong>Original workbook:</strong> ${esc(wb.original_workbook_name || 'Company source workbook')}<br>
        <strong>Month:</strong> ${esc(ev.month_label || i.period)}<br>
        <strong>Source reference:</strong> ${esc(i.source_reference || 'Workbook/monthly control')}<br>
        <strong>Parsed sheet:</strong> ${esc(parsed.sheet || '')}<br>
        <strong>Parsed row/cell:</strong> ${esc(parsed.row || parsed.cell || '')}<br>
        <div class="toolbar evidence-actions"><a class="btn primary" href="${esc(wb.pdf_report || '#')}" target="_blank">Open PDF report</a><a class="btn" href="${esc(wb.review_workbook || '#')}" target="_blank">Download review workbook</a></div>
      </div>
      <div class="source-box"><h3>2. Review item created by app</h3>${smallObjRows(i, ['review_item_id','period','severity','issue_type','description','details','suggested_action','related_object_id','source_reference'])}</div>
    </div>
    ${linkedTx ? `<div class="source-box"><h3>3. Normalized transaction created by app</h3>${smallObjRows(linkedTx, ['transaction_id','date','direction','payment_method','currency_original','amount_rsd_equivalent','raw_category','suggested_category','report_group','counterparty','counterparty_type','description','source_sheet','source_row','confidence','review_status'])}</div>` : ''}
    ${linkedVal ? `<div class="source-box"><h3>3. Validation/control record created by app</h3>${smallObjRows(linkedVal, ['validation_id','check_name','status','severity','expected_value','actual_value','difference','explanation','action_required'])}</div>` : ''}
    ${firstSourceRows || '<div class="source-box"><h3>Original Excel row context</h3><div class="mini">No direct row context was found. This usually happens for workbook-level validation/control issues. Use the monthly status and validation sections below.</div></div>'}
    <div class="source-box"><h3>Daily balance control for referenced sheet</h3>${dailyTable}</div>
    <div class="source-box"><h3>Monthly status rows from generated review document</h3>${monthlyEvidenceTable(monthlyRows)}</div>
    ${evidenceTxTable(sameCat, 'Same-category transactions sorted from generated import data')}
  </div>`;
}
function openModal(id) {
  currentId = id;
  const i = itemById(id), tx = txById(i.related_object_id), t = taskFor(id), d = STATE.decisions[id];
  document.getElementById('modalTitle').textContent = `${role() === 'owner' ? 'Owner task' : 'Review item'}: ${id}`;
  document.getElementById('modalSubtitle').innerHTML = `${esc(i.period)} • ${pill(i.severity)} ${pill(generalIssueType(i), 'purple')} ${pill(statusLabel(itemStatus(i)), STATUS_META[itemStatus(i)]?.tone)}`;
  const taskBlock = t ? `<div class="source-box"><h3>Owner clarification task</h3><strong>Question:</strong> ${esc(t.question || '')}<br><strong>Status:</strong> ${pill(t.status)}<br><strong>Owner answer:</strong> ${esc(t.owner_answer || '')}</div>` : '';
  const source = evidencePanel(i, tx) + `<div class="source-box"><h3>Issue summary</h3><strong>Description:</strong> ${esc(i.description)}<br><strong>Details:</strong> ${esc(i.details || '')}<br><strong>Suggested action:</strong> ${esc(i.suggested_action || '')}<br><strong>Source:</strong> ${esc(i.source_reference || 'Workbook/control issue')}<br><strong>Related object:</strong> ${esc(i.related_object_id || '')}</div>`;
  const evidence = sourceEvidencePanel(i, tx);
  const editable = tx ? transactionForm(i, tx) : controlForm(i);
  const roleFields = role() === 'accountant'
    ? `<div class="field-grid"><div class="field"><label>Reviewer / accountant name</label><input id="reviewerName" value="${esc(d?.reviewer || 'Accountant Reviewer')}"></div><div class="field"><label>Question to owner, if clarification is needed</label><input id="ownerQuestion" value="${esc(t?.question || '')}" placeholder="Example: Was this legal cost, loan repayment, or owner withdrawal?"></div><div class="field" style="grid-column:1/-1"><label>Review note / accountant explanation</label><textarea id="reviewNote">${esc(d?.note || '')}</textarea></div></div>`
    : `<div class="field-grid"><div class="field"><label>Owner name</label><input id="reviewerName" value="${esc(d?.reviewer || 'Business Owner')}"></div><div class="field" style="grid-column:1/-1"><label>Owner answer / business context</label><textarea id="reviewNote" placeholder="Give business context. You do not need to make a final accounting judgment.">${esc(t?.owner_answer || d?.note || '')}</textarea></div></div>`;
  const buttons = role() === 'accountant' ? accountantButtons() : ownerButtons();
  document.getElementById('modalBody').innerHTML = `${taskBlock}${source}${evidence}<div class="divider"></div>${editable}<div class="divider"></div>${roleFields}<div class="actions modal-actions">${buttons}</div><div class="mini">Current demo stores this decision in browser localStorage. Export your package before clearing browser data.</div>`;
  document.getElementById('reviewModal').classList.add('open');
}
function transactionForm(i, tx) {
  const e = effTx(tx);
  const lists = suggestionDatalists();
  return `<div><h3>Editable transaction fields</h3>${lists}<input type="hidden" id="txId" value="${esc(e.transaction_id)}"><div class="field-grid"><div class="field"><label>Date</label><input id="txDate" value="${esc(e.date)}"></div><div class="field"><label>Direction</label><select id="txDirection"><option value="inflow" ${e.direction === 'inflow' ? 'selected' : ''}>inflow</option><option value="outflow" ${e.direction === 'outflow' ? 'selected' : ''}>outflow</option></select></div><div class="field"><label>Payment method</label><input id="txPayment" value="${esc(e.payment_method)}"></div><div class="field"><label>Currency</label><input id="txCurrency" value="${esc(e.currency_original)}"></div><div class="field"><label>Amount RSD equivalent</label><input id="txAmount" type="number" step="0.01" value="${Number(e.amount_rsd_equivalent || 0)}"></div><div class="field"><label>Raw category</label><input id="txRaw" value="${esc(e.raw_category)}"></div>${categoryInput('txCat', 'Mapped / management category', e.suggested_category)}${reportGroupInput('txGroup', 'Report group', e.report_group)}<div class="field"><label>Counterparty type</label><input id="txCpType" value="${esc(e.counterparty_type)}"></div><div class="field"><label>Counterparty token</label><input id="txCp" value="${esc(e.counterparty)}"></div><div class="field" style="grid-column:1/-1"><label>Description</label><textarea id="txDesc">${esc(e.description)}</textarea></div></div><label class="checkbox"><input type="checkbox" id="createMappingRule"><span>Create/update a reusable mapping rule from this row. The mapped category is saved to suggestions even if this is not ticked.</span></label></div>`;
}
function controlForm(i) {
  const val = valById(i.related_object_id), c = STATE.controlOverrides[i.review_item_id] || {}, a = STATE.manualAdjustments[i.review_item_id] || {};
  const diff = (c.difference_amount ?? Number(val?.difference || 0)) || 0;
  const lists = suggestionDatalists();
  return `<div><h3>Editable validation / control issue</h3>${lists}<input type="hidden" id="ctrlIssueId" value="${esc(i.review_item_id)}"><div class="field-grid"><div class="field"><label>Review area</label><select id="ctrlArea">${generalOptions(c.review_area || generalIssueType(i))}</select></div><div class="field"><label>Trusted source</label><select id="ctrlTrustedSource"><option ${c.trusted_source === 'Daily extracted transactions' ? 'selected' : ''}>Daily extracted transactions</option><option ${c.trusted_source === 'Monthly status sheet' ? 'selected' : ''}>Monthly status sheet</option><option ${c.trusted_source === 'Manual accountant review' ? 'selected' : ''}>Manual accountant review</option><option ${c.trusted_source === 'Owner clarification' ? 'selected' : ''}>Owner clarification</option></select></div><div class="field"><label>Accounting treatment / provisional treatment</label><input id="ctrlTreatment" value="${esc(c.accounting_treatment || '')}" placeholder="Example: formula issue acknowledged; use daily extracted transactions"></div>${categoryInput('ctrlAffectedLine', 'Affected report line / category', c.affected_report_line || '', 'Choose existing or type affected category/report line')}<div class="field"><label>Expected value</label><input id="ctrlExpected" value="${esc(c.expected_value || val?.expected_value || '')}"></div><div class="field"><label>Actual value</label><input id="ctrlActual" value="${esc(c.actual_value || val?.actual_value || '')}"></div><div class="field"><label>Difference amount</label><input id="ctrlDifference" type="number" step="0.01" value="${Number(diff || 0)}"></div><div class="field"><label>Corrected amount RSD</label><input id="ctrlCorrectedAmount" type="number" step="0.01" value="${Number(c.corrected_amount_rsd || 0)}"></div><div class="field" style="grid-column:1/-1"><label>Root cause / issue found</label><textarea id="ctrlRootCause">${esc(c.root_cause || '')}</textarea></div><div class="field" style="grid-column:1/-1"><label>Corrective action / certification explanation</label><textarea id="ctrlCorrectiveAction">${esc(c.corrective_action || '')}</textarea></div></div><label class="checkbox"><input type="checkbox" id="ctrlCreateAdjustment" ${a.transaction_id ? 'checked' : ''}><span>Create a manual adjustment transaction from this control issue.</span></label><div class="source-box"><h3>Optional manual adjustment</h3><div class="field-grid"><div class="field"><label>Adjustment date</label><input id="adjDate" value="${esc(a.date || `${i.period}-28`)}"></div><div class="field"><label>Direction</label><select id="adjDirection"><option value="inflow" ${a.direction === 'inflow' ? 'selected' : ''}>inflow</option><option value="outflow" ${a.direction === 'outflow' ? 'selected' : ''}>outflow</option></select></div><div class="field"><label>Amount RSD</label><input id="adjAmount" type="number" step="0.01" value="${Number(a.amount_rsd_equivalent || 0)}"></div>${categoryInput('adjCategory', 'Mapped category', a.suggested_category || c.affected_report_line || '')}${reportGroupInput('adjGroup', 'Report group', a.report_group || '')}<div class="field"><label>Description</label><input id="adjDesc" value="${esc(a.description || i.description || '')}"></div></div></div>${val ? `<div class="source-box"><strong>Original validation:</strong><br>${esc(val.check_name)}<br>Expected: ${esc(val.expected_value)}<br>Actual: ${esc(val.actual_value)}<br>Difference: ${esc(val.difference)}<br>${esc(val.explanation || '')}</div>` : ''}</div>`;
}
function accountantButtons() {
  return [
    '<button class="btn" onclick="saveDecision(\'in_review\')">Mark in review</button>',
    '<button class="btn warn" onclick="saveDecision(\'needs_owner_clarification\')">Ask owner / clarification needed</button>',
    '<button class="btn purpleBtn" onclick="saveDecision(\'management_ready\')">Accept for management analysis</button>',
    '<button class="btn warn" onclick="saveDecision(\'needs_accountant_review\')">Needs accountant review</button>',
    '<button class="btn good" onclick="saveDecision(\'accountant_certified\')">Accountant certify item</button>',
    '<button class="btn bad" onclick="saveDecision(\'rejected\')">Reject / exclude</button>',
    '<button class="btn primary" onclick="saveDecision(\'management_ready\', true)">Save + open next</button>'
  ].join('');
}
function ownerButtons() {
  return [
    '<button class="btn primary" onclick="saveDecision(\'owner_reviewed\')">Owner reviewed / answer sent</button>',
    '<button class="btn warn" onclick="saveDecision(\'needs_accountant_review\')">Not sure - accountant review needed</button>',
    '<button class="btn bad" onclick="saveDecision(\'escalated\')">Flag as incorrect / urgent</button>'
  ].join('');
}
function closeModal() { document.getElementById('reviewModal').classList.remove('open'); currentId = null; }
function readControlForm() {
  const id = document.getElementById('ctrlIssueId')?.value; if (!id) return null;
  return { kind: 'control', review_item_id: id, review_area: document.getElementById('ctrlArea').value, accounting_treatment: document.getElementById('ctrlTreatment').value, affected_report_line: document.getElementById('ctrlAffectedLine').value, trusted_source: document.getElementById('ctrlTrustedSource').value, expected_value: document.getElementById('ctrlExpected').value, actual_value: document.getElementById('ctrlActual').value, difference_amount: Number(document.getElementById('ctrlDifference').value || 0), corrected_amount_rsd: Number(document.getElementById('ctrlCorrectedAmount').value || 0), root_cause: document.getElementById('ctrlRootCause').value, corrective_action: document.getElementById('ctrlCorrectiveAction').value, create_adjustment: document.getElementById('ctrlCreateAdjustment').checked, adjustment: { date: document.getElementById('adjDate').value, direction: document.getElementById('adjDirection').value, amount_rsd_equivalent: Number(document.getElementById('adjAmount').value || 0), suggested_category: document.getElementById('adjCategory').value, report_group: document.getElementById('adjGroup').value, description: document.getElementById('adjDesc').value } };
}
function readForm() {
  const id = document.getElementById('txId')?.value;
  if (id) return { kind: 'transaction', transaction_id: id, date: document.getElementById('txDate').value, direction: document.getElementById('txDirection').value, payment_method: document.getElementById('txPayment').value, currency_original: document.getElementById('txCurrency').value, amount_rsd_equivalent: Number(document.getElementById('txAmount').value || 0), raw_category: document.getElementById('txRaw').value, suggested_category: document.getElementById('txCat').value, report_group: document.getElementById('txGroup').value, counterparty_type: document.getElementById('txCpType').value, counterparty: document.getElementById('txCp').value, description: document.getElementById('txDesc').value };
  return readControlForm();
}
function resolutionType(status, form) {
  if (status === 'owner_reviewed') return 'owner_context_confirmed';
  if (status === 'management_ready') return form?.kind === 'control' ? 'control_resolved_for_management_analysis' : 'accepted_for_management_analysis';
  if (status === 'needs_owner_clarification') return 'sent_to_owner_for_clarification';
  if (status === 'needs_accountant_review') return 'requires_accountant_review';
  if (status === 'accountant_certified') return form?.kind === 'control' ? 'accountant_certified_control_issue' : 'accountant_certified_transaction';
  if (status === 'rejected') return 'rejected_excluded_from_import';
  if (status === 'escalated') return 'escalated';
  return 'marked_in_review';
}
function saveDecision(status, next = false) {
  captureReviewUI();
  const i = itemById(currentId), tx = txById(i.related_object_id), before = tx ? effTx(tx) : { review_item: i, validation: valById(i.related_object_id), existing_control: STATE.controlOverrides[i.review_item_id] || null };
  const form = readForm(), reviewer = document.getElementById('reviewerName')?.value || (role() === 'owner' ? 'Business Owner' : 'Accountant Reviewer'), note = document.getElementById('reviewNote')?.value || '', ownerQuestion = document.getElementById('ownerQuestion')?.value || '';
  if (form?.kind === 'transaction' && tx) {
    const over = { ...form }; delete over.kind;
    over.review_status = statusLabel(status);
    if (status === 'rejected') over.excluded = true;
    STATE.txOverrides[tx.transaction_id] = over;
    if (status !== 'rejected') {
      addLocalSuggestion(form.suggested_category, 'category', { source: 'manual transaction review', review_item_id: i.review_item_id, reviewer });
      addLocalSuggestion(form.report_group, 'report_group', { source: 'manual transaction review', review_item_id: i.review_item_id, reviewer });
    }
    if (document.getElementById('createMappingRule')?.checked && status !== 'rejected') {
      STATE.mappingRules.push({ mapping_rule_id: `LOCAL-MAP-${Date.now()}`, period: 'ALL', direction: form.direction, raw_category: form.raw_category, suggested_category: form.suggested_category, report_group: form.report_group, created_from_review_item_id: i.review_item_id, reviewer, created_at: now() });
    }
  } else if (form?.kind === 'control') {
    const control = { ...form }; delete control.kind; delete control.adjustment;
    control.updated_at = now(); control.reviewer = reviewer;
    STATE.controlOverrides[i.review_item_id] = control;
    if (status !== 'rejected') {
      addLocalSuggestion(form.affected_report_line, 'category', { source: 'validation/control review', review_item_id: i.review_item_id, reviewer });
      addLocalSuggestion(form.adjustment?.suggested_category, 'category', { source: 'manual adjustment review', review_item_id: i.review_item_id, reviewer });
      addLocalSuggestion(form.adjustment?.report_group, 'report_group', { source: 'manual adjustment review', review_item_id: i.review_item_id, reviewer });
    }
    if (form.create_adjustment && Number(form.adjustment.amount_rsd_equivalent || 0) !== 0 && !['rejected', 'escalated', 'needs_owner_clarification'].includes(status)) {
      STATE.manualAdjustments[i.review_item_id] = { transaction_id: `ADJ-${i.review_item_id}`, review_item_id: i.review_item_id, period: i.period, date: form.adjustment.date, direction: form.adjustment.direction, payment_method: 'manual_adjustment', currency_original: 'RSD', amount_rsd_equivalent: Number(form.adjustment.amount_rsd_equivalent || 0), raw_category: 'Manual control adjustment', suggested_category: form.adjustment.suggested_category, report_group: form.adjustment.report_group, counterparty_type: 'internal_review', counterparty: 'CONTROL_REVIEW', description: form.adjustment.description || `Manual adjustment from ${i.review_item_id}`, source_reference: i.source_reference || 'Manual control review', review_status: 'Manual adjustment - local review' };
    } else {
      delete STATE.manualAdjustments[i.review_item_id];
    }
  }

  if (status === 'needs_owner_clarification') {
    STATE.clarificationTasks[i.review_item_id] = { task_id: `TASK-${i.review_item_id}`, period: i.period, review_item_id: i.review_item_id, status: 'open', question: ownerQuestion || i.suggested_action || i.description, asked_by: reviewer, asked_at: now(), owner_answer: taskFor(i.review_item_id)?.owner_answer || '' };
  } else if (status === 'owner_reviewed') {
    const existing = taskFor(i.review_item_id) || { task_id: `TASK-${i.review_item_id}`, period: i.period, review_item_id: i.review_item_id };
    STATE.clarificationTasks[i.review_item_id] = { ...existing, status: 'answered', owner_answer: note, answered_by: reviewer, answered_at: now() };
  } else if (['management_ready', 'accountant_certified', 'rejected'].includes(status) && STATE.clarificationTasks[i.review_item_id]) {
    STATE.clarificationTasks[i.review_item_id].status = 'closed';
    STATE.clarificationTasks[i.review_item_id].closed_by = reviewer;
    STATE.clarificationTasks[i.review_item_id].closed_at = now();
  }

  STATE.decisions[i.review_item_id] = {
    review_item_id: i.review_item_id,
    period: i.period,
    status,
    resolution_type: resolutionType(status, form),
    role: role(),
    review_area: generalIssueType(i),
    reviewer,
    note,
    owner_question: ownerQuestion,
    decided_at: now(),
    before,
    after: form || { review_item_status: status },
    source_reference: i.source_reference,
    related_object_id: i.related_object_id,
    management_analysis_status: ['owner_reviewed', 'management_ready', 'needs_accountant_review', 'accountant_certified'].includes(status) ? 'usable_with_limitations' : status === 'rejected' ? 'excluded' : 'not_ready',
    accountant_review_status: status === 'accountant_certified' ? 'certified' : ['needs_accountant_review', 'owner_reviewed', 'management_ready'].includes(status) ? 'pending_or_reviewed' : status,
    formal_reporting_status: status === 'accountant_certified' ? 'accountant_certified' : status === 'rejected' ? 'excluded' : 'not_certified'
  };
  audit('review_decision_saved', { review_item_id: i.review_item_id, period: i.period, status, role: role(), reviewer });
  const ui = { ...REVIEW_UI };
  saveState();
  renderAll();
  restoreReviewScroll();
  setActiveSection(role() === 'owner' ? 'shared' : 'review');
  if (next) { const n = nextMatchingReviewItem(i.review_item_id); n ? openModal(n.review_item_id) : closeModal(); } else closeModal();
}

function markMonthManagementReady() {
  if (selected === 'ALL') return alert('Choose one month first.');
  const s = stats(selected);
  if (s.management_blocking > 0) return alert(`${s.management_blocking} management blocker(s) remain. Resolve or send them to the correct workflow first.`);
  STATE.batchStatuses[selected] = { ...(STATE.batchStatuses[selected] || {}), management_ready: true, management_ready_by: 'Accountant Reviewer', management_ready_at: now(), local_quality_score: localQuality(selected) };
  audit('month_marked_management_ready', { period: selected });
  saveState(); renderAll(); alert(`${month().label} marked management-ready.`);
}
function certifyMonthAccountant() {
  if (selected === 'ALL') return alert('Choose one month first.');
  const s = stats(selected);
  if (s.formal_blocking > 0) return alert(`${s.formal_blocking} item(s) still need accountant certification or rejection before formal certification.`);
  STATE.batchStatuses[selected] = { ...(STATE.batchStatuses[selected] || {}), management_ready: true, management_ready_by: STATE.batchStatuses[selected]?.management_ready_by || 'Accountant Reviewer', management_ready_at: STATE.batchStatuses[selected]?.management_ready_at || now(), accountant_certified: true, accountant_certified_by: 'Accountant Reviewer', accountant_certified_at: now(), local_quality_score: localQuality(selected) };
  audit('month_accountant_certified', { period: selected });
  saveState(); renderAll(); alert(`${month().label} accountant-certified locally.`);
}

function exportObj() {
  return { exported_at: now(), app: 'Ask Your Business Certified Data View Demo', version: APP_VERSION, warning: 'Browser-local demo export. Move to PostgreSQL in Option 2 for persistent storage.', company_code: DATA.company_code, selected_scope: selected, review_state: STATE, months: DATA.months.map(m => ({ period: m.period, label: m.label, management_status: managementStatus(m.period), formal_status: formalStatus(m.period), adjusted_summary: adjSummary(m.period), certified_summary: certifiedSummary(m.period), review_stats: stats(m.period), local_quality_score: localQuality(m.period), batch_status: STATE.batchStatuses[m.period] || null, transactions: displayTxRows(m.period).map(t => effTx(t)), management_ready_transactions: certifiedSummary(m.period).managementRows, accountant_certified_transactions: certifiedSummary(m.period).accountingRows, rejected_rows: certifiedSummary(m.period).rejectedRows, pending_rows: certifiedSummary(m.period).pendingRows, unresolved_review_items: m.manual_review_queue.filter(i => blocksManagement(i) || blocksFormal(i)), review_decisions: Object.values(STATE.decisions).filter(d => d.period === m.period), clarification_tasks: Object.values(STATE.clarificationTasks).filter(t => t.period === m.period), control_resolutions: Object.entries(STATE.controlOverrides).filter(([id, c]) => itemById(id)?.period === m.period).map(([review_item_id, c]) => ({ review_item_id, ...c })), manual_adjustments: Object.values(STATE.manualAdjustments).filter(a => a.period === m.period) })) };
}
function managementAnalysisPackObj() {
  return { exported_at: now(), pack_type: 'Management Analysis Pack', app: 'Ask Your Business Certified Data View Demo', version: APP_VERSION, company_code: DATA.company_code, selected_scope: selected, purpose: 'Owner/management-ready dataset with limitations. Not a formal accountant-certified report unless the period status says accountant-certified.', scope_status: { management: managementStatus(), formal: formalStatus(), limitations: certificationLimitations().limits }, months: scopeMonths().map(m => ({ period: m.period, label: m.label, management_status: managementStatus(m.period), formal_status: formalStatus(m.period), limitations: certificationLimitations(m.period).limits, summary: certifiedSummary(m.period).management, category_breakdown: categoryBreakdownRows(certifiedSummary(m.period).managementRows, 999), transactions: certifiedSummary(m.period).managementRows, manual_adjustments: certifiedSummary(m.period).adjustments, rejected_rows: certifiedSummary(m.period).rejectedRows })) };
}
function accountantPackObj() {
  return { exported_at: now(), pack_type: 'Accountant Operations Review Pack', version: APP_VERSION, company_code: DATA.company_code, selected_scope: selected, purpose: 'Send to accountant/internal finance team for unresolved certification, owner clarifications, validation-control decisions, and formal reporting sign-off.', months: scopeMonths().map(m => ({ period: m.period, label: m.label, management_status: managementStatus(m.period), formal_status: formalStatus(m.period), pending_accountant_items: m.manual_review_queue.filter(needsAccountantWork), owner_clarification_tasks: Object.values(STATE.clarificationTasks).filter(t => t.period === m.period), validation_results: m.validation_results, owner_reviewed_decisions: Object.values(STATE.decisions).filter(d => d.period === m.period && ['owner_reviewed', 'needs_accountant_review'].includes(d.status)), mapping_rules: STATE.mappingRules.filter(r => !r.period || r.period === 'ALL' || r.period === m.period), manual_adjustments: Object.values(STATE.manualAdjustments).filter(a => a.period === m.period) })) };
}
function download(name, obj) { const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' }), url = URL.createObjectURL(blob), a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }
function exportPackage() { download(`certified_import_package_${new Date().toISOString().slice(0, 10)}.json`, exportObj()); }
function exportAccountantPack() { download(`accountant_review_pack_${new Date().toISOString().slice(0, 10)}.json`, accountantPackObj()); }
function exportManagementAnalysisPack() { download(`management_analysis_pack_${new Date().toISOString().slice(0, 10)}.json`, managementAnalysisPackObj()); }
function exportState() { download(`review_decisions_${new Date().toISOString().slice(0, 10)}.json`, STATE); }
function resetProgress() { if (confirm('Reset all local decisions, edits, clarification tasks, and batch statuses in this browser?')) { STATE = blankState(); saveState(); renderAll(); } }

window.openLink = openLink;
window.bulkMarkLowRiskManagementReady = bulkMarkLowRiskManagementReady;
window.sendVisibleOwnerQuestions = sendVisibleOwnerQuestions;
window.setRole = setRole;
window.setActiveSection = setActiveSection;
window.openModal = openModal;
window.closeModal = closeModal;
window.saveDecision = saveDecision;
window.exportPackage = exportPackage;
window.exportAccountantPack = exportAccountantPack;
window.exportManagementAnalysisPack = exportManagementAnalysisPack;
window.exportState = exportState;
window.resetProgress = resetProgress;
window.markMonthManagementReady = markMonthManagementReady;
window.certifyMonthAccountant = certifyMonthAccountant;
window.ask = ask;
window.renderReviewTable = renderReviewTable;
window.renderTxTable = renderTxTable;

document.querySelectorAll('.tab').forEach(b => b.onclick = () => setActiveSection(b.dataset.section));
Promise.all([
  fetch('/data/university_mar_apr_may_import_data.json').then(r => r.json()),
  fetch('/data/source_evidence_lookup.json').then(r => r.ok ? r.json() : {}).catch(() => ({}))
])
  .then(([d, evidence]) => {
    DATA = d;
    SOURCE_EVIDENCE = evidence || {};
    loadState();
    const sel = document.getElementById('monthSelect');
    sel.innerHTML = '<option value="ALL">All months combined</option>' + d.months.map(m => `<option value="${m.period}">${m.label}</option>`).join('');
    sel.onchange = () => { selected = sel.value; renderAll(); };
    renderAll();
    setActiveSection(role() === 'owner' ? 'owner' : 'accountant');
  })
  .catch(e => { document.getElementById('owner').innerHTML = `<div class="card"><h2>Data load error</h2><p>${esc(e.message)}</p></div>`; });
