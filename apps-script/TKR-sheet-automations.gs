/**
 * TKR agent - sheet automations. No AI, no Claude credits.
 *
 * Install (once):
 *   1. Open the sheet "TKR Recruitment Agent - Data".
 *   2. Extensions > Apps Script. Add a new script file (+ > Script) called "TKR automations", paste this whole file, Save.
 *      Every name in this file starts with tkr or TKR_, so it cannot clash with other scripts in the project (such as Adzuna).
 *   3. Pick "tkrSetup" in the function list and press Run. Allow the permissions it asks for
 *      (this sheet, and read-only Gmail for LinkedIn "accepted your invitation" emails).
 *   4. Reload the sheet. A "TKR" menu appears with "Run sheet automations now".
 *
 * What it does every 15 minutes (each step only writes when something has changed):
 *   1. LinkedIn Queue: fills Sent date, Accepted date or Withdrawn date where Status says so and the date is blank.
 *   2. Gmail: reads LinkedIn "accepted your invitation" emails from the last 7 days and sets the matching
 *      Queue row (Status Sent or Queued) to Accepted with the email date. Matches on the profile link first,
 *      then on first and last name. Two or more possible matches: writes nothing and logs it.
 *   3. LinkedIn Queue: rows still Sent 35 or more days after the Sent date get "Withdraw" added to Notes.
 *      Nothing is ever withdrawn automatically.
 *   4. Pipeline: copies each TKR agent Queue row's status onto its Pipeline row (LinkedIn status and date).
 *      Sent > Pending, Accepted > Accepted, Withdrawn > Withdrawn, Already connected > Already connected,
 *      Error > Not found. Only when the company on both rows matches, and never steps a status backwards.
 *   5. Pipeline: firms you approved (Status "approved", Stage "1 Qualified") whose "Cleared to run" says Yes
 *      move to Stage "2 In sequence", Sequence day 1.
 *   6. Pipeline: Sequence day for firms this script moved into the sequence = working days since they entered + 1.
 *
 * Everything it changes is listed on a "Script log" tab. It never writes into a cell that holds a formula.
 * Columns are found by header name, never by letter.
 */

const TKR_TZ = 'Europe/London';
const TKR_QUEUE = 'LinkedIn Queue';
const TKR_PIPE = 'Pipeline';
const TKR_LOG = 'Script log';
const TKR_PROJECT_NAMES = ['recruitment agent', 'tkr agent'];
const TKR_WITHDRAW_DAYS = 35;

function tkrSetup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'tkrRunAll')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tkrRunAll').timeBased().everyMinutes(15).create();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tkrOnOpen').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tkrOnOpen').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  tkr_log('tkrSetup', 'Installed: tkrRunAll every 15 minutes');
  tkrRunAll();
}

function tkrOnOpen() {
  SpreadsheetApp.getUi().createMenu('TKR')
    .addItem('Run sheet automations now', 'tkrRunAll')
    .addItem('Install 15-minute trigger', 'tkrSetup')
    .addToUi();
}

function tkrRunAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = SpreadsheetApp.getActive();
    const notes = [
      tkr_safe('queue dates', () => tkr_stampQueueDates(ss)),
      tkr_safe('accept emails', () => tkr_readAcceptEmails(ss)),
      tkr_safe('35-day flag', () => tkr_flagWithdrawals(ss)),
      tkr_safe('queue to pipeline', () => tkr_syncQueueToPipeline(ss)),
      tkr_safe('into sequence', () => tkr_moveApprovedIntoSequence(ss)),
      tkr_safe('sequence day', () => tkr_advanceSequenceDay(ss)),
    ].filter(Boolean);
    const text = notes.join(' | ');
    const props = PropertiesService.getScriptProperties();
    if (text && text !== props.getProperty('lastLog')) tkr_log('run', text);   // a repeat of the last message is not logged again
    props.setProperty('lastLog', text);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- steps ---------- */

function tkr_stampQueueDates(ss) {
  const q = tkr_table(ss, TKR_QUEUE);
  const cS = q.col('status'), cSent = q.col('sent date'), cAcc = q.col('accepted date'), cW = q.col('withdrawn date');
  const t = tkr_today();
  let n = 0;
  for (let r = 1; r < q.disp.length; r++) {
    const s = tkr_low(q.disp[r][cS]);
    if (s === 'sent' && !q.disp[r][cSent] && q.set(r, cSent, t)) n++;
    if (s === 'accepted' && !q.disp[r][cAcc] && q.set(r, cAcc, t)) n++;
    if (s === 'withdrawn' && !q.disp[r][cW] && q.set(r, cW, t)) n++;
  }
  return n ? n + ' Queue dates filled' : '';
}

function tkr_readAcceptEmails(ss) {
  const threads = GmailApp.search('from:invitations@linkedin.com subject:"accepted your invitation" newer_than:7d', 0, 100);
  if (!threads.length) return '';
  const q = tkr_table(ss, TKR_QUEUE);
  const cN = q.col('name'), cS = q.col('status'), cAcc = q.col('accepted date'), cP = q.col('profile url');
  let n = 0;
  const unsure = [];
  threads.forEach(th => th.getMessages().forEach(m => {
    if (!/accepted your invitation/i.test(m.getSubject())) return;
    const body = m.getBody();
    const nameM = body.match(/See ([^<>]{2,80}?)(?:’|&#8217;|&rsquo;|')s connections/);
    const slugM = body.match(/linkedin\.com\/(?:comm\/)?in\/([A-Za-z0-9\-_%]+)/);
    const fullName = nameM ? nameM[1].trim() : '';
    const slug = slugM ? slugM[1].toLowerCase() : '';
    const open = (r) => /^(sent|queued)$/.test(tkr_low(q.disp[r][cS]));
    let hits = [];
    if (slug) for (let r = 1; r < q.disp.length; r++) if (open(r) && tkr_low(q.disp[r][cP]).indexOf('/in/' + slug) >= 0) hits.push(r);
    if (!hits.length && fullName) for (let r = 1; r < q.disp.length; r++) if (open(r) && tkr_nameKey(q.disp[r][cN]) === tkr_nameKey(fullName)) hits.push(r);
    if (hits.length === 1) {
      const r = hits[0];
      if (q.set(r, cS, 'Accepted')) { q.set(r, cAcc, tkr_ymd(m.getDate())); n++; }
    } else if (hits.length > 1) {
      unsure.push((fullName || slug) + ' (' + hits.length + ' rows)');
    }
  }));
  let out = n ? n + ' acceptances recorded from LinkedIn emails' : '';
  if (unsure.length) out += (out ? '; ' : '') + 'not recorded, more than one match: ' + unsure.join(', ');
  return out;
}

function tkr_flagWithdrawals(ss) {
  const q = tkr_table(ss, TKR_QUEUE);
  const cS = q.col('status'), cSent = q.col('sent date'), cNotes = q.col('notes');
  const now = tkr_toDate(tkr_today());
  let n = 0;
  for (let r = 1; r < q.disp.length; r++) {
    if (tkr_low(q.disp[r][cS]) !== 'sent') continue;
    const sent = tkr_toDate(q.vals[r][cSent]) || tkr_toDate(q.disp[r][cSent]);
    if (!sent || tkr_daysBetween(sent, now) < TKR_WITHDRAW_DAYS) continue;
    const notes = String(q.disp[r][cNotes] || '');
    if (/withdraw\b/i.test(notes) && !/consider withdraw/i.test(notes)) continue;
    if (q.set(r, cNotes, notes ? notes + '; Withdraw' : 'Withdraw')) n++;
  }
  return n ? n + ' invites flagged to withdraw (35+ days)' : '';
}

const TKR_LI_MAP = { 'queued': 'Queued', 'sent': 'Pending', 'accepted': 'Accepted', 'withdrawn': 'Withdrawn', 'already connected': 'Already connected', 'error': 'Not found' };
const TKR_LI_RANK = { '': 0, 'not invited': 0, 'queued': 1, 'pending': 2, 'not found': 2, 'withdrawn': 3, 'accepted': 4, 'already connected': 4 };

function tkr_syncQueueToPipeline(ss) {
  const q = tkr_table(ss, TKR_QUEUE), p = tkr_table(ss, TKR_PIPE);
  const qProj = q.col('project'), qCo = q.col('company'), qS = q.col('status'), qSent = q.col('sent date'),
        qAcc = q.col('accepted date'), qW = q.col('withdrawn date'), qRow = q.col('pipeline row');
  const pCo = p.col('company'), pLS = p.col('linkedin status'), pLD = p.col('linkedin date');
  let n = 0;
  const mismatch = [];
  for (let r = 1; r < q.disp.length; r++) {
    if (TKR_PROJECT_NAMES.indexOf(tkr_low(q.disp[r][qProj])) < 0) continue;
    const status = tkr_low(q.disp[r][qS]);
    const target = TKR_LI_MAP[status];
    const pr = parseInt(q.disp[r][qRow], 10) - 1;
    if (!target || !(pr >= 1) || pr >= p.disp.length) continue;
    if (tkr_norm(p.disp[pr][pCo]) !== tkr_norm(q.disp[r][qCo])) { mismatch.push(q.disp[r][qCo] + ' (Queue row ' + (r + 1) + ')'); continue; }
    const cur = tkr_low(p.disp[pr][pLS]);
    if (cur === tkr_low(target)) continue;
    if ((TKR_LI_RANK[tkr_low(target)] || 0) < (TKR_LI_RANK[cur] || 0)) continue;
    const date = status === 'sent' ? q.disp[r][qSent] : status === 'accepted' ? q.disp[r][qAcc] : status === 'withdrawn' ? q.disp[r][qW] : tkr_today();
    if (p.set(pr, pLS, target)) { if (date) p.set(pr, pLD, date); n++; }
  }
  let out = n ? n + ' Pipeline LinkedIn statuses updated' : '';
  if (mismatch.length) out += (out ? '; ' : '') + 'company does not match its Pipeline row: ' + mismatch.slice(0, 10).join(', ');
  return out;
}

function tkr_moveApprovedIntoSequence(ss) {
  const p = tkr_table(ss, TKR_PIPE);
  const cCo = p.col('company'), cS = p.col('status'), cSt = p.col('stage'), cDay = p.col('sequence day');
  const cClr = p.colMaybe('cleared to run');
  const props = PropertiesService.getScriptProperties();
  const moved = [];
  for (let r = 1; r < p.disp.length; r++) {
    if (tkr_low(p.disp[r][cS]) !== 'approved' || !/^1 qualified/.test(tkr_low(p.disp[r][cSt]))) continue;
    if (cClr >= 0 && !/^yes/.test(tkr_low(p.disp[r][cClr]))) continue;
    if (p.set(r, cSt, '2 In sequence')) {
      p.set(r, cDay, 1);
      props.setProperty('seq:' + tkr_norm(p.disp[r][cCo]), tkr_today());
      moved.push(p.disp[r][cCo]);
    }
  }
  return moved.length ? moved.length + ' approved firms moved into the sequence: ' + moved.join(', ') : '';
}

function tkr_advanceSequenceDay(ss) {
  const p = tkr_table(ss, TKR_PIPE);
  const cCo = p.col('company'), cSt = p.col('stage'), cDay = p.col('sequence day');
  const props = PropertiesService.getScriptProperties();
  const now = tkr_toDate(tkr_today());
  let n = 0;
  for (let r = 1; r < p.disp.length; r++) {
    if (!/^2 in sequence/.test(tkr_low(p.disp[r][cSt]))) continue;
    const start = props.getProperty('seq:' + tkr_norm(p.disp[r][cCo]));
    if (!start) continue;
    const day = tkr_workingDays(tkr_toDate(start), now) + 1;
    if (String(p.disp[r][cDay]) !== String(day) && p.set(r, cDay, day)) n++;
  }
  return n ? n + ' sequence days moved on' : '';
}

/* ---------- helpers ---------- */

function tkr_table(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Tab not found: ' + name);
  const rng = sh.getDataRange();
  const disp = rng.getDisplayValues(), vals = rng.getValues(), forms = rng.getFormulas();
  const head = disp[0].map(h => tkr_low(h));
  const colMaybe = (start) => head.findIndex(h => h.indexOf(start) === 0);
  const col = (start) => { const i = colMaybe(start); if (i < 0) throw new Error(name + ': no column starting "' + start + '"'); return i; };
  const set = (r, c, v) => {
    if (forms[r] && forms[r][c]) { tkr_log('skipped', name + ' row ' + (r + 1) + ', column "' + disp[0][c] + '" holds a formula'); return false; }
    const cell = sh.getRange(r + 1, c + 1);
    cell.setValue(v);
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(v))) cell.setNumberFormat('yyyy-mm-dd');   // keep dates as 2026-11-22, which the dashboard reads
    disp[r][c] = String(v); vals[r][c] = v;
    return true;
  };
  return { sh, disp, vals, col, colMaybe, set };
}

function tkr_safe(step, fn) {
  try { return fn(); } catch (e) { tkr_log('error', step + ': ' + (e && e.message || e)); return ''; }
}

function tkr_log(step, detail) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TKR_LOG);
  if (!sh) { sh = ss.insertSheet(TKR_LOG); sh.appendRow(['When', 'Step', 'Detail']); sh.setFrozenRows(1); }
  sh.appendRow([Utilities.formatDate(new Date(), TKR_TZ, 'yyyy-MM-dd HH:mm'), step, String(detail).slice(0, 4000)]);
}

function tkr_today() { return Utilities.formatDate(new Date(), TKR_TZ, 'yyyy-MM-dd'); }
function tkr_ymd(d) { return Utilities.formatDate(d, TKR_TZ, 'yyyy-MM-dd'); }
function tkr_low(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

function tkr_toDate(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const u = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);   // UK display format dd/mm/yyyy
  return u ? new Date(+u[3], +u[2] - 1, +u[1]) : null;
}

function tkr_daysBetween(a, b) { return Math.round((b - a) / 86400000); }

function tkr_workingDays(a, b) {
  let n = 0;
  const d = new Date(a);
  while (d < b) { d.setDate(d.getDate() + 1); const w = d.getDay(); if (w !== 0 && w !== 6) n++; }
  return n;
}

function tkr_norm(s) {
  return tkr_low(s).replace(/\(.*?\)/g, ' ').replace(/&/g, ' ')
    .replace(/\b(ltd|limited|plc|llp|the|and|co)\b/g, ' ').replace(/[^a-z0-9]+/g, '');
}

function tkr_nameKey(s) {
  const w = tkr_low(s).replace(/[^a-z\s'\-]/g, ' ').split(/\s+/).filter(Boolean);
  return w.length ? w[0] + ' ' + w[w.length - 1] : '';
}
