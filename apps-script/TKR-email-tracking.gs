/**
 * TKR agent - email tracking. Second file in the same Apps Script project as TKR-sheet-automations.gs.
 * No AI, no Claude credits. All names start with tkr or TKR_ so they cannot clash with other scripts.
 *
 * What it does
 *   1. tkrEmailHit (the web app): the Netlify pixel calls it each time a tracked email's image loads. It writes a line
 *      to the "Email events" tab, marking whether the open looks real (see tkr_judgeOpen).
 *   2. tkrEmailStats (every 15 minutes, called from tkrRunAll in the other file): for each row on "Email log" it
 *        - counts real opens, first and last open;
 *        - reads the Gmail thread for replies and bounces;
 *        - writes the result back to that row, and to the Pipeline row's Email status and Email date.
 *      Pipeline Email status only ever moves forward: Sent > Opened 2+ > Replied, and Bounced.
 *
 * Install (once, on a laptop):
 *   1. In the TKR Apps Script project add a second file called "TKR email tracking" and paste this in. Save.
 *   2. Project Settings (cog) > Script properties > add TKR_SECRET = a long random string (keep a copy, Netlify needs it).
 *   3. Deploy > New deployment > type Web app. Execute as: Me. Who has access: Anyone. Deploy and copy the /exec URL.
 *      (Anyone can open the URL, but every request must carry TKR_SECRET or it is ignored.)
 *   4. Netlify site environment variables: TKR_SCRIPT_URL = that /exec URL, TKR_SECRET = the same secret.
 *   5. In tkrRunAll (the other file) add the line:  tkr_safe('email stats', () => tkrEmailStats(ss)),  to the list of steps,
 *      or simply run tkrEmailStatsNow once from the editor to test it.
 */

const TKR_EMAIL_LOG = 'Email log';
const TKR_EMAIL_EVENTS = 'Email events';
const TKR_MIN_SECONDS_FOR_REAL_OPEN = 60;     // anything faster than this after sending is a scanner or pre-fetch
const TKR_SECOND_OPEN_GAP_MINUTES = 30;       // "Opened 2+" needs two real opens at least this far apart
const TKR_BOT_AGENTS = /(barracuda|mimecast|proofpoint|symantec|messagelabs|forcepoint|trendmicro|trend micro|sophos|fireeye|mcafee|ironport|cisco|zscaler|microsoft office|outlook-ios|bingpreview|urlscan|python-requests|curl\/|headlesschrome|go-http-client|java\/|okhttp|wget)/i;

/* ---------- 1. the web app the pixel calls ---------- */

function doGet(e) { return tkrEmailHit(e); }

function tkrEmailHit(e) {
  const p = (e && e.parameter) || {};
  const secret = PropertiesService.getScriptProperties().getProperty('TKR_SECRET');
  if (!secret || p.k !== secret) return ContentService.createTextOutput('no');
  const id = String(p.id || '').slice(0, 40);
  if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) return ContentService.createTextOutput('no');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return ContentService.createTextOutput('busy');
  try {
    const ss = tkr_ss();
    const logRow = tkr_findEmailRow_(ss, id);
    if (!logRow) return ContentService.createTextOutput('unknown');
    const sentAt = logRow.sentAt;
    const secs = sentAt ? Math.round((Date.now() - sentAt.getTime()) / 1000) : '';
    const verdict = tkr_judgeOpen(String(p.ua || ''), secs);
    const sh = ss.getSheetByName(TKR_EMAIL_EVENTS);
    sh.appendRow([new Date(), id, p.e || 'open', verdict.real ? 'Yes' : 'No', verdict.why, String(p.ua || '').slice(0, 200), secs]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

function tkr_judgeOpen(ua, secs) {
  if (secs !== '' && secs < TKR_MIN_SECONDS_FOR_REAL_OPEN) return { real: false, why: 'within ' + TKR_MIN_SECONDS_FOR_REAL_OPEN + 's of sending (scanner or pre-fetch)' };
  if (TKR_BOT_AGENTS.test(ua)) return { real: false, why: 'security scanner or tool' };
  if (!ua) return { real: false, why: 'no user agent' };
  if (/GoogleImageProxy/i.test(ua)) return { real: true, why: 'Gmail (image proxy, repeat opens may be cached)' };
  return { real: true, why: 'looks like a mail app' };
}

/* ---------- 2. stats, replies, bounces, Pipeline ---------- */

function tkrEmailStatsNow() { tkrEmailStats(tkr_ss()); }

function tkrEmailStats(ss) {
  const log = tkr_table(ss, TKR_EMAIL_LOG);
  if (log.disp.length < 2) return '';
  const ev = ss.getSheetByName(TKR_EMAIL_EVENTS);
  const evVals = ev.getLastRow() > 1 ? ev.getRange(2, 1, ev.getLastRow() - 1, 7).getValues() : [];
  const opens = {};
  evVals.forEach(r => {
    if (r[3] !== 'Yes') return;
    (opens[r[1]] = opens[r[1]] || []).push(new Date(r[0]));
  });
  const c = {
    id: log.col('email id'), sent: log.col('sent at'), to: log.col('to'), co: log.col('company'), prow: log.col('pipeline row'),
    thread: log.col('thread id'), op: log.col('opens'), first: log.col('first open'), last: log.col('last open'),
    rep: log.col('replied'), repd: log.col('reply date'), bnc: log.col('bounced'), st: log.col('status'),
  };
  const pipe = tkr_table(ss, 'Pipeline');
  const pCo = pipe.col('company'), pES = pipe.col('email status'), pED = pipe.col('email date');
  const RANK = { '': 0, 'not sent': 0, 'no email': 0, 'sent': 1, 'opened 2+': 2, 'replied': 4, 'bounced': 3, 'unsubscribed': 5 };
  const changes = [];
  for (let r = 1; r < log.disp.length; r++) {
    const id = String(log.disp[r][c.id] || '').trim();
    if (!id) continue;
    const times = (opens[id] || []).sort((a, b) => a - b);
    const first = times[0] || '', last = times[times.length - 1] || '';
    const twoOpens = times.length >= 2 && (last - first) >= TKR_SECOND_OPEN_GAP_MINUTES * 60000;
    let replied = String(log.disp[r][c.rep] || ''), replyDate = String(log.disp[r][c.repd] || ''), bounced = String(log.disp[r][c.bnc] || '');
    const threadId = String(log.disp[r][c.thread] || '').trim();
    if (threadId && !/^yes/i.test(replied) && !/^yes/i.test(bounced)) {
      try {
        const th = GmailApp.getThreadById(threadId);
        if (th) {
          const me = Session.getEffectiveUser().getEmail().toLowerCase();
          th.getMessages().forEach(m => {
            const from = m.getFrom().toLowerCase();
            if (/mailer-daemon|postmaster/.test(from)) bounced = 'Yes';
            else if (from.indexOf(me) < 0 && !/^yes/i.test(replied)) { replied = 'Yes'; replyDate = tkr_ymd(m.getDate()); }
          });
        }
      } catch (e) { /* thread not readable: leave as is */ }
    }
    const status = /^yes/i.test(bounced) ? 'Bounced' : /^yes/i.test(replied) ? 'Replied' : twoOpens ? 'Opened 2+' : 'Sent';
    log.set(r, c.op, times.length);
    if (first) log.set(r, c.first, Utilities.formatDate(first, TKR_TZ, 'yyyy-MM-dd HH:mm'));
    if (last) log.set(r, c.last, Utilities.formatDate(last, TKR_TZ, 'yyyy-MM-dd HH:mm'));
    if (replied) { log.set(r, c.rep, replied); if (replyDate) log.set(r, c.repd, replyDate); }
    if (bounced) log.set(r, c.bnc, bounced);
    if (String(log.disp[r][c.st]) !== status) log.set(r, c.st, status);
    // Pipeline (only rows whose company matches, and only ever forwards)
    const pr = parseInt(log.disp[r][c.prow], 10) - 1;
    if (pr >= 1 && pr < pipe.disp.length && tkr_norm(pipe.disp[pr][pCo]) === tkr_norm(log.disp[r][c.co])) {
      const cur = tkr_low(pipe.disp[pr][pES]);
      if ((RANK[tkr_low(status)] || 0) > (RANK[cur] || 0)) {
        if (pipe.set(pr, pES, status)) {
          const dt = status === 'Replied' && replyDate ? replyDate : tkr_today();
          pipe.set(pr, pED, dt);
          changes.push(log.disp[r][c.co] + ' > ' + status);
        }
      }
    }
  }
  return changes.length ? 'Email status: ' + changes.join(', ') : '';
}

function tkr_findEmailRow_(ss, id) {
  const sh = ss.getSheetByName(TKR_EMAIL_LOG);
  if (!sh || sh.getLastRow() < 2) return null;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const ci = head.indexOf('email id'), cs = head.indexOf('sent at');
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][ci]).trim() === id) {
      const s = vals[i][cs];
      const d = s instanceof Date ? s : new Date(String(s));
      return { row: i + 2, sentAt: isNaN(d) ? null : d };
    }
  }
  return null;
}
