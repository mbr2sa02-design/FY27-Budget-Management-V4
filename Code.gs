/**
 * MBR2 Scope Tracker - shared backend (Google Apps Script)
 * Stores team entries in the Google Sheet this script is attached to.
 * Sheets: Users (name, pin, role) | Data (k, v, u, t) | Log (t, u, k, v)
 */
const STAFF = ['Soy Bonthoeurt','Hien Piphup','Leap Sophy','Loum Sopheara','Leak Lida','Phorn Sokchea'];

// Run this ONCE from the editor. It creates the sheets and gives everyone a random 4-digit PIN.
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (n, h) => { let s = ss.getSheetByName(n) || ss.insertSheet(n); if (s.getLastRow() === 0) s.appendRow(h); return s; };
  const u = mk('Users', ['name', 'pin', 'role']);
  mk('Data', ['k', 'v', 'u', 't']);
  mk('Log', ['t', 'u', 'k', 'v']);
  u.getRange('B:B').setNumberFormat('@');
  if (u.getLastRow() === 1) {
    const pin = () => String(Math.floor(1000 + Math.random() * 9000));
    STAFF.forEach(n => u.appendRow([n, pin(), 'staff']));
    u.appendRow(['Manager', pin(), 'manager']);
  }
}

const sh = n => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const rows = n => { const v = sh(n).getDataRange().getValues(); v.shift(); return v; };

function auth(name, pin) {
  const r = rows('Users').find(x => String(x[0]) === String(name) && String(x[1]) === String(pin));
  return r ? { name: String(r[0]), role: String(r[2]) } : null;
}
const parse = s => { try { return JSON.parse(s); } catch (e) { return null; } };
const rec = r => ({ k: r[0], v: parse(r[1]), u: r[2], t: r[3] });

function doGet(e) {
  const p = e.parameter || {}, user = auth(p.name, p.pin);
  if (!user) return out({ ok: false });
  const recs = rows('Data').map(rec).sort((a, b) => (a.t > b.t ? 1 : -1));
  const lg = sh('Log'), n = lg.getLastRow() - 1, take = Math.min(12, Math.max(n, 0));
  const recent = take ? lg.getRange(lg.getLastRow() - take + 1, 1, take, 4).getValues()
    .reverse().map(r => ({ t: r[0], u: r[1], k: r[2], v: parse(r[3]) })) : [];
  return out({ ok: true, role: user.role, recs, recent });
}

function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Bad request' }); }
  const user = auth(b.name, b.pin);
  if (!user) return out({ ok: false, error: 'Wrong name or PIN' });
  if (b.a === 'login') return out({ ok: true, role: user.role });
  if (b.a !== 'set') return out({ ok: false, error: 'Unknown action' });

  const key = parse(b.k);
  if (!Array.isArray(key) || key.length < 2) return out({ ok: false, error: 'Bad key' });
  if (user.role !== 'manager' && key[1] !== user.name) return out({ ok: false, error: 'You can only edit your own lines' });

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const t = new Date().toISOString(), v = JSON.stringify(b.v === undefined ? null : b.v), data = sh('Data');
    const ks = data.getLastRow() > 1 ? data.getRange(2, 1, data.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    const i = ks.indexOf(b.k);
    if (i >= 0) data.getRange(i + 2, 1, 1, 4).setValues([[b.k, v, user.name, t]]);
    else data.appendRow([b.k, v, user.name, t]);
    sh('Log').appendRow([t, user.name, b.k, v]);
    return out({ ok: true });
  } finally { lock.release(); }
}
