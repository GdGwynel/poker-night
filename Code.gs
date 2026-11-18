/**
 * Poker Night — Google Apps Script backend.
 * Paste this whole file into Extensions → Apps Script of a new Google Sheet.
 * 1. Change ADMIN_PIN below to your own host PIN.
 * 2. Run importPastGames once (optional) to load the earlier games.
 * 3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 * 4. Copy the /exec URL into API_URL at the top of index.html's script.
 */

const ADMIN_PIN = '0000'; // CHANGE THIS before deploying. Only the host should know it.

const POLL_H = ['id', 'created', 'data'];
const VOTE_H = ['pollId', 'key', 'name', 'data', 'updated'];
const SESS_H = ['id', 'date', 'created', 'data'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, headers) {
  let s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.appendRow(headers); s.setFrozenRows(1); }
  return s;
}
function rows_(name, headers) {
  const v = sheet_(name, headers).getDataRange().getValues();
  return v.slice(1).map((r, i) => ({ row: i + 2, vals: r }));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function key_(n) { return clean_(n, 40).toLowerCase().replace(/\s+/g, ' '); }
function money_(x) {
  const n = Number(x);
  if (!isFinite(n) || n < 0 || n > 10000000) throw new Error('Amounts must be positive numbers.');
  return Math.round(n * 100) / 100;
}
function needPin_(b) { if (String(b.pin || '') !== String(ADMIN_PIN)) throw new Error('Wrong host PIN.'); }

function doGet() { return out_({ ok: true, service: 'poker-night' }); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Bad request.' }); }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return out_({ ok: false, error: 'Busy. Try again in a moment.' }); }
  try { return out_(handle_(body || {})); }
  catch (err) { return out_({ ok: false, error: String((err && err.message) || err) }); }
  finally { lock.releaseLock(); }
}

function handle_(b) {
  switch (b.action) {
    case 'state': return state_();
    case 'checkPin': needPin_(b); return { ok: true };
    case 'vote': return vote_(b);
    case 'createPoll': needPin_(b); return createPoll_(b);
    case 'lockPoll': needPin_(b); return setLock_(b, true);
    case 'reopenPoll': needPin_(b); return setLock_(b, false);
    case 'saveSession': needPin_(b); return saveSession_(b);
    case 'deleteSession': needPin_(b); return deleteSession_(b);
    case 'markPaid': return markPaid_(b);
  }
  throw new Error('Unknown action.');
}

function activePoll_() {
  const r = rows_('Polls', POLL_H);
  if (!r.length) return null;
  const last = r[r.length - 1];
  return { row: last.row, poll: JSON.parse(last.vals[2]) };
}

function state_() {
  const ap = activePoll_();
  const poll = ap ? ap.poll : null;
  const votes = poll
    ? rows_('Votes', VOTE_H).filter(r => r.vals[0] === poll.id).map(r => Object.assign({ name: String(r.vals[2]) }, JSON.parse(r.vals[3])))
    : [];
  const sessions = rows_('Sessions', SESS_H).filter(r => r.vals[3]).map(r => JSON.parse(r.vals[3]));
  return { ok: true, poll: poll, votes: votes, sessions: sessions };
}

function vote_(b) {
  const ap = activePoll_();
  if (!ap) throw new Error('There is no open poll.');
  const poll = ap.poll;
  if (b.pollId !== poll.id) throw new Error('The poll changed. Refresh the page.');
  if (poll.locked) throw new Error('The date is already locked in.');
  const name = clean_(b.name, 30);
  if (!name) throw new Error('Enter your name.');
  const slots = {};
  poll.slots.forEach(s => { const v = b.slots && b.slots[s.id]; if (v === 'yes' || v === 'maybe' || v === 'no') slots[s.id] = v; });
  const host = poll.hosts.indexOf(b.host) >= 0 ? b.host : '';
  const data = JSON.stringify({ slots: slots, host: host });
  const k = key_(name);
  const s = sheet_('Votes', VOTE_H);
  const hit = rows_('Votes', VOTE_H).find(r => r.vals[0] === poll.id && r.vals[1] === k);
  if (hit) s.getRange(hit.row, 1, 1, 5).setValues([[poll.id, k, name, data, new Date()]]);
  else s.appendRow([poll.id, k, name, data, new Date()]);
  return state_();
}

function createPoll_(b) {
  const title = clean_(b.title, 80) || 'Poker night';
  const slots = (b.slots || []).slice(0, 12)
    .map(w => clean_(w, 16)).filter(w => /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(w))
    .map((w, i) => ({ id: 's' + (i + 1), when: w }));
  if (!slots.length) throw new Error('Add at least one date.');
  const hosts = (b.hosts || []).map(h => clean_(h, 30)).filter(Boolean).slice(0, 10);
  const poll = { id: 'p' + Date.now(), title: title, slots: slots, hosts: hosts, locked: null, created: new Date().toISOString() };
  sheet_('Polls', POLL_H).appendRow([poll.id, new Date(), JSON.stringify(poll)]);
  return state_();
}

function setLock_(b, on) {
  const ap = activePoll_();
  if (!ap) throw new Error('There is no poll to lock.');
  const p = ap.poll;
  if (on) {
    if (!p.slots.some(s => s.id === b.slotId)) throw new Error('Pick a date to lock in.');
    p.locked = { slotId: b.slotId, host: clean_(b.host, 30) };
  } else p.locked = null;
  sheet_('Polls', POLL_H).getRange(ap.row, 3).setValue(JSON.stringify(p));
  return state_();
}

function saveSession_(b) {
  const s = b.session || {};
  const date = /^\d{4}-\d{2}-\d{2}$/.test(s.date) ? s.date : '';
  const players = (s.players || []).slice(0, 30).map(p => ({
    name: clean_(p.name, 30),
    buyins: (p.buyins || []).slice(0, 40).map(money_),
    winnings: money_(p.winnings || 0),
    prepaid: money_(p.prepaid || 0),
    prepaidTo: clean_(p.prepaidTo, 30)
  })).filter(p => p.name);
  if (players.length < 2) throw new Error('A game needs at least two players.');
  const tin = players.reduce((a, p) => a + p.buyins.reduce((x, y) => x + y, 0), 0);
  const tout = players.reduce((a, p) => a + p.winnings, 0);
  if (Math.abs(tin - tout) > 0.009) throw new Error('Buy-ins and winnings don\'t match, so this game can\'t be saved yet.');
  const sess = { id: 'g' + Date.now(), date: date, note: clean_(s.note, 80), players: players, paid: {}, created: new Date().toISOString() };
  sheet_('Sessions', SESS_H).appendRow([sess.id, date ? "'" + date : '', new Date(), JSON.stringify(sess)]);
  return state_();
}

function deleteSession_(b) {
  const hit = rows_('Sessions', SESS_H).find(r => r.vals[0] === b.id);
  if (!hit) throw new Error('That game was already deleted.');
  sheet_('Sessions', SESS_H).deleteRow(hit.row);
  return state_();
}

function markPaid_(b) {
  const hit = rows_('Sessions', SESS_H).find(r => r.vals[0] === b.id);
  if (!hit) throw new Error('That game no longer exists. Refresh the page.');
  const sess = JSON.parse(hit.vals[3]);
  sess.paid = sess.paid || {};
  sess.paid[clean_(b.key, 120)] = !!b.paid;
  sheet_('Sessions', SESS_H).getRange(hit.row, 4).setValue(JSON.stringify(sess));
  return state_();
}

/**
 * Optional, run once from the Apps Script editor (select it in the toolbar, then Run).
 * Loads the 6 earlier games from the ChatGPT history so the leaderboard starts with them.
 * They are undated and marked as settled. Running it twice does nothing the second time.
 * Assumes "A" and "Allainah" are the same player, saved as "A".
 */
function importPastGames() {
  const P = (name, buyins, winnings, prepaid, prepaidTo) => ({ name: name, buyins: buyins, winnings: winnings, prepaid: prepaid || 0, prepaidTo: prepaidTo || '' });
  const games = [
    [P('Cali', [300], 1145), P('GD', [300, 700], 1540), P('Kyle', [300, 200, 500], 370), P('Bill', [300, 200], 1640),
     P('Mariel', [300, 200], 130), P('Miles', [300], 240), P('Mama', [300, 200], 535), P('Jha', [300, 200, 200], 0),
     P('A', [300], 0), P('Ejay', [300, 200], 0)],
    [P('Kyle', [300, 300], 2000), P('Ejay', [300, 300, 400, 500], 885), P('Bill', [300], 2430), P('Mama', [300, 300, 400, 300, 200], 2270),
     P('A', [300, 200, 300, 200], 365), P('GD', [300, 700, 1000, 500], 430), P('Cali', [300, 700, 500, 500], 1020)],
    [P('Kyle', [300, 500, 500], 955), P('Ejay', [300, 700], 185), P('Bill', [300], 300), P('Mama', [300, 200, 500], 480),
     P('GD', [300, 300], 1680), P('Cali', [300, 500], 2300), P('Khelvin', [100, 200], 0, 300, 'Cali'), P('Skyler', [300, 300], 0, 600, 'Cali')],
    [P('Kyle', [300, 500], 1190), P('Bill', [300], 1400), P('Mama', [300], 0), P('Cali', [300, 200, 500, 500], 400), P('GD', [300], 210)],
    [P('Kyle', [300, 300], 710), P('Bill', [300, 300], 1130), P('Mama', [300], 300), P('Cali', [300, 300], 0),
     P('GD', [300, 300, 300], 390), P('Ejay', [300, 300, 400], 1110), P('A', [300, 200], 860)],
    [P('Kyle', [300, 300, 400, 500], 400), P('Bill', [300, 300], 3400), P('Mama', [300], 1985), P('Cali', [300, 300, 400, 500, 500], 0),
     P('GD', [300, 300, 400], 0), P('Ejay', [300], 415), P('A', [300, 200], 0)]
  ];
  const existing = rows_('Sessions', SESS_H).map(r => String(r.vals[0]));
  const sh = sheet_('Sessions', SESS_H);
  games.forEach((players, i) => {
    const id = 'past' + (i + 1);
    if (existing.indexOf(id) >= 0) return;
    const sess = { id: id, date: '', note: 'Earlier game ' + (i + 1), players: players, paid: {}, settled: true, created: '2000-01-0' + (i + 1) };
    sh.appendRow([id, '', new Date(), JSON.stringify(sess)]);
  });
}
