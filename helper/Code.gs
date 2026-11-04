/**
 * Ramarc Weight Scanner — shared loads helper (Google Apps Script)
 *
 * Keeps every phone's weight sheets (drafts and completed) in one Google Sheet
 * in your Google account, and sends weight-sheet PDFs by email with one tap.
 *
 * ONE-TIME SETUP (about 5 minutes, on a computer):
 *  1. Go to script.google.com and click "New project". Delete what is there and paste this whole file.
 *  2. Click Save. (Optional: choose "setup" in the function menu at the top and click Run. If you skip it,
 *     the Google Sheet "Ramarc Weight Scanner - Shared Loads" is created the first time a phone connects.)
 *  3. Click Deploy > New deployment > gear icon > Web app.
 *       Execute as: Me      Who has access: Anyone
 *     Click Deploy and copy the "Web app URL" (ends in /exec).
 *  4. Give that URL to the app (Connect on the start screen). Keep it private: anyone with it can see the loads.
 *
 * Email recipients for the one-tap "Send to office" button are listed on the sheet's "Settings" tab.
 */

var VERSION = 1;
var CHUNK = 45000;          // a Google Sheets cell holds 50,000 characters; big documents are split across cells
var MAX_CHUNKS = 10;          // up to 450,000 characters per document (a 300-scan part is ~90,000)
var LOAD_HEAD = ['Load ID', 'Updated', 'Status', 'Mode', 'Product', 'Vendor', 'BOL#', 'LOT', 'Customer#', 'Date', 'Count', 'Net lb', 'Started by', 'Finished by', 'Data'];
var PART_HEAD = ['Path', 'Updated', 'Data'];

/* ---------------- setup ---------------- */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Ramarc Weight Scanner - Shared Loads');
    props.setProperty('SHEET_ID', ss.getId());
  }
  ensureTabs_(ss);
  Logger.log('Ready. Shared sheet: ' + ss.getUrl());
  Logger.log('Next: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), then copy the Web app URL.');
  return ss.getUrl();
}

function ensureTabs_(ss) {
  var loads = ss.getSheetByName('Loads');
  if (!loads) {
    loads = ss.getSheets()[0];
    if (loads.getName() !== 'Sheet1' || loads.getLastRow() > 0) loads = ss.insertSheet('Loads'); else loads.setName('Loads');
    loads.getRange(1, 1, 1, LOAD_HEAD.length).setValues([LOAD_HEAD]).setFontWeight('bold');
    loads.setFrozenRows(1);
  }
  var parts = ss.getSheetByName('Scans');
  if (!parts) {
    parts = ss.insertSheet('Scans');
    parts.getRange(1, 1, 1, PART_HEAD.length).setValues([PART_HEAD]).setFontWeight('bold');
    parts.setFrozenRows(1);
  }
  var st = ss.getSheetByName('Settings');
  if (!st) {
    st = ss.insertSheet('Settings');
    st.getRange(1, 1, 4, 2).setValues([
      ['Email recipients (one per row below)', 'Notes'],
      [Session.getEffectiveUser().getEmail() || '', 'Weight sheets sent with "Send to office" go to every address in column A.'],
      ['', ''],
      ['', '']
    ]);
    st.getRange(1, 1, 1, 2).setFontWeight('bold');
    st.setColumnWidth(1, 320); st.setColumnWidth(2, 520);
  }
  return { loads: loads, parts: parts, settings: st };
}

function book_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) { setup(); id = PropertiesService.getScriptProperties().getProperty('SHEET_ID'); }   // first use creates the sheet, even if setup was never run
  var ss = SpreadsheetApp.openById(id);
  return ensureTabs_(ss);
}

/* ---------------- web app ---------------- */
function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, code: 'invalid_argument', msg: 'bad json' }); }
  return handle_(body);
}

function handle_(q) {
  try {
    var op = q.op || 'ping';
    if (op === 'ping') return out_({ ok: true, v: VERSION, now: Date.now() });
    if (op === 'poll') return out_(poll_(q));
    if (op === 'set' || op === 'update') return out_(write_(q.op, String(q.p || ''), q.d || {}));
    if (op === 'config') return out_({ ok: true, recipients: recipients_() });
    if (op === 'mail') return out_(mail_(q));
    return out_({ ok: false, code: 'invalid_argument', msg: 'unknown op' });
  } catch (err) {
    var m = String(err && err.message || err);
    return out_({ ok: false, code: m === 'not_setup' ? 'failed_precondition' : 'internal', msg: m });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* paths: loads/<id>  (load info, one row on "Loads")
          loads/<id>/parts/<phone>-<n>  (one phone's scans, rows on "Scans") */
function kind_(p) {
  var s = p.split('/');
  if (s.length === 2 && s[0] === 'loads' && s[1]) return 'load';
  if (s.length === 4 && s[0] === 'loads' && s[1] && s[2] === 'parts' && s[3]) return 'part';
  return null;
}

function joinChunks_(row, from) {
  var s = '';
  for (var i = from; i < row.length; i++) { if (row[i] === '' || row[i] == null) break; s += String(row[i]); }
  return s;
}
function splitChunks_(json) {
  var a = [];
  for (var i = 0; i < json.length; i += CHUNK) a.push(json.substring(i, i + CHUNK));
  if (a.length > MAX_CHUNKS) throw new Error('document too large');
  return a;
}

function findRow_(sh, key) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  var keys = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(keys[i][0]) === key) return i + 2;
  return -1;
}

function write_(op, p, d) {
  var k = kind_(p);
  if (!k) return { ok: false, code: 'invalid_argument', msg: 'bad path' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, code: 'unavailable', msg: 'busy' };
  try {
    var b = book_();
    var now = Date.now();
    if (k === 'load') {
      var sh = b.loads, id = p.split('/')[1];
      var r = findRow_(sh, id), cur = null;
      if (r > 0) cur = JSON.parse(joinChunks_(sh.getRange(r, 15, 1, MAX_CHUNKS).getValues()[0], 0) || '{}');
      if (op === 'update' && !cur) return { ok: false, code: 'not_found', msg: 'no such load' };
      var data = op === 'update' ? merge_(cur, d) : d;
      var info = data.info || {};
      var chunks = splitChunks_(JSON.stringify(data));
      var row = [id, now, data.status || 'open', data.mode || 'boxes', info.product || '', info.vendor || '', info.bol || '', info.lot || '', info.cust || '', info.date || '',
                 data.count == null ? '' : data.count, data.netLb == null ? '' : data.netLb, data.createdByName || '', data.closedByName || ''].concat(chunks);
      if (r < 0) r = sh.getLastRow() + 1;
      sh.getRange(r, 15, 1, MAX_CHUNKS).clearContent();
      sh.getRange(r, 1, 1, row.length).setValues([row]);
      return { ok: true, u: now };
    } else {
      var shp = b.parts;
      var rp = findRow_(shp, p), curp = null;
      if (op === 'update') {
        if (rp < 0) return { ok: false, code: 'not_found', msg: 'no such document' };
        curp = JSON.parse(joinChunks_(shp.getRange(rp, 3, 1, MAX_CHUNKS).getValues()[0], 0) || '{}');
      }
      var datap = op === 'update' ? merge_(curp, d) : d;
      var rowp = [p, now].concat(splitChunks_(JSON.stringify(datap)));
      if (rp < 0) rp = shp.getLastRow() + 1;
      shp.getRange(rp, 3, 1, MAX_CHUNKS).clearContent();
      shp.getRange(rp, 1, 1, rowp.length).setValues([rowp]);
      return { ok: true, u: now };
    }
  } finally {
    lock.releaseLock();
  }
}

function merge_(a, b) { var o = {}; var k; for (k in a) o[k] = a[k]; for (k in b) o[k] = b[k]; return o; }

/* poll: load info changed since `since`, plus the scans of one load changed since `psince` */
function poll_(q) {
  var b = book_();
  var now = Date.now();
  var since = Number(q.since || 0) - 3000, psince = Number(q.psince || 0) - 3000;   // 3 s overlap: nothing slips between polls
  var docs = [];
  var sh = b.loads, n = sh.getLastRow() - 1;
  if (n > 0) {
    var ub = sh.getRange(2, 1, n, 2).getValues();
    var hits = [];
    for (var i = 0; i < n; i++) if (Number(ub[i][1]) > since) hits.push(i);
    if (hits.length) {
      var all = hits.length > 15 ? sh.getRange(2, 15, n, MAX_CHUNKS).getValues() : null;
      for (var h = 0; h < hits.length; h++) {
        var ri = hits[h];
        var rowv = all ? all[ri] : sh.getRange(ri + 2, 15, 1, MAX_CHUNKS).getValues()[0];
        var js = joinChunks_(rowv, 0);
        if (js) docs.push({ p: 'loads/' + ub[ri][0], u: Number(ub[ri][1]), d: JSON.parse(js) });
      }
    }
  }
  if (q.load) {
    var pre = 'loads/' + q.load + '/parts/';
    var sp = b.parts, m = sp.getLastRow() - 1;
    if (m > 0) {
      var pu = sp.getRange(2, 1, m, 2).getValues();
      for (var j = 0; j < m; j++) {
        var key = String(pu[j][0]);
        if (key.indexOf(pre) === 0 && Number(pu[j][1]) > psince) {
          var jsp = joinChunks_(sp.getRange(j + 2, 3, 1, MAX_CHUNKS).getValues()[0], 0);
          if (jsp) docs.push({ p: key, u: Number(pu[j][1]), d: JSON.parse(jsp) });
        }
      }
    }
  }
  return { ok: true, now: now, docs: docs };
}

/* ---------------- email ---------------- */
function recipients_() {
  var st = book_().settings, n = st.getLastRow() - 1;
  if (n < 1) return [];
  return st.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
}

function mail_(q) {
  var to = recipients_();
  if (!to.length) return { ok: false, code: 'failed_precondition', msg: 'No email recipients on the Settings tab of the shared sheet.' };
  var name = String(q.filename || 'Weight Sheet.pdf').replace(/[\\\/:*?"<>|]+/g, '-');
  var blob = Utilities.newBlob(Utilities.base64Decode(String(q.pdf || '')), 'application/pdf', name);
  var link = '';
  try {
    var folder = folder_();
    link = folder.createFile(blob.copyBlob()).getUrl();
  } catch (e) { link = ''; }
  var body = String(q.body || '') + (link ? '\n\nCopy saved in Google Drive: ' + link : '');
  MailApp.sendEmail({ to: to.join(','), subject: String(q.subject || name), body: body, name: 'Ramarc Weight Scanner', attachments: [blob] });
  return { ok: true, to: to, drive: link };
}

function folder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID'), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder('Ramarc Weight Sheets (PDF)'); props.setProperty('FOLDER_ID', f.getId()); }
  return f;
}
