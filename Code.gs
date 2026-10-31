/**
 * Little Bloom — Google Sheets backend (Apps Script web app).
 *
 * Each "kind" (tasks, entries) is a tab. Row 1 holds the headers, one row per item, keyed by `id`.
 * The sheet is dynamic: if the app sends a field the sheet has never seen, a new column is added for it,
 * so you can grow the data model in index.html without touching this file.
 *
 * Setup:
 *   1. Create a Google Sheet → Extensions → Apps Script → paste this file as Code.gs.
 *   2. Change TOKEN below (and SHEET_TOKEN in index.html to the same value).
 *   3. Run `setup` once from the editor and approve the permissions.
 *   4. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
 *   5. Copy the /exec URL into SHEET_API_URL in index.html.
 * After editing this file, use Deploy → Manage deployments → Edit → New version so the URL keeps working.
 */

var TOKEN = 'pBbsH2RBEOQh1hh1DxB5GbCEATpk';

// Known fields and their types. Order here is the column order on a fresh sheet.
// Fields not listed are still stored; they get a column on first sight and keep the value as written.
var SCHEMA = {
  tasks: {
    id: 'string', title: 'string', due: 'string', imp: 'string', mins: 'number',
    effort: 'string', done: 'boolean', doneAt: 'number', createdAt: 'number',
    // `time` is "HH:mm" on the due date; `endAt` is that moment as epoch ms, worked out on the phone.
    // `warnedFor` is the endAt the 10-minute push was last sent for.
    time: 'string', endAt: 'number', warnedFor: 'number',
    desc: 'string'
  },
  entries: {
    id: 'string', type: 'string', mood: 'string', text: 'string', createdAt: 'number'
  },
  // Phones that turned on reminders. `id` is the FCM registration token.
  devices: {
    id: 'string', ua: 'string', createdAt: 'number', lastSeen: 'number'
  }
};

// Kinds the app may read and write directly. Devices only change through register/unregister.
var CLIENT_KINDS = ['tasks', 'entries'];

/* ---------- reminders (Firebase Cloud Messaging) ----------
 * 1. Firebase console → Project settings → Service accounts → Generate new private key (downloads a JSON file).
 * 2. Apps Script → Project Settings → Script properties → add FCM_SERVICE_ACCOUNT with the whole JSON file as the value.
 * 3. Project Settings → set the time zone to yours, so "due today" and the reminder hour match your day.
 * 4. Run `sendTestPush` to check, then run `installReminders` once.
 * Pushes go out for: the daily summary, every newly added task, and WARN_MINUTES before a task's finish time.
 */
var APP_URL = 'https://elhadiysolution.github.io/seedsy/';  // Where the app is hosted, ending in "/", e.g. https://you.github.io/little-bloom/
var REMINDER_HOUR = 8;   // 0–23, in the script time zone.
var WARN_MINUTES = 10;   // How long before a task's finish time the "almost time" push goes out.

/* ---------- web app entry points ---------- */

function doPost(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  if (body.token !== TOKEN) return json_({ ok: false, error: 'unauthorized' });

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    switch (body.action) {
      case 'list':
        return json_({ ok: true, tasks: readAll_('tasks'), entries: readAll_('entries') });
      case 'upsert':
        checkKind_(body.kind);
        if (!body.item || typeof body.item.id !== 'string' || !body.item.id) return json_({ ok: false, error: 'bad_item' });
        var isNew = rowOf_(sheet_(body.kind), body.item.id) < 0;
        upsert_(body.kind, body.item);
        if (body.kind === 'tasks' && body.notify && isNew) tryPush_(newTaskMessage_(body.item));
        return json_({ ok: true });
      case 'delete':
        checkKind_(body.kind);
        remove_(body.kind, String(body.id || ''));
        return json_({ ok: true });
      case 'register':
        var dev = body.device || {};
        if (typeof dev.token !== 'string' || dev.token.length < 20) return json_({ ok: false, error: 'bad_device' });
        var item = { id: dev.token, ua: String(dev.ua || '').slice(0, 200), lastSeen: Date.now() };
        if (rowOf_(sheet_('devices'), dev.token) < 0) item.createdAt = Date.now();
        upsert_('devices', item);
        return json_({ ok: true });
      case 'unregister':
        remove_('devices', String(body.token || ''));
        return json_({ ok: true });
      default:
        return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// Visiting the /exec URL in a browser is a quick health check.
function doGet() {
  return json_({ ok: true, service: 'little-bloom', kinds: Object.keys(SCHEMA) });
}

/* ---------- one-time setup ---------- */

function setup() {
  Object.keys(SCHEMA).forEach(function (kind) {
    var sh = sheet_(kind);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#FDE3EC');
    sh.autoResizeColumns(1, sh.getLastColumn());
  });
  // Drop the default empty "Sheet1" if it is still around.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

/* ---------- sheet helpers ---------- */

function checkKind_(kind) {
  if (CLIENT_KINDS.indexOf(kind) < 0) throw new Error('unknown_kind');
}

// Returns the tab for a kind, creating it with schema headers if needed.
function sheet_(kind) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(kind);
  if (!sh) {
    sh = ss.insertSheet(kind);
    var cols = Object.keys(SCHEMA[kind]);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    cols.forEach(function (c, i) { formatColumn_(sh, kind, c, i + 1); });
  }
  return sh;
}

function headers_(sh) {
  var n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(String) : [];
}

// Adds a column for every key in `item` the sheet does not have yet. Returns the updated header list.
function ensureColumns_(sh, kind, item) {
  var heads = headers_(sh);
  Object.keys(item).forEach(function (key) {
    if (heads.indexOf(key) >= 0) return;
    var col = heads.length + 1;
    sh.getRange(1, col).setValue(key).setFontWeight('bold').setBackground('#FDE3EC');
    formatColumn_(sh, kind, key, col);
    heads.push(key);
  });
  return heads;
}

// String columns are plain text so Sheets never turns "2026-10-27" or "1e5" into a date or number.
function formatColumn_(sh, kind, key, col) {
  var type = SCHEMA[kind][key];
  if (type === 'string') sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
}

function rowOf_(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  var hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : -1;
}

/* ---------- value conversion ---------- */

function toCell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

function fromCell_(kind, key, v) {
  var type = SCHEMA[kind][key];
  if (v === '' || v === null) {
    if (type === 'boolean') return false;
    return type === 'string' ? '' : null;
  }
  if (type === 'number') return Number(v);
  if (type === 'boolean') return v === true || String(v).toUpperCase() === 'TRUE';
  if (type === 'string') return String(v);
  // Dynamic column: bring back JSON objects/arrays, otherwise keep what Sheets holds.
  if (typeof v === 'string' && /^[\[{]/.test(v)) {
    try { return JSON.parse(v); } catch (err) {}
  }
  if (v instanceof Date) return v.getTime();
  return v;
}

/* ---------- CRUD ---------- */

function readAll_(kind) {
  var sh = sheet_(kind);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var heads = headers_(sh);
  var rows = sh.getRange(2, 1, last - 1, heads.length).getValues();
  return rows
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      var o = {};
      heads.forEach(function (h, i) { if (h) o[h] = fromCell_(kind, h, r[i]); });
      return o;
    });
}

function upsert_(kind, item) {
  var sh = sheet_(kind);
  var heads = ensureColumns_(sh, kind, item);
  var row = rowOf_(sh, item.id);
  var current = row > 0 ? sh.getRange(row, 1, 1, heads.length).getValues()[0] : heads.map(function () { return ''; });
  // Fields the item does not mention keep their old value, so older app versions don't wipe newer columns.
  var values = heads.map(function (h, i) { return item.hasOwnProperty(h) ? toCell_(item[h]) : current[i]; });
  if (row > 0) sh.getRange(row, 1, 1, heads.length).setValues([values]);
  else sh.appendRow(values);
}

function remove_(kind, id) {
  if (!id) return;
  var sh = sheet_(kind);
  var row = rowOf_(sh, id);
  if (row > 0) sh.deleteRow(row);
}

/* ---------- push reminders ---------- */

// Run once from the editor. Replaces any earlier reminder triggers.
function installReminders() {
  var mine = ['sendDailyReminders', 'checkEndingTasks'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (mine.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyDays(1).atHour(REMINDER_HOUR).create();
  ScriptApp.newTrigger('checkEndingTasks').timeBased().everyMinutes(1).create();
}

function sendDailyReminders() {
  var msg = dailyMessage_();
  if (msg) pushAll_(msg.title, msg.body);
}

// Runs every minute: one push per open task whose finish time is WARN_MINUTES away or less.
function checkEndingTasks() {
  if (!readAll_('devices').length) return;
  var now = Date.now();
  var soon = readAll_('tasks').filter(function (t) {
    return !t.done && t.endAt && t.endAt > now && t.endAt - now <= WARN_MINUTES * 60000 && t.warnedFor !== t.endAt;
  });
  if (!soon.length) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;  // Try again next minute.
  try {
    soon.forEach(function (t) {
      var left = Math.max(1, Math.round((t.endAt - now) / 60000));
      tryPush_({
        title: '⏰ ' + left + (left === 1 ? ' minute' : ' minutes') + ' left',
        body: t.title + ' · finish by ' + Utilities.formatDate(new Date(t.endAt), Session.getScriptTimeZone(), 'h:mm a')
      });
      // The task may have been deleted while we were sending; never re-create it.
      if (rowOf_(sheet_('tasks'), t.id) > 0) upsert_('tasks', { id: t.id, warnedFor: t.endAt });
    });
  } finally {
    lock.releaseLock();
  }
}

function newTaskMessage_(t) {
  var tz = Session.getScriptTimeZone();
  var when = '';
  if (t.endAt) when = 'Finish by ' + Utilities.formatDate(new Date(t.endAt), tz, 'EEE d MMM, h:mm a');
  else if (t.due) when = 'Due ' + Utilities.formatDate(Utilities.parseDate(t.due, tz, 'yyyy-MM-dd'), tz, 'EEE d MMM');
  return { title: 'New task planted 🌱', body: String(t.title || '') + (when ? '\n' + when : '') };
}

// A push that fails (no devices, FCM not set up yet) must never fail the save that triggered it.
function tryPush_(msg) {
  try {
    pushAll_(msg.title, msg.body);
  } catch (err) {
    console.warn('Push skipped: ' + (err && err.message || err));
  }
}

// Run from the editor to check the setup. The execution log shows how many phones got it.
function sendTestPush() {
  console.log('Sent to ' + pushAll_('Little Bloom 🌸', 'Reminders are working.') + ' device(s).');
}

function dailyMessage_() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var open = readAll_('tasks').filter(function (t) { return !t.done && t.due && t.due <= today; });
  if (!open.length) return null;
  var late = open.filter(function (t) { return t.due < today; }).length;
  var names = open.slice(0, 3).map(function (t) { return t.title; }).join(', ');
  if (open.length > 3) names += ' and ' + (open.length - 3) + ' more';
  return {
    title: open.length === 1 ? '1 task is waiting for you 🌷' : open.length + ' tasks are waiting for you 🌷',
    body: names + (late ? ' (' + late + ' overdue)' : '')
  };
}

// Sends one notification to every registered phone. Returns how many were delivered.
function pushAll_(title, body) {
  var devices = readAll_('devices');
  if (!devices.length) return 0;
  var sa = serviceAccount_();
  var auth = 'Bearer ' + accessToken_(sa);
  var url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
  var requests = devices.map(function (d) {
    return {
      url: url, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: auth },
      payload: JSON.stringify({
        message: {
          token: d.id,
          notification: { title: title, body: body },
          webpush: {
            notification: APP_URL ? { icon: APP_URL + 'icons/icon-192.png' } : {},
            fcm_options: APP_URL ? { link: APP_URL } : undefined
          }
        }
      })
    };
  });
  var sent = 0;
  UrlFetchApp.fetchAll(requests).forEach(function (res, i) {
    var code = res.getResponseCode();
    if (code === 200) { sent++; return; }
    // The phone uninstalled the app or turned notifications off: forget it.
    if (code === 404 || /UNREGISTERED/.test(res.getContentText())) remove_('devices', devices[i].id);
    else console.warn('FCM ' + code + ': ' + res.getContentText());
  });
  return sent;
}

function serviceAccount_() {
  var raw = PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT');
  if (!raw) throw new Error('Add the FCM_SERVICE_ACCOUNT script property first.');
  return JSON.parse(raw);
}

// OAuth access token for FCM, from a service-account JWT. Cached for 50 minutes.
function accessToken_(sa) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('fcm_access_token');
  if (hit) return hit;
  var b64 = function (s) { return Utilities.base64EncodeWebSafe(s).replace(/=+$/, ''); };
  var now = Math.floor(Date.now() / 1000);
  var unsigned = b64(JSON.stringify({ alg: 'RS256', typ: 'JWT' })) + '.' + b64(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  }));
  var jwt = unsigned + '.' + b64(Utilities.computeRsaSha256Signature(unsigned, sa.private_key));
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt }
  });
  var data = JSON.parse(res.getContentText());
  if (!data.access_token) throw new Error('FCM auth failed: ' + res.getContentText());
  cache.put('fcm_access_token', data.access_token, 3000);
  return data.access_token;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
