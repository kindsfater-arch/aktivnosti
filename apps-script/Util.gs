/** Общие помощники. */

var _ss = null;
function ss_() {
  if (!_ss) _ss = SpreadsheetApp.openById(prop_('SPREADSHEET_ID') || APP.SPREADSHEET_ID);
  return _ss;
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Нет листа «' + name + '». Запустите setup() в редакторе скрипта.');
  return sh;
}

function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value == null ? '' : value));
}

// Excel теряет ведущий ноль у ИНН, поэтому 9 и 11 знаков дополняем до 10 и 12.
function normInn_(v) {
  if (typeof v === 'number') v = Math.round(v).toString();
  var s = String(v == null ? '' : v).replace(/\D/g, '');
  if (s.length === 9 || s.length === 11) s = '0' + s;
  return s;
}

function isInn_(s) {
  return /^(\d{10}|\d{12})$/.test(s);
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v == null ? '' : v).replace(/[\s ₽]/g, '');
  if (!s) return 0;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/,/g, '');
  else s = s.replace(',', '.');
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = str_(v), m;
  if ((m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(s))) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function ymd_(v) {
  var d = toDate_(v);
  return d ? Utilities.formatDate(d, APP.TZ, 'yyyy-MM-dd') : '';
}

function stamp_(v) {
  return v instanceof Date ? Utilities.formatDate(v, APP.TZ, "yyyy-MM-dd'T'HH:mm:ss") : '';
}

function today_() {
  return toDate_(Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'));
}

function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function ruDate_(s) {
  return s ? s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(0, 4) : '';
}

function money_(n) {
  return Math.round(n || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

// Текст, начинающийся с =, +, -, @, Sheets воспринял бы как формулу.
function safeText_(s) {
  s = str_(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function yes_(v) {
  return v === true || /^(да|yes|true|1|y)$/i.test(str_(v));
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function ensureRows_(sh, rowNum) {
  var max = sh.getMaxRows();
  if (rowNum > max) sh.insertRowsAfter(max, Math.max(200, rowNum - max));
}

function log_(user, action, id, details) {
  try {
    sheet_(APP.SHEET.LOG).appendRow([
      new Date(), user ? user.fio : 'система', action, id || '', details ? String(details).slice(0, 500) : ''
    ]);
  } catch (e) {
    console.warn('log_: ' + e);
  }
}

function tempFolder_() {
  var id = prop_('TEMP_FOLDER_ID');
  if (id) {
    try {
      var f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* папку удалили — создадим заново */ }
  }
  var folder = DriveApp.getFolderById(prop_('DATA_FOLDER_ID') || APP.DATA_FOLDER_ID).createFolder('Активности — служебное');
  setProp_('TEMP_FOLDER_ID', folder.getId());
  return folder;
}
