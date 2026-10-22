/**
 * Первичная настройка. Запустите setup() один раз из редактора Apps Script.
 * Повторный запуск безопасен: существующие данные не трогаются.
 */
function setup() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(APP.TZ);

  setupActivities_(ss);
  ensureSheet_(ss, APP.SHEET.USERS, USER_HEADERS);
  ensureSheet_(ss, APP.SHEET.DICT, DICT_HEADERS);
  ensureSheet_(ss, APP.SHEET.TURN, TURN_HEADERS);
  ensureSheet_(ss, APP.SHEET.COMP, COMP_HEADERS);
  plansSheet_();
  ensureSheet_(ss, APP.SHEET.LOG, LOG_HEADERS);

  var dict = ss.getSheetByName(APP.SHEET.DICT);
  if (dict.getLastRow() < 2) writeDicts_(DEFAULT_DICTS);
  else applyValidations_();

  seedUsers_();
  readActivities_(); // выдаст ID строкам, уже заполненным в таблице
  installTriggers_();

  var admin = readUsers_().filter(function (u) { return u.role === ROLES.ADMIN; })[0];
  console.log('Готово. Ссылка администратора: ' + (linkFor_(admin.token) ||
    '(появится после развёртывания веб-приложения; затем запустите refreshLinks_ или сохраните APP_URL в админке)'));
  console.log('Первый импорт оборотов: запустите importTurnovers() или нажмите «Обновить обороты» в админке.');
}

function setupActivities_(ss) {
  var sh = ss.getSheetByName(APP.SHEET.ACT);
  if (!sh) {
    var first = ss.getSheetByName('Лист1') || ss.getSheets()[0];
    sh = first.getRange(1, 1).getValue() === ACT_HEADERS[0] ? first.setName(APP.SHEET.ACT) : ss.insertSheet(APP.SHEET.ACT, 0);
  }
  sh.getRange(1, 1, 1, NCOL).setValues([ACT_HEADERS]).setFontWeight('bold').setWrap(true);
  sh.getRange(1, 1, 1, COL.DEAL_SUM).setBackground('#D9EAD3');
  sh.getRange(1, COL.ID, 1, NCOL - COL.ID + 1).setBackground('#EFEFEF').setFontColor('#666666');
  sh.setFrozenRows(1);
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, COL.INN, rows, 1).setNumberFormat('@');
  sh.getRange(2, COL.POTENTIAL, rows, 1).setNumberFormat('#,##0');
  sh.getRange(2, COL.DEAL_SUM, rows, 1).setNumberFormat('#,##0');
  sh.getRange(2, COL.DATE, rows, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, COL.NEXT, rows, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, COL.CREATED, rows, 2).setNumberFormat('dd.mm.yyyy hh:mm');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// Первые пользователи: администратор и все ФИО, уже встречающиеся в журнале активностей.
function seedUsers_() {
  var sh = sheet_(APP.SHEET.USERS);
  if (sh.getLastRow() > 1) return;
  var act = sheet_(APP.SHEET.ACT), last = act.getLastRow();
  var pairs = last > 1 ? act.getRange(2, 1, last - 1, 2).getValues() : [];
  var people = {}, order = [];
  var add = function (fio, role, no) {
    fio = str_(fio);
    if (!fio || people[fio]) return;
    people[fio] = { role: role, no: no || '' };
    order.push(fio);
  };
  add(APP.ADMIN_FIO, ROLES.ADMIN, '');
  pairs.forEach(function (p) {
    add(p[0], ROLES.EMPLOYEE, str_(p[1]));
    add(p[1], ROLES.HEAD, '');
  });
  if (people[APP.ADMIN_FIO] && !people[APP.ADMIN_FIO].no) {
    var own = pairs.filter(function (p) { return str_(p[0]) === APP.ADMIN_FIO; })[0];
    if (own) people[APP.ADMIN_FIO].no = str_(own[1]);
  }
  var rows = order.map(function (fio) {
    var p = people[fio], token = newToken_();
    return [newId_(), fio, '', p.no, p.role, p.role === ROLES.EMPLOYEE ? 'нет' : 'да', '', 'да', token, linkFor_(token)];
  });
  sh.getRange(2, 1, rows.length, USER_HEADERS.length).setValues(rows);
  CacheService.getScriptCache().remove('users');
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importTurnovers') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importTurnovers').timeBased().everyDays(1).atHour(6).inTimezone(APP.TZ).create();
}

// Обновить персональные ссылки после первого развёртывания веб-приложения.
function refreshLinks() {
  refreshLinks_();
  CacheService.getScriptCache().remove('users');
  readUsers_().forEach(function (u) { console.log(u.fio + ' (' + u.role + '): ' + linkFor_(u.token)); });
}
