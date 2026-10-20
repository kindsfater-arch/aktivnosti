/** Пользователи, роли и вход по персональной ссылке. */

function readUsers_() {
  var sh = sheet_(APP.SHEET.USERS), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, USER_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2,
      id: str_(r[UCOL.ID - 1]),
      fio: str_(r[UCOL.FIO - 1]),
      type: str_(r[UCOL.TYPE - 1]),
      no: str_(r[UCOL.NO - 1]),
      role: str_(r[UCOL.ROLE - 1]) || ROLES.EMPLOYEE,
      reports: yes_(r[UCOL.REPORTS - 1]),
      extra: String(r[UCOL.EXTRA - 1]).split(/[;\n]/).map(str_).filter(Boolean),
      active: str_(r[UCOL.ACTIVE - 1]) === '' ? true : yes_(r[UCOL.ACTIVE - 1]),
      token: str_(r[UCOL.TOKEN - 1])
    };
  }).filter(function (u) { return u.fio; });
}

function usersCached_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('users');
  if (cached) return JSON.parse(cached);
  var users = readUsers_();
  cache.put('users', JSON.stringify(users), 300);
  return users;
}

function decorate_(u) {
  u.isAdmin = u.role === ROLES.ADMIN;
  u.seeAll = u.isAdmin || u.role === ROLES.HEAD;
  u.canReport = u.isAdmin || u.reports;
  return u;
}

function auth_(token) {
  token = str_(token);
  if (token.length < 32) {
    throw new Error('Откройте сервис по своей персональной ссылке. Ссылку выдаёт администратор.');
  }
  var u = usersCached_().filter(function (x) { return x.token === token; })[0];
  if (!u || !u.active) {
    throw new Error('Ссылка недействительна или доступ отключён. Обратитесь к администратору.');
  }
  return decorate_(u);
}

function publicUser_(u) {
  return {
    id: u.id, fio: u.fio, type: u.type, no: u.no, role: u.role,
    isAdmin: u.isAdmin, seeAll: u.seeAll, canReport: u.canReport
  };
}

function noFor_(fio) {
  var u = usersCached_().filter(function (x) { return x.fio === fio; })[0];
  return u ? u.no : '';
}

function typeByFio_() {
  var m = {};
  usersCached_().forEach(function (u) { m[u.fio] = u.type; });
  return m;
}

function appUrl_() {
  var url = prop_('APP_URL') || APP.PUBLIC_URL || APP.WEB_APP_URL;
  if (url) return url;
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function linkFor_(token) {
  var base = appUrl_();
  if (!base || !token) return '';
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'k=' + token;
}

function listUsers_() {
  return readUsers_().map(function (u) {
    return {
      id: u.id, fio: u.fio, type: u.type, no: u.no, role: u.role, reports: u.reports,
      extra: u.extra, active: u.active, link: linkFor_(u.token)
    };
  });
}

function saveUser_(admin, input) {
  var fio = str_(input.fio).replace(/\s+/g, ' ');
  if (fio.split(' ').length < 2) throw new Error('Укажите ФИО полностью');
  var role = str_(input.role) || ROLES.EMPLOYEE;
  if ([ROLES.ADMIN, ROLES.HEAD, ROLES.EMPLOYEE].indexOf(role) < 0) throw new Error('Неизвестная роль');
  var type = str_(input.type);
  if (type && EMP_TYPES.indexOf(type) < 0) throw new Error('Тип сотрудника: ' + EMP_TYPES.join(', '));
  var extra = (Array.isArray(input.extra) ? input.extra : String(input.extra || '').split(/[;\n]/))
    .map(str_).filter(Boolean);

  var res = withLock_(function () {
    var users = readUsers_();
    var dup = users.filter(function (u) { return u.fio === fio && u.id !== input.id; })[0];
    if (dup) throw new Error('Пользователь с таким ФИО уже есть');
    var sh = sheet_(APP.SHEET.USERS);
    var cur = input.id ? users.filter(function (u) { return u.id === input.id; })[0] : null;
    if (input.id && !cur) throw new Error('Пользователь не найден');
    if (cur && cur.id === admin.id && role !== ROLES.ADMIN) throw new Error('Нельзя снять с себя роль администратора');

    var id = cur ? cur.id : newId_();
    var token = cur ? cur.token : newToken_();
    var rowNum = cur ? cur.row : sh.getLastRow() + 1;
    ensureRows_(sh, rowNum);
    sh.getRange(rowNum, 1, 1, USER_HEADERS.length).setValues([[
      id, safeText_(fio), type, safeText_(str_(input.no)), role,
      input.reports ? 'да' : 'нет', safeText_(extra.join('; ')),
      input.active === false ? 'нет' : 'да', token, linkFor_(token)
    ]]);
    return { id: id, isNew: !cur };
  });
  CacheService.getScriptCache().remove('users');
  log_(admin, res.isNew ? 'Добавлен пользователь' : 'Изменён пользователь', res.id, fio + ', ' + role);
  return listUsers_();
}

function regenToken_(admin, args) {
  withLock_(function () {
    var u = readUsers_().filter(function (x) { return x.id === args.id; })[0];
    if (!u) throw new Error('Пользователь не найден');
    var token = newToken_();
    var sh = sheet_(APP.SHEET.USERS);
    sh.getRange(u.row, UCOL.TOKEN).setValue(token);
    sh.getRange(u.row, UCOL.LINK).setValue(linkFor_(token));
  });
  CacheService.getScriptCache().remove('users');
  log_(admin, 'Перевыпущена ссылка', args.id);
  return listUsers_();
}

function refreshLinks_() {
  var sh = sheet_(APP.SHEET.USERS), last = sh.getLastRow();
  if (last < 2) return;
  var tokens = sh.getRange(2, UCOL.TOKEN, last - 1, 1).getValues();
  sh.getRange(2, UCOL.LINK, last - 1, 1).setValues(tokens.map(function (r) { return [linkFor_(str_(r[0]))]; }));
}
