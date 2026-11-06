/** Журнал активностей: чтение, права, создание, изменение, удаление. */

function rowToAct_(r, rowNum) {
  return {
    row: rowNum,
    id: str_(r[COL.ID - 1]),
    fio: str_(r[COL.FIO - 1]),
    no: str_(r[COL.NO - 1]),
    inn: normInn_(r[COL.INN - 1]),
    short: str_(r[COL.SHORT - 1]),
    full: str_(r[COL.FULL - 1]),
    channel: str_(r[COL.CHANNEL - 1]),
    pirate: /^(да|нет)$/.test(str_(r[COL.PIRATE - 1])) ? str_(r[COL.PIRATE - 1]) : pirateFor_(normInn_(r[COL.INN - 1])),
    status: str_(r[COL.STATUS - 1]),
    date: ymd_(r[COL.DATE - 1]),
    type: str_(r[COL.TYPE - 1]),
    next: ymd_(r[COL.NEXT - 1]),
    comment: str_(r[COL.COMMENT - 1]),
    dealSum: num_(r[COL.DEAL_SUM - 1]),
    author: str_(r[COL.AUTHOR - 1]),
    created: stamp_(r[COL.CREATED - 1]),
    updated: stamp_(r[COL.UPDATED - 1]),
    priority: str_(r[COL.PRIORITY - 1])
  };
}

var _acts = null;

// Все неудалённые активности. Строкам, добавленным в таблицу вручную, выдаётся ID.
function readActivities_() {
  if (_acts) return _acts;
  var sh = sheet_(APP.SHEET.ACT), last = sh.getLastRow();
  ensureActColumns_(sh);
  if (last < 2) return (_acts = []);
  var vals = sh.getRange(2, 1, last - 1, NCOL).getValues();
  var out = [], missing = [];
  vals.forEach(function (r, i) {
    if (!str_(r[COL.ID - 1])) {
      var hasData = str_(r[COL.FIO - 1]) && (str_(r[COL.INN - 1]) || r[COL.DATE - 1]);
      if (!hasData) return;
      r[COL.ID - 1] = newId_();
      missing.push({ row: i + 2, id: r[COL.ID - 1] });
    }
    if (yes_(r[COL.DELETED - 1])) return;
    out.push(rowToAct_(r, i + 2));
  });
  if (missing.length) assignIds_(sh, missing, out);
  return (_acts = out);
}

// Колонки, добавленные после первого запуска (T — приоритет клиента): заголовок и выпадающий список.
// Делается один раз, дальше проверка — одно чтение свойства.
function ensureActColumns_(sh) {
  if (prop_('ACT_VOLUME_LABEL') !== '1') {
    sh.getRange(1, COL.DEAL_SUM).setValue('Объём оборота, ₽/месяц');
    setProp_('ACT_VOLUME_LABEL', '1');
  }
  if (prop_('ACT_NCOL') === String(NCOL)) return;
  if (sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
  sh.getRange(1, 1, 1, NCOL).setValues([ACT_HEADERS]);
  sh.getRange(1, COL.PRIORITY).setFontWeight('bold').setBackground('#D9EAD3');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(PRIORITIES, true).setAllowInvalid(false).build();
  sh.getRange(2, COL.PRIORITY, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  setProp_('ACT_NCOL', NCOL);
}

function assignIds_(sh, missing, acts) {
  withLock_(function () {
    var now = new Date();
    missing.forEach(function (m) {
      var cell = sh.getRange(m.row, COL.ID);
      var existing = str_(cell.getValue());
      if (existing) {
        // Параллельный запрос успел выдать ID раньше — используем его.
        acts.forEach(function (a) { if (a.row === m.row) a.id = existing; });
        return;
      }
      cell.setValue(m.id);
      sh.getRange(m.row, COL.AUTHOR).setValue('таблица');
      sh.getRange(m.row, COL.CREATED).setValue(now);
    });
  });
}

function canSee_(u, a) {
  return u.seeAll || a.fio === u.fio || u.extra.indexOf(a.fio) >= 0;
}

function canEdit_(u, a) {
  return u.seeAll || a.fio === u.fio;
}

function visibleActs_(u, fio) {
  return readActivities_().filter(function (a) {
    return canSee_(u, a) && (!fio || a.fio === fio);
  });
}

function pubAct_(u, a) {
  var o = {};
  Object.keys(a).forEach(function (k) { if (k !== 'row') o[k] = a[k]; });
  o.canEdit = canEdit_(u, a);
  return o;
}

function byDateDesc_(x, y) {
  return (y.date + y.created).localeCompare(x.date + x.created);
}

// Последняя активность по каждой паре «сотрудник + ИНН» — текущее состояние клиента.
function latestByClient_(acts) {
  var m = {};
  acts.forEach(function (a) {
    var k = a.fio + '|' + a.inn;
    var cur = m[k];
    if (!cur || a.date > cur.date || (a.date === cur.date && a.created > cur.created)) m[k] = a;
  });
  return Object.keys(m).map(function (k) { return m[k]; });
}

function listActivities_(u, args) {
  var q = str_(args.q).toLowerCase();
  var visible = visibleActs_(u, args.fio);
  // Поиск не должен подтверждать скрытые ФИО и контакты по числу совпадений.
  var names = q ? identityNames_(visible) : [], labels = q ? displayNames_(names) : {};
  var list = visible.filter(function (a) {
    if (args.from && a.date < args.from) return false;
    if (args.to && a.date > args.to) return false;
    if (args.status && a.status !== args.status) return false;
    if (args.type && a.type !== args.type) return false;
    if (args.channel && a.channel !== args.channel) return false;
    if (q && (a.inn + ' ' + a.short + ' ' + a.full + ' ' + privateComment_(a.comment, names, labels)).toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
  list.sort(byDateDesc_);
  return {
    total: list.length,
    sum: actsSummary_(list),
    rows: list.slice(0, 1000).map(function (a) { return pubAct_(u, a); })
  };
}

// Итоги по всем отобранным активностям (не только по показанным строкам) — по каждому полю, кроме комментария.
// Признак «Пиратка» — по уникальным клиентам в отобранных активностях.
function actsSummary_(list) {
  var today = ymd_(today_()), kinds = statusKinds_();
  var s = { count: list.length, days: 0, fios: 0, clients: 0, key: 0, byType: {}, byStatus: {}, byChannel: {},
    pirateClients: 0, deals: 0, dealSum: 0, next: 0, nextAhead: 0 };
  var days = {}, fios = {}, inns = {};
  var inc = function (o, k) { o[k || '—'] = (o[k || '—'] || 0) + 1; };
  list.forEach(function (a) {
    days[a.date] = 1;
    fios[a.fio] = 1;
    if (!inns[a.inn]) {
      inns[a.inn] = 1;
      if (a.pirate === 'да') s.pirateClients++;
      if (a.priority === PRIORITIES[0]) s.key++;
    }
    inc(s.byType, a.type);
    inc(s.byStatus, a.status);
    inc(s.byChannel, a.channel);
    if (kinds[a.status] === 'deal' && a.dealSum > 0) { s.deals++; s.dealSum += a.dealSum; }
    if (a.next) { s.next++; if (a.next >= today) s.nextAhead++; }
  });
  s.days = Object.keys(days).length;
  s.fios = Object.keys(fios).length;
  s.clients = Object.keys(inns).length;
  return s;
}

function findActRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var hit = sh.getRange(2, COL.ID, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function validateActivity_(u, input) {
  var d = getDicts_(), kinds = statusKinds_(d);
  var v = {};

  v.fio = storedName_(u.seeAll ? input.fio || u.fio : u.fio);
  if (!v.fio) throw new Error('Не указан сотрудник');

  v.inn = normInn_(input.inn);
  if (!isInn_(v.inn)) throw new Error('ИНН должен содержать 10 цифр (юрлицо) или 12 цифр (ИП)');

  v.priority = str_(input.priority);
  if (PRIORITIES.indexOf(v.priority) < 0) throw new Error('Укажите приоритет клиента: ключевой или неключевой');

  v.channel = str_(input.channel);
  if (d.channels.indexOf(v.channel) < 0) throw new Error('Выберите канал поступления');

  v.status = str_(input.status);
  if (!kinds[v.status]) throw new Error('Выберите статус');
  v.statusActive = activeStatuses_(d).some(function (s) { return s.name === v.status; });

  v.date = toDate_(input.date);
  if (!v.date) throw new Error('Укажите дату активности');

  v.type = str_(input.type);
  if (d.types.indexOf(v.type) < 0) throw new Error('Выберите вид активности');

  v.next = null;
  if (str_(input.next)) {
    v.next = toDate_(input.next);
    if (!v.next) throw new Error('Неверная дата следующей активности');
    if (v.next < v.date) throw new Error('Дата следующей активности не может быть раньше даты активности');
  }

  v.comment = storedComment_(input.comment);
  if (v.comment.length < APP.MIN_COMMENT) {
    throw new Error('Опишите в комментарии суть активности (не короче ' + APP.MIN_COMMENT + ' символов)');
  }

  v.dealSum = '';
  if (needsVolume_(v.status, kinds)) {
    v.dealSum = Math.round(num_(input.dealSum));
    if (!(v.dealSum > 0)) throw new Error('Для статуса «' + v.status + '» укажите объём оборота в месяц');
  }

  v.pirate = pirateFor_(v.inn);

  v.short = str_(input.short);
  v.full = str_(input.full);
  if (!v.short || !v.full) {
    var comp = companyFromCache_(v.inn);
    if (comp) {
      v.short = v.short || comp.short;
      v.full = v.full || comp.full;
    }
  }
  v.no = noFor_(v.fio) || storedName_(input.no);
  return v;
}

function saveActivity_(u, input) {
  // Никаких чтений из кэша внутри проверки дублей: конкурирующие сохранения сериализованы.
  var saved = withLock_(function () {
    var sh = sheet_(APP.SHEET.ACT), now = new Date();
    ensureActColumns_(sh);
    var journal = contactsSheet_(), eventLast = journal.getLastRow();
    var existingEvents = eventLast > 1 ? journal.getRange(2, 1, eventLast - 1, CONTACT_HEADERS.length).getValues() : [];
    var operation = str_(input.operationId) || newId_();
    if (!/^[a-zA-Z0-9_-]{8,80}$/.test(operation)) throw new Error('Неверный идентификатор операции');
    var retry = existingEvents.filter(function (r) { return r[0] === operation && r[9] === u.fio; })[0];
    if (retry) {
      var retriedRow = findActRow_(sh, retry[1]);
      if (!retriedRow) throw new Error('Запись операции не найдена');
      var retried = sh.getRange(retriedRow, 1, 1, NCOL).getValues()[0];
      if (yes_(retried[COL.DELETED - 1]) || !canEdit_(u, rowToAct_(retried, retriedRow))) throw new Error('Нет доступа к активности');
      return { row: retried, rowNum: retriedRow, repeated: true };
    }
    var rowNum = input.id ? findActRow_(sh, input.id) : 0;
    if (input.id && !rowNum) throw new Error('Активность не найдена');
    var old = rowNum ? sh.getRange(rowNum, 1, 1, NCOL).getValues()[0] : null;
    var cur = old ? rowToAct_(old, rowNum) : null;
    if (cur && (yes_(old[COL.DELETED - 1]) || !canEdit_(u, cur))) throw new Error('Нет прав на изменение этой активности');
    if (cur && input.expectedUpdated != null && input.expectedUpdated !== cur.updated) throw new Error('Активность уже изменена. Откройте её заново');
    var payload = {}; Object.keys(input).forEach(function (k) { payload[k] = input[k]; });
    var resultDate = '';
    if (input.recordResult) {
      if (!cur || completedStatus_(cur.status)) throw new Error('Нельзя записать результат: активность уже завершена или не найдена');
      resultDate = str_(input.resultDate) || ymd_(today_());
      if (ymd_(resultDate) !== resultDate || resultDate > ymd_(today_()) || resultDate < cur.date) throw new Error('Укажите дату результата между датой активности и сегодня');
      if (commentBody_(input.comment).length < APP.MIN_COMMENT) throw new Error('Опишите результат контакта');
      payload.fio = cur.fio; payload.inn = cur.inn; payload.date = cur.date;
      payload.comment = cur.comment + '\n\n' + datedComment_(input.comment, str_(input.type));
    }
    if (!cur && !input.recordResult) {
      if (commentBody_(input.comment).length < APP.MIN_COMMENT) throw new Error('Опишите в комментарии суть активности');
      payload.comment = datedComment_(input.comment);
    }
    if (cur && !input.recordResult && input.historyEdited === false) payload.comment = cur.comment;
    if (cur && !input.recordResult && commentBody_(input.appendComment)) {
      if (commentBody_(input.appendComment).length < APP.MIN_COMMENT) throw new Error('Опишите новый комментарий подробнее');
      payload.comment = str_(payload.comment == null ? cur.comment : payload.comment) + '\n\n' + datedComment_(input.appendComment);
    }
    var v = validateActivity_(u, payload);
    var last = sh.getLastRow(), values = last > 1 ? sh.getRange(2, 1, last - 1, NCOL).getValues() : [];
    if (input.recordResult) {
      var newer = values.map(function (r, i) { return yes_(r[COL.DELETED - 1]) ? null : rowToAct_(r, i + 2); })
        .filter(function (a) { return a && a.fio === cur.fio && a.inn === cur.inn && !completedStatus_(a.status); }).sort(byDateDesc_)[0];
      if (newer && newer.id !== cur.id) throw new Error('У клиента есть более поздняя незавершённая активность. Откройте план заново');
    }
    if ((!cur || cur.inn !== v.inn || (completedStatus_(cur.status) && !completedStatus_(v.status))) && unfinishedRows_(values, v.inn, rowNum).length) throw new Error(unfinishedMessage_());
    if (!v.statusActive && (!cur || v.status !== cur.status)) throw new Error(retiredStatusMsg_(v.status));
    if (input.recordResult && v.next && ymd_(v.next) < resultDate) throw new Error('Следующий шаг не может быть раньше результата');
    if (completedStatus_(v.status)) v.next = null;
    var id = cur ? cur.id : newId_();
    if (!rowNum) { rowNum = last + 1; ensureRows_(sh, rowNum); }
    var row = [safeText_(v.fio), safeText_(v.no), v.inn, safeText_(v.short), safeText_(v.full), v.channel,
      v.pirate, v.status, v.date, v.type, v.next || '', safeText_(v.comment), v.dealSum,
      id, '', cur ? cur.author : storedName_(u.fio), cur ? old[COL.CREATED - 1] : now, now, '', v.priority];
    var logs = [];
    if (cur && !existingEvents.some(function (r) { return r[1] === cur.id && (r[10] === 'legacy' || r[10] === 'create'); })) {
      logs.push(['legacy_' + cur.id, cur.id, toDate_(cur.date), safeText_(cur.fio), cur.inn, cur.type, cur.status, '', old[COL.CREATED - 1] || now, safeText_(cur.author), 'legacy']);
    }
    logs.push([operation, id, input.recordResult ? toDate_(resultDate) : cur ? today_() : v.date,
      safeText_(v.fio), v.inn, v.type, v.status, cur ? cur.status : '', now, safeText_(storedName_(u.fio)), input.recordResult ? 'result' : cur ? 'edit' : 'create']);
    ensureRows_(journal, eventLast + logs.length);
    try {
      sh.getRange(rowNum, COL.INN).setNumberFormat('@');
      sh.getRange(rowNum, 1, 1, NCOL).setValues([row]);
      journal.getRange(eventLast + 1, 5, logs.length, 1).setNumberFormat('@');
      journal.getRange(eventLast + 1, 1, logs.length, CONTACT_HEADERS.length).setValues(logs);
      SpreadsheetApp.flush();
    } catch (err) {
      // Откат обеих записей при ошибке, чтобы повтор не добавил лишний контакт.
      if (old) sh.getRange(rowNum, 1, 1, NCOL).setValues([old]);
      else sh.getRange(rowNum, 1, 1, NCOL).clearContent();
      journal.getRange(eventLast + 1, 1, logs.length, CONTACT_HEADERS.length).clearContent();
      SpreadsheetApp.flush();
      throw err;
    }
    return { row: row, rowNum: rowNum, isNew: !cur, result: !!input.recordResult };
  });
  _acts = null;
  var a = rowToAct_(saved.row, saved.rowNum);
  if (!saved.repeated) {
    if (a.short || a.full) rememberCompany_({ inn: a.inn, short: a.short, full: a.full, source: 'активность' });
    log_(u, saved.result ? 'Записан результат' : saved.isNew ? 'Создана активность' : 'Изменена активность', a.id, a.inn + ' · ' + a.status);
  }
  return pubAct_(u, a);
}

function retiredStatusMsg_(status) {
  return 'Статус «' + status + '» больше не используется — выберите другой';
}

// Подсказка приоритета: из последней активности по клиенту, иначе из файла продаж
// («Ключевой» там — ключевой, любой другой приоритет — неключевой).
function priorityFor_(inn, acts, u) {
  var last = acts.filter(function (a) { return a.priority; }).sort(byDateDesc_)[0];
  if (last) return { value: last.priority, source: 'по последней активности' };
  var sale = readSales_().filter(function (x) {
    return x.inn === inn && x.priority && (u.seeAll || x.fio === u.fio || u.extra.indexOf(x.fio) >= 0);
  })[0];
  if (sale) return { value: sale.priority === 'Ключевой' ? 'Ключевой' : 'Неключевой', source: 'по файлу продаж: ' + sale.priority };
  return null;
}

function deleteActivity_(u, args) {
  withLock_(function () {
    var sh = sheet_(APP.SHEET.ACT);
    ensureActColumns_(sh);
    var rowNum = findActRow_(sh, args.id);
    if (!rowNum) throw new Error('Активность не найдена');
    var cur = rowToAct_(sh.getRange(rowNum, 1, 1, NCOL).getValues()[0], rowNum);
    if (!canEdit_(u, cur)) throw new Error('Нет прав на удаление этой активности');
    sh.getRange(rowNum, COL.UPDATED, 1, 2).setValues([[new Date(), 'да']]);
  });
  _acts = null;
  log_(u, 'Удалена активность', args.id);
  return true;
}

function lookupInn_(u, args) {
  var inn = normInn_(args.inn);
  if (!isInn_(inn)) throw new Error('ИНН должен содержать 10 или 12 цифр');
  var all = readActivities_().filter(function (a) { return a.inn === inn; });
  var visible = all.filter(function (a) { return canSee_(u, a); });
  var inWork = latestByClient_(visible).map(function (a) {
    return { fio: a.fio, status: a.status, date: a.date, next: a.next };
  });
  // Вся история по клиенту, которую сотруднику можно видеть, — от новых к старым, с полными комментариями.
  var history = visible.sort(byDateDesc_)
    .map(function (a) { return pubAct_(u, a); });
  var unfinished = visible.filter(function (a) { return !completedStatus_(a.status); });
  return {
    inn: inn,
    priority: priorityFor_(inn, visible, u),
    company: visible.length || u.seeAll ? findCompany_(inn) : null,
    pirate: visible.length || u.seeAll ? pirateFor_(inn) : '',
    inWork: inWork,
    // Нужен только признак занятости ИНН: детали чужих активностей не раскрываем.
    blocked: all.some(function (a) { return !completedStatus_(a.status); }),
    unfinishedCount: unfinished.length,
    unfinished: unfinished.sort(byDateDesc_).map(function (a) { return pubAct_(u, a); }),
    history: history
  };
}
/**
 * Веб-приложение и единая точка входа API.
 *  doGet  — отдаёт интерфейс (персональная ссылка: …/exec?k=<токен>).
 *  doPost — JSON API для интерфейса, размещённого на GitHub Pages.
 *  api    — тот же API для google.script.run, когда интерфейс открыт с адреса Apps Script.
 */

function doGet(e) {
  var token = e && e.parameter ? str_(e.parameter.k).replace(/[^0-9a-f]/gi, '') : '';
  // Интерфейс живёт на GitHub Pages: старые ссылки на Apps Script отправляем туда.
  if (!/script\.google\.com/.test(appUrl_())) {
    var target = token ? linkFor_(token) : appUrl_();
    return HtmlService.createHtmlOutput(
      '<div style="font:16px system-ui,sans-serif;max-width:480px;margin:15vh auto;text-align:center">' +
      '<p>Сервис «Активности» открывается по новому адресу.</p>' +
      '<p><a target="_top" href="' + target.replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }) + '" style="display:inline-block;padding:10px 18px;background:#1a7f37;' +
      'color:#fff;border-radius:9px;text-decoration:none">Перейти</a></p></div>'
    ).setTitle('Активности').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  var t = HtmlService.createTemplateFromFile('Index');
  t.bootJson = JSON.stringify({ mode: 'gas', token: token, assetsUrl: APP.PUBLIC_URL }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Активности')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var req = {};
  try {
    if (!e || !e.postData || e.postData.contents.length > 5000000) throw new Error('Размер запроса');
    req = JSON.parse(e.postData.contents);
    if (!req || Array.isArray(req) || typeof req !== 'object') req = {};
  } catch (err) {
    req = {};
  }
  return ContentService.createTextOutput(JSON.stringify(api(req.k, req.fn, req.args, req.rid)))
    .setMimeType(ContentService.MimeType.JSON);
}

// rid — идентификатор запроса. Google иногда теряет ответ на POST (лишний редирект),
// и браузер повторяет запрос: по rid отдаём сохранённый результат, не выполняя действие второй раз.
function api(token, fn, args, rid) {
  try {
    // Авторизация до чтения кэша: отозванные ссылки и изменённые роли действуют сразу.
    var user = auth_(token), cache = CacheService.getScriptCache();
    var scope = [token, fn, args || {}, rid, user.id, user.fio, user.role, user.reports, user.extra];
    var key = rid ? 'private-v4:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(scope))).slice(0, 44) : '';
    if (key) { var hit = cache.get(key); if (hit) return JSON.parse(hit); }
    var res = callRoute_(token, fn, args, user);
    if (key && res.ok) {
      var json = JSON.stringify(res);
      if (json.length < 90000) cache.put(key, json, 600);
    }
    return res;
  } catch (err) { return { ok: false, error: String(err.message || err) }; }
}

function callRoute_(token, fn, args, authenticated) {
  try {
    var user = authenticated || auth_(token);
    var routes = routes_();
    var route = typeof fn === 'string' && Object.prototype.hasOwnProperty.call(routes, fn) ? routes[fn] : null;
    if (!route) throw new Error('Неизвестный метод');
    if (route.need && !user[route.need]) throw new Error('Недостаточно прав для этого действия');
    if (prop_('STORED_PRIVACY_PENDING') && /^(save|delete|import|run)/.test(fn)) throw new Error('Обезличивание данных: повторите сохранение через минуту');
    validateApiArgs_(args || {});
    var result = privacyResult_(user, route.fn(user, privateArgs_(user, args || {})));
    result.ok = true;
    return result;
  } catch (err) {
    console.error(fn + ': ' + (err && err.stack || err));
    return { ok: false, error: user ? privacyResult_(user, { error: String(err && err.message || err) }).data.error : 'Ссылка недействительна или доступ отключён' };
  }
}

function routes_() {
  return {
    me: { fn: me_ },
    plan: { fn: getPlan_ },
    listActivities: { fn: listActivities_ },
    saveActivity: { fn: saveActivity_ },
    activityForResult: { fn: activityForResult_ },
    deleteActivity: { fn: deleteActivity_ },
    lookupInn: { fn: lookupInn_ },
    lookupCompanyOnline: { fn: lookupCompanyOnline_ },
    disableCompanyNamesBackground: { fn: disableCompanyNamesBackground_, need: 'isAdmin' },
    reportData: { fn: buildReport_, need: 'canReport' },
    planProgress: { fn: planProgress_ },
    confirmedSalesReport: { fn: confirmedSalesReport_ },
    previewConfirmedSales: { fn: previewConfirmedSales_, need: 'isAdmin' },
    importConfirmedSales: { fn: importConfirmedSales_, need: 'isAdmin' },
    plans: { fn: getPlans_, need: 'seeAll' },
    savePlans: { fn: savePlans_, need: 'seeAll' },
    depersonalizeStorage: { fn: depersonalizeStorage_, need: 'isAdmin' },
    secureMaintenance: { fn: secureMaintenance_, need: 'isAdmin' },
    userLink: { fn: userLink_, need: 'isAdmin' },
    users: { fn: function () { return listUsers_(); }, need: 'isAdmin' },
    saveUser: { fn: saveUser_, need: 'isAdmin' },
    regenToken: { fn: regenToken_, need: 'isAdmin' },
    saveDicts: { fn: saveDicts_, need: 'isAdmin' },
    adminInfo: { fn: adminInfo_, need: 'isAdmin' },
    migratePrivacy: { fn: migratePrivacy_, need: 'isAdmin' },
    runImport: { fn: function () { return importTurnovers_(true); }, need: 'isAdmin' },
    runSalesImport: { fn: function () { return importSales_(); }, need: 'isAdmin' },
    saveSettings: { fn: saveSettings_, need: 'isAdmin' }
  };
}

function me_(u) {
  var users = usersCached_().filter(function (x) { return x.active; });
  var fios = u.seeAll
    ? users.map(function (x) { return x.fio; })
    : [u.fio].concat(u.extra);
  return {
    user: publicUser_(u),
    dicts: getDicts_(),
    fios: fios.sort(),
    roles: [ROLES.EMPLOYEE, ROLES.HEAD, ROLES.ADMIN],
    empTypes: EMP_TYPES,
    minComment: APP.MIN_COMMENT,
    data: importStatus_()
  };
}

function adminInfo_() {
  return {
    import: importStatus_(),
    sales: salesStatus_(),
    confirmed: confirmedBatches_(),
    appUrl: prop_('APP_URL') || '',
    scriptUrl: (function () { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })(),
    dadata: !!prop_('DADATA_KEY'),
    companyLookup: { mode: 'manual', backgroundTriggers: companyNamesBackgroundTriggers_().length },
    sheetUrl: ss_().getUrl()
  };
}

function saveSettings_(u, args) {
  if (args.appUrl != null) {
    var url = str_(args.appUrl);
    if (url && !/^https:\/\/[a-z0-9.-]+(?::\d+)?(?:[/?#][^\s<>"'`\\]*)?$/i.test(url)) throw new Error('Адрес должен начинаться с https://');
    setProp_('APP_URL', url);
    refreshLinks_();
  }
  if (args.dadataKey != null) setProp_('DADATA_KEY', str_(args.dadataKey));
  log_(u, 'Изменены настройки');
  return adminInfo_();
}

function validateApiArgs_(args, depth) {
  depth = depth || 0;
  if (depth > 30) throw new Error('Слишком сложный запрос');
  if (!args || typeof args !== 'object') return;
  Object.keys(args).forEach(function (key) {
    if (['__proto__', 'prototype', 'constructor'].indexOf(key) >= 0) throw new Error('Недопустимое поле запроса');
    validateApiArgs_(args[key], depth + 1);
  });
}
/** Дата записи нового комментария, независимо от даты контакта. */
function commentBody_(text) {
  return str_(text).replace(/^\d{2}\.\d{2}\.\d{4}\s*[—–-]\s*/, '').trim();
}
function datedComment_(text, type) {
  var date = ruDate_(ymd_(today_())), body = commentBody_(text);
  return type ? '[' + date + ' · ' + type + '] ' + body : date + ' — ' + body;
}
/** Названия из уже сохранённых данных. Ввод ИНН никогда не отправляет запрос во внешний источник. */
function findCompany_(inn) {
  return companyFromCache_(inn) || companyFromActivities_(inn);
}

function companyFromCache_(inn) {
  var sh = sheet_(APP.SHEET.COMP), last = sh.getLastRow();
  if (last < 2) return null;
  var hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(inn).matchEntireCell(true).findNext();
  if (!hit) return null;
  var r = sh.getRange(hit.getRow(), 1, 1, COMP_HEADERS.length).getValues()[0];
  return { inn: inn, short: str_(r[1]), full: str_(r[2]), ogrn: str_(r[3]), region: str_(r[4]), status: str_(r[5]), source: str_(r[6]) };
}

function companyFromActivities_(inn) {
  var a = readActivities_().filter(function (x) { return x.inn === inn && (x.short || x.full); })[0];
  return a ? { inn: inn, short: a.short, full: a.full, source: 'активность' } : null;
}

// Новые названия кэшируем; при дополнении существующей строки перечитываем её под блокировкой.
function rememberCompany_(c) {
  if (!c || !(c.short || c.full)) return null;
  try {
    var existing = companyFromCache_(c.inn);
    if (existing) {
      // Успешный поиск дополняет неполный справочник, не затирая введённые названия.
      if ((!existing.short && c.short) || (!existing.full && c.full)) withLock_(function () {
        var sh = sheet_(APP.SHEET.COMP), last = sh.getLastRow();
        var hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(c.inn).matchEntireCell(true).findNext();
        if (!hit) return;
        var row = hit.getRow(), names = sh.getRange(row, 2, 1, 2).getValues()[0];
        if (!str_(names[0]) && c.short) sh.getRange(row, 2).setValue(safeText_(c.short));
        if (!str_(names[1]) && c.full) sh.getRange(row, 3).setValue(safeText_(c.full));
        SpreadsheetApp.flush();
      });
      return c;
    }
    sheet_(APP.SHEET.COMP).appendRow([
      "'" + c.inn, safeText_(c.short), safeText_(c.full), c.ogrn || '', c.region || '', c.status || '', c.source || '', new Date()
    ]);
    CacheService.getScriptCache().remove('nf:' + c.inn);
  } catch (e) {
    console.warn('rememberCompany_: ' + e);
  }
  return c;
}

function companiesMap_() {
  var sh = sheet_(APP.SHEET.COMP), last = sh.getLastRow(), m = {};
  if (last < 2) return m;
  sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
    m[normInn_(r[0])] = str_(r[1]) || str_(r[2]);
  });
  return m;
}
/**
 * Константы сервиса «Активности».
 * ID таблицы и папки с оборотами можно переопределить в свойствах скрипта
 * (SPREADSHEET_ID, DATA_FOLDER_ID), не меняя код.
 */
var APP = {
  SPREADSHEET_ID: '',
  DATA_FOLDER_ID: '',
  // Таблица фактических продаж (строка — одна продажа: ИНН, MID, дата регистрации, МПП). Свойство SALES_SHEET_ID имеет приоритет.
  SALES_SHEET_ID: '',
  ADMIN_FIO: '',
  // Адрес интерфейса на GitHub Pages — на него ведут персональные ссылки.
  PUBLIC_URL: '',
  // Адрес веб-приложения Apps Script (API). Свойство скрипта APP_URL или настройка в админке имеют приоритет.
  WEB_APP_URL: '',
  TZ: 'Europe/Kaliningrad',
  MIN_COMMENT: 10,
  SHEET: {
    ACT: 'Активности',
    USERS: 'Пользователи',
    DICT: 'Справочники',
    TURN: 'Обороты',
    COMP: 'Компании',
    PLANS: 'Планы',
    SALES: 'Продажи',
    LOG: 'Журнал'
  }
};

// Колонки листа «Активности». A–L — исходная структура файла, M–S — служебные, T — добавлена позже.
var COL = {
  FIO: 1, NO: 2, INN: 3, SHORT: 4, FULL: 5, CHANNEL: 6, PIRATE: 7, STATUS: 8,
  DATE: 9, TYPE: 10, NEXT: 11, COMMENT: 12, DEAL_SUM: 13,
  ID: 14, POT_SRC: 15, AUTHOR: 16, CREATED: 17, UPDATED: 18, DELETED: 19, PRIORITY: 20
};
var NCOL = 20;
var PRIORITIES = ['Ключевой', 'Неключевой'];
var ACT_HEADERS = [
  'ФИО МПП/МПС/МРГС/НУ', 'ФИО НО', 'ИНН', 'Краткое название ЮЛ', 'Полное название ЮЛ',
  'Канал поступления', 'Пиратка', 'Статус', 'Дата активности', 'Вид активности',
  'Дата следующей активности', 'Комментарий', 'Объём оборота, ₽/месяц',
  'ID', 'Резерв', 'Автор', 'Создано', 'Изменено', 'Удалено', 'Приоритет клиента'
];

var ROLES = { ADMIN: 'Администратор', HEAD: 'Начальник отдела', EMPLOYEE: 'Сотрудник' };
var EMP_TYPES = ['МПП', 'МПС', 'МРГК', 'НУ'];

var UCOL = { ID: 1, FIO: 2, TYPE: 3, NO: 4, ROLE: 5, REPORTS: 6, EXTRA: 7, ACTIVE: 8, TOKEN: 9, LINK: 10 };
var USER_HEADERS = [
  'ID', 'ФИО', 'Тип сотрудника', 'ФИО НО', 'Роль', 'Отчёты', 'Доп. доступ к ФИО (через ;)',
  'Активен', 'Токен', 'Персональная ссылка'
];

// Типы статусов: open — клиент в работе, deal — сделка (нужна сумма), lost — закрыт без сделки.
// active: false — статус остаётся в отчётах и старых записях, но выбрать его в новой или изменённой активности нельзя.
var KIND_LABEL = { open: 'Открыт', deal: 'Сделка', lost: 'Закрыт' };
var DEFAULT_DICTS = {
  channels: ['Кампания продаж', 'Пиратка', 'Полевой обход', 'Личное привлечение', 'Другое'],
  statuses: [
    { name: 'Недозвон', kind: 'open', active: true },
    { name: 'Отказ', kind: 'lost', active: true },
    { name: 'Договор подписан', kind: 'deal', active: true },
    { name: 'Выполнено расширение', kind: 'deal', active: true },
    { name: 'Нет контактов', kind: 'open', active: true },
    { name: 'Назначено', kind: 'open', active: true },
    { name: 'Получено согласие', kind: 'open', active: true },
    { name: 'Клиент думает', kind: 'open', active: true }
  ],
  types: ['Звонок', 'Встреча', 'Сообщение']
};
// Увеличьте, чтобы статусы из DEFAULT_DICTS (тип и доступность) один раз заново записались в справочник.
var DICT_VERSION = 2;
var DICT_HEADERS = ['Канал поступления', 'Статус', 'Тип статуса (Открыт/Сделка/Закрыт)', 'Вид активности', 'Статус доступен для выбора (Да/Нет)'];

var TURN_HEADERS = ['ИНН', 'Пиратка'];

var COMP_HEADERS = ['ИНН', 'Краткое название', 'Полное название', 'ОГРН', 'Регион', 'Статус', 'Источник', 'Обновлено'];
var LOG_HEADERS = ['Время', 'Пользователь', 'Действие', 'ID', 'Детали'];
var SALES_HEADERS = ['Дата регистрации', 'ИНН', 'MID', 'МПП', 'Приоритет', 'Тип привлечения', 'Продукт', 'Загружено'];
var PLAN_HEADERS = ['Месяц', 'ФИО', 'План сделок, шт', 'План объёма сделок, ₽', 'Изменено', 'Кто изменил'];
/** Подтверждённые продажи: снимки месяца. Исходные строки не смешиваются с активностями/MID. */
var CONFIRMED_HEADERS = ['Пакет', 'Месяц', 'Данные на дату', 'Сотрудник', 'ИНН', 'Канал поступления', 'Тип привлечения', 'Сегмент код', 'Приоритет код', 'Сегмент', 'Приоритет', 'Категория', 'Коды продуктов', 'Продукты'];
var CONFIRMED_BATCH_HEADERS = ['Пакет', 'Месяц', 'Данные на дату', 'Хеш', 'Строк', 'Загружено', 'Автор'];
var CONFIRMED_SHEET = 'Подтверждённые продажи';
var CONFIRMED_BATCH_SHEET = 'Загрузки подтверждений';

function confirmedKey_(v) { return str_(v).toLowerCase().replace(/ё/g, 'е').replace(/[\s_]+/g, ' '); }
function confirmedMonth_(v) {
  var s = str_(v);
  if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(s)) throw new Error('Укажите месяц продаж в формате ГГГГ-ММ');
  return s;
}
function confirmedEnd_(month) {
  return month + '-' + new Date(+month.slice(0, 4), +month.slice(5, 7), 0).getDate();
}

// TSV из буфера, CSV с ; или запятой, кавычки и многострочные ячейки.
function confirmedParseText_(text) {
  text = String(text || '').replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  var first = text.split('\n')[0], sep = first.indexOf('\t') >= 0 ? '\t' : first.indexOf(';') >= 0 ? ';' : ',';
  var rows = [], row = [], cell = '', quoted = false;
  for (var i = 0; i < text.length; i++) {
    var c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') { cell += '"'; i++; }
      else if (quoted || !cell) quoted = !quoted;
      else cell += c;
    } else if (!quoted && (c === sep || c === '\n')) {
      row.push(cell); cell = '';
      if (c === '\n') { rows.push(row); row = []; }
    } else cell += c;
  }
  if (quoted) throw new Error('Незакрытая кавычка в таблице');
  row.push(cell); rows.push(row);
  return rows;
}

function confirmedBatches_() {
  var sh = ss_().getSheetByName(CONFIRMED_BATCH_SHEET), latest = {};
  if (!sh || sh.getLastRow() < 2) return latest;
  sh.getRange(2, 1, sh.getLastRow() - 1, CONFIRMED_BATCH_HEADERS.length).getValues().forEach(function (r) {
    var month = monthKey_(r[1]);
    if (month && r[0]) latest[month] = { id: str_(r[0]), month: month, asOf: r[2] ? ymd_(r[2]) : '', hash: str_(r[3]), count: +r[4], at: stamp_(r[5]), author: str_(r[6]) };
  });
  return latest;
}
function readConfirmed_(batches) {
  batches = batches || confirmedBatches_();
  var sh = ss_().getSheetByName(CONFIRMED_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CONFIRMED_HEADERS.length).getValues().filter(function (r) {
    var b = batches[monthKey_(r[1])]; return b && b.id === str_(r[0]);
  }).map(function (r) {
    return { month: monthKey_(r[1]), asOf: r[2] ? ymd_(r[2]) : '', fio: str_(r[3]), inn: normInn_(r[4]), channel: str_(r[5]), attraction: str_(r[6]), segmentCode: str_(r[7]), priorityCode: str_(r[8]), segment: str_(r[9]), priority: str_(r[10]), category: str_(r[11]), productCodes: str_(r[12]), products: str_(r[13]) };
  });
}

function prepareConfirmed_(args) {
  var month = confirmedMonth_(args.month), asOf = str_(args.asOf);
  if (asOf && (!/^\d{4}-\d{2}-\d{2}$/.test(asOf) || ymd_(asOf) !== asOf || asOf < month + '-01' || asOf > confirmedEnd_(month) || asOf > ymd_(today_()))) {
    throw new Error('Дата снимка должна быть существующей датой выбранного месяца, не позднее сегодня');
  }
  if (month > ymd_(today_()).slice(0, 7)) throw new Error('Нельзя загрузить продажи будущего месяца');
  var input = args.rows || confirmedParseText_(args.text);
  if (!Array.isArray(input) || input.length > 10001 || JSON.stringify(input).length > 4000000) throw new Error('Лимит загрузки — 10 000 строк и 4 МБ текста');
  var aliases = {
    fio: ['сотрудник', 'фио', 'мпп', 'sales fio'], inn: ['инн', 'inn'], channel: ['канал поступления'], attraction: ['тип привлечения'],
    segmentCode: ['сегмент код'], priorityCode: ['приоритет код'], segment: ['сегмент'], priority: ['приоритет'], category: ['категория'], productCodes: ['коды продуктов'], products: ['продукты', 'продукт'], period: ['период', 'месяц']
  };
  var idx = {}, header = -1;
  for (var i = 0; i < Math.min(input.length, 20); i++) {
    if (!Array.isArray(input[i])) throw new Error('Ожидается таблица строк');
    var keys = input[i].map(confirmedKey_);
    if (keys.indexOf('инн') < 0 && keys.indexOf('inn') < 0) continue;
    Object.keys(aliases).forEach(function (k) {
      var found = [];
      keys.forEach(function (name, j) { if (aliases[k].indexOf(name) >= 0) found.push(j); });
      if (found.length > 1) throw new Error('Повторяющийся столбец: ' + k);
      if (found.length) idx[k] = found[0];
    });
    header = i; break;
  }
  if (header < 0 || idx.fio == null || idx.inn == null) throw new Error('Нужна строка заголовков со столбцами «Сотрудник» и «ИНН»');
  var users = {}, unknown = {}, seen = {}, data = [], errors = [], duplicates = 0;
  usersCached_().forEach(function (u) { users[confirmedKey_(u.fio)] = u.fio; });
  for (var n = header + 1; n < input.length; n++) {
    var row = input[n];
    if (!Array.isArray(row)) throw new Error('Неверная строка ' + (n + 1));
    if (!row.some(function (v) { return str_(v); })) continue;
    var get = function (k) { return idx[k] == null ? '' : str_(row[idx[k]]); };
    var sourceFio = get('fio').replace(/\s+/g, ' '), fio = users[confirmedKey_(sourceFio)] || storedName_(sourceFio), rawInn = get('inn');
    if (!users[confirmedKey_(fio)] && sourceFio) unknown[sourceFio] = true;
    var inn = /^[\d\s]+$/.test(rawInn) ? normInn_(rawInn) : '';
    if (fio.split(' ').length < 2 || !isInn_(inn)) { errors.push('Строка ' + (n + 1) + ': ' + (fio.split(' ').length < 2 ? 'укажите полное ФИО сотрудника' : 'некорректный ИНН')); continue; }
    var period = get('period');
    if (period) {
      var pm = /^(\d{2})[.\/](\d{4})$/.exec(period);
      if ((pm ? pm[2] + '-' + pm[1] : monthKey_(period)) !== month) { errors.push('Строка ' + (n + 1) + ': период отличается от выбранного месяца'); continue; }
    }
    var values = [fio, inn, get('channel'), get('attraction'), get('segmentCode'), get('priorityCode'), get('segment'), get('priority'), get('category'), get('productCodes'), get('products')];
    if (values.some(function (v) { return v.length > 1000; })) { errors.push('Строка ' + (n + 1) + ': ячейка длиннее 1000 символов'); continue; }
    var key = JSON.stringify(values);
    if (seen[key]) { duplicates++; continue; }
    seen[key] = true; data.push(values);
  }
  if (errors.length) throw new Error('Загрузка отменена, исправьте ошибки (' + errors.length + '):\n' + errors.slice(0, 15).join('\n'));
  if (!data.length) throw new Error('Нет строк продаж для загрузки');
  data.sort(function (a, b) { return JSON.stringify(a).localeCompare(JSON.stringify(b)); });
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify([month, asOf, data]), Utilities.Charset.UTF_8);
  var hash = digest.map(function (x) { return ('0' + ((x + 256) % 256).toString(16)).slice(-2); }).join('');
  var counts = {}, inns = {};
  data.forEach(function (r) { (counts[r[0]] || (counts[r[0]] = {}))[r[1]] = true; inns[r[1]] = true; });
  return { month: month, asOf: asOf, rows: data, hash: hash, unknown: Object.keys(unknown), duplicates: duplicates, clients: Object.keys(inns).length,
    managers: Object.keys(counts).map(function (f) { return { fio: f, confirmed: Object.keys(counts[f]).length }; }) };
}
function previewConfirmedSales_(u, args) {
  var p = prepareConfirmed_(args), current = confirmedBatches_()[p.month];
  return { month: p.month, asOf: p.asOf, rows: p.rows.length, clients: p.clients, unknown: p.unknown, duplicates: p.duplicates, managers: p.managers,
    expectedBatch: current ? current.id : '', previous: current || null, unchanged: !!current && current.hash === p.hash };
}
function importConfirmedSales_(u, args) {
  if (!u.isAdmin) throw new Error('Загрузка доступна только администратору');
  var p = prepareConfirmed_(args);
  return withLock_(function () {
    var current = confirmedBatches_()[p.month];
    if (current && current.hash === p.hash) return { unchanged: true, count: p.rows.length, clients: p.clients, month: p.month };
    if (args.expectedBatch == null || args.expectedBatch !== (current ? current.id : '')) throw new Error('Снимок месяца изменился. Выполните проверку загрузки заново');
    var ss = ss_(), sh = ss.getSheetByName(CONFIRMED_SHEET) || ensureSheet_(ss, CONFIRMED_SHEET, CONFIRMED_HEADERS);
    var batchSheet = ss.getSheetByName(CONFIRMED_BATCH_SHEET) || ensureSheet_(ss, CONFIRMED_BATCH_SHEET, CONFIRMED_BATCH_HEADERS);
    var id = Utilities.getUuid(), rows = p.rows.map(function (r) { r[0] = storedName_(r[0]); return [id, p.month, p.asOf].concat(r).map(safeText_); });
    var start = sh.getLastRow() + 1;
    ensureRows_(sh, start + rows.length - 1);
    sh.getRange(start, 1, rows.length, CONFIRMED_HEADERS.length).setNumberFormat('@').setValues(rows);
    SpreadsheetApp.flush();
    // Пакет становится виден только после полной записи строк. Старые снимки сохраняются для аудита.
    batchSheet.appendRow([id, "'" + p.month, p.asOf ? "'" + p.asOf : '', p.hash, rows.length, new Date(), safeText_(storedName_(u.fio))]);
    SpreadsheetApp.flush();
    log_(u, 'Подтверждённые продажи', id, p.month + ': ' + rows.length + ' строк, ' + p.clients + ' ИНН; на дату ' + (p.asOf || 'не задана'));
    return { count: rows.length, clients: p.clients, month: p.month, asOf: p.asOf };
  });
}

// Пятидневка РФ. Проверено по https://www.cbr.ru/other/holidays/ (02.11.2026).
// Для неизвестного календаря прогноз не выдаётся как точный.
function salesWorkdays_(month, asOf) {
  var year = month.slice(0, 4), extra = {
    '2025': ['05-02', '05-08', '06-13', '11-03', '12-31'],
    '2026': ['01-09', '03-09', '05-11', '12-31'],
    '2027': ['02-22', '05-03', '05-10', '06-14', '11-05', '12-31']
  };
  if (!extra[year]) return { total: null, passed: null, known: false };
  var holidays = ['01-01','01-02','01-03','01-04','01-05','01-06','01-07','01-08','02-23','03-08','05-01','05-09','06-12','11-04'].concat(extra[year]);
  var days = +confirmedEnd_(month).slice(8), total = 0, passed = 0;
  for (var d = 1; d <= days; d++) {
    var date = month + '-' + ('0' + d).slice(-2), dow = new Date(+year, +month.slice(5) - 1, d).getDay();
    var work = date === '2025-11-01' || (dow !== 0 && dow !== 6 && holidays.indexOf(date.slice(5)) < 0);
    if (work) { total++; if (asOf && date <= asOf) passed++; }
  }
  return { total: total, passed: asOf ? passed : null, known: true };
}

function confirmedComparison_(u, month, batches, confirmed, activities) {
  var b = batches[month], kinds = statusKinds_(), map = {}, types = typeByFio_();
  var rowFor = function (fio, inn) {
    var key = fio + '|' + inn;
    return map[key] || (map[key] = { month: month, fio: fio, inn: inn, short: '', manager: 0, confirmed: 0, sourceRows: 0, products: [], asOf: b ? b.asOf : '' });
  };
  activities.forEach(function (a) {
    if (a.date.slice(0, 7) !== month || kinds[a.status] !== 'deal') return;
    var r = rowFor(a.fio, a.inn); r.manager = 1; r.short = a.short || r.short;
  });
  confirmed.forEach(function (s) {
    if (s.month !== month) return;
    var r = rowFor(s.fio, s.inn); r.confirmed = 1; r.sourceRows++;
    s.products.split(';').map(str_).filter(Boolean).forEach(function (p) { if (r.products.indexOf(p) < 0) r.products.push(p); });
  });
  var all = Object.keys(map).map(function (k) { return map[k]; }).filter(function (r) {
    return u.canReport || canSee_(u, r);
  }), byFio = {}, totals = { manager: {}, confirmed: {} };
  all.forEach(function (r) {
    var m = byFio[r.fio] || (byFio[r.fio] = { fio: r.fio, type: types[r.fio] || '', manager: 0, confirmed: 0 });
    m.manager += r.manager; m.confirmed += r.confirmed;
    if (r.manager) totals.manager[r.inn] = true;
    if (r.confirmed) totals.confirmed[r.inn] = true;
  });
  var wd = salesWorkdays_(month, b ? b.asOf : ''), forecast = function (n) { return b && b.asOf && wd.known && wd.passed > 0 ? Math.round(n / wd.passed * wd.total * 10) / 10 : null; };
  var managers = Object.keys(byFio).sort().map(function (f) { var m = byFio[f]; m.forecast = forecast(m.confirmed); return m; });
  return { month: month, loaded: !!b, asOf: b ? b.asOf : '', at: b ? b.at : '',
    sourceRows: u.canReport || u.seeAll ? (b ? b.count : 0) : all.reduce(function (n, r) { return n + r.sourceRows; }, 0), workdays: wd,
    manager: Object.keys(totals.manager).length, confirmed: b ? Object.keys(totals.confirmed).length : null,
    forecast: forecast(Object.keys(totals.confirmed).length), managers: managers,
    details: all.sort(function (a, b) { return a.fio.localeCompare(b.fio) || a.inn.localeCompare(b.inn); }) };
}
function confirmedSalesReport_(u, args) {
  var month = confirmedMonth_(args.month), batches = confirmedBatches_();
  return confirmedComparison_(u, month, batches, readConfirmed_(batches), readActivities_());
}
/** Справочники: каналы, статусы, виды активности. Хранятся на листе «Справочники». */

function getDicts_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dicts3');
  if (cached) return JSON.parse(cached);

  var sh = sheet_(APP.SHEET.DICT);
  var d = readDictSheet_(sh);
  if (d.statuses.length && !d.hasActiveCol) {
    migrateActiveStatuses_(d);
    d = readDictSheet_(sh);
  }
  if (Number(prop_('DICT_VERSION') || 0) < DICT_VERSION) {
    syncDefaultStatuses_(d);
    d = readDictSheet_(sh);
  }
  delete d.hasActiveCol;
  if (!d.channels.length) d.channels = DEFAULT_DICTS.channels.slice();
  if (!d.statuses.length) d.statuses = DEFAULT_DICTS.statuses.slice();
  if (!d.types.length) d.types = DEFAULT_DICTS.types.slice();

  cache.put('dicts3', JSON.stringify(d), 600);
  return d;
}

// Пустая ячейка «доступен для выбора» означает «Да»: статус, дописанный в таблицу руками, сразу доступен.
// Если колонка пуста у всех статусов, справочник ещё в старом формате (см. migrateActiveStatuses_).
function readDictSheet_(sh) {
  var kindByLabel = {};
  Object.keys(KIND_LABEL).forEach(function (k) { kindByLabel[KIND_LABEL[k].toLowerCase()] = k; });

  var last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
  var d = { channels: [], statuses: [], types: [], hasActiveCol: false };
  vals.forEach(function (r) {
    if (str_(r[0])) d.channels.push(str_(r[0]));
    if (str_(r[1])) {
      if (str_(r[4])) d.hasActiveCol = true;
      d.statuses.push({
        name: str_(r[1]),
        kind: kindByLabel[str_(r[2]).toLowerCase()] || 'open',
        active: !/^нет$/i.test(str_(r[4]))
      });
    }
    if (str_(r[3])) d.types.push(str_(r[3]));
  });
  return d;
}

// Разовый переход на колонку «доступен для выбора»: доступными остаются только статусы из DEFAULT_DICTS,
// остальные сохраняются для старых записей и отчётов. Типы статусов, заданные в таблице, не меняются.
function migrateActiveStatuses_(d) {
  var allowed = {}, have = {};
  DEFAULT_DICTS.statuses.forEach(function (s) { allowed[s.name] = true; });
  d.statuses.forEach(function (s) {
    s.active = !!allowed[s.name];
    have[s.name] = true;
  });
  DEFAULT_DICTS.statuses.forEach(function (s) {
    if (!have[s.name]) d.statuses.push({ name: s.name, kind: s.kind, active: true });
  });
  writeDicts_(d);
}

// Статусы из DEFAULT_DICTS получают тип и доступность из кода, недостающие дописываются в конец.
// Остальные статусы справочника не трогаем. Версию записываем до writeDicts_: он снова вызывает getDicts_.
function syncDefaultStatuses_(d) {
  setProp_('DICT_VERSION', DICT_VERSION);
  var byName = {};
  d.statuses.forEach(function (s) { byName[s.name] = s; });
  DEFAULT_DICTS.statuses.forEach(function (def) {
    var s = byName[def.name];
    if (s) {
      s.kind = def.kind;
      s.active = def.active;
    } else {
      d.statuses.push({ name: def.name, kind: def.kind, active: def.active });
    }
  });
  writeDicts_(d);
}

function activeStatuses_(d) {
  return (d || getDicts_()).statuses.filter(function (s) { return s.active !== false; });
}

function statusKinds_(d) {
  var m = {};
  (d || getDicts_()).statuses.forEach(function (s) { m[s.name] = s.kind; });
  return m;
}

function writeDicts_(d) {
  var sh = sheet_(APP.SHEET.DICT);
  var n = Math.max(d.channels.length, d.statuses.length, d.types.length);
  var rows = [];
  for (var i = 0; i < n; i++) {
    var s = d.statuses[i];
    rows.push([
      d.channels[i] || '',
      s ? s.name : '',
      s ? KIND_LABEL[s.kind] || KIND_LABEL.open : '',
      d.types[i] || '',
      s ? (s.active === false ? 'Нет' : 'Да') : ''
    ]);
  }
  var w = DICT_HEADERS.length;
  sh.getRange(1, 1, 1, w).setValues([DICT_HEADERS]).setFontWeight('bold');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, w).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, w).setValues(rows);
  CacheService.getScriptCache().remove('dicts3');
  applyValidations_();
}

function saveDicts_(user, args) {
  var clean = function (list) {
    var seen = {};
    return (list || []).map(str_).filter(function (x) {
      if (!x || seen[x]) return false;
      seen[x] = true;
      return true;
    });
  };
  var d = {
    channels: clean(args.channels),
    types: clean(args.types),
    statuses: (args.statuses || []).map(function (s) {
      return { name: str_(s.name), kind: KIND_LABEL[s.kind] ? s.kind : 'open', active: s.active !== false };
    }).filter(function (s) { return s.name; })
  };
  if (!d.channels.length || !d.types.length || !d.statuses.length) {
    throw new Error('Каждый справочник должен содержать хотя бы одно значение');
  }
  if (!activeStatuses_(d).length) throw new Error('Оставьте доступным для выбора хотя бы один статус');
  writeDicts_(d);
  log_(user, 'Справочники изменены');
  return getDicts_();
}

// Выпадающие списки и проверка дат прямо в таблице — на случай правки руками.
// В статусе — только доступные для выбора; старые значения в строках остаются, таблица лишь помечает их.
function applyValidations_() {
  var d = getDicts_();
  var sh = sheet_(APP.SHEET.ACT);
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  var list = function (values) {
    return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  };
  sh.getRange(2, COL.CHANNEL, rows, 1).setDataValidation(list(d.channels));
  sh.getRange(2, COL.STATUS, rows, 1).setDataValidation(list(activeStatuses_(d).map(function (s) { return s.name; })));
  sh.getRange(2, COL.TYPE, rows, 1).setDataValidation(list(d.types));
  var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
    .setHelpText('Введите дату в формате ДД.ММ.ГГГГ').build();
  sh.getRange(2, COL.DATE, rows, 1).setDataValidation(dateRule);
  sh.getRange(2, COL.NEXT, rows, 1).setDataValidation(dateRule);
}
/** Из PS сохраняется только признак наличия оборота; PL/US не импортируются. */
function importTurnovers_(force) {
  return withLock_(function () { return runImport_(force === true); });
}
function runImport_(force) {
  if (prop_('TURN_PRIVACY_VERSION') !== '1') throw new Error('Сначала выполните очистку прежних оборотов');
  var file = latestFile_(DriveApp.getFolderById(prop_('DATA_FOLDER_ID') || APP.DATA_FOLDER_ID), 'PS');
  if (!file) throw new Error('В папке нет файла PS');
  var sig = 'presence-v1:' + file.getId() + ':' + file.getLastUpdated().getTime();
  if (!force && sig === prop_('IMPORT_SIG')) return 'Признак «Пиратка» уже актуален';
  var map = parsePs_(readXlsx_(file)), sh = sheet_(APP.SHEET.TURN);
  var rows = Object.keys(map).sort().map(function (inn) { return [inn, map[inn] ? 'да' : 'нет']; });
  var old = sh.getLastRow();
  ensureRows_(sh, rows.length + 1);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  if (old > rows.length + 1) sh.getRange(rows.length + 2, 1, old - rows.length - 1, 2).clearContent();
  sh.getRange(1, 1, 1, 2).setValues([TURN_HEADERS]);
  _pirates = map;
  refreshPirates_(map);
  var msg = 'Признак «Пиратка» обновлён: ' + rows.length + ' ИНН';
  setProp_('IMPORT_SIG', sig); setProp_('IMPORT_AT', new Date().toISOString()); setProp_('IMPORT_MSG', msg);
  log_(null, 'Обновление признака Пиратка', '', msg);
  return msg;
}

function latestFile_(folder, prefix) {
  var re = new RegExp('^' + prefix + '_\\d+', 'i');
  var it = folder.getFiles(), best = null;
  while (it.hasNext()) {
    var f = it.next();
    if (!re.test(f.getName())) continue;
    var mime = f.getMimeType();
    if (mime !== MimeType.GOOGLE_SHEETS && !/\.xlsx?$/i.test(f.getName())) continue;
    if (!best || f.getDateCreated() > best.getDateCreated()) best = f;
  }
  return best;
}

// xlsx конвертируется во временную Google-таблицу, читается и сразу удаляется.
function readXlsx_(file) {
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues();
  }
  var tmp = Drive.Files.create(
    { name: 'tmp_import_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [tempFolder_().getId()] },
    file.getBlob()
  );
  try {
    return SpreadsheetApp.openById(tmp.id).getSheets()[0].getDataRange().getValues();
  } finally {
    Drive.Files.remove(tmp.id);
  }
}

function headerKey_(v) {
  return str_(v).toLowerCase().replace(/\\/g, '');
}

function findHeaderRow_(values, test) {
  for (var i = 0; i < Math.min(values.length, 30); i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (test(headerKey_(values[i][j]))) return { row: i, col: j };
    }
  }
  return null;
}

function parsePs_(values) {
  var h = findHeaderRow_(values, function (c) { return c === 'inn' || c === 'инн'; });
  if (!h) throw new Error('В файле PS не найдена колонка inn');
  var idx = {}; values[h.row].forEach(function (c, j) { idx[headerKey_(c)] = j; });
  if (idx.vol == null) throw new Error('В файле PS не найдена колонка vol');
  var totals = {}, flags = {};
  for (var i = h.row + 1; i < values.length; i++) {
    var inn = normInn_(values[i][h.col]);
    if (isInn_(inn)) totals[inn] = (totals[inn] || 0) + num_(values[i][idx.vol]);
  }
  Object.keys(totals).forEach(function (inn) { flags[inn] = totals[inn] > 0; });
  return flags;
}
function pirateValue_(v) {
  return /^(да|нет|true|false)$/i.test(str_(v)) ? yes_(v) : num_(v) > 0;
}
var _pirates = null;
function pirateMap_() {
  if (_pirates) return _pirates;
  var sh = sheet_(APP.SHEET.TURN), map = {}, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    var inn = normInn_(r[0]);
    if (isInn_(inn)) map[inn] = pirateValue_(r[1]);
  });
  return (_pirates = map);
}
function pirateFor_(inn) { return pirateMap_()[inn] ? 'да' : 'нет'; }
function refreshPirates_(map) {
  var sh = sheet_(APP.SHEET.ACT), last = sh.getLastRow();
  if (last < 2) return;
  var flags = sh.getRange(2, COL.INN, last - 1, 1).getValues().map(function (r) {
    return [map[normInn_(r[0])] ? 'да' : 'нет'];
  });
  sh.getRange(2, COL.PIRATE, flags.length, 1).clearDataValidations().setNumberFormat('@').setValues(flags);
  sh.getRange(2, COL.POT_SRC, flags.length, 1).clearContent();
  _acts = null;
}
function importStatus_() {
  return { at: prop_('IMPORT_AT') || '', msg: prop_('TURN_PRIVACY_VERSION') === '1' ? (prop_('IMPORT_MSG') || 'Сохранён только признак «Пиратка»') : 'Требуется очистка прежних оборотов' };
}

// Явная однократная миграция администратором; повтор безопасен, персональные ссылки не меняет.
function migratePrivacy_(u) {
  if (!u.isAdmin) throw new Error('Недостаточно прав');
  // Ключ создаём до основной блокировки.
  privacySecret_();
  return withLock_(function () {
    var sh = sheet_(APP.SHEET.TURN), vals = sh.getDataRange().getValues();
    var width = Math.max(sh.getLastColumn(), 2), map = {}, rows = [];
    vals.slice(1).forEach(function (r) {
      var inn = normInn_(r[0]);
      if (isInn_(inn)) map[inn] = pirateValue_(r[1]);
    });
    var header = TURN_HEADERS.slice(); while (header.length < width) header.push('');
    rows.push(header);
    Object.keys(map).sort().forEach(function (inn) {
      var r = [inn, map[inn] ? 'да' : 'нет']; while (r.length < width) r.push(''); rows.push(r);
    });
    while (rows.length < vals.length) rows.push(Array(width).fill(''));
    sh.getRange(1, 1, rows.length, width).clearDataValidations().setNumberFormat('@').setValues(rows);
    _pirates = map;
    refreshPirates_(map);
    var act = sheet_(APP.SHEET.ACT);
    act.getRange(1, COL.PIRATE).setValue('Пиратка'); act.getRange(1, COL.POT_SRC).setValue('Резерв');
    SpreadsheetApp.flush();
    var invalid = sh.getRange(2, 2, Math.max(sh.getLastRow() - 1, 1), 1).getValues().some(function (r) { return r[0] && r[0] !== 'да' && r[0] !== 'нет'; });
    if (invalid) throw new Error('Проверка очистки не пройдена');
    ['PS_FILE', 'SBER_FILE', 'IMPORT_SIG'].forEach(function (k) { setProp_(k, ''); });
    setProp_('IMPORT_MSG', 'Сохранён только признак «Пиратка»: да/нет');
    setProp_('TURN_PRIVACY_VERSION', '1');
    log_(u, 'Удалены импортированные обороты', '', 'Оставлен признак Пиратка; ' + Object.keys(map).length + ' ИНН');
    return { migrated: true, clients: Object.keys(map).length, pirate: Object.keys(map).filter(function (k) { return map[k]; }).length };
  });
}
/** Планы сотрудников на месяц: количество сделок и их объём. Хранятся на листе «Планы». */

function plansSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(APP.SHEET.PLANS);
  if (sh) return sh;
  sh = ensureSheet_(ss, APP.SHEET.PLANS, PLAN_HEADERS);
  sh.getRange('A:A').setNumberFormat('@'); // «2026-10» иначе станет датой
  return sh;
}

// Месяц в ячейке — текст «yyyy-MM»; если его всё же превратили в дату, читаем и её.
function monthKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, 'yyyy-MM');
  var m = /^(\d{4})-(\d{2})/.exec(str_(v));
  return m ? m[1] + '-' + m[2] : '';
}

function prevMonth_(month) {
  var y = +month.slice(0, 4), m = +month.slice(5, 7) - 1;
  return Utilities.formatDate(new Date(y, m - 1, 1), APP.TZ, 'yyyy-MM');
}

// { 'yyyy-MM': { fio: { deals, volume } } }
function readPlans_() {
  var sh = plansSheet_(), last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  var plans = {};
  vals.forEach(function (r) {
    var month = monthKey_(r[0]), fio = str_(r[1]);
    if (!month || !fio) return;
    (plans[month] || (plans[month] = {}))[fio] = { deals: num_(r[2]), volume: num_(r[3]) };
  });
  return plans;
}

// Сотрудники, которым ставится план: активные с ролью «Сотрудник», плюс те, у кого план уже есть.
function getPlans_(u, args) {
  var month = monthKey_(args.month);
  if (!month) throw new Error('Укажите месяц');
  var plans = readPlans_();
  var cur = plans[month] || {}, prev = plans[prevMonth_(month)] || {};
  var staff = usersCached_().filter(function (x) { return x.active && x.role === ROLES.EMPLOYEE; });
  var seen = {};
  var rows = staff.map(function (x) {
    seen[x.fio] = true;
    return { fio: x.fio, type: x.type, no: x.no };
  });
  Object.keys(cur).forEach(function (fio) { if (!seen[fio]) rows.push({ fio: fio, type: '', no: '' }); });
  rows.forEach(function (r) {
    r.deals = cur[r.fio] ? cur[r.fio].deals : '';
    r.volume = cur[r.fio] ? cur[r.fio].volume : '';
    r.prevDeals = prev[r.fio] ? prev[r.fio].deals : '';
    r.prevVolume = prev[r.fio] ? prev[r.fio].volume : '';
  });
  rows.sort(function (a, b) { return a.fio.localeCompare(b.fio); });
  return { month: month, rows: rows };
}

// Месяц перезаписывается целиком; строки с нулевым планом не храним.
function savePlans_(u, args) {
  var month = monthKey_(args.month);
  if (!month) throw new Error('Укажите месяц');
  var now = new Date();
  var fresh = (args.rows || []).map(function (r) {
    return { fio: storedName_(r.fio), deals: Math.max(0, Math.round(num_(r.deals))), volume: Math.max(0, Math.round(num_(r.volume))) };
  }).filter(function (r) { return r.fio && (r.deals || r.volume); });

  withLock_(function () {
    var sh = plansSheet_(), last = sh.getLastRow(), w = PLAN_HEADERS.length;
    var keep = (last > 1 ? sh.getRange(2, 1, last - 1, w).getValues() : []).filter(function (r) {
      return monthKey_(r[0]) && monthKey_(r[0]) !== month;
    }).map(function (r) {
      r[0] = monthKey_(r[0]);
      return r;
    });
    var rows = keep.concat(fresh.map(function (r) { return [month, storedName_(r.fio), r.deals, r.volume, now, storedName_(u.fio)]; }));
    rows.sort(function (a, b) { return a[0].localeCompare(b[0]) || String(a[1]).localeCompare(String(b[1])); });
    if (last > 1) sh.getRange(2, 1, last - 1, w).clearContent();
    if (rows.length) {
      ensureRows_(sh, rows.length + 1);
      sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, rows.length, w).setValues(rows);
      sh.getRange(2, 3, rows.length, 2).setNumberFormat('#,##0');
      sh.getRange(2, 5, rows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    }
  });
  log_(u, 'Планы изменены', '', month + ': ' + fresh.length + ' сотрудников');
  return getPlans_(u, { month: month });
}

// План на период: месячный план умножается на долю дней месяца, попавших в период.
// Неделя получает 7/30 месячного плана, целый месяц — весь план.
function planForRange_(from, to) {
  var plans = readPlans_(), byFio = {}, months = [], partial = false;
  Object.keys(plans).sort().forEach(function (month) {
    var y = +month.slice(0, 4), m = +month.slice(5, 7);
    var dim = new Date(y, m, 0).getDate();
    var first = month + '-01', last = month + '-' + (dim < 10 ? '0' : '') + dim;
    var s = from > first ? from : first, e = to < last ? to : last;
    if (s > e) return;
    var share = (+e.slice(8, 10) - +s.slice(8, 10) + 1) / dim;
    if (share < 1) partial = true;
    months.push(month);
    Object.keys(plans[month]).forEach(function (fio) {
      var p = plans[month][fio], o = byFio[fio] || (byFio[fio] = { deals: 0, volume: 0 });
      o.deals += p.deals * share;
      o.volume += p.volume * share;
    });
  });
  Object.keys(byFio).forEach(function (fio) {
    byFio[fio].deals = Math.round(byFio[fio].deals * 10) / 10;
    byFio[fio].volume = Math.round(byFio[fio].volume);
  });
  return { byFio: byFio, months: months, partial: partial };
}

// Выполнение плана за месяц — видят все сотрудники.
// Продажи (шт) — из листа «Продажи» по дате регистрации; продажи без МПП идут строкой с пустым ФИО.
// Объём (₽) — из активностей: уникальные ИНН со статусом типа «Сделка», сумма по последней такой активности клиента.
// Рабочие дни (пн–пт, без праздников) нужны интерфейсу для прогноза RR и расчёта «сколько нужно в день».
function planProgress_(u, args) {
  var month = monthKey_(args.month) || Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM');
  var y = +month.slice(0, 4), m = +month.slice(5, 7), dim = new Date(y, m, 0).getDate();
  var from = month + '-01', to = month + '-' + (dim < 10 ? '0' : '') + dim;
  var plan = readPlans_()[month] || {};
  var kinds = statusKinds_();

  // Первая загрузка продаж — сразу, не дожидаясь ночного запуска; заодно ставится ежедневный триггер.
  if (!prop_('SALES_AT')) {
    try { importSales_(); } catch (e) { console.error('importSales: ' + e.message); }
  }
  var sales = {};
  readSales_().forEach(function (x) {
    if (x.date >= from && x.date <= to) sales[x.fio] = (sales[x.fio] || 0) + 1;
  });

  var deals = {};
  readActivities_().forEach(function (a) {
    if (a.date < from || a.date > to || kinds[a.status] !== 'deal') return;
    var byInn = deals[a.fio] || (deals[a.fio] = {});
    var cur = byInn[a.inn];
    if (!cur || a.date + a.created > cur.date + cur.created) byInn[a.inn] = a;
  });

  var today = ymd_(today_()), workdays = { total: 0, passed: 0 };
  for (var d = 1; d <= dim; d++) {
    var dt = new Date(y, m - 1, d), dow = dt.getDay();
    if (dow === 0 || dow === 6) continue;
    workdays.total++;
    if (ymd_(dt) <= today) workdays.passed++;
  }

  var types = typeByFio_(), fios = {};
  Object.keys(plan).concat(Object.keys(deals), Object.keys(sales)).forEach(function (f) {
    if (u.canReport || canSee_(u, { fio: f })) fios[f] = true;
  });
  if (u.role === ROLES.EMPLOYEE) fios[u.fio] = true;
  var rows = Object.keys(fios).map(function (fio) {
    var list = Object.keys(deals[fio] || {}).map(function (inn) { return deals[fio][inn]; });
    return {
      fio: fio,
      type: types[fio] || '',
      plan: plan[fio] || null,
      fact: { sales: sales[fio] || 0, volume: list.reduce(function (s, a) { return s + (a.dealSum || 0); }, 0) }
    };
  });
  return { month: month, today: today, workdays: workdays, rows: rows, sales: salesStatus_() };
}
/** Сокращённые имена и непрозрачные ссылки на сотрудников для всех ролей. */
function displayName_(fio, surnameLength) {
  var saved = storedIdentityIndex_().refs[str_(fio)];
  if (saved || / \[[A-Za-z0-9_-]{8}\]$/.test(str_(fio))) return str_(fio).replace(/ \[[^\]]+\]$/, '');
  var p = str_(fio).split(/\s+/);
  return p.length > 1 ? p[1] + ' ' + p[0].slice(0, surnameLength || 1) : (p[0] ? p[0].charAt(0) : '');
}
function displayNames_(names) {
  var labels = Object.create(null);
  names.forEach(function (fio) { labels[fio] = displayName_(fio, 3); });
  return labels;
}

// Только свободный текст. ИНН и названия компаний в отдельных полях не меняются.
function privateComment_(text, names, labels) {
  var value = storedText_(String(text == null ? '' : text));
  var escape = function (s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); };
  (names || []).forEach(function (name) {
    var p = str_(name).split(/\s+/), variants = [name];
    if (p.length >= 2) {
      variants.push(p[0] + ' ' + p[1], p[1] + ' ' + p[0]);
      if (p.length >= 3) variants.push(p[1] + ' ' + p[2] + ' ' + p[0]);
    }
    variants.sort(function (a, b) { return b.length - a.length; }).forEach(function (v) {
      var re = new RegExp('(^|[^А-Яа-яЁёA-Za-z])' + v.split(/\s+/).map(escape).join('\\s+') + '(?=$|[^А-Яа-яЁёA-Za-z])', 'gi');
      value = value.replace(re, function (_, before) { return before + labels[name]; });
    });
  });
  // Инициалы известных сотрудников также заменяются подписью.
  (names || []).forEach(function (name) {
    var p = str_(name).split(/\s+/);
    if (p.length < 3) return;
    var initials = escape(p[1][0]) + '\\.\\s*' + escape(p[2][0]) + '\\.';
    var re = new RegExp('(^|[^А-Яа-яЁё])(?:' + escape(p[0]) + '\\s+' + initials + '|' + initials + '\\s*' + escape(p[0]) + ')(?=$|[^А-Яа-яЁё])', 'gi');
    value = value.replace(re, function (_, before) { return before + labels[name]; });
  });
  // Неизвестные справочнику ФИО: три слова с русским отчеством. Это не распознавание любого имени.
  value = value.replace(/(^|[^А-Яа-яЁё])([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]*(?:вич|вна|ична|ич))\s+([А-ЯЁ][а-яё-]+)(?=$|[^А-Яа-яЁё])/gi,
    function (_, before, first, patronymic, surname) { return before + first + ' ' + surname.slice(0, 3); });
  value = value.replace(/(^|[^А-Яа-яЁё])([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]+)\s+([А-ЯЁ][а-яё-]*(?:вич|вна|ична|ич))(?=$|[^А-Яа-яЁё])/gi,
    function (_, before, surname, first) { return before + first + ' ' + surname.slice(0, 3); });
  value = value.replace(/[A-ZА-ЯЁ0-9.!#$%&'*+/=?^_`{|}~-]+@[A-ZА-ЯЁ0-9-]+(?:\.[A-ZА-ЯЁ0-9-]+)+/gi, '[email скрыт]');
  value = value.replace(/(^|[^\d])(?:\+?7|8)[\s().-]*\d{3}[\s().-]*\d{3}[\s().-]*\d{2}[\s().-]*\d{2}(?!\d)/g, '$1[телефон скрыт]');
  value = value.replace(/\+\d(?:[ ().-]*\d){9,14}(?!\d)/g, '[телефон скрыт]');
  value = value.replace(/(^|[^\d])\(?\d{3}\)?[ .-]+\d{3}[ .-]+\d{2}[ .-]+\d{2}(?!\d)/g, '$1[телефон скрыт]');
  value = value.replace(/((?:тел(?:ефон)?|моб(?:ильный)?|phone|whatsapp|ватсап)\.?\s*:?\s*)\d{10}(?!\d)/gi, '$1[телефон скрыт]');
  value = value.replace(/([?&]k=)[a-f0-9]{32,}/gi, '$1[ссылка скрыта]');
  return value;
}
var _privacySecret = null;
function privacySecret_() {
  if (_privacySecret) return _privacySecret;
  var key = prop_('PERSON_REF_SECRET');
  if (key) return (_privacySecret = key);
  return withLock_(function () {
    var saved = prop_('PERSON_REF_SECRET');
    if (!saved) { saved = newToken_(); setProp_('PERSON_REF_SECRET', saved); }
    return (_privacySecret = saved);
  });
}
function rawPersonRef_(fio) {
  var bytes = Utilities.computeHmacSha256Signature(str_(fio), privacySecret_(), Utilities.Charset.UTF_8);
  return 'person_' + Utilities.base64EncodeWebSafe(bytes).replace(/=/g, '').slice(0, 24);
}
function personRef_(fio) { return storedIdentityIndex_().refs[str_(fio)] || rawPersonRef_(fio); }
function identityNames_(data) {
  var names = Object.create(null);
  var add = function (s) { if (typeof s === 'string' && str_(s).split(/\s+/).length >= 2) names[s] = true; };
  usersCached_().forEach(function (u) { add(u.fio); add(u.no); (u.extra || []).forEach(add); });
  var scan = function (v, field) {
    if (Array.isArray(v)) { v.forEach(function (x) { scan(x, field); }); return; }
    if (v && typeof v === 'object') {
      Object.keys(v).forEach(function (k) {
        if (field === 'byFio') add(k);
        scan(v[k], k);
      });
    } else if (['fio', 'no', 'author', 'fios', 'idle', 'extra', 'unknown'].indexOf(field) >= 0) add(v);
  };
  scan(data, '');
  return Object.keys(names);
}
function privacyResult_(u, data) {
  var refs = Object.create(null), people = Object.create(null);
  var names = identityNames_(data).sort(function (a, b) { return b.length - a.length; });
  var labels = displayNames_(names);
  names.forEach(function (f) { var ref = personRef_(f); refs[f] = ref; people[ref] = labels[f]; });
  var transform = function (v, field) {
    if (typeof v === 'string') {
      if (['inn', 'short', 'full'].indexOf(field) >= 0) return v;
      if (/comment|note|error|message/i.test(field || '')) return privateComment_(v, names, labels);
      if (refs[v]) return refs[v];
      // В свободном тексте заменяем известные ФИО подписью, не служебным ID.
      names.forEach(function (name) { if (v.indexOf(name) >= 0) v = v.split(name).join(labels[name]); });
      return v;
    }
    if (Array.isArray(v)) return v.map(function (x) { return transform(x, field); });
    if (v && typeof v === 'object') {
      var out = {};
      Object.keys(v).forEach(function (k) { out[refs[k] || k] = transform(v[k], k); });
      return out;
    }
    return v;
  };
  var safe = transform(data);
  // Отдаём только подписи идентификаторов, которые действительно присутствуют в ответе.
  var json = JSON.stringify(safe), used = {};
  Object.keys(people).forEach(function (ref) { if (json.indexOf(ref) >= 0) used[ref] = people[ref]; });
  return { data: safe, people: used };
}
function privateArgs_(u, args) {
  if (JSON.stringify(args).indexOf('person_') < 0) return args;
  var payload = { acts: readActivities_(), plans: { byFio: {} } };
  var plans = readPlans_();
  Object.keys(plans).forEach(function (m) { Object.keys(plans[m]).forEach(function (f) { payload.plans.byFio[f] = true; }); });
  var reverse = {};
  identityNames_(payload).forEach(function (f) { reverse[personRef_(f)] = f; });
  var convert = function (v, field) {
    if (typeof v === 'string' && ['fio', 'no', 'extra'].indexOf(field) >= 0 && /^person_/.test(v)) {
      if (!reverse[v]) throw new Error('Сотрудник не найден. Обновите страницу');
      return reverse[v];
    }
    if (Array.isArray(v)) return v.map(function (x) { return convert(x, field); });
    if (v && typeof v === 'object') {
      var out = {}; Object.keys(v).forEach(function (k) { out[k] = convert(v[k], k); }); return out;
    }
    return v;
  };
  return convert(args, '');
}
/** Названия из общедоступных HTML-карточек egrul.org. JSON/XML API не используется. */
function registryCompanyFromHtml_(inn, html) {
  var re = /<script\b[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script\s*>/gi;
  var match, nodes = [];
  function collect(x) {
    if (!x || typeof x !== 'object') return;
    if (Array.isArray(x)) { x.forEach(collect); return; }
    nodes.push(x); if (x['@graph']) collect(x['@graph']);
  }
  while ((match = re.exec(html))) {
    try { collect(JSON.parse(match[1])); } catch (e) { /* повреждённый блок не используем */ }
  }
  for (var i = 0; i < nodes.length; i++) {
    var x = nodes[i], id = x.identifier, ogrn = '', short = '', full = '';
    if (x['@type'] === 'Organization' && String(x.taxID || '') === inn && inn.length === 10) {
      short = str_(x.name); full = str_(x.legalName);
      ogrn = /^\d{13}$/.test(String(id)) ? String(id) : '';
    } else if (x['@type'] === 'Person' && inn.length === 12 && id && id.propertyID === 'ИНН' && String(id.value) === inn) {
      var ip = x.identifier2;
      // Физлицо без ОГРНИП не выдаём за индивидуального предпринимателя.
      if (!ip || ip.propertyID !== 'ОГРНИП' || !/^\d{15}$/.test(String(ip.value)) || !str_(x.name)) continue;
      ogrn = String(ip.value); short = 'ИП ' + str_(x.name); full = 'Индивидуальный предприниматель ' + str_(x.name);
    }
    if (short && full && short.length <= 2000 && full.length <= 2000) {
      return { inn: inn, short: short, full: full, ogrn: ogrn, source: 'egrul.org', region: '', status: '' };
    }
  }
  return null;
}
function companyFromPublicRegistry_(inn) {
  if (!isInn_(inn)) return null;
  var day = ymd_(today_());
  if (prop_('PUBLIC_NAMES_PAUSED_DAY') === day) throw new Error('egrul.org ограничил доступ. Повторите поиск завтра');
  var response = UrlFetchApp.fetch('https://egrul.org/' + inn, {
    method: 'get', headers: { Accept: 'text/html' }, muteHttpExceptions: true, followRedirects: true
  });
  var code = response.getResponseCode(), html = response.getContentText('UTF-8');
  if (code === 404) return null;
  var title = (/<title[^>]*>([\s\S]*?)<\/title>/i.exec(html) || [])[1] || '';
  if (code === 403 || code === 429 || /captcha|access denied|too many requests|превышен.*лимит|лимит.*исчерпан/i.test(title)) {
    setProp_('PUBLIC_NAMES_PAUSED_DAY', day);
    throw new Error('egrul.org ограничил доступ (HTTP ' + code + '). Повторите поиск завтра');
  }
  if (code !== 200) throw new Error('egrul.org временно недоступен: HTTP ' + code);
  return registryCompanyFromHtml_(inn, html);
}

// Только явное действие пользователя. Данные возвращаются в форму; активность сохраняется обычной кнопкой.
function lookupCompanyOnline_(u, args) {
  if (args.manual !== true) throw new Error('Нажмите «Найти название по ИНН» в форме активности');
  var inn = str_(args.inn);
  if (!isInn_(inn)) throw new Error('ИНН должен содержать 10 или 12 цифр');
  var visible = readActivities_().filter(function (a) { return a.inn === inn && canSee_(u, a); });
  var cached = companyFromCache_(inn);
  if (cached && cached.source !== 'egrul.org' && !visible.length && !u.seeAll) cached = null;
  var previous = visible.filter(function (a) { return a.short || a.full; })[0], local = null;
  if (cached || previous) {
    local = { inn: inn, short: (cached && cached.short) || (previous && previous.short) || '',
      full: (cached && cached.full) || (previous && previous.full) || '', source: 'Сохранённые данные' };
    if (local.short && local.full) return { inn: inn, company: local, source: local.source, fromCache: true };
  }
  try {
    var c = companyFromPublicRegistry_(inn);
    if (c) rememberCompany_(c);
    return { inn: inn, company: c || local, source: c ? c.source : 'Сохранённые данные' };
  } catch (e) {
    return { inn: inn, company: local, source: 'Сохранённые данные', unavailable: true,
      message: 'Внешний источник сейчас ограничивает поиск или недоступен. Сохранённые названия доступны; остальные можно внести вручную или проверить в ФНС.' };
  }
}
function companyNamesBackgroundTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function (t) {
    return ['fillMissingCompanyNames', 'continueCompanyNames'].indexOf(t.getHandlerFunction()) >= 0;
  });
}
function disableCompanyNamesBackground_() {
  return withLock_(function () {
    var list = companyNamesBackgroundTriggers_();
    list.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    setProp_('PUBLIC_COMPANY_NAMES', 'нет');
    return { mode: 'manual', removed: list.length, backgroundTriggers: companyNamesBackgroundTriggers_().length };
  });
}
/** Отчёты за день, неделю, месяц или произвольный период во всех разрезах. */

var MONTHS_RU = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];

var REPORT_GROUPS = [
  { key: 'fio', title: 'По сотрудникам' },
  { key: 'no', title: 'По начальникам отделов' },
  { key: 'empType', title: 'По типам сотрудников' },
  { key: 'channel', title: 'По каналам поступления' },
  { key: 'status', title: 'По статусам' },
  { key: 'type', title: 'По видам активности' },
  { key: 'day', title: 'По дням' }
];

function periodRange_(a) {
  if (a.period === 'all') return { from: '0000-00-00', to: '9999-12-31', label: 'Всё время' };
  var d = toDate_(a.date) || today_();
  var from, to, label;
  if (a.period === 'day') {
    from = to = d;
    label = 'День ' + ruDate_(ymd_(d));
  } else if (a.period === 'week') {
    from = addDays_(d, -((d.getDay() + 6) % 7));
    to = addDays_(from, 6);
    label = 'Неделя ' + ruDate_(ymd_(from)) + ' – ' + ruDate_(ymd_(to));
  } else if (a.period === 'month') {
    from = new Date(d.getFullYear(), d.getMonth(), 1);
    to = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    label = 'Месяц: ' + MONTHS_RU[d.getMonth()] + ' ' + d.getFullYear();
  } else {
    from = toDate_(a.from);
    to = toDate_(a.to);
    if (!from || !to || from > to) throw new Error('Укажите корректный период');
    label = 'Период ' + ruDate_(ymd_(from)) + ' – ' + ruDate_(ymd_(to));
  }
  return { from: ymd_(from), to: ymd_(to), label: label };
}

function emptyAgg_(key) {
  return { key: key, acts: 0, byType: {}, clients: 0, newClients: 0, deals: 0, dealSum: 0, lost: 0 };
}

// Сделки считаются по уникальным ИНН, чтобы повторный статус клиента не удваивал сумму.
function aggregate_(key, list, ctx) {
  var r = emptyAgg_(key);
  var pot = {}, fresh = {}, deal = {};
  list.forEach(function (a) {
    r.acts++;
    r.byType[a.type] = (r.byType[a.type] || 0) + 1;
    var kind = ctx.kinds[a.status];
    if (kind === 'lost') r.lost++;
    if (kind === 'deal') {
      var cur = deal[a.inn];
      if (!cur || a.date + a.created > cur.date + cur.created) deal[a.inn] = a;
    }
    pot[a.inn] = true;
    if (ctx.first[a.inn] >= ctx.from) fresh[a.inn] = true;
  });
  r.clients = Object.keys(pot).length;
  r.newClients = Object.keys(fresh).length;
  r.deals = Object.keys(deal).length;
  r.dealSum = Object.keys(deal).reduce(function (s, k) { return s + (deal[k].dealSum || 0); }, 0);
  return r;
}

function buildReport_(u, args) {
  var range = periodRange_(args);
  var dicts = getDicts_();
  var kinds = statusKinds_(dicts);
  var types = typeByFio_();
  var all = readActivities_();

  var first = {};
  all.forEach(function (a) { if (!first[a.inn] || a.date < first[a.inn]) first[a.inn] = a.date; });
  var ctx = { kinds: kinds, first: first, from: range.from };

  var list = all.filter(function (a) { return a.date >= range.from && a.date <= range.to; });
  list.forEach(function (a) { a.empType = types[a.fio] || 'не указан'; });

  var groups = {};
  REPORT_GROUPS.forEach(function (g) {
    var buckets = {};
    list.forEach(function (a) {
      var k = g.key === 'day' ? a.date : (a[g.key] || 'не указано');
      (buckets[k] || (buckets[k] = [])).push(a);
    });
    var rows = Object.keys(buckets).map(function (k) { return aggregate_(k, buckets[k], ctx); });
    if (g.key === 'day') rows.sort(function (x, y) { return x.key.localeCompare(y.key); });
    else rows.sort(function (x, y) { return y.acts - x.acts || x.key.localeCompare(y.key); });
    groups[g.key] = rows;
  });

  // Сотрудники без единой активности за период.
  var active = {};
  list.forEach(function (a) { active[a.fio] = true; });
  var idle = usersCached_().filter(function (x) {
    return x.active && x.role === ROLES.EMPLOYEE && !active[x.fio];
  }).map(function (x) { return x.fio; });

  // Просроченные следующие шаги и открытая воронка — на дату формирования.
  var ts = ymd_(today_());
  var overdue = {}, pipeline = { count: 0, next30: 0 };
  var in30 = ymd_(addDays_(today_(), 30));
  latestByClient_(all).forEach(function (a) {
    if ((kinds[a.status] || 'open') !== 'open') return;
    pipeline.count++;
    if (a.next && a.next >= ts && a.next <= in30) {
      pipeline.next30++;
    }
    if (a.next && a.next < ts) overdue[a.fio] = (overdue[a.fio] || 0) + 1;
  });

  // Текущее состояние клиентов, по которым были активности за период.
  var perClient = {};
  list.forEach(function (a) { var k = a.fio + '|' + a.inn; perClient[k] = (perClient[k] || 0) + 1; });
  var clients = latestByClient_(all).filter(function (a) { return perClient[a.fio + '|' + a.inn]; })
    .sort(function (x, y) { return x.inn.localeCompare(y.inn); })
    .map(function (a) {
      return {
        fio: a.fio, no: a.no, inn: a.inn, short: a.short, priority: a.priority, channel: a.channel, status: a.status,
        kind: KIND_LABEL[kinds[a.status]] || '', pirate: a.pirate, dealSum: a.dealSum,
        date: a.date, next: a.next, count: perClient[a.fio + '|' + a.inn], comment: a.comment
      };
    });

  // За «всё время» план считаем по текущий месяц включительно, будущие месяцы не добавляем.
  var t = today_();
  var plans = planForRange_(range.from, args.period === 'all' ? ymd_(new Date(t.getFullYear(), t.getMonth() + 1, 0)) : range.to);

  // Состав команды: для строк «без активностей» и фильтра по типу сотрудника.
  var staff = usersCached_().filter(function (x) { return x.active; })
    .map(function (x) { return { fio: x.fio, type: x.type, role: x.role, no: x.no }; });

  return {
    mpp: mppSnapshot_(all, readContacts_(), usersCached_(), ymd_(today_())),
    label: range.label,
    confirmedSales: (function () {
      var batches = confirmedBatches_(), confirmed = readConfirmed_(batches);
      return Object.keys(batches).sort().filter(function (month) {
        return month + '-01' >= range.from && confirmedEnd_(month) <= range.to;
      }).map(function (month) { return confirmedComparison_(u, month, batches, confirmed, all); });
    })(),
    staff: staff,
    plans: plans,
    clients: clients,
    from: range.from,
    to: range.to,
    generatedAt: Utilities.formatDate(new Date(), APP.TZ, 'dd.MM.yyyy HH:mm'),
    author: u.fio,
    types: dicts.types,
    groupsMeta: REPORT_GROUPS,
    summary: aggregate_('Итого', list, ctx),
    groups: groups,
    idle: idle,
    overdue: Object.keys(overdue).map(function (f) { return { fio: f, count: overdue[f] }; })
      .sort(function (a, b) { return b.count - a.count; }),
    pipeline: pipeline,
    details: list.sort(byDateDesc_).map(function (a) {
      return {
        date: a.date, fio: a.fio, no: a.no, empType: a.empType, inn: a.inn, short: a.short, priority: a.priority, channel: a.channel,
        type: a.type, status: a.status, pirate: a.pirate, dealSum: a.dealSum, next: a.next, comment: a.comment
      };
    })
  };
}
/**
 * Фактические продажи из таблицы продаж: одна строка — одна продажа (MID).
 * Загружаются раз в сутки на лист «Продажи»; факт «Продажи, шт» в плане считается по дате регистрации.
 */

// Точка входа для ежедневного триггера и кнопки в админке.
function importSales_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('sales_running')) return 'Загрузка продаж уже выполняется';
  cache.put('sales_running', '1', 300);
  try {
    return runSalesImport_();
  } catch (e) {
    setProp_('SALES_MSG', 'Ошибка: ' + e.message);
    setProp_('SALES_AT', new Date().toISOString());
    throw e;
  } finally {
    cache.remove('sales_running');
  }
}

// Триггер ставится сам при первой загрузке, отдельно запускать setup_() не нужно.
function ensureSalesTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'importSales_'; });
  if (!has) ScriptApp.newTrigger('importSales_').timeBased().everyDays(1).atHour(7).inTimezone(APP.TZ).create();
}

function salesSheet_() {
  var ss = ss_();
  return ss.getSheetByName(APP.SHEET.SALES) || ensureSheet_(ss, APP.SHEET.SALES, SALES_HEADERS);
}

// Источник отвечает за те месяцы регистрации, которые в нём есть: их строки заменяем целиком,
// остальные месяцы (например, прошлые, которых в источнике уже нет) сохраняем.
function runSalesImport_() {
  ensureSalesTrigger_();
  return withLock_(function () {
  if (prop_('STORED_PRIVACY_PENDING')) return 'Обезличивание данных: повторите импорт позже';
  var src = SpreadsheetApp.openById(prop_('SALES_SHEET_ID') || APP.SALES_SHEET_ID).getSheets()[0].getDataRange().getValues();
  var h = findHeaderRow_(src, function (c) { return c === 'inn' || c === 'инн'; });
  if (!h) throw new Error('В таблице продаж не найдена колонка inn');
  var idx = {};
  src[h.row].forEach(function (c, j) { idx[headerKey_(c)] = j; });
  if (idx.merch_date_reg == null) throw new Error('В таблице продаж не найдена колонка merch_date_reg');
  var col = function (r, name) { return idx[name] == null ? '' : str_(r[idx[name]]); };

  var now = new Date(), fresh = [], months = {};
  for (var i = h.row + 1; i < src.length; i++) {
    var r = src[i];
    var inn = normInn_(r[h.col]);
    var date = toDate_(r[idx.merch_date_reg]);
    if (!isInn_(inn) || !date) continue;
    months[ymd_(date).slice(0, 7)] = true;
    fresh.push([date, inn, col(r, 'mid'), storedName_(col(r, 'sales_fio')), col(r, 'priority_name'), col(r, 'attraction_type'), col(r, 'product_name'), now]);
  }
  if (!fresh.length) throw new Error('В таблице продаж нет строк с ИНН и датой регистрации');

  var sh = salesSheet_(), last = sh.getLastRow(), w = SALES_HEADERS.length;
  var keep = (last > 1 ? sh.getRange(2, 1, last - 1, w).getValues() : []).filter(function (r) {
    var d = toDate_(r[0]);
    return d && !months[ymd_(d).slice(0, 7)];
  });
  var rows = keep.concat(fresh).sort(function (a, b) { return toDate_(b[0]) - toDate_(a[0]); });
  if (last > 1) sh.getRange(2, 1, last - 1, w).clearContent();
  ensureRows_(sh, rows.length + 1);
  sh.getRange(2, 2, rows.length, 2).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, w).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, w, rows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm');

  var msg = 'Продажи: ' + fresh.length + ' шт за ' + Object.keys(months).sort().join(', ') + '; всего на листе ' + rows.length;
  setProp_('SALES_MSG', msg);
  setProp_('SALES_AT', now.toISOString());
  log_(null, 'Загрузка продаж', '', msg);
  return msg;
  });
}

// [{ date, inn, fio, priority }] — все продажи с листа.
function readSales_() {
  var sh = ss_().getSheetByName(APP.SHEET.SALES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map(function (r) {
    return { date: ymd_(r[0]), inn: normInn_(r[1]), fio: str_(r[3]), priority: str_(r[4]) };
  }).filter(function (s) { return s.date && s.inn; });
}

function salesStatus_() {
  return { at: prop_('SALES_AT') || '', msg: prop_('SALES_MSG') || 'Продажи ещё не загружались' };
}
/**
 * Первичная настройка. Запустите setup_() один раз из редактора Apps Script.
 * Повторный запуск безопасен: существующие данные не трогаются.
 */
function setup_() {
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

  console.log('Настройка завершена. Персональные ссылки доступны администратору в интерфейсе.');
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
  sh.getRange(2, COL.PIRATE, rows, 1).setNumberFormat('@');
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
  var adminFio = prop_('ADMIN_FIO') || APP.ADMIN_FIO;
  if (!adminFio) throw new Error('Для первой настройки задайте свойство ADMIN_FIO в Apps Script');
  add(adminFio, ROLES.ADMIN, '');
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
    if (t.getHandlerFunction() === 'importTurnovers' || t.getHandlerFunction() === 'importTurnovers_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importTurnovers_').timeBased().everyDays(1).atHour(6).inTimezone(APP.TZ).create();
}


// Разовая миграция существующих расписаний, без повторного импорта данных.
function secureMaintenance_(u) {
  if (!u.isAdmin) throw new Error('Недостаточно прав');
  return withLock_(function () {
    var triggers = ScriptApp.getProjectTriggers(), moved = [];
    ['importTurnovers', 'importSales'].forEach(function (name) {
      var old = triggers.filter(function (t) { return t.getHandlerFunction() === name; });
      if (!old.length) return;
      if (!triggers.some(function (t) { return t.getHandlerFunction() === name + '_'; })) {
        ScriptApp.newTrigger(name + '_').timeBased().everyDays(1).atHour(name === 'importSales' ? 7 : 6).inTimezone(APP.TZ).create();
      }
      old.forEach(function (t) { ScriptApp.deleteTrigger(t); });
      moved.push(name);
    });
    return { migrated: moved };
  });
}
/** Обезличивание хранимых данных. Индекс содержит только HMAC и сокращённые подписи. */
var _storedIdentityIndex = null, _storedKeyCache = Object.create(null);
function storedIdentityIndex_() {
  if (_storedIdentityIndex) return _storedIdentityIndex;
  var count = prop_('STORED_NAMES_ENCODING') === 'UTF-8' ? Number(prop_('STORED_NAME_PARTS') || 0) : 0, json = '';
  for (var i = 0; i < count; i++) json += prop_('STORED_NAME_' + i) || '';
  return (_storedIdentityIndex = json ? JSON.parse(json) : { sources: {}, texts: {}, refs: {} });
}
function storeIdentityIndex_(index) {
  var json = JSON.stringify(index), props = PropertiesService.getScriptProperties();
  var old = Number(prop_('STORED_NAME_PARTS') || 0), count = Math.ceil(json.length / 2000), values = {};
  for (var i = 0; i < count; i++) values['STORED_NAME_' + i] = json.slice(i * 2000, (i + 1) * 2000);
  values.STORED_NAME_PARTS = String(count); values.STORED_NAMES_ENCODING = 'UTF-8';
  props.setProperties(values);
  for (var j = count; j < old; j++) props.deleteProperty('STORED_NAME_' + j);
  _storedIdentityIndex = index;
}
function storedSourceKey_(name) {
  var normalized = str_(name).replace(/\./g, '').replace(/\s+/g, ' ').toLowerCase().replace(/ё/g, 'е');
  return _storedKeyCache[normalized] || (_storedKeyCache[normalized] = rawPersonRef_(normalized));
}
function addStoredIdentity_(index, name) {
  name = str_(name);
  if (!name || name.split(/\s+/).length < 2 || index.refs[name]) return name;
  var key = storedSourceKey_(name);
  if (index.sources[key]) return index.sources[key];
  var ref = rawPersonRef_(name), p = name.split(/\s+/), label = p[1] + ' ' + p[0].slice(0, 3), canonical = label;
  if (prop_('STORED_PRIVACY_VERSION') === '1') canonical += ' [' + key.slice(7, 15) + ']';
  else if (index.refs[canonical] && index.refs[canonical] !== ref) canonical += ' [' + ref.slice(7, 15) + ']';
  index.sources[key] = canonical; index.refs[canonical] = ref;
  var variants = [name, p[0] + ' ' + p[1], p[1] + ' ' + p[0]];
  if (p.length >= 3) variants.push(p[1] + ' ' + p[2] + ' ' + p[0], p[0] + ' ' + p[1][0] + ' ' + p[2][0], p[1][0] + ' ' + p[2][0] + ' ' + p[0]);
  variants.forEach(function (v) {
    var k = storedSourceKey_(v);
    // Не угадываем неоднозначное имя без отчества.
    if (index.texts[k] && index.texts[k] !== label) index.texts[k] = '';
    else if (!(k in index.texts)) index.texts[k] = label;
  });
  return canonical;
}
function storedName_(name) {
  name = str_(name);
  if (prop_('STORED_PRIVACY_VERSION') !== '1' || !name) return name;
  var index = storedIdentityIndex_();
  if (index.refs[name]) return name;
  var key = storedSourceKey_(name);
  if (index.sources[key]) return index.sources[key];
  var p = name.split(/\s+/);
  if (p.length < 2) return name;
  // Неизвестные сотрудники получают различимые ключи даже при одинаковых подписях.
  return p[1] + ' ' + p[0].slice(0, 3) + ' [' + key.slice(7, 15) + ']';
}
function storedText_(text) {
  var index = storedIdentityIndex_(), value = String(text == null ? '' : text), tokens = [], match;
  if (!Object.keys(index.texts).length) return value;
  var re = /[А-Яа-яЁёA-Za-z-]+/g;
  while ((match = re.exec(value))) tokens.push({ start: match.index, end: re.lastIndex });
  var edits = [];
  for (var i = 0; i < tokens.length; i++) {
    for (var n = 3; n >= 2; n--) {
      if (i + n > tokens.length) continue;
      var start = tokens[i].start, end = tokens[i + n - 1].end, part = value.slice(start, end);
      if (!/^[А-Яа-яЁёA-Za-z.\s-]+$/.test(part)) continue;
      var replacement = index.texts[storedSourceKey_(part)];
      if (replacement) { edits.push({ start: start, end: end, value: replacement }); i += n - 1; break; }
    }
  }
  edits.reverse().forEach(function (e) { value = value.slice(0, e.start) + e.value + value.slice(e.end); });
  return value;
}
function storedComment_(text) {
  return prop_('STORED_PRIVACY_VERSION') === '1' ? privateComment_(text, [], {}) : str_(text);
}
function storedPrivacySpecs_() {
  return {
    'Активности': { names: [0, 1, 15], text: [11], keep: [2, 3, 4] },
    'Действия по активностям': { names: [3, 9], text: [], keep: [4] },
    'Подтверждённые продажи': { names: [3], text: [], keep: [4] },
    'Загрузки подтверждений': { names: [6], text: [], keep: [] },
    'Продажи': { names: [3], text: [], keep: [1] },
    'Планы': { names: [1, 5], text: [], keep: [] },
    'Пользователи': { names: [1, 3], text: [], extra: 6, keep: [8, 9] },
    'Журнал': { names: [1], text: [4], keep: [] }
  };
}
function depersonalizeStorage_(u, args) {
  if (!u.isAdmin) throw new Error('Недостаточно прав');
  return withLock_(function () {
    var ss = ss_(), specs = storedPrivacySpecs_(), tables = [], names = {}, index = JSON.parse(JSON.stringify(storedIdentityIndex_()));
    Object.keys(specs).forEach(function (title) {
      var sh = ss.getSheetByName(title);
      if (!sh || sh.getLastRow() < 2) return;
      var range = sh.getDataRange(), values = range.getValues();
      tables.push({ sh: sh, spec: specs[title], values: values, notes: range.getNotes(), rules: range.getDataValidations() });
      values.slice(1).forEach(function (row) {
        specs[title].names.forEach(function (c) { var f = str_(row[c]); if (f.split(/\s+/).length >= 2) names[f] = true; });
        if (specs[title].extra != null) String(row[specs[title].extra] || '').split(/[;\n]/).forEach(function (f) { f = str_(f); if (f.split(/\s+/).length >= 2) names[f] = true; });
      });
    });
    var aliases = {}; Object.keys(names).sort().forEach(function (f) { aliases[f] = addStoredIdentity_(index, f); });
    var labels = {}; Object.keys(names).forEach(function (f) { labels[f] = aliases[f].replace(/ \[[^\]]+\]$/, ''); });
    var requests = [], counts = [];
    tables.forEach(function (table) {
      var values = table.values, spec = table.spec, changed = 0, sheetId = table.sh.getSheetId();
      for (var col = 0; col < values[0].length; col++) {
        var run = [], first = 0;
        var flush = function () {
          if (!run.length) return;
          requests.push({ updateCells: { start: { sheetId: sheetId, rowIndex: first, columnIndex: col }, rows: run.map(function (v) { return { values: [{ userEnteredValue: { stringValue: v } }] }; }), fields: 'userEnteredValue' } }); run = [];
        };
        for (var row = 1; row < values.length; row++) {
          var original = values[row][col], next = original;
          if (typeof original === 'string') {
            if (spec.names.indexOf(col) >= 0) next = aliases[str_(original)] || original;
            else if (spec.extra === col) next = original.split(/[;\n]/).map(function (f) { return aliases[str_(f)] || str_(f); }).filter(Boolean).join('; ');
            else if (spec.text.indexOf(col) >= 0) next = privateComment_(original, Object.keys(names), labels);
          }
          if (next !== original) { if (!run.length) first = row; run.push(next); changed++; } else flush();
          var note = table.notes[row] && table.notes[row][col];
          if (note) {
            var safe = privateComment_(note, Object.keys(names), labels);
            if (safe !== note) { requests.push({ updateCells: { start: { sheetId: sheetId, rowIndex: row, columnIndex: col }, rows: [{ values: [{ note: safe }] }], fields: 'note' } }); changed++; }
          }
          var rule = table.rules[row] && table.rules[row][col];
          if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            var options = rule.getCriteriaValues()[0], updated = options.map(function (f) { return aliases[str_(f)] || f; });
            if (JSON.stringify(options) !== JSON.stringify(updated)) {
              requests.push({ setDataValidation: { range: { sheetId: sheetId, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: col, endColumnIndex: col + 1 }, rule: { condition: { type: 'ONE_OF_LIST', values: updated.map(function (s) { return { userEnteredValue: s }; }) }, strict: !rule.getAllowInvalid(), showCustomUi: true } } }); changed++;
            }
          }
        }
        flush();
      }
      counts.push({ sheet: table.sh.getName(), rows: values.length - 1, changed: changed });
    });
    var summary = { applied: false, identities: Object.keys(names).length, sheets: counts, changes: counts.reduce(function (s, x) { return s + x.changed; }, 0) };
    summary.hashProbeDistinct = rawPersonRef_('Иванов Иван') !== rawPersonRef_('Петров Пётр');
    summary.neutralProbeUnchanged = storedText_('ошибка атомарной записи Google Sheets') === 'ошибка атомарной записи Google Sheets';
    if (!summary.hashProbeDistinct || !summary.neutralProbeUnchanged) throw new Error('Storage identity self-check failed');
    if (args.finalize) {
      if (!prop_('STORED_PRIVACY_PENDING') || args.lease !== prop_('STORED_PRIVACY_PENDING')) throw new Error('Storage lease mismatch');
      if (summary.changes) throw new Error('Storage verification still finds personal data');
      setProp_('STORED_PRIVACY_VERSION', '1');
      PropertiesService.getScriptProperties().deleteProperty('STORED_PRIVACY_PENDING');
      CacheService.getScriptCache().remove('users'); _acts = null;
      summary.applied = true;
      return summary;
    }
    if (args.exportRequests) {
      storeIdentityIndex_(index);
      var lease = prop_('STORED_PRIVACY_PENDING') || Utilities.getUuid();
      setProp_('STORED_PRIVACY_PENDING', lease);
      return { summary: summary, requests: requests, lease: lease };
    }
    if (!args.apply) return summary;
    // Индекс сохраняется до атомарной записи: после обрыва запрос можно безопасно повторить.
    storeIdentityIndex_(index);
    if (requests.length) {
      var response = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + ss.getId() + ':batchUpdate', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: JSON.stringify({ requests: requests }), muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) throw new Error('Storage HTTP ' + response.getResponseCode() + ': ' + JSON.parse(response.getContentText()).error.message);
    }
    setProp_('STORED_PRIVACY_VERSION', '1');
    CacheService.getScriptCache().remove('users'); _acts = null;
    summary.applied = true;
    return summary;
  });
}
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
  var u = readUsers_().filter(function (x) { return x.token === token; })[0];
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
      extra: u.extra, active: u.active
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
    .map(storedName_).filter(Boolean);

  var res = withLock_(function () {
    var users = readUsers_();
    extra = extra.map(storedName_);
    if (prop_('STORED_PRIVACY_VERSION') === '1') {
      var index = storedIdentityIndex_();
      fio = index.refs[fio] ? fio : addStoredIdentity_(index, fio);
      storeIdentityIndex_(index);
    }
    var dup = users.filter(function (u) { return u.fio === fio && u.id !== input.id; })[0];
    if (dup) throw new Error('Пользователь с таким ФИО уже есть');
    var sh = sheet_(APP.SHEET.USERS);
    var cur = input.id ? users.filter(function (u) { return u.id === input.id; })[0] : null;
    if (input.id && !cur) throw new Error('Пользователь не найден');
    if (cur && cur.id === admin.id && (role !== ROLES.ADMIN || input.active === false)) throw new Error('Нельзя снять с себя роль администратора');

    var id = cur ? cur.id : newId_();
    var token = cur ? cur.token : newToken_();
    var rowNum = cur ? cur.row : sh.getLastRow() + 1;
    ensureRows_(sh, rowNum);
    sh.getRange(rowNum, 1, 1, USER_HEADERS.length).setValues([[
      id, safeText_(fio), type, safeText_(storedName_(input.no)), role,
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

// Персональная ссылка выдаётся по явному запросу администратора, не вместе со всем справочником.
function userLink_(admin, args) {
  if (!admin.isAdmin) throw new Error('Недостаточно прав');
  var user = readUsers_().filter(function (u) { return u.id === args.id; })[0];
  if (!user) throw new Error('Пользователь не найден');
  return { link: linkFor_(user.token) };
}
/** Общие помощники. */

var _ss = null;
function ss_() {
  if (!_ss) _ss = SpreadsheetApp.openById(prop_('SPREADSHEET_ID') || APP.SPREADSHEET_ID);
  return _ss;
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Нет листа «' + name + '». Запустите setup_() в редакторе скрипта.');
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
      new Date(), user ? storedName_(user.fio) : 'система', action, id || '', details ? storedComment_(String(details)).slice(0, 500) : ''
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
/** Данные для экранов «План». */

function slimAct_(u, a, kinds) {
  return {
    id: a.id, fio: a.fio, inn: a.inn, short: a.short, full: a.full, channel: a.channel,
    priority: a.priority, updated: a.updated, status: a.status, kind: kinds[a.status] || 'open', pirate: a.pirate, dealSum: a.dealSum,
    date: a.date, type: a.type, next: a.next, comment: a.comment, canEdit: canEdit_(u, a)
  };
}

function getPlan_(u, args) {
  var kinds = statusKinds_();
  var latest = latestByClient_(visibleActs_(u, args.fio).filter(function (a) { return !completedStatus_(a.status); }));
  var t = today_(), ts = ymd_(t);
  var dow = (t.getDay() + 6) % 7; // 0 — понедельник
  var weekEnd = ymd_(addDays_(t, 6 - dow));
  var nextWeekEnd = ymd_(addDays_(t, 13 - dow));
  var res = { today: ts, overdue: [], todayList: [], week: [], nextWeek: [], later: [], noNext: [] };

  latest.forEach(function (a) {
    var item = slimAct_(u, a, kinds);
    if (!a.next) {
      res.noNext.push(item);
      return;
    }
    if (a.next < ts) res.overdue.push(item);
    else if (a.next === ts) res.todayList.push(item);
    else if (a.next <= weekEnd) res.week.push(item);
    else if (a.next <= nextWeekEnd) res.nextWeek.push(item);
    else res.later.push(item);
  });
  var byNext = function (x, y) { return x.next.localeCompare(y.next); };
  ['overdue', 'todayList', 'week', 'nextWeek', 'later'].forEach(function (k) { res[k].sort(byNext); });
  res.noNext.sort(function (x, y) { return x.date.localeCompare(y.date); });
  return res;
}
/** Правила завершения и журнал контактов: результат изменяет существующую активность. */
var CONTACT_SHEET = 'Действия по активностям';
var CONTACT_HEADERS = ['ID операции', 'ID активности', 'Дата контакта', 'Сотрудник', 'ИНН', 'Вид', 'Статус', 'Предыдущий статус', 'Записано', 'Автор', 'Операция'];
function completedStatus_(status) { return status === 'Отказ' || status === 'Договор подписан'; }
function needsVolume_(status, kinds) { return status === 'Получено согласие' || (kinds || statusKinds_())[status] === 'deal'; }
function unfinishedRows_(values, inn, exceptRow) {
  return values.filter(function (r, i) {
    return i + 2 !== exceptRow && !yes_(r[COL.DELETED - 1]) && normInn_(r[COL.INN - 1]) === inn && !completedStatus_(str_(r[COL.STATUS - 1]));
  });
}
function unfinishedMessage_() { return 'По этому ИНН есть незавершённые активности. Сначала завершите предыдущие: установите «Отказ» или «Договор подписан». Если запись принадлежит другому сотруднику, обратитесь к нему или администратору.'; }
function activityForResult_(u, args) {
  var all = readActivities_(), selected = all.filter(function (a) { return a.id === args.id; })[0];
  if (!selected || !canEdit_(u, selected)) throw new Error('Нет доступа к активности');
  var latest = all.filter(function (a) { return a.fio === selected.fio && a.inn === selected.inn && !completedStatus_(a.status); }).sort(byDateDesc_)[0];
  if (!latest) throw new Error('Активности клиента уже завершены. Обновите план');
  return pubAct_(u, latest);
}
function contactsSheet_() {
  return ss_().getSheetByName(CONTACT_SHEET) || ensureSheet_(ss_(), CONTACT_SHEET, CONTACT_HEADERS);
}
function contactFromRow_(r) {
  return { operationId: str_(r[0]), activityId: str_(r[1]), date: ymd_(r[2]), fio: str_(r[3]), inn: normInn_(r[4]), type: str_(r[5]), status: str_(r[6]), previous: str_(r[7]), at: stamp_(r[8]), author: str_(r[9]), kind: str_(r[10]) };
}
function readContacts_() {
  var sh = ss_().getSheetByName(CONTACT_SHEET);
  return !sh || sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, CONTACT_HEADERS.length).getValues().map(contactFromRow_);
}
// Контакты старых записей считаем по исходной дате. После первого изменения исходный контакт фиксируется в журнале.
function contactTimeline_(activities, events) {
  var ids = {}, baseline = {}, out = [];
  activities.forEach(function (a) { ids[a.id] = true; });
  events.forEach(function (e) {
    if (!ids[e.activityId]) return;
    if (e.kind === 'create' || e.kind === 'legacy') baseline[e.activityId] = true;
    out.push(e);
  });
  activities.forEach(function (a) {
    if (!baseline[a.id]) out.push({ activityId: a.id, date: a.date, fio: a.fio, inn: a.inn, type: a.type, status: a.status, previous: '', at: a.created, kind: 'legacy' });
  });
  return out;
}

function mppSnapshot_(activities, events, staff, asOf) {
  var from = ymd_(addDays_(toDate_(asOf), -29)), month = asOf.slice(0, 7), rows = {}, states = {}, contracts = {};
  staff.filter(function (u) { return u.active && u.type === 'МПП'; }).forEach(function (u) {
    rows[u.fio] = { fio: u.fio, contracts: 0, agreed: 0, thinking: 0, calls: 0, meetings: 0, messages: 0, contacts: 0, last30: 0 };
    contracts[u.fio] = {};
  });
  contactTimeline_(activities, events).forEach(function (e) {
    var r = rows[e.fio]; if (!r || !e.date || e.date > asOf) return;
    var isContact = e.kind !== 'edit';
    if (isContact && e.date >= from && ['Звонок','Встреча','Сообщение'].indexOf(e.type) >= 0) r.last30++;
    if (isContact && e.date === asOf) {
      if (e.type === 'Звонок') r.calls++;
      if (e.type === 'Встреча') r.meetings++;
      if (e.type === 'Сообщение') r.messages++;
    }
    if (e.date.slice(0, 7) === month && e.status === 'Договор подписан' && e.previous !== 'Договор подписан') contracts[e.fio][e.inn] = true;
    var key = e.fio + '|' + e.inn, rank = e.date + (e.at || '') + (e.kind === 'legacy' ? '0' : '1');
    if (!states[key] || rank >= states[key].rank) states[key] = { rank: rank, event: e };
  });
  Object.keys(states).forEach(function (key) {
    var e = states[key].event, r = rows[e.fio];
    if (e.date < from) return;
    if (e.status === 'Получено согласие') r.agreed++;
    if (e.status === 'Клиент думает') r.thinking++;
  });
  var list = Object.keys(rows).sort().map(function (f) {
    var r = rows[f]; r.contracts = Object.keys(contracts[f]).length;
    r.contacts = r.calls + r.meetings + r.messages;
    r.met = { agreed: r.agreed >= 5, thinking: r.thinking >= 20, calls: r.calls >= 20 };
    return r;
  });
  return { asOf: asOf, from: from, month: month, targets: { agreed: 5, thinking: 20, calls: 20 }, rows: list };
}
