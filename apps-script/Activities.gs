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
    potential: num_(r[COL.POTENTIAL - 1]),
    status: str_(r[COL.STATUS - 1]),
    date: ymd_(r[COL.DATE - 1]),
    type: str_(r[COL.TYPE - 1]),
    next: ymd_(r[COL.NEXT - 1]),
    comment: str_(r[COL.COMMENT - 1]),
    dealSum: num_(r[COL.DEAL_SUM - 1]),
    potSrc: str_(r[COL.POT_SRC - 1]),
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
  var list = visibleActs_(u, args.fio).filter(function (a) {
    if (args.from && a.date < args.from) return false;
    if (args.to && a.date > args.to) return false;
    if (args.status && a.status !== args.status) return false;
    if (args.type && a.type !== args.type) return false;
    if (args.channel && a.channel !== args.channel) return false;
    if (q && (a.inn + ' ' + a.short + ' ' + a.full + ' ' + a.comment).toLowerCase().indexOf(q) < 0) return false;
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
// Потенциал — по клиентам: у каждого ИНН берётся значение из последней активности за период (список уже отсортирован).
function actsSummary_(list) {
  var today = ymd_(today_());
  var s = { count: list.length, days: 0, fios: 0, clients: 0, key: 0, byType: {}, byStatus: {}, byChannel: {},
    potential: 0, deals: 0, dealSum: 0, next: 0, nextAhead: 0 };
  var days = {}, fios = {}, inns = {};
  var inc = function (o, k) { o[k || '—'] = (o[k || '—'] || 0) + 1; };
  list.forEach(function (a) {
    days[a.date] = 1;
    fios[a.fio] = 1;
    if (!inns[a.inn]) {
      inns[a.inn] = 1;
      s.potential += a.potential || 0;
      if (a.priority === PRIORITIES[0]) s.key++;
    }
    inc(s.byType, a.type);
    inc(s.byStatus, a.status);
    inc(s.byChannel, a.channel);
    if (a.dealSum > 0) { s.deals++; s.dealSum += a.dealSum; }
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

  v.fio = u.seeAll ? str_(input.fio || u.fio) : u.fio;
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

  v.comment = str_(input.comment);
  if (v.comment.length < APP.MIN_COMMENT) {
    throw new Error('Опишите в комментарии суть активности (не короче ' + APP.MIN_COMMENT + ' символов)');
  }

  v.dealSum = '';
  if (kinds[v.status] === 'deal') {
    v.dealSum = Math.round(num_(input.dealSum));
    if (!(v.dealSum > 0)) throw new Error('Для статуса «' + v.status + '» укажите сумму сделки');
  }

  // Потенциал = оборот у конкурентов из PS. Если сотрудник ввёл другое значение, оно сохраняется как ручное.
  var t = turnoverFor_(v.inn);
  var ps = t ? Math.round(t.ps) : 0;
  var raw = input.potential;
  var manual = raw === '' || raw == null ? null : Math.round(num_(raw));
  if (manual == null || (ps && manual === ps)) {
    v.potential = ps || '';
    v.potSrc = ps ? 'PS' : '';
  } else {
    v.potential = manual;
    v.potSrc = 'вручную';
  }

  v.short = str_(input.short);
  v.full = str_(input.full);
  if (!v.short || !v.full) {
    var comp = companyFromCache_(v.inn);
    if (comp) {
      v.short = v.short || comp.short;
      v.full = v.full || comp.full;
    }
  }
  v.no = noFor_(v.fio) || str_(input.no);
  return v;
}

function saveActivity_(u, input) {
  var v = validateActivity_(u, input);
  var saved = withLock_(function () {
    var sh = sheet_(APP.SHEET.ACT), now = new Date();
    ensureActColumns_(sh);
    var rowNum, id, author, created;
    if (input.id) {
      rowNum = findActRow_(sh, input.id);
      if (!rowNum) throw new Error('Активность не найдена — возможно, её удалили');
      var curRow = sh.getRange(rowNum, 1, 1, NCOL).getValues()[0];
      var cur = rowToAct_(curRow, rowNum);
      if (yes_(curRow[COL.DELETED - 1])) throw new Error('Активность удалена');
      if (!canEdit_(u, cur)) throw new Error('Нет прав на изменение этой активности');
      // Снятый с выбора статус можно оставить у старой записи, но нельзя выбрать заново.
      if (!v.statusActive && v.status !== cur.status) throw new Error(retiredStatusMsg_(v.status));
      id = cur.id;
      author = cur.author;
      created = curRow[COL.CREATED - 1] || now;
    } else {
      if (!v.statusActive) throw new Error(retiredStatusMsg_(v.status));
      id = newId_();
      author = u.fio;
      created = now;
      rowNum = sh.getLastRow() + 1;
      ensureRows_(sh, rowNum);
    }
    var row = [
      safeText_(v.fio), safeText_(v.no), v.inn, safeText_(v.short), safeText_(v.full), v.channel,
      v.potential, v.status, v.date, v.type, v.next || '', safeText_(v.comment), v.dealSum,
      id, v.potSrc, author, created, now, '', v.priority
    ];
    sh.getRange(rowNum, COL.INN).setNumberFormat('@');
    sh.getRange(rowNum, 1, 1, NCOL).setValues([row]);
    return { row: row, rowNum: rowNum, isNew: !input.id };
  });
  _acts = null;
  if (v.short || v.full) rememberCompany_({ inn: v.inn, short: v.short, full: v.full, source: 'активность' });
  var a = rowToAct_(saved.row, saved.rowNum);
  log_(u, saved.isNew ? 'Создана активность' : 'Изменена активность', a.id, a.inn + ' · ' + a.status);
  return pubAct_(u, a);
}

function retiredStatusMsg_(status) {
  return 'Статус «' + status + '» больше не используется — выберите другой';
}

// Подсказка приоритета: из последней активности по клиенту, иначе из файла продаж
// («Ключевой» там — ключевой, любой другой приоритет — неключевой).
function priorityFor_(inn, acts) {
  var last = acts.filter(function (a) { return a.priority; }).sort(byDateDesc_)[0];
  if (last) return { value: last.priority, source: 'по последней активности' };
  var sale = readSales_().filter(function (x) { return x.inn === inn && x.priority; })[0];
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
  var t = turnoverFor_(inn);
  var all = readActivities_().filter(function (a) { return a.inn === inn; });
  var inWork = latestByClient_(all).map(function (a) {
    return { fio: a.fio, status: a.status, date: a.date, next: a.next };
  });
  // Вся история по клиенту, которую сотруднику можно видеть, — от новых к старым, с полными комментариями.
  var history = all.filter(function (a) { return canSee_(u, a); }).sort(byDateDesc_)
    .map(function (a) { return pubAct_(u, a); });
  return {
    inn: inn,
    priority: priorityFor_(inn, all),
    company: findCompany_(inn),
    potential: t && t.ps ? Math.round(t.ps) : null,
    ps: t && t.ps ? { turnover: Math.round(t.ps), banks: t.banks, sbp: t.sbp, ecom: t.ecom, file: t.psFile } : null,
    sber: t && t.sber ? { turnover: Math.round(t.sber), income: Math.round(t.income), file: t.sberFile } : null,
    inWork: inWork,
    history: history
  };
}
