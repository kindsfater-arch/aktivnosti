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
    return { fio: str_(r.fio), deals: Math.max(0, Math.round(num_(r.deals))), volume: Math.max(0, Math.round(num_(r.volume))) };
  }).filter(function (r) { return r.fio && (r.deals || r.volume); });

  withLock_(function () {
    var sh = plansSheet_(), last = sh.getLastRow(), w = PLAN_HEADERS.length;
    var keep = (last > 1 ? sh.getRange(2, 1, last - 1, w).getValues() : []).filter(function (r) {
      return monthKey_(r[0]) && monthKey_(r[0]) !== month;
    }).map(function (r) {
      r[0] = monthKey_(r[0]);
      return r;
    });
    var rows = keep.concat(fresh.map(function (r) { return [month, r.fio, r.deals, r.volume, now, u.fio]; }));
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

// Выполнение плана за месяц — видят все сотрудники. Факт сделок считается как в отчёте: уникальные ИНН
// со статусом типа «Сделка», сумма по последней такой активности клиента. Рабочие дни (пн–пт, без праздников)
// нужны интерфейсу для прогноза RR и расчёта «сколько нужно в день».
function planProgress_(u, args) {
  var month = monthKey_(args.month) || Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM');
  var y = +month.slice(0, 4), m = +month.slice(5, 7), dim = new Date(y, m, 0).getDate();
  var from = month + '-01', to = month + '-' + (dim < 10 ? '0' : '') + dim;
  var plan = readPlans_()[month] || {};
  var kinds = statusKinds_();

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
  Object.keys(plan).concat(Object.keys(deals)).forEach(function (f) { fios[f] = true; });
  if (u.role === ROLES.EMPLOYEE) fios[u.fio] = true;
  var rows = Object.keys(fios).map(function (fio) {
    var list = Object.keys(deals[fio] || {}).map(function (inn) { return deals[fio][inn]; });
    return {
      fio: fio,
      type: types[fio] || '',
      plan: plan[fio] || null,
      fact: { deals: list.length, volume: list.reduce(function (s, a) { return s + (a.dealSum || 0); }, 0) }
    };
  });
  return { month: month, today: today, workdays: workdays, rows: rows };
}
