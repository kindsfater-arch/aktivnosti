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
  return { key: key, acts: 0, byType: {}, clients: 0, newClients: 0, deals: 0, dealSum: 0, lost: 0, potential: 0 };
}

// Сделки и потенциал считаются по уникальным ИНН, чтобы «подписан» и «запущен» одного клиента не удваивали сумму.
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
    pot[a.inn] = Math.max(pot[a.inn] || 0, a.potential || 0);
    if (ctx.first[a.inn] >= ctx.from) fresh[a.inn] = true;
  });
  r.clients = Object.keys(pot).length;
  r.newClients = Object.keys(fresh).length;
  r.potential = Object.keys(pot).reduce(function (s, k) { return s + pot[k]; }, 0);
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
  var overdue = {}, pipeline = { count: 0, potential: 0, next30: 0, next30Potential: 0 };
  var in30 = ymd_(addDays_(today_(), 30));
  latestByClient_(all).forEach(function (a) {
    if ((kinds[a.status] || 'open') !== 'open') return;
    pipeline.count++;
    pipeline.potential += a.potential || 0;
    if (a.next && a.next >= ts && a.next <= in30) {
      pipeline.next30++;
      pipeline.next30Potential += a.potential || 0;
    }
    if (a.next && a.next < ts) overdue[a.fio] = (overdue[a.fio] || 0) + 1;
  });

  // Текущее состояние клиентов, по которым были активности за период.
  var perClient = {};
  list.forEach(function (a) { var k = a.fio + '|' + a.inn; perClient[k] = (perClient[k] || 0) + 1; });
  var clients = latestByClient_(all).filter(function (a) { return perClient[a.fio + '|' + a.inn]; })
    .sort(function (x, y) { return (y.potential || 0) - (x.potential || 0); })
    .map(function (a) {
      return {
        fio: a.fio, no: a.no, inn: a.inn, short: a.short, priority: a.priority, channel: a.channel, status: a.status,
        kind: KIND_LABEL[kinds[a.status]] || '', potential: a.potential, dealSum: a.dealSum,
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
    label: range.label,
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
        type: a.type, status: a.status, potential: a.potential, dealSum: a.dealSum, next: a.next, comment: a.comment
      };
    })
  };
}

