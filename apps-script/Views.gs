/** Данные для экранов «План», «Сделки», «Лиды PS». */

function slimAct_(u, a, kinds) {
  return {
    id: a.id, fio: a.fio, inn: a.inn, short: a.short, full: a.full, channel: a.channel,
    status: a.status, kind: kinds[a.status] || 'open', potential: a.potential, potSrc: a.potSrc, dealSum: a.dealSum,
    date: a.date, type: a.type, next: a.next, comment: a.comment, canEdit: canEdit_(u, a)
  };
}

function getPlan_(u, args) {
  var kinds = statusKinds_();
  var latest = latestByClient_(visibleActs_(u, args.fio));
  var t = today_(), ts = ymd_(t);
  var dow = (t.getDay() + 6) % 7; // 0 — понедельник
  var weekEnd = ymd_(addDays_(t, 6 - dow));
  var nextWeekEnd = ymd_(addDays_(t, 13 - dow));
  var res = { today: ts, overdue: [], todayList: [], week: [], nextWeek: [], later: [], noNext: [] };

  latest.forEach(function (a) {
    var item = slimAct_(u, a, kinds);
    if (!a.next) {
      if (item.kind === 'open') res.noNext.push(item);
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

// Воронка: открытые клиенты по месяцу следующего шага. Сделки: активности со статусом «Сделка» по месяцам.
function getDeals_(u, args) {
  var kinds = statusKinds_();
  var acts = visibleActs_(u, args.fio);
  var ts = ymd_(today_());

  var pipe = {};
  latestByClient_(acts).forEach(function (a) {
    if ((kinds[a.status] || 'open') !== 'open') return;
    var key = !a.next ? 'none' : (a.next < ts ? 'overdue' : a.next.slice(0, 7));
    var g = pipe[key] || (pipe[key] = { key: key, count: 0, potential: 0, items: [] });
    g.count++;
    g.potential += a.potential || 0;
    g.items.push(slimAct_(u, a, kinds));
  });
  var order = function (k) { return k === 'overdue' ? '0' : (k === 'none' ? '9' : '1' + k); };
  var pipeline = Object.keys(pipe).sort(function (a, b) { return order(a).localeCompare(order(b)); })
    .map(function (k) {
      pipe[k].items.sort(function (x, y) { return (x.next || '9').localeCompare(y.next || '9'); });
      return pipe[k];
    });

  var dl = {};
  acts.forEach(function (a) {
    if (kinds[a.status] !== 'deal') return;
    var key = a.date.slice(0, 7);
    var g = dl[key] || (dl[key] = { key: key, count: 0, sum: 0, byStatus: {}, items: [] });
    g.count++;
    g.sum += a.dealSum || 0;
    var s = g.byStatus[a.status] || (g.byStatus[a.status] = { count: 0, sum: 0 });
    s.count++;
    s.sum += a.dealSum || 0;
    g.items.push(slimAct_(u, a, kinds));
  });
  var deals = Object.keys(dl).sort().reverse().map(function (k) {
    dl[k].items.sort(byDateDesc_);
    return dl[k];
  });

  return { today: ts, pipeline: pipeline, deals: deals };
}

// Крупные ИНН из PS, по которым ещё никто не ведёт работу.
function getLeads_(u, args) {
  var taken = {};
  readActivities_().forEach(function (a) { taken[a.inn] = true; });
  var names = companiesMap_();
  var min = num_(args.min) || 0;
  var onlyNew = !!args.onlyNew;
  var q = str_(args.q).toLowerCase();
  var list = readTurnovers_().filter(function (t) {
    if (!t.ps || taken[t.inn] || t.ps < min) return false;
    if (onlyNew && t.sber > 0) return false;
    if (q && (t.inn + ' ' + (names[t.inn] || '') + ' ' + t.banks).toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
  list.sort(function (a, b) { return b.ps - a.ps; });
  var limit = Math.min(num_(args.limit) || 100, 500);
  return {
    total: list.length,
    psFile: prop_('PS_FILE') || '',
    rows: list.slice(0, limit).map(function (t) {
      return {
        inn: t.inn, name: names[t.inn] || '', ps: Math.round(t.ps), banks: t.banks,
        sbp: t.sbp, ecom: t.ecom, sber: Math.round(t.sber)
      };
    })
  };
}
