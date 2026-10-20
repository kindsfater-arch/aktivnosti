/**
 * Веб-приложение и единая точка входа API.
 *  doGet  — отдаёт интерфейс (персональная ссылка: …/exec?k=<токен>).
 *  doPost — JSON API для интерфейса, размещённого на GitHub Pages.
 *  api    — тот же API для google.script.run, когда интерфейс открыт с адреса Apps Script.
 */

function doGet(e) {
  var t = HtmlService.createTemplateFromFile('Index');
  var token = e && e.parameter ? str_(e.parameter.k) : '';
  t.bootJson = JSON.stringify({ mode: 'gas', token: token }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Активности')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var req = {};
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    req = {};
  }
  return ContentService.createTextOutput(JSON.stringify(api(req.k, req.fn, req.args)))
    .setMimeType(ContentService.MimeType.JSON);
}

function api(token, fn, args) {
  try {
    var user = auth_(token);
    var route = routes_()[fn];
    if (!route) throw new Error('Неизвестный метод: ' + fn);
    if (route.need && !user[route.need]) throw new Error('Недостаточно прав для этого действия');
    return { ok: true, data: route.fn(user, args || {}) };
  } catch (err) {
    console.error(fn + ': ' + (err && err.stack || err));
    return { ok: false, error: String(err && err.message || err) };
  }
}

function routes_() {
  return {
    me: { fn: me_ },
    plan: { fn: getPlan_ },
    listActivities: { fn: listActivities_ },
    saveActivity: { fn: saveActivity_ },
    deleteActivity: { fn: deleteActivity_ },
    lookupInn: { fn: lookupInn_ },
    deals: { fn: getDeals_ },
    leads: { fn: getLeads_ },
    report: { fn: reportForClient_, need: 'canReport' },
    exportReport: { fn: exportReport_, need: 'canReport' },
    exportAll: { fn: exportAll_, need: 'canExportAll' },
    users: { fn: function () { return listUsers_(); }, need: 'isAdmin' },
    saveUser: { fn: saveUser_, need: 'isAdmin' },
    regenToken: { fn: regenToken_, need: 'isAdmin' },
    saveDicts: { fn: saveDicts_, need: 'isAdmin' },
    adminInfo: { fn: adminInfo_, need: 'isAdmin' },
    runImport: { fn: function () { return importTurnovers(true); }, need: 'isAdmin' },
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
    appUrl: prop_('APP_URL') || '',
    scriptUrl: (function () { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })(),
    dadata: !!prop_('DADATA_KEY'),
    sheetUrl: ss_().getUrl()
  };
}

function saveSettings_(u, args) {
  if (args.appUrl != null) {
    var url = str_(args.appUrl);
    if (url && !/^https:\/\//.test(url)) throw new Error('Адрес должен начинаться с https://');
    setProp_('APP_URL', url);
    refreshLinks_();
  }
  if (args.dadataKey != null) setProp_('DADATA_KEY', str_(args.dadataKey));
  log_(u, 'Изменены настройки');
  return adminInfo_();
}
