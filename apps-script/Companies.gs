/**
 * Наименования организаций по ИНН.
 * Порядок поиска: кэш «Компании» → уже заполненные строки «Активностей» → ФНС (egrul.nalog.ru, без ключа)
 * → DaData (только если в свойствах скрипта задан DADATA_KEY). Найденное сохраняется в кэш,
 * повторно в интернет за тем же ИНН не ходим.
 */

function findCompany_(inn) {
  return companyFromCache_(inn) ||
    rememberCompany_(companyFromActivities_(inn)) ||
    rememberCompany_(companyFromFns_(inn)) ||
    rememberCompany_(companyFromDadata_(inn));
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
  return a ? { inn: inn, short: a.short || a.full, full: a.full || a.short, source: 'активность' } : null;
}

function companyFromFns_(inn) {
  try {
    var opts = { muteHttpExceptions: true, followRedirects: true };
    var r1 = UrlFetchApp.fetch('https://egrul.nalog.ru/', Object.assign({ method: 'post', payload: { query: inn } }, opts));
    if (r1.getResponseCode() !== 200) return null;
    var t = JSON.parse(r1.getContentText()).t;
    if (!t) return null;
    for (var attempt = 0; attempt < 4; attempt++) {
      Utilities.sleep(600);
      var r2 = UrlFetchApp.fetch('https://egrul.nalog.ru/search-result/' + t, opts);
      if (r2.getResponseCode() !== 200) continue;
      var j = JSON.parse(r2.getContentText());
      if (j.status === 'wait') continue;
      var rows = j.rows || [];
      var row = rows.filter(function (x) { return x.i === inn; })[0] || rows[0];
      if (!row) return null;
      var isIp = row.k === 'fl';
      return {
        inn: inn,
        short: isIp ? 'ИП ' + titleCase_(row.n) : (row.c || row.n),
        full: isIp ? 'Индивидуальный предприниматель ' + titleCase_(row.n) : row.n,
        ogrn: row.o || '',
        region: row.rn || '',
        status: row.e ? 'Прекращена ' + row.e : 'Действует',
        source: 'ФНС'
      };
    }
  } catch (e) {
    console.warn('ФНС: ' + e);
  }
  return null;
}

function companyFromDadata_(inn) {
  var key = prop_('DADATA_KEY');
  if (!key) return null;
  try {
    var resp = UrlFetchApp.fetch('https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Token ' + key, Accept: 'application/json' },
      payload: JSON.stringify({ query: inn }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return null;
    var s = (JSON.parse(resp.getContentText()).suggestions || [])[0];
    if (!s) return null;
    var d = s.data || {}, name = d.name || {};
    var state = d.state && d.state.status;
    return {
      inn: inn,
      short: name.short_with_opf || s.value,
      full: name.full_with_opf || s.value,
      ogrn: d.ogrn || '',
      region: d.address && d.address.data ? d.address.data.region_with_type || '' : '',
      status: state === 'ACTIVE' ? 'Действует' : (state ? 'Статус: ' + state : ''),
      source: 'DaData'
    };
  } catch (e) {
    console.warn('DaData: ' + e);
    return null;
  }
}

function rememberCompany_(c) {
  if (!c || !(c.short || c.full)) return null;
  try {
    withLock_(function () {
      if (companyFromCache_(c.inn)) return;
      var sh = sheet_(APP.SHEET.COMP), rowNum = sh.getLastRow() + 1;
      ensureRows_(sh, rowNum);
      sh.getRange(rowNum, 1).setNumberFormat('@');
      sh.getRange(rowNum, 1, 1, COMP_HEADERS.length).setValues([[
        c.inn, safeText_(c.short), safeText_(c.full), c.ogrn || '', c.region || '', c.status || '', c.source || '', new Date()
      ]]);
    });
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

function titleCase_(s) {
  return str_(s).toLowerCase().replace(/(^|[\s-])([а-яёa-z])/g, function (m, p, ch) { return p + ch.toUpperCase(); });
}
