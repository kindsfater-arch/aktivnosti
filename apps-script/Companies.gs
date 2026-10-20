/**
 * Наименования организаций по ИНН.
 * Порядок поиска: кэш «Компании» → уже заполненные строки «Активностей» → DaData (если в настройках есть ключ).
 * ФНС (egrul.nalog.ru) не используется: она недоступна с серверов Google, запрос просто зависает.
 * Найденное сохраняется в кэш, повторно за тем же ИНН не ходим; ненайденное не ищем 6 часов.
 */

function findCompany_(inn) {
  var c = companyFromCache_(inn) || rememberCompany_(companyFromActivities_(inn));
  if (c) return c;
  var cache = CacheService.getScriptCache();
  if (cache.get('nf:' + inn)) return null;
  c = rememberCompany_(companyFromDadata_(inn));
  if (!c) cache.put('nf:' + inn, '1', 21600);
  return c;
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

// Без блокировки: редкий дубль строки в кэше безвреден, а ожидание блокировки тормозило сохранение.
function rememberCompany_(c) {
  if (!c || !(c.short || c.full)) return null;
  try {
    if (companyFromCache_(c.inn)) return c;
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
