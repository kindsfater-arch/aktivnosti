/**
 * Фактические продажи из таблицы продаж: одна строка — одна продажа (MID).
 * Загружаются раз в сутки на лист «Продажи»; факт «Продажи, шт» в плане считается по дате регистрации.
 */

// Точка входа для ежедневного триггера и кнопки в админке.
function importSales() {
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

// Триггер ставится сам при первой загрузке, отдельно запускать setup() не нужно.
function ensureSalesTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'importSales'; });
  if (!has) ScriptApp.newTrigger('importSales').timeBased().everyDays(1).atHour(7).inTimezone(APP.TZ).create();
}

function salesSheet_() {
  var ss = ss_();
  return ss.getSheetByName(APP.SHEET.SALES) || ensureSheet_(ss, APP.SHEET.SALES, SALES_HEADERS);
}

// Источник отвечает за те месяцы регистрации, которые в нём есть: их строки заменяем целиком,
// остальные месяцы (например, прошлые, которых в источнике уже нет) сохраняем.
function runSalesImport_() {
  ensureSalesTrigger_();
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
    fresh.push([date, inn, col(r, 'mid'), col(r, 'sales_fio'), col(r, 'priority_name'), col(r, 'attraction_type'), col(r, 'product_name'), now]);
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
