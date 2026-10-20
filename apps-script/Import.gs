/**
 * Импорт оборотов из папки на Drive.
 *  PS_xx — оборот клиентов у банков-конкурентов (по терминалам), суммируем vol по ИНН.
 *  PL_xx — оборот в Сбере (колонка «Оборот» = US + NON-US). Если PL нет, берём US_xx и складываем обе суммы.
 * Из каждой группы берётся самый свежий файл (по дате создания).
 */

// Точка входа для ежедневного триггера и кнопки в админке.
// Флаг в кэше вместо LockService: внутри импорта блокировка нужна для записи в «Активности».
function importTurnovers(force) {
  var cache = CacheService.getScriptCache();
  if (cache.get('import_running')) return 'Импорт уже выполняется';
  cache.put('import_running', '1', 600);
  try {
    return runImport_(force === true);
  } catch (e) {
    setProp_('IMPORT_MSG', 'Ошибка: ' + e.message);
    setProp_('IMPORT_AT', new Date().toISOString());
    throw e;
  } finally {
    cache.remove('import_running');
  }
}

function runImport_(force) {
  var folder = DriveApp.getFolderById(prop_('DATA_FOLDER_ID') || APP.DATA_FOLDER_ID);
  var psFile = latestFile_(folder, 'PS');
  var sberFile = latestFile_(folder, 'PL') || latestFile_(folder, 'US');
  if (!psFile && !sberFile) throw new Error('В папке нет файлов PS_xx, PL_xx или US_xx');

  var sig = [psFile, sberFile].map(function (f) {
    return f ? f.getId() + ':' + f.getLastUpdated().getTime() : '-';
  }).join('|');
  if (!force && sig === prop_('IMPORT_SIG')) return 'Файлы не изменились — импорт не нужен';

  var ps = psFile ? parsePs_(readXlsx_(psFile)) : {};
  var sber = sberFile ? parseSber_(readXlsx_(sberFile)) : { map: {}, period: '' };
  var psName = psFile ? psFile.getName() : '';
  var sberName = sberFile ? sberFile.getName() + (sber.period ? ' (' + sber.period + ')' : '') : '';

  var inns = {};
  Object.keys(ps).forEach(function (k) { inns[k] = 1; });
  Object.keys(sber.map).forEach(function (k) { inns[k] = 1; });

  var now = new Date();
  var rows = Object.keys(inns).map(function (inn) {
    var p = ps[inn], s = sber.map[inn];
    return [
      inn,
      p ? Math.round(p.total) : '',
      p ? topBanks_(p.banks) : '',
      p && p.total ? p.sbp / p.total : '',
      p && p.total ? p.ecom / p.total : '',
      s ? Math.round(s.turnover) : '',
      s ? Math.round(s.income) : '',
      p ? psName : '',
      s ? sberName : '',
      now
    ];
  });
  rows.sort(function (a, b) { return (b[1] || 0) - (a[1] || 0); });

  var sh = sheet_(APP.SHEET.TURN);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TURN_HEADERS.length).clearContent();
  if (rows.length) {
    ensureRows_(sh, rows.length + 1);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, TURN_HEADERS.length).setValues(rows);
    sh.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');
    sh.getRange(2, 4, rows.length, 2).setNumberFormat('0%');
    sh.getRange(2, 6, rows.length, 2).setNumberFormat('#,##0');
    sh.getRange(2, 10, rows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm');
  }

  var updated = refreshPotentials_(ps);
  var msg = 'PS: ' + (psName || '—') + ' (' + Object.keys(ps).length + ' ИНН); Сбер: ' +
    (sberName || '—') + ' (' + Object.keys(sber.map).length + ' ИНН); обновлён потенциал в ' + updated + ' строках';
  setProp_('IMPORT_SIG', sig);
  setProp_('IMPORT_MSG', msg);
  setProp_('IMPORT_AT', now.toISOString());
  setProp_('PS_FILE', psName);
  setProp_('SBER_FILE', sberName);
  log_(null, 'Импорт оборотов', '', msg);
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
  var idx = {};
  values[h.row].forEach(function (c, j) { idx[headerKey_(c)] = j; });
  if (idx.vol == null) throw new Error('В файле PS не найдена колонка vol');

  var map = {};
  for (var i = h.row + 1; i < values.length; i++) {
    var r = values[i];
    var inn = normInn_(r[h.col]);
    if (!isInn_(inn)) continue;
    var vol = num_(r[idx.vol]);
    if (!vol) continue;
    var m = map[inn] || (map[inn] = { total: 0, banks: {}, sbp: 0, ecom: 0 });
    m.total += vol;
    var bank = idx.acq_bank != null ? str_(r[idx.acq_bank]) || 'не указан' : 'не указан';
    m.banks[bank] = (m.banks[bank] || 0) + vol;
    if (idx.card_or_sbp != null && /sbp|сбп/i.test(str_(r[idx.card_or_sbp]))) m.sbp += vol;
    if (idx.pos_or_ecom != null && !/^pos$/i.test(str_(r[idx.pos_or_ecom]))) m.ecom += vol;
  }
  return map;
}

function parseSber_(values) {
  var h = findHeaderRow_(values, function (c) { return c.indexOf('инн') === 0; });
  if (!h) throw new Error('В файле оборотов Сбера не найдена колонка ИНН');

  // Показатель указан в строках над заголовком: в PL — одна колонка «Оборот», в US — две (NON-US и US).
  var turnoverCols = [], incomeCol = -1, period = '';
  var width = values[h.row].length;
  for (var j = 0; j < width; j++) {
    if (j === h.col) continue;
    for (var i = 0; i < h.row; i++) {
      var label = headerKey_(values[i][j]);
      if (label === 'оборот') { turnoverCols.push(j); break; }
      if (label === 'комиссионный доход') { incomeCol = j; break; }
    }
  }
  if (!turnoverCols.length) turnoverCols = [h.col + 1];
  for (var p = 0; p < h.row; p++) {
    var at = values[p].map(headerKey_).indexOf('отчетная дата');
    if (at >= 0) period = str_(values[p][at + 1]);
  }

  var map = {};
  for (var r = h.row + 1; r < values.length; r++) {
    var inn = normInn_(values[r][h.col]);
    if (!isInn_(inn)) continue;
    var m = map[inn] || (map[inn] = { turnover: 0, income: 0 });
    turnoverCols.forEach(function (c) { m.turnover += num_(values[r][c]); });
    if (incomeCol >= 0) m.income += num_(values[r][incomeCol]);
  }
  return { map: map, period: period };
}

function topBanks_(banks) {
  return Object.keys(banks)
    .sort(function (a, b) { return banks[b] - banks[a]; })
    .slice(0, 3)
    .map(function (b) { return b + ' ' + money_(banks[b]); })
    .join('; ');
}

// Автоматический потенциал (источник «PS» или пусто) обновляется по свежему файлу, ручной — не трогаем.
function refreshPotentials_(ps) {
  return withLock_(function () {
    var sh = sheet_(APP.SHEET.ACT), last = sh.getLastRow();
    if (last < 2) return 0;
    var inns = sh.getRange(2, COL.INN, last - 1, 1).getValues();
    var pot = sh.getRange(2, COL.POTENTIAL, last - 1, 1).getValues();
    var src = sh.getRange(2, COL.POT_SRC, last - 1, 1).getValues();
    var changed = 0;
    for (var i = 0; i < inns.length; i++) {
      if (str_(src[i][0]) === 'вручную') continue;
      var p = ps[normInn_(inns[i][0])];
      if (!p) continue;
      var v = Math.round(p.total);
      if (num_(pot[i][0]) !== v) changed++;
      pot[i][0] = v;
      src[i][0] = 'PS';
    }
    if (changed) {
      sh.getRange(2, COL.POTENTIAL, last - 1, 1).setValues(pot);
      sh.getRange(2, COL.POT_SRC, last - 1, 1).setValues(src);
    }
    return changed;
  });
}

/** Обороты по ИНН из листа «Обороты». */
function turnRowToObj_(r) {
  return {
    inn: normInn_(r[0]),
    ps: num_(r[1]),
    banks: str_(r[2]),
    sbp: num_(r[3]),
    ecom: num_(r[4]),
    sber: num_(r[5]),
    income: num_(r[6]),
    psFile: str_(r[7]),
    sberFile: str_(r[8])
  };
}

function turnoverFor_(inn) {
  var sh = sheet_(APP.SHEET.TURN), last = sh.getLastRow();
  if (last < 2) return null;
  var hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(inn).matchEntireCell(true).findNext();
  if (!hit) return null;
  return turnRowToObj_(sh.getRange(hit.getRow(), 1, 1, TURN_HEADERS.length).getValues()[0]);
}

function readTurnovers_() {
  var sh = sheet_(APP.SHEET.TURN), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TURN_HEADERS.length).getValues().map(turnRowToObj_);
}

function importStatus_() {
  return {
    at: prop_('IMPORT_AT') || '',
    msg: prop_('IMPORT_MSG') || 'Импорт ещё не запускался',
    psFile: prop_('PS_FILE') || '',
    sberFile: prop_('SBER_FILE') || ''
  };
}
