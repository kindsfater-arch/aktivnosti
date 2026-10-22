/** Справочники: каналы, статусы, виды активности. Хранятся на листе «Справочники». */

function getDicts_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dicts2');
  if (cached) return JSON.parse(cached);

  var sh = sheet_(APP.SHEET.DICT);
  var d = readDictSheet_(sh);
  if (d.statuses.length && !d.hasActiveCol) {
    migrateActiveStatuses_(d);
    d = readDictSheet_(sh);
  }
  delete d.hasActiveCol;
  if (!d.channels.length) d.channels = DEFAULT_DICTS.channels.slice();
  if (!d.statuses.length) d.statuses = DEFAULT_DICTS.statuses.slice();
  if (!d.types.length) d.types = DEFAULT_DICTS.types.slice();

  cache.put('dicts2', JSON.stringify(d), 600);
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
  CacheService.getScriptCache().remove('dicts2');
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
