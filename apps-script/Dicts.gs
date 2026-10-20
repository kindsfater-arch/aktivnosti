/** Справочники: каналы, статусы, виды активности. Хранятся на листе «Справочники». */

function getDicts_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dicts');
  if (cached) return JSON.parse(cached);

  var kindByLabel = {};
  Object.keys(KIND_LABEL).forEach(function (k) { kindByLabel[KIND_LABEL[k].toLowerCase()] = k; });

  var sh = sheet_(APP.SHEET.DICT), last = sh.getLastRow();
  var vals = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
  var d = { channels: [], statuses: [], types: [] };
  vals.forEach(function (r) {
    if (str_(r[0])) d.channels.push(str_(r[0]));
    if (str_(r[1])) d.statuses.push({ name: str_(r[1]), kind: kindByLabel[str_(r[2]).toLowerCase()] || 'open' });
    if (str_(r[3])) d.types.push(str_(r[3]));
  });
  if (!d.channels.length) d.channels = DEFAULT_DICTS.channels.slice();
  if (!d.statuses.length) d.statuses = DEFAULT_DICTS.statuses.slice();
  if (!d.types.length) d.types = DEFAULT_DICTS.types.slice();

  cache.put('dicts', JSON.stringify(d), 600);
  return d;
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
      d.types[i] || ''
    ]);
  }
  sh.getRange(1, 1, 1, 4).setValues([DICT_HEADERS]).setFontWeight('bold');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  CacheService.getScriptCache().remove('dicts');
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
      return { name: str_(s.name), kind: KIND_LABEL[s.kind] ? s.kind : 'open' };
    }).filter(function (s) { return s.name; })
  };
  if (!d.channels.length || !d.types.length || !d.statuses.length) {
    throw new Error('Каждый справочник должен содержать хотя бы одно значение');
  }
  writeDicts_(d);
  log_(user, 'Справочники изменены');
  return getDicts_();
}

// Выпадающие списки и проверка дат прямо в таблице — на случай правки руками.
function applyValidations_() {
  var d = getDicts_();
  var sh = sheet_(APP.SHEET.ACT);
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  var list = function (values) {
    return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  };
  sh.getRange(2, COL.CHANNEL, rows, 1).setDataValidation(list(d.channels));
  sh.getRange(2, COL.STATUS, rows, 1).setDataValidation(list(d.statuses.map(function (s) { return s.name; })));
  sh.getRange(2, COL.TYPE, rows, 1).setDataValidation(list(d.types));
  var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
    .setHelpText('Введите дату в формате ДД.ММ.ГГГГ').build();
  sh.getRange(2, COL.DATE, rows, 1).setDataValidation(dateRule);
  sh.getRange(2, COL.NEXT, rows, 1).setDataValidation(dateRule);
}
