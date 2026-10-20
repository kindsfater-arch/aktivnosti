/**
 * Выгрузки: отчёт в Excel / PDF / PowerPoint и полный сводный файл в Excel.
 * Файлы собираются во временных Google-таблицах и презентациях, экспортируются и сразу удаляются.
 */

var ACCENT = '#1A7F37';

function exportReport_(u, args) {
  var rep = buildReport_(u, args);
  var base = 'Отчёт ' + rep.from + (rep.from === rep.to ? '' : '_' + rep.to);
  var ss = reportSpreadsheet_(rep);
  var cleanup = [ss.getId()];
  try {
    if (args.format === 'xlsx') return fileResult_(exportUrl_(ss.getId(), 'xlsx', base + '.xlsx'));
    var deck = reportDeck_(rep, ss);
    cleanup.push(deck.getId());
    if (args.format === 'pptx') {
      return fileResult_(exportUrl_(deck.getId(), 'pptx', base + '.pptx'));
    }
    return fileResult_(DriveApp.getFileById(deck.getId()).getAs(MimeType.PDF).setName(base + '.pdf'));
  } finally {
    cleanup.forEach(trashQuietly_);
  }
}

function exportAll_(u) {
  var kinds = statusKinds_();
  var types = typeByFio_();
  var acts = readActivities_().slice().sort(byDateDesc_);
  var ss = SpreadsheetApp.create('Активности — сводный ' + ymd_(today_()));
  DriveApp.getFileById(ss.getId()).moveTo(tempFolder_());
  try {
    var sh = ss.getSheets()[0].setName('Активности');
    var head = ACT_HEADERS.slice(0, COL.DEAL_SUM).concat(['Тип сотрудника', 'Тип статуса', 'ID', 'Автор', 'Создано']);
    var rows = acts.map(function (a) {
      return [a.fio, a.no, a.inn, a.short, a.full, a.channel, a.potential || '', a.status, toDate_(a.date) || '',
        a.type, toDate_(a.next) || '', a.comment, a.dealSum || '', types[a.fio] || '',
        KIND_LABEL[kinds[a.status]] || '', a.id, a.author, a.created.replace('T', ' ')];
    });
    writeTable_(sh, head, rows, { money: [7, 13], dates: [9, 11], text: [3] });

    var cl = ss.insertSheet('Клиенты');
    var latest = latestByClient_(acts).sort(function (x, y) { return (y.potential || 0) - (x.potential || 0); });
    var count = {};
    acts.forEach(function (a) { var k = a.fio + '|' + a.inn; count[k] = (count[k] || 0) + 1; });
    writeTable_(cl,
      ['ФИО', 'ФИО НО', 'ИНН', 'Краткое название', 'Канал', 'Текущий статус', 'Тип статуса', 'Потенциал оборота в мес',
        'Сумма сделки', 'Последняя активность', 'Следующая активность', 'Всего активностей', 'Последний комментарий'],
      latest.map(function (a) {
        return [a.fio, a.no, a.inn, a.short, a.channel, a.status, KIND_LABEL[kinds[a.status]] || '', a.potential || '',
          a.dealSum || '', toDate_(a.date) || '', toDate_(a.next) || '', count[a.fio + '|' + a.inn], a.comment];
      }),
      { money: [8, 9], dates: [10, 11], text: [3] });
    SpreadsheetApp.flush();
    log_(u, 'Выгрузка сводного файла', '', acts.length + ' строк');
    return fileResult_(exportUrl_(ss.getId(), 'xlsx', 'Активности_сводный_' + ymd_(today_()) + '.xlsx'));
  } finally {
    trashQuietly_(ss.getId());
  }
}

function groupHeaders_(rep) {
  return ['Разрез', 'Активностей'].concat(rep.types)
    .concat(['Клиентов', 'Новых клиентов', 'Сделок', 'Сумма сделок', 'Отказы и блокировки', 'Потенциал, мес']);
}

function groupRow_(rep, r, keyLabel) {
  return [keyLabel || r.key, r.acts].concat(rep.types.map(function (t) { return r.byType[t] || 0; }))
    .concat([r.clients, r.newClients, r.deals, r.dealSum, r.lost, r.potential]);
}

function reportSpreadsheet_(rep) {
  var ss = SpreadsheetApp.create('Отчёт — ' + rep.label);
  DriveApp.getFileById(ss.getId()).moveTo(tempFolder_());
  var s = rep.summary;

  var sum = ss.getSheets()[0].setName('Сводка');
  var kv = [
    ['Отчёт', rep.label],
    ['Сформирован', rep.generatedAt + ', ' + rep.author],
    ['Активностей', s.acts]
  ];
  rep.types.forEach(function (t) { kv.push(['  ' + t, s.byType[t] || 0]); });
  kv = kv.concat([
    ['Клиентов в работе за период', s.clients],
    ['Новых клиентов', s.newClients],
    ['Сделок (клиентов)', s.deals],
    ['Сумма сделок, ₽', s.dealSum],
    ['Отказы и блокировки', s.lost],
    ['Потенциал клиентов за период, ₽/мес', s.potential],
    ['Открытая воронка (клиентов)', rep.pipeline.count],
    ['Потенциал открытой воронки, ₽/мес', rep.pipeline.potential],
    ['Следующие шаги в ближайшие 30 дней', rep.pipeline.next30],
    ['Сотрудники без активностей', rep.idle.join(', ') || '—']
  ]);
  sum.getRange(1, 1, kv.length, 2).setValues(kv);
  sum.getRange(1, 1, kv.length, 1).setFontWeight('bold');
  sum.getRange(1, 2, kv.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('left');
  sum.setColumnWidth(1, 320);
  sum.setColumnWidth(2, 360);

  var head = groupHeaders_(rep);
  var moneyCols = [head.length - 2, head.length];
  rep.groupsMeta.forEach(function (g) {
    var sh = ss.insertSheet(g.title);
    var rows = rep.groups[g.key].map(function (r) {
      return groupRow_(rep, r, g.key === 'day' ? ruDate_(r.key) : r.key);
    });
    rows.push(groupRow_(rep, s, 'Итого'));
    writeTable_(sh, head, rows, { money: moneyCols });
    sh.getRange(rows.length + 1, 1, 1, head.length).setFontWeight('bold');
  });

  var od = ss.insertSheet('Просроченные шаги');
  writeTable_(od, ['ФИО', 'Просроченных следующих активностей'], rep.overdue.map(function (o) { return [o.fio, o.count]; }), {});

  var det = ss.insertSheet('Активности за период');
  writeTable_(det,
    ['Дата', 'ФИО', 'ФИО НО', 'Тип сотрудника', 'ИНН', 'Клиент', 'Канал', 'Вид', 'Статус', 'Потенциал', 'Сумма сделки', 'След. активность', 'Комментарий'],
    rep.details.map(function (a) {
      return [toDate_(a.date) || '', a.fio, a.no, a.empType, a.inn, a.short, a.channel, a.type, a.status,
        a.potential || '', a.dealSum || '', toDate_(a.next) || '', a.comment];
    }),
    { money: [10, 11], dates: [1, 12], text: [5] });

  // График для презентации: активности по сотрудникам.
  var bySh = ss.getSheetByName('По сотрудникам');
  var n = rep.groups.fio.length;
  if (bySh && n) {
    var chart = bySh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(bySh.getRange(1, 1, n + 1, 1))
      .addRange(bySh.getRange(1, 3, n + 1, rep.types.length))
      .setOption('isStacked', true)
      .setOption('title', 'Активности по сотрудникам')
      .setOption('legend', { position: 'bottom' })
      .setOption('colors', ['#1A7F37', '#5BB974', '#A8DAB5', '#F29900'])
      .setOption('width', 900)
      .setOption('height', Math.max(320, 40 + n * 28))
      .setPosition(n + 4, 1, 0, 0)
      .build();
    bySh.insertChart(chart);
  }
  SpreadsheetApp.flush();
  return ss;
}

function writeTable_(sh, head, rows, fmt) {
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(ACCENT).setWrap(true);
  if (rows.length) {
    (fmt.text || []).forEach(function (c) { sh.getRange(2, c, rows.length, 1).setNumberFormat('@'); });
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    (fmt.money || []).forEach(function (c) { sh.getRange(2, c, rows.length, 1).setNumberFormat('#,##0'); });
    (fmt.dates || []).forEach(function (c) { sh.getRange(2, c, rows.length, 1).setNumberFormat('dd.mm.yyyy'); });
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, head.length);
  for (var c = 1; c <= head.length; c++) {
    if (sh.getColumnWidth(c) > 360) sh.setColumnWidth(c, 360);
  }
}

/** Презентация отчёта: титул, ключевые цифры, график, таблицы по разрезам. */
function reportDeck_(rep, ss) {
  var deck = SlidesApp.create('Отчёт — ' + rep.label);
  DriveApp.getFileById(deck.getId()).moveTo(tempFolder_());
  var W = deck.getPageWidth(), H = deck.getPageHeight(), M = 28;
  var s = rep.summary;

  var title = deck.getSlides()[0];
  setPlaceholder_(title, SlidesApp.PlaceholderType.CENTERED_TITLE, 'Активности по эквайрингу');
  setPlaceholder_(title, SlidesApp.PlaceholderType.SUBTITLE, rep.label + '\nСформирован ' + rep.generatedAt);

  var kpi = deck.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
  setPlaceholder_(kpi, SlidesApp.PlaceholderType.TITLE, 'Итоги: ' + rep.label);
  var tiles = [
    [money_(s.acts), 'активностей'],
    [money_(s.clients), 'клиентов в работе'],
    [money_(s.newClients), 'новых клиентов'],
    [money_(s.deals), 'сделок'],
    [money_(s.dealSum) + ' ₽', 'сумма сделок'],
    [money_(s.potential) + ' ₽', 'потенциал клиентов, мес'],
    [money_(s.lost), 'отказы и блокировки'],
    [money_(rep.pipeline.potential) + ' ₽', 'потенциал открытой воронки']
  ];
  var cols = 4, gap = 12, tw = (W - 2 * M - gap * (cols - 1)) / cols, th = 92, top = 100;
  tiles.forEach(function (t, i) {
    var x = M + (i % cols) * (tw + gap), y = top + Math.floor(i / cols) * (th + gap);
    var box = kpi.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, tw, th);
    box.getFill().setSolidFill('#EEF6F0');
    box.getBorder().setTransparent();
    var tr = box.getText();
    tr.setText(t[0] + '\n' + t[1]);
    tr.getTextStyle().setFontSize(11).setForegroundColor('#4A5A50');
    tr.getRange(0, t[0].length).getTextStyle().setFontSize(22).setBold(true).setForegroundColor(ACCENT);
  });
  var byType = rep.types.map(function (t) { return t + ': ' + (s.byType[t] || 0); }).join('   ·   ');
  var note = kpi.insertTextBox(byType + (rep.idle.length ? '\nБез активностей: ' + rep.idle.join(', ') : ''),
    M, top + 2 * (th + gap) + 4, W - 2 * M, 60);
  note.getText().getTextStyle().setFontSize(11).setForegroundColor('#4A5A50');

  var chartSheet = ss.getSheetByName('По сотрудникам');
  var charts = chartSheet ? chartSheet.getCharts() : [];
  if (charts.length) {
    var cs = deck.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    setPlaceholder_(cs, SlidesApp.PlaceholderType.TITLE, 'Активности по сотрудникам');
    cs.insertSheetsChartAsImage(charts[0], M, 80, W - 2 * M, H - 100);
  }

  var head = groupHeaders_(rep);
  rep.groupsMeta.forEach(function (g) {
    var rows = rep.groups[g.key].map(function (r) {
      return groupRow_(rep, r, g.key === 'day' ? ruDate_(r.key) : r.key);
    });
    if (!rows.length) return;
    rows.push(groupRow_(rep, s, 'Итого'));
    var perSlide = 12;
    for (var start = 0; start < rows.length; start += perSlide) {
      var part = rows.slice(start, start + perSlide);
      var sl = deck.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
      setPlaceholder_(sl, SlidesApp.PlaceholderType.TITLE,
        g.title + (rows.length > perSlide ? ' (' + (start / perSlide + 1) + ')' : ''));
      slideTable_(sl, head, part, M, 78, W - 2 * M);
    }
  });

  if (rep.overdue.length) {
    var os = deck.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    setPlaceholder_(os, SlidesApp.PlaceholderType.TITLE, 'Просроченные следующие шаги');
    slideTable_(os, ['ФИО', 'Просрочено'], rep.overdue.slice(0, 14).map(function (o) { return [o.fio, o.count]; }),
      M, 78, (W - 2 * M) / 2);
  }
  deck.saveAndClose();
  return deck;
}

function slideTable_(slide, head, rows, left, top, width) {
  var table = slide.insertTable(rows.length + 1, head.length, left, top, width, 24 * (rows.length + 1));
  var put = function (r, c, v, isHead) {
    var cell = table.getCell(r, c);
    var text = typeof v === 'number' ? money_(v) : String(v == null ? '' : v);
    cell.getText().setText(text);
    var st = cell.getText().getTextStyle().setFontSize(head.length > 6 ? 8 : 10);
    if (isHead) {
      cell.getFill().setSolidFill(ACCENT);
      st.setBold(true).setForegroundColor('#FFFFFF');
    }
    if (c > 0) cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  };
  head.forEach(function (h, c) { put(0, c, h, true); });
  rows.forEach(function (row, r) { row.forEach(function (v, c) { put(r + 1, c, v, false); }); });
  return table;
}

function setPlaceholder_(slide, type, text) {
  var ph = slide.getPlaceholder(type);
  if (ph) ph.asShape().getText().setText(text);
}

function exportUrl_(id, format, name) {
  SpreadsheetApp.flush();
  var url = format === 'pptx'
    ? 'https://docs.google.com/presentation/d/' + id + '/export/pptx'
    : 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=' + format;
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('Не удалось сформировать файл (код ' + resp.getResponseCode() + ')');
  return resp.getBlob().setName(name);
}

function fileResult_(blob) {
  return { name: blob.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

function trashQuietly_(id) {
  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (e) {
    console.warn('trash ' + id + ': ' + e);
  }
}
