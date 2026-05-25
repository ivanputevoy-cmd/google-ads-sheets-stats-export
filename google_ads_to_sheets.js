/**
 * Google Ads → Google Sheets
 * Раз в 4 часа дописывает снимок статистики ЗА ВСЁ ВРЕМЯ (накопленный итог)
 * в иерархическом виде. Старые строки не стираются — история копится сверху вниз.
 *
 * Хранение: каждый ПЕРИОД (по умолчанию месяц) пишется в ОТДЕЛЬНУЮ таблицу-файл.
 * Файлы создаются автоматически, ссылки на них собираются в лист «Файлы» в этой
 * (стартовой) таблице — она работает как реестр. Место в Google не кончится.
 *
 * В каждом файле три листа:
 *   «Сводка»          — открывается первой: итоги по аккаунту + топ кампаний (на момент скана)
 *   «Статистика»      — Кампания → её Группы → их Объявления
 *   «Ключевые слова»  — Кампания → её Группы → их Ключевые слова
 *
 * Оформление для лёгкого чтения: разделители тысяч и валюта в деньгах, проценты со знаком,
 * замороженные шапка и столбец названий, цвет строк по уровню (кампания/группа/объявление),
 * цветная подсветка ROAS и конверсий.
 *
 * Новые кампании, группы, объявления и ключи подхватываются автоматически.
 *
 * Дополнительно (необязательно): актуальная сводка (итог по аккаунту + все активные
 * кампании) может отправляться в Airtable для CRM команды — см. настройки AIRTABLE_*.
 *
 * Куда вставлять: кабинет Google Ads → Инструменты и настройки →
 *   Массовые действия → Скрипты → «+» → вставить код → Авторизовать →
 *   Просмотр → расписание «Ежечасно». «Время скана» — всегда по Москве.
 *
 * ─────────────────────────────────────────────────────────────────────────
 * Как устроен код (для тех, кто читает впервые):
 *   main()                 — точка входа, Google вызывает её по расписанию
 *   findOrCreatePeriodFile — находит или создаёт таблицу-файл текущего периода
 *   fetchTree()            — тянет данные из Google Ads и строит дерево кампаний
 *   emitRows()             — разворачивает дерево в строки таблицы с отступами
 *   appendTree()           — дописывает строки и раскрашивает их по уровню
 *   writeSummary()         — перерисовывает лист «Сводка»
 *
 * Ключевой инвариант: метрики везде идут в одном и том же порядке — его задаёт
 *   metricRow(): [показы, клики, CTR %, ср. CPC, расход, конверсии, CR %, CPA,
 *   ценность, ROAS]. На этот порядок завязаны buildHeader() (подписи столбцов),
 *   applyDetailFormat() (форматы столбцов E…N) и writeSummary(). Меняешь порядок —
 *   меняй во всех четырёх местах.
 *
 * Перед использованием впиши свой REGISTRY_ID (см. ниже) — ID Google-таблицы,
 * которая станет реестром создаваемых файлов.
 * ─────────────────────────────────────────────────────────────────────────
 */

// ── Настройки ──────────────────────────────────────────────────────────────
var REGISTRY_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE'; // ID Google-таблицы-реестра (часть ссылки между /d/ и /edit)
var REGISTRY_SHEET = 'Файлы';
var SUMMARY_SHEET = 'Сводка';
var STATS_SHEET = 'Статистика';
var KEYWORDS_SHEET = 'Ключевые слова';

var PERIOD = 'month';              // 'month' = новый файл каждый месяц, 'year' = каждый год
var FILE_PREFIX = 'Google Ads — '; // имя создаваемых файлов: «Google Ads — 05.2026»
var DRIVE_FOLDER_ID = '';          // пусто = корень Google Диска; ID папки = складывать туда
var SHARE_PUBLIC_EDIT = true;      // true = новые файлы сразу открыты на редактирование всем по ссылке
var TOP_N = 10;                    // сколько кампаний показывать в «Сводке»

// ── Airtable (необязательно; оставь плейсхолдеры — и отправка выключена) ──
// Раз в скан в Airtable отправляются итог по аккаунту и все активные кампании
// (обновлением существующих записей, без дублей). История остаётся в Google Sheets.
var AIRTABLE_TOKEN = 'PASTE_YOUR_AIRTABLE_TOKEN_HERE';     // Personal Access Token, право data.records:write
var AIRTABLE_BASE_ID = 'PASTE_YOUR_AIRTABLE_BASE_ID_HERE'; // ID базы (начинается с «app...»)
var AIRTABLE_TABLE = 'Google Ads';                         // название таблицы в Airtable

var EVERY_N_HOURS = 4;
var TIMEZONE = 'Europe/Moscow';

var STATUS_RU = { ENABLED: 'Активна', PAUSED: 'Пауза', REMOVED: 'Удалена' };
var MATCH_RU  = { EXACT: 'Точное', PHRASE: 'Фразовое', BROAD: 'Широкое' };

var IND_GROUP = '   ↳ ';
var IND_LEAF  = '      ↳ ';

// Цвета
var C_HEADER = '#b30000';
var LEVEL_BG = { 'Кампания': '#fadbd8', 'Группа': '#fdebd0', 'Объявление': '#ffffff', 'Ключ': '#ffffff' };
var C_GOOD = '#c6efce', C_MID = '#ffe599', C_BAD = '#f8cbad', C_ZERO = '#efefef';

// Форматы чисел
var F_INT = '#,##0', F_MONEY = '#,##0.00', F_PCT = '0.00"%"', F_X = '0.00';

// ── Точка входа ──────────────────────────────────────────────────────────────
function main() {
  var account = AdsApp.currentAccount();
  var now = new Date();
  var stamp = Utilities.formatDate(now, TIMEZONE, 'dd.MM.yyyy HH:mm');
  var periodKey = Utilities.formatDate(now, TIMEZONE, PERIOD === 'year' ? 'yyyy' : 'MM.yyyy');

  var registry = SpreadsheetApp.openById(REGISTRY_ID);
  var fileId = findOrCreatePeriodFile(registry, periodKey, stamp);

  var ss = SpreadsheetApp.openById(fileId);
  var stats = ss.getSheetByName(STATS_SHEET) || ss.insertSheet(STATS_SHEET);

  if (!shouldWrite(stats, stamp)) {
    Logger.log('Пропуск: с прошлой записи прошло меньше ' + EVERY_N_HOURS + ' ч. (файл ' + periodKey + ').');
    return;
  }

  var cur = account.getCurrencyCode();
  var tree = fetchTree();

  var keywords = ss.getSheetByName(KEYWORDS_SHEET) || ss.insertSheet(KEYWORDS_SHEET);
  ensureHeader(stats, buildHeader(cur));
  ensureHeader(keywords, buildHeader(cur));

  appendTree(stats, emitRows(stamp, tree, false));
  appendTree(keywords, emitRows(stamp, tree, true));
  stats.autoResizeColumns(1, 4);
  keywords.autoResizeColumns(1, 4);

  writeSummary(ss, stamp, cur, tree);

  // Отправка в Airtable не должна ронять основную выгрузку — поэтому в отдельном try.
  try { pushToAirtable(tree); } catch (e) { Logger.log('Airtable: ошибка — ' + e); }

  Logger.log('Файл ' + periodKey + ' обновлён: ' + stamp + ' (' + cur + ').');
}

// ── Реестр периодных файлов ─────────────────────────────────────────────────────
function findOrCreatePeriodFile(registry, periodKey, stamp) {
  var reg = registry.getSheetByName(REGISTRY_SHEET) || registry.insertSheet(REGISTRY_SHEET);
  if (reg.getLastRow() === 0) {
    reg.getRange(1, 1, 1, 4).setValues([['Период', 'Файл', 'ID файла', 'Создан']])
       .setFontWeight('bold').setBackground(C_HEADER).setFontColor('#ffffff');
    reg.setFrozenRows(1);
  }

  var last = reg.getLastRow();
  if (last >= 2) {
    var data = reg.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === periodKey && data[i][2]) {
        return String(data[i][2]);
      }
    }
  }

  var created = SpreadsheetApp.create(FILE_PREFIX + periodKey);
  var def = created.getSheets()[0];
  created.insertSheet(STATS_SHEET);
  created.insertSheet(KEYWORDS_SHEET);
  created.deleteSheet(def);
  moveToFolder(created.getId());
  shareFile(created.getId());

  var row = reg.getLastRow() + 1;
  reg.getRange(row, 1).setValue(periodKey);
  reg.getRange(row, 2).setFormula('=HYPERLINK("' + created.getUrl() + '";"открыть")');
  reg.getRange(row, 3).setValue(created.getId());
  reg.getRange(row, 4).setValue(stamp);

  Logger.log('Создан новый файл за период ' + periodKey + ': ' + created.getUrl());
  return created.getId();
}

function moveToFolder(fileId) {
  if (!DRIVE_FOLDER_ID) { return; }
  var file = DriveApp.getFileById(fileId);
  DriveApp.getFolderById(DRIVE_FOLDER_ID).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

// Открывает файл на редактирование всем по ссылке. Если домен/настройки это запрещают,
// файл всё равно создан — пишем предупреждение в журнал, а не валим весь запуск.
function shareFile(fileId) {
  if (!SHARE_PUBLIC_EDIT) { return; }
  try {
    DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (e) {
    Logger.log('Не удалось открыть доступ по ссылке для файла ' + fileId + ': ' + e);
  }
}

// ── Сбор дерева из Google Ads ───────────────────────────────────────────────────
// Делает четыре запроса (кампании, группы, объявления, ключи) и связывает их в дерево
// campaigns[id] → groups[id] → ads[]/keywords[]. Период в запросах НЕ указан — значит
// Google отдаёт метрики за всё время. Фильтр только по статусу != REMOVED, поэтому новые
// сущности попадают автоматически. Заодно копит totals — суммы по аккаунту для «Сводки».
// Возвращает { campaigns, order (порядок кампаний по расходу), totals }.
function fetchTree() {
  var campaigns = {};
  var order = [];
  var totals = { imp: 0, clicks: 0, cost: 0, conv: 0, value: 0 };

  iterate('SELECT campaign.id, campaign.name, campaign.status, ' + metricFields() +
          ' FROM campaign WHERE campaign.status != "REMOVED"' +
          ' ORDER BY metrics.cost_micros DESC', function (r) {
    campaigns[r.campaign.id] = {
      name: r.campaign.name, status: r.campaign.status, metrics: metricRow(r.metrics),
      groups: {}, groupOrder: []
    };
    order.push(r.campaign.id);
    totals.imp += num(r.metrics.impressions);
    totals.clicks += num(r.metrics.clicks);
    totals.cost += num(r.metrics.costMicros) / 1e6;
    totals.conv += num(r.metrics.conversions);
    totals.value += num(r.metrics.conversionsValue);
  });

  iterate('SELECT campaign.id, ad_group.id, ad_group.name, ad_group.status, ' + metricFields() +
          ' FROM ad_group WHERE ad_group.status != "REMOVED"' +
          ' ORDER BY metrics.cost_micros DESC', function (r) {
    var c = campaigns[r.campaign.id];
    if (!c) { return; }
    c.groups[r.adGroup.id] = {
      name: r.adGroup.name, status: r.adGroup.status, metrics: metricRow(r.metrics),
      ads: [], keywords: []
    };
    c.groupOrder.push(r.adGroup.id);
  });

  iterate('SELECT campaign.id, ad_group.id, ad_group_ad.ad.id, ad_group_ad.ad.type, ' +
          'ad_group_ad.status, ' + metricFields() +
          ' FROM ad_group_ad WHERE ad_group_ad.status != "REMOVED"' +
          ' ORDER BY metrics.cost_micros DESC', function (r) {
    var g = groupOf(campaigns, r);
    if (!g) { return; }
    g.ads.push({
      level: 'Объявление',
      label: 'Объявление #' + r.adGroupAd.ad.id + ' (' + r.adGroupAd.ad.type + ')',
      status: r.adGroupAd.status, metrics: metricRow(r.metrics)
    });
  });

  iterate('SELECT campaign.id, ad_group.id, ad_group_criterion.keyword.text, ' +
          'ad_group_criterion.keyword.match_type, ad_group_criterion.status, ' + metricFields() +
          ' FROM keyword_view WHERE ad_group_criterion.status != "REMOVED"' +
          ' ORDER BY metrics.cost_micros DESC', function (r) {
    var g = groupOf(campaigns, r);
    if (!g) { return; }
    var kw = r.adGroupCriterion.keyword;
    g.keywords.push({
      level: 'Ключ',
      label: '"' + kw.text + '" [' + matchRu(kw.matchType) + ']',
      status: r.adGroupCriterion.status, metrics: metricRow(r.metrics)
    });
  });

  return { campaigns: campaigns, order: order, totals: totals };
}

// Находит в дереве группу, к которой относится строка ответа (по id кампании и группы).
// Возвращает null, если кампания/группа не попали в дерево (например, удалены).
function groupOf(campaigns, r) {
  var c = campaigns[r.campaign.id];
  return c ? c.groups[r.adGroup.id] : null;
}

// ── Разворачивание дерева в строки с отступами ────────────────────────────────
// isKeywords = false → листья это объявления (для листа «Статистика»);
//              true  → листья это ключевые слова (для листа «Ключевые слова»).
// Кампании и группы без листьев нужного типа пропускаются, чтобы не было пустых веток
// (например, у медийной кампании нет ключей — на листе ключей она не появится).
function emitRows(stamp, tree, isKeywords) {
  var rows = [];
  tree.order.forEach(function (cid) {
    var c = tree.campaigns[cid];
    var campaignRows = [];

    c.groupOrder.forEach(function (gid) {
      var g = c.groups[gid];
      var leaves = isKeywords ? g.keywords : g.ads;
      if (!leaves.length) { return; }
      campaignRows.push([stamp, 'Группа', IND_GROUP + g.name, statusRu(g.status)].concat(g.metrics));
      leaves.forEach(function (x) {
        campaignRows.push([stamp, x.level, IND_LEAF + x.label, statusRu(x.status)].concat(x.metrics));
      });
    });

    if (!campaignRows.length) { return; }
    rows.push([stamp, 'Кампания', c.name, statusRu(c.status)].concat(c.metrics));
    rows = rows.concat(campaignRows);
  });
  return rows;
}

// ── Лист «Сводка» (перезаписывается каждый скан) ────────────────────────────────
function writeSummary(ss, stamp, cur, tree) {
  var sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  sh.clear();
  sh.getRange('A1:J1').breakApart();

  var t = tree.totals;
  var ctr  = t.imp > 0 ? t.clicks / t.imp * 100 : 0;
  var cpc  = t.clicks > 0 ? t.cost / t.clicks : 0;
  var cr   = t.clicks > 0 ? t.conv / t.clicks * 100 : 0;
  var cpa  = t.conv > 0 ? t.cost / t.conv : 0;
  var roas = t.cost > 0 ? t.value / t.cost : 0;

  sh.getRange('A1').setValue('СВОДКА ПО АККАУНТУ');
  sh.getRange('A1:J1').merge().setBackground(C_HEADER).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sh.getRange('A2').setValue('Обновлено: ' + stamp + ' (МСК)').setFontWeight('bold');
  sh.getRange('A3').setValue('Валюта: ' + cur);

  sh.getRange('A5').setValue('Итого за всё время').setFontWeight('bold').setFontSize(12);
  var totHead = ['Показы', 'Клики', 'CTR %', 'Ср. CPC', 'Расход', 'Конверсии', 'CR %', 'CPA', 'Ценность', 'ROAS'];
  sh.getRange(6, 1, 1, totHead.length).setValues([totHead]).setFontWeight('bold').setBackground('#f2f2f2');
  sh.getRange(7, 1, 1, 10).setValues([[t.imp, t.clicks, r2(ctr), r2(cpc), r2(t.cost),
                                       r2(t.conv), r2(cr), r2(cpa), r2(t.value), r2(roas)]]);
  setRowFormats(sh, 7, [F_INT, F_INT, F_PCT, F_MONEY, F_MONEY, F_MONEY, F_PCT, F_MONEY, F_MONEY, F_X]);

  sh.getRange('A9').setValue('Топ кампаний по расходу').setFontWeight('bold').setFontSize(12);
  var topHead = ['Кампания', 'Статус', 'Показы', 'Клики', 'CTR %', 'Расход', 'Конверсии', 'CPA', 'ROAS'];
  sh.getRange(10, 1, 1, topHead.length).setValues([topHead]).setFontWeight('bold').setBackground('#f2f2f2');

  var n = Math.min(TOP_N, tree.order.length);
  if (n > 0) {
    var data = [];
    for (var i = 0; i < n; i++) {
      var c = tree.campaigns[tree.order[i]];
      var m = c.metrics; // [imp,clicks,ctr,cpc,cost,conv,cr,cpa,value,roas]
      data.push([c.name, statusRu(c.status), m[0], m[1], m[2], m[4], m[5], m[7], m[9]]);
    }
    sh.getRange(11, 1, n, topHead.length).setValues(data);
    var fmt = [null, null, F_INT, F_INT, F_PCT, F_MONEY, F_MONEY, F_MONEY, F_X];
    for (var col = 3; col <= 9; col++) {
      sh.getRange(11, col, n, 1).setNumberFormat(fmt[col - 1]);
    }
  }

  sh.autoResizeColumns(1, 10);
  sh.setColumnWidth(1, 260);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}

function setRowFormats(sheet, row, formats) {
  for (var i = 0; i < formats.length; i++) {
    sheet.getRange(row, i + 1).setNumberFormat(formats[i]);
  }
}

// ── Запись детальных листов ───────────────────────────────────────────────────
function buildHeader(cur) {
  return ['Время скана', 'Уровень', 'Название', 'Статус',
          'Показы', 'Клики', 'CTR %', 'Ср. CPC, ' + cur, 'Расход, ' + cur,
          'Конверсии', 'CR %', 'CPA, ' + cur, 'Ценность конв., ' + cur, 'ROAS'];
}

function ensureHeader(sheet, header) {
  if (sheet.getLastRow() > 0) { return; }
  sheet.getRange(1, 1, 1, header.length).setValues([header])
       .setFontWeight('bold').setBackground(C_HEADER).setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  applyDetailFormat(sheet);
}

// Форматы чисел и подсветка — задаются один раз на открытые диапазоны столбцов,
// поэтому действуют и на все будущие строки.
function applyDetailFormat(sheet) {
  sheet.setFrozenColumns(3);
  sheet.getRange('E2:E').setNumberFormat(F_INT);    // Показы
  sheet.getRange('F2:F').setNumberFormat(F_INT);    // Клики
  sheet.getRange('G2:G').setNumberFormat(F_PCT);    // CTR %
  sheet.getRange('H2:H').setNumberFormat(F_MONEY);  // Ср. CPC
  sheet.getRange('I2:I').setNumberFormat(F_MONEY);  // Расход
  sheet.getRange('J2:J').setNumberFormat(F_MONEY);  // Конверсии
  sheet.getRange('K2:K').setNumberFormat(F_PCT);    // CR %
  sheet.getRange('L2:L').setNumberFormat(F_MONEY);  // CPA
  sheet.getRange('M2:M').setNumberFormat(F_MONEY);  // Ценность
  sheet.getRange('N2:N').setNumberFormat(F_X);      // ROAS

  var roas = sheet.getRange('N2:N');
  var conv = sheet.getRange('J2:J');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setBackground(C_BAD).setRanges([roas]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1, 2).setBackground(C_MID).setRanges([roas]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(2).setBackground(C_GOOD).setRanges([roas]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground(C_ZERO).setRanges([conv]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C_GOOD).setRanges([conv]).build()
  ]);
}

function appendTree(sheet, rows) {
  if (!rows.length) { return; }
  var startRow = sheet.getLastRow() + 1;
  var width = rows[0].length;
  sheet.getRange(startRow, 1, rows.length, width).setValues(rows);
  paintLevels(sheet, startRow, rows, width);
}

// Заливка строк по уровню + жирные строки кампаний.
function paintLevels(sheet, startRow, rows, width) {
  var bg = [], fw = [];
  for (var i = 0; i < rows.length; i++) {
    var lvl = rows[i][1];
    var color = LEVEL_BG[lvl] || '#ffffff';
    var weight = lvl === 'Кампания' ? 'bold' : 'normal';
    var bgLine = [], fwLine = [];
    for (var j = 0; j < width; j++) { bgLine.push(color); fwLine.push(weight); }
    bg.push(bgLine); fw.push(fwLine);
  }
  var rng = sheet.getRange(startRow, 1, rows.length, width);
  rng.setBackgrounds(bg);
  rng.setFontWeights(fw);
}

// Пишем, если данных ещё нет или с прошлой записи прошло >= (EVERY_N_HOURS ч − 10 мин).
function shouldWrite(sheet, stamp) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { return true; }

  var minGap = EVERY_N_HOURS * 60 - 10;
  var raw = sheet.getRange(lastRow, 1).getValue();
  var lastMin, nowMin;
  if (raw instanceof Date) {
    lastMin = Math.floor(raw.getTime() / 60000);
    nowMin = Math.floor(new Date().getTime() / 60000);
  } else {
    lastMin = wallMinutes(String(raw));
    nowMin = wallMinutes(stamp);
  }
  if (lastMin === null || nowMin === null) { return true; }
  return (nowMin - lastMin) >= minGap;
}

// Разбирает метку «дд.мм.гггг чч:мм» в число минут (для сравнения интервала между сканами).
// Возвращает null, если строка не похожа на дату.
function wallMinutes(s) {
  var m = String(s).match(/(\d{2})\.(\d{2})\.(\d{4})\D+(\d{2}):(\d{2})/);
  if (!m) { return null; }
  return Date.UTC(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]) / 60000;
}

// ── Airtable (живая сводка для CRM команды) ─────────────────────────────────────
// Отправляет в Airtable итог по аккаунту и все активные кампании, ОБНОВЛЯЯ записи
// (upsert по полю «ID»), а не плодя новые. Кампании, пропавшие из выгрузки, помечает
// статусом «Неактивна». Если токен/база не заданы (плейсхолдеры) — тихо пропускает.
function pushToAirtable(tree) {
  if (!airtableEnabled()) {
    Logger.log('Airtable выключен (не заданы токен/база) — пропускаю.');
    return;
  }
  var iso = new Date().toISOString(); // Airtable хранит время в UTC и показывает в поясе пользователя

  var records = [totalFields(tree.totals, iso)];
  var seen = { 'ACCOUNT': true };
  tree.order.forEach(function (cid) {
    records.push(campaignFields(cid, tree.campaigns[cid], iso));
    seen[String(cid)] = true;
  });

  airtableUpsert(records);
  airtableMarkMissing(seen);
  Logger.log('Airtable обновлён: ' + records.length + ' записей.');
}

function airtableEnabled() {
  return AIRTABLE_TOKEN.indexOf('PASTE_') !== 0 && AIRTABLE_BASE_ID.indexOf('PASTE_') !== 0;
}

// Поля одной кампании. Порядок метрик — из metricRow(): [imp,clicks,ctr,cpc,cost,conv,cr,cpa,value,roas].
function campaignFields(cid, c, iso) {
  var m = c.metrics;
  return {
    'ID': String(cid), 'Уровень': 'Кампания', 'Кампания': c.name, 'Статус': statusRu(c.status),
    'Показы': m[0], 'Клики': m[1], 'CTR %': m[2], 'Ср. CPC': m[3], 'Расход': m[4],
    'Конверсии': m[5], 'CR %': m[6], 'CPA': m[7], 'Ценность': m[8], 'ROAS': m[9],
    'Обновлено': iso
  };
}

// Одна запись с итогом по аккаунту (производные считаем из сырых сумм totals).
function totalFields(t, iso) {
  var ctr  = t.imp > 0 ? t.clicks / t.imp * 100 : 0;
  var cpc  = t.clicks > 0 ? t.cost / t.clicks : 0;
  var cr   = t.clicks > 0 ? t.conv / t.clicks * 100 : 0;
  var cpa  = t.conv > 0 ? t.cost / t.conv : 0;
  var roas = t.cost > 0 ? t.value / t.cost : 0;
  return {
    'ID': 'ACCOUNT', 'Уровень': 'Итого', 'Кампания': 'ИТОГО ПО АККАУНТУ', 'Статус': 'Активна',
    'Показы': t.imp, 'Клики': t.clicks, 'CTR %': r2(ctr), 'Ср. CPC': r2(cpc), 'Расход': r2(t.cost),
    'Конверсии': r2(t.conv), 'CR %': r2(cr), 'CPA': r2(cpa), 'Ценность': r2(t.value), 'ROAS': r2(roas),
    'Обновлено': iso
  };
}

// Создаёт/обновляет записи пачками по 10 (лимит Airtable на запрос).
function airtableUpsert(records) {
  var url = airtableUrl();
  for (var i = 0; i < records.length; i += 10) {
    var batch = records.slice(i, i + 10).map(function (f) { return { fields: f }; });
    airtableFetch('patch', url, { performUpsert: { fieldsToMergeOn: ['ID'] }, records: batch, typecast: true });
    Utilities.sleep(250); // держимся в пределах лимита Airtable (~5 запросов/сек)
  }
}

// Кампании, которых нет в свежей выгрузке, помечаем «Неактивна» (не удаляем).
function airtableMarkMissing(seenIds) {
  var stale = [];
  airtableList().forEach(function (rec) {
    var f = rec.fields || {};
    if (f['Уровень'] === 'Кампания' && !seenIds[String(f['ID'])] && f['Статус'] !== 'Неактивна') {
      stale.push({ id: rec.id, fields: { 'Статус': 'Неактивна' } });
    }
  });
  var url = airtableUrl();
  for (var i = 0; i < stale.length; i += 10) {
    airtableFetch('patch', url, { records: stale.slice(i, i + 10), typecast: true });
    Utilities.sleep(250);
  }
  if (stale.length) { Logger.log('Airtable: помечено неактивными ' + stale.length + ' кампаний.'); }
}

// Читает все записи таблицы (постранично по 100).
function airtableList() {
  var out = [];
  var offset = null;
  do {
    var url = airtableUrl() + '?pageSize=100' + (offset ? '&offset=' + encodeURIComponent(offset) : '');
    var data = airtableFetch('get', url, null);
    if (!data) { break; }
    (data.records || []).forEach(function (r) { out.push(r); });
    offset = data.offset;
  } while (offset);
  return out;
}

function airtableUrl() {
  return 'https://api.airtable.com/v0/' + AIRTABLE_BASE_ID + '/' + encodeURIComponent(AIRTABLE_TABLE);
}

// Запрос к Airtable. Ошибки логируем и возвращаем null (не валим скрипт).
function airtableFetch(method, url, payload) {
  var options = {
    method: method, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + AIRTABLE_TOKEN },
    muteHttpExceptions: true
  };
  if (payload) { options.payload = JSON.stringify(payload); }
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('Airtable ' + method.toUpperCase() + ' ' + code + ': ' + res.getContentText());
    return null;
  }
  var text = res.getContentText();
  return text ? JSON.parse(text) : {};
}

// ── Метрики ───────────────────────────────────────────────────────────────────
// Поля метрик, которые запрашиваются у Google Ads (общие для всех четырёх запросов).
function metricFields() {
  return 'metrics.impressions, metrics.clicks, metrics.ctr, metrics.average_cpc, ' +
         'metrics.cost_micros, metrics.conversions, metrics.conversions_value';
}

// Превращает сырые метрики Google в строку из 10 чисел в ФИКСИРОВАННОМ порядке:
//   [показы, клики, CTR %, ср. CPC, расход, конверсии, CR %, CPA, ценность, ROAS].
// Деньги Google отдаёт в «микро-единицах» (1/1 000 000), поэтому делим на 1e6.
// CTR/CR/CPA/ROAS — производные, считаем сами. Этот порядок — общий контракт со всеми,
// кто читает метрики (buildHeader, applyDetailFormat, writeSummary).
function metricRow(m) {
  var clicks = num(m.clicks);
  var cost   = num(m.costMicros) / 1e6;
  var conv   = num(m.conversions);
  var value  = num(m.conversionsValue);
  return [
    num(m.impressions),
    clicks,
    r2(num(m.ctr) * 100),
    r2(num(m.averageCpc) / 1e6),
    r2(cost),
    r2(conv),
    r2(clicks > 0 ? (conv / clicks) * 100 : 0),
    r2(conv > 0 ? cost / conv : 0),
    r2(value),
    r2(cost > 0 ? value / cost : 0)
  ];
}

// Выполняет GAQL-запрос к Google Ads и вызывает onRow для каждой строки результата.
function iterate(query, onRow) {
  var it = AdsApp.search(query);
  while (it.hasNext()) { onRow(it.next()); }
}

function statusRu(s) { return STATUS_RU[s] || s; }     // ENABLED → «Активна» и т.д.
function matchRu(s) { return MATCH_RU[s] || s; }       // EXACT → «Точное» и т.д.
function num(v) { var n = Number(v); return isFinite(n) ? n : 0; } // безопасное число (NaN/undefined → 0)
function r2(n) { return Math.round(n * 100) / 100; }   // округление до 2 знаков
