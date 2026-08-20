/**
 * FOCUS 新規参加者アンケート - 完成版 Google Apps Script
 * setup() を1回実行すると、回答・集計シートとグラフを自動構築します。
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1bNckj3lWxxKUa50QDxtJZ6-eJra_cUftSird_6bywLY',
  RESPONSE_SHEET: '回答',
  SUMMARY_SHEET: '集計',
  FORM_TYPE: 'FOCUS新規参加者アンケート',
  MULTI_SEPARATOR: /\s*\/\s*/,
  TIME_ZONE: 'Asia/Tokyo'
});

const HEADERS = Object.freeze([
  '回答日時', 'Discordプロフィール名', '性別', '年齢', '現在の職業・立場',
  'FOCUSを知ったきっかけ', '現在の生活・状態', '現在の悩み・課題',
  'FOCUSに入った理由', '入会の決め手', '入会前の不安',
  '特に変えたい・伸ばしたい分野', '一番変えたいこと', '3ヶ月後の目標',
  '数字で表せる目標', '行動力', '継続力', '自己規律', '自信',
  '現在の生活への満足度', '自由記述', 'submission_id', 'クライアント送信日時'
]);

const VISIBLE_COLUMN_COUNT = 21;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const p = e && e.parameter ? e.parameter : {};
    if (clean_(p.form_type) !== CONFIG.FORM_TYPE) throw new Error('不正なフォーム種別です。');
    const submissionId = clean_(p.submission_id);
    if (!submissionId) throw new Error('submission_id がありません。');

    const ss = openSpreadsheet_();
    const sheet = getOrCreateResponseSheet_(ss, false);
    const existingRow = findSubmissionRow_(sheet, submissionId);
    if (existingRow) return output_({ ok: true, duplicate: true, row: existingRow }, p.callback);

    const row = [
      new Date(), clean_(p.profile_name), clean_(p.sex), numberOrBlank_(p.age),
      clean_(p.current_job), clean_(p.discovery), clean_(p.current_situation),
      clean_(p.current_problems), clean_(p.join_reasons), clean_(p.deciding_factor),
      clean_(p.before_concern), clean_(p.focus_area), clean_(p.main_change),
      clean_(p.three_month_goal), clean_(p.numeric_goal), rating_(p.rating_action),
      rating_(p.rating_consistency), rating_(p.rating_discipline),
      rating_(p.rating_confidence), rating_(p.rating_life), clean_(p.free_message),
      submissionId, clean_(p.submitted_at)
    ];
    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();
    formatNewResponseRow_(sheet, rowNumber);
    updateSummary_(ss, false);
    SpreadsheetApp.flush();
    return output_({ ok: true, duplicate: false, row: rowNumber }, p.callback);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return output_({ ok: false, message: error.message || '保存に失敗しました。' });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** GETは稼働確認と、no-cors POST後の保存確認に使用します。 */
function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const submissionId = clean_(p.submission_id);
    let saved = false;
    if (submissionId) {
      const sheet = getOrCreateResponseSheet_(openSpreadsheet_(), false);
      saved = Boolean(findSubmissionRow_(sheet, submissionId));
    }
    return output_({ ok: true, service: 'FOCUS Survey API', saved: saved }, p.callback);
  } catch (error) {
    return output_({ ok: false, saved: false, message: error.message }, e && e.parameter && e.parameter.callback);
  }
}

/** 初回と、レイアウトを再構築したいときに実行。既存回答は削除しません。 */
function setup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = openSpreadsheet_();
    // 既存の回答シートは重い全体再書式・フィルター再作成を行わない。
    getOrCreateResponseSheet_(ss, false);
    updateSummary_(ss, true);
    SpreadsheetApp.flush();
    return 'セットアップ完了：回答データは保持されています。';
  } finally {
    lock.releaseLock();
  }
}

/** モバイル表示不調の診断用。シート内容は変更しません。 */
function auditSpreadsheet() {
  const ss = openSpreadsheet_();
  const result = ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn(),
    charts: sheet.getCharts().length,
    drawings: sheet.getDrawings().length,
    mergedRanges: sheet.getDataRange().getMergedRanges().length,
    conditionalFormatRules: sheet.getConditionalFormatRules().length,
    hasFilter: Boolean(sheet.getFilter()),
    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns()
  }));
  console.log(JSON.stringify(result));
  return result;
}

/** 重複した条件付き書式だけを除去し、回答データと通常書式は保持します。 */
function repairSpreadsheetPerformance() {
  const ss = openSpreadsheet_();
  const responseSheet = ss.getSheetByName(CONFIG.RESPONSE_SHEET);
  const summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (responseSheet) responseSheet.setConditionalFormatRules([]);
  if (summarySheet) summarySheet.setConditionalFormatRules([]);
  SpreadsheetApp.flush();
  return auditSpreadsheet();
}

function openSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateResponseSheet_(ss, applyFormatting) {
  let sheet = ss.getSheetByName(CONFIG.RESPONSE_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.RESPONSE_SHEET);
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (applyFormatting !== false) formatResponseSheet_(sheet);
  return sheet;
}

function formatResponseSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  header.setFontWeight('bold').setBackground('#312e81').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);
  const oldFilter = sheet.getFilter();
  if (oldFilter) oldFilter.remove();
  sheet.getRange(1, 1, lastRow, VISIBLE_COLUMN_COUNT).createFilter();
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.getRange(2, 4, Math.max(lastRow - 1, 1), 1).setNumberFormat('0');
  sheet.getRange(2, 16, Math.max(lastRow - 1, 1), 5).setNumberFormat('0');
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), VISIBLE_COLUMN_COUNT)
    .setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  const widths = [145, 180, 85, 65, 170, 180, 260, 260, 260, 260, 260, 170, 260, 260, 180, 70, 70, 70, 70, 110, 280];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  sheet.hideColumns(VISIBLE_COLUMN_COUNT + 1, HEADERS.length - VISIBLE_COLUMN_COUNT);
  // 条件付き書式はモバイル表示を重くするため使用しない。
  sheet.setConditionalFormatRules([]);
}

function formatNewResponseRow_(sheet, row) {
  sheet.getRange(row, 1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  sheet.getRange(row, 4).setNumberFormat('0');
  sheet.getRange(row, 16, 1, 5).setNumberFormat('0');
  sheet.getRange(row, 1, 1, VISIBLE_COLUMN_COUNT).setVerticalAlignment('top').setWrap(true);
}

function updateSummary_(ss, rebuildLayout) {
  const formatLayout = rebuildLayout !== false;
  const responseSheet = getOrCreateResponseSheet_(ss, false);
  let sheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SUMMARY_SHEET);
    rebuildLayout = true;
  }

  if (rebuildLayout !== false) {
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    sheet.clear();
    sheet.setConditionalFormatRules([]);
  } else {
    // 回答ごとにシート全体・グラフを作り直さず、値の範囲だけ更新する。
    sheet.getRange('A1:H120').clearContent();
  }

  const total = Math.max(responseSheet.getLastRow() - 1, 0);
  const values = total ? responseSheet.getRange(2, 1, total, HEADERS.length).getValues() : [];
  const titleRange = sheet.getRange('A1:H1');
  if (formatLayout) {
    titleRange.merge().setValue('FOCUS 新規参加者アンケート｜集計ダッシュボード')
      .setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#312e81')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 44);
  } else {
    sheet.getRange('A1').setValue('FOCUS 新規参加者アンケート｜集計ダッシュボード');
  }
  sheet.getRange('A3:B3').setValues([['総回答数', total]]);
  sheet.getRange('D3:E3').setValues([['平均年齢', average_(values.map(r => Number(r[3])).filter(Number.isFinite))]]);
  if (formatLayout) {
    styleKpi_(sheet.getRange('A3:B3'));
    styleKpi_(sheet.getRange('D3:E3'));
    sheet.getRange('E3').setNumberFormat('0.0');
  }

  const ratings = [['行動力', 15], ['継続力', 16], ['自己規律', 17], ['自信', 18], ['生活への満足度', 19]];
  const ratingData = ratings.map(([label, index]) => [label, average_(values.map(r => Number(r[index])).filter(n => Number.isFinite(n) && n >= 1 && n <= 5))]);
  writeTable_(sheet, 6, 1, '現在の自己評価（1〜5点）', ['項目', '平均'], ratingData, 2, formatLayout);
  if (formatLayout) sheet.getRange(8, 2, ratingData.length, 1).setNumberFormat('0.00');

  let row = 15;
  const sections = [];
  row = writeCountSection_(sheet, values, 2, '性別', row, total, false, null, sections, formatLayout);
  row += 2;
  row = writeCountSection_(sheet, values, 4, '職業・立場', row, total, false, null, sections, formatLayout);
  row += 2;
  row = writeCountSection_(sheet, values, 5, 'FOCUSを知ったきっかけ', row, total, false, null, sections, formatLayout);
  row += 2;
  row = writeCountSection_(sheet, values, 7, '現在の悩み・課題', row, total, true, null, sections, formatLayout);
  row += 2;
  row = writeCountSection_(sheet, values, 8, 'FOCUSに入った理由', row, total, true, null, sections, formatLayout);
  row += 2;
  writeCountSection_(sheet, values, 11, '特に変えたい・伸ばしたい分野', row, total, false,
    ['BUSINESS', 'MINDSET', 'BODY', 'APPEARANCE', 'ENGLISH', 'DISCIPLINE'], sections, formatLayout);

  if (formatLayout) {
    sheet.setFrozenRows(1);
    [300, 100, 105, 30, 300, 100, 105, 30].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.getRange('A1:E120').setVerticalAlignment('middle');
    addCharts_(sheet, sections, ratingData.length);
  }
}

function writeCountSection_(sheet, values, columnIndex, title, startRow, total, isMulti, fixedOrder, sections, formatLayout) {
  const counts = {};
  (fixedOrder || []).forEach(label => counts[label] = 0);
  values.forEach(row => {
    const raw = clean_(row[columnIndex]);
    if (!raw) return;
    const items = isMulti ? raw.split(CONFIG.MULTI_SEPARATOR) : [raw];
    const uniqueItems = [...new Set(items.map(normalizeCategory_).filter(Boolean))];
    uniqueItems.forEach(label => counts[label] = (counts[label] || 0) + 1);
  });
  let entries = Object.entries(counts);
  if (!fixedOrder) entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'ja'));
  const rows = entries.map(([label, count]) => [label, count, total ? count / total : 0]);
  writeTable_(sheet, startRow, 1, title,
    ['項目', isMulti ? '選択人数' : '人数', isMulti ? '選択率' : '割合'], rows, 3, formatLayout);
  if (formatLayout && rows.length) sheet.getRange(startRow + 2, 3, rows.length, 1).setNumberFormat('0.0%');
  sections.push({ title: title, headerRow: startRow + 1, dataRows: rows.length });
  return startRow + 2 + Math.max(rows.length, 1);
}

function writeTable_(sheet, row, col, title, headers, rows, width, formatLayout) {
  const titleRange = sheet.getRange(row, col, 1, width);
  const headerRange = sheet.getRange(row + 1, col, 1, headers.length);
  if (formatLayout) {
    titleRange.merge().setValue(title).setFontWeight('bold')
      .setFontColor('#ffffff').setBackground('#4f46e5');
    headerRange.setValues([headers]).setFontWeight('bold')
      .setBackground('#e0e7ff').setFontColor('#1e1b4b');
  } else {
    sheet.getRange(row, col).setValue(title);
    headerRange.setValues([headers]);
  }
  if (rows.length) {
    const dataRange = sheet.getRange(row + 2, col, rows.length, headers.length).setValues(rows);
    if (formatLayout) dataRange.setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    const emptyCell = sheet.getRange(row + 2, col).setValue('回答なし');
    if (formatLayout) emptyCell.setFontColor('#64748b');
  }
}

function styleKpi_(range) {
  range.setBackground('#eef2ff').setFontWeight('bold').setFontSize(14)
    .setBorder(true, true, true, true, false, false, '#818cf8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function addCharts_(sheet, sections, ratingCount) {
  const ratingChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(7, 1, ratingCount + 1, 2)).setPosition(3, 5, 0, 0)
    .setOption('title', '自己評価平均').setOption('legend', { position: 'none' })
    .setOption('vAxis', { viewWindow: { min: 0, max: 5 } }).build();
  sheet.insertChart(ratingChart);
  let chartRow = 21;
  const chartTitles = new Set(['FOCUSを知ったきっかけ', '現在の悩み・課題', 'FOCUSに入った理由', '特に変えたい・伸ばしたい分野']);
  sections.filter(s => chartTitles.has(s.title) && s.dataRows > 0).forEach(section => {
    const chart = sheet.newChart().asBarChart()
      .addRange(sheet.getRange(section.headerRow, 1, section.dataRows + 1, 2))
      .setPosition(chartRow, 5, 0, 0).setOption('title', section.title)
      .setOption('legend', { position: 'none' }).build();
    sheet.insertChart(chart);
    chartRow += 18;
  });
}

function findSubmissionRow_(sheet, submissionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const match = sheet.getRange(2, 22, lastRow - 1, 1).createTextFinder(submissionId)
    .matchEntireCell(true).matchCase(true).findNext();
  return match ? match.getRow() : 0;
}

function normalizeCategory_(value) {
  const text = clean_(value);
  return /^(その他|その他[:：])/.test(text) ? 'その他' : text;
}

function numberOrBlank_(value) {
  const text = clean_(value);
  if (!text) return '';
  const number = Number(text);
  return Number.isFinite(number) ? number : '';
}

function rating_(value) {
  const number = Number(clean_(value));
  return Number.isInteger(number) && number >= 1 && number <= 5 ? number : '';
}

function average_(numbers) {
  if (!numbers.length) return '';
  return Math.round(numbers.reduce((sum, n) => sum + n, 0) / numbers.length * 100) / 100;
}

function clean_(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function output_(data, callback) {
  const json = JSON.stringify(data);
  const safeCallback = /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(clean_(callback)) ? clean_(callback) : '';
  return ContentService.createTextOutput(safeCallback ? safeCallback + '(' + json + ');' : json)
    .setMimeType(safeCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
