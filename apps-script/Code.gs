/**
 * 売上目標トラッカー API（Google Apps Script）
 *
 * スプレッドシートの「売上目標」シートを読み書きして、Web アプリに JSON を返す。
 * シートのレイアウトはこのファイルの setupSheet() が作る（唯一の定義元）。
 * レイアウトを変えたら、スプシのメニュー「売上目標 → シートを作成・更新」を実行する。
 * 入力済みの売上・本数・設定・休業日は引き継がれる。
 *
 * 必須: 「プロジェクトの設定 → スクリプト プロパティ」に PASSCODE を追加すること。
 * 閲覧も書き込みもパスコードが合っているときだけ応答する。未設定なら何も返さない。
 * パスコードは URL に載らないよう、すべて POST の本文で受け取る。
 */
const SHEET_NAME = '売上目標';
const HEADER_ROW = 11;
const FIRST_ROW = 12;
const LAST_ROW = 42;
const ROWS = LAST_ROW - FIRST_ROW + 1;
const COL = { date: 1, weekday: 2, weight: 3, sales: 4, salesTarget: 5, salesDiff: 6, qty: 7, qtyTarget: 8, qtyDiff: 9 };
const CELL = { goalSales: 'B3', goalQty: 'B4', month: 'B6', weights: 'G4:M4' };

// ======================= Web アプリの窓口 =======================

/** 動作確認用。データは返さない。 */
function doGet() {
  return json_({ ok: true, passcodeConfigured: !!passcode_() });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'bad_request' });
  }
  const passcode = passcode_();
  if (!passcode) return json_({ error: 'no_passcode_configured' });
  if (String(req.passcode || '') !== passcode) return json_({ error: 'passcode' });

  try {
    if (req.action === 'get') return json_(readState_());
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    switch (req.action) {
      case 'setEntry':
        setEntry_(sh, req.date, req.sales, req.qty);
        break;
      case 'setHoliday':
        setHoliday_(sh, req.date, !!req.holiday);
        break;
      case 'setSettings':
        setSettings_(sh, req);
        break;
      case 'startMonth':
        startMonth_(sh, req.month);
        break;
      default:
        return json_({ error: 'unknown_action' });
    }
    SpreadsheetApp.flush();
  } catch (err) {
    return json_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
  return json_(readState_());
}

// ======================= 読み込み =======================

function readState_() {
  const sh = sheet_();
  if (sh.getRange('A3').getValue() !== '月間目標売上') {
    throw new Error('シートが古い形式です。スプレッドシートのメニュー「売上目標 → シートを作成・更新」を実行してください');
  }
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const month = sh.getRange(CELL.month).getValue();
  const rows = sh.getRange(FIRST_ROW, 1, ROWS, COL.qty).getValues();
  const weightFormulas = sh.getRange(FIRST_ROW, COL.weight, ROWS, 1).getFormulas();
  const days = [];
  rows.forEach(function (r, i) {
    if (!(r[COL.date - 1] instanceof Date)) return;
    const w = Number(r[COL.weight - 1]) || 0;
    days.push({
      date: Utilities.formatDate(r[COL.date - 1], tz, 'yyyy-MM-dd'),
      weight: w,
      holiday: !weightFormulas[i][0] && w === 0, // 重みを手で 0 にした日
      sales: numOrNull_(r[COL.sales - 1]),
      qty: numOrNull_(r[COL.qty - 1]),
    });
  });
  return {
    goalSales: Number(sh.getRange(CELL.goalSales).getValue()) || 0,
    goalQty: Number(sh.getRange(CELL.goalQty).getValue()) || 0,
    month: month instanceof Date ? Utilities.formatDate(month, tz, 'yyyy-MM') : '',
    weights: sh.getRange(CELL.weights).getValues()[0].map(Number),
    days: days,
    updatedAt: new Date().toISOString(),
  };
}

// ======================= 書き込み =======================

/** sales / qty は「数値=保存」「null=消す」「undefined=変更しない」。 */
function setEntry_(sh, date, sales, qty) {
  const row = rowOf_(sh, date);
  writeNumber_(sh.getRange(row, COL.sales), sales, '売上金額');
  writeNumber_(sh.getRange(row, COL.qty), qty, '本数');
}

function setHoliday_(sh, date, holiday) {
  const row = rowOf_(sh, date);
  const cell = sh.getRange(row, COL.weight);
  if (holiday) cell.setValue(0);
  else cell.setFormula(weightFormula_(row));
}

function setSettings_(sh, req) {
  if (req.goalSales !== undefined) sh.getRange(CELL.goalSales).setValue(nonNegative_(req.goalSales, '月間目標売上', false));
  if (req.goalQty !== undefined) sh.getRange(CELL.goalQty).setValue(nonNegative_(req.goalQty, '月間目標本数', false));
  if (req.weights !== undefined) {
    if (!Array.isArray(req.weights) || req.weights.length !== 7) throw new Error('曜日の重みが正しくありません');
    sh.getRange(CELL.weights).setValues([req.weights.map(function (w) { return nonNegative_(w, '重み', true); })]);
  }
}

/** 今のシートを「yyyy-MM」という名前で値だけ保存し、新しい月を始める。 */
function startMonth_(sh, month) {
  if (!/^\d{4}-\d{2}$/.test(String(month))) throw new Error('月の形式が正しくありません');
  const ss = sh.getParent();
  const tz = ss.getSpreadsheetTimeZone();
  const current = sh.getRange(CELL.month).getValue();
  if (current instanceof Date) {
    const name = Utilities.formatDate(current, tz, 'yyyy-MM');
    if (!ss.getSheetByName(name)) {
      const copy = sh.copyTo(ss).setName(name);
      const range = copy.getDataRange();
      range.copyTo(range, { contentsOnly: true });
    }
  }
  const p = month.split('-');
  sh.getRange(CELL.month).setFormula('=DATE(' + Number(p[0]) + ',' + Number(p[1]) + ',1)');
  sh.getRange(FIRST_ROW, COL.sales, ROWS, 1).clearContent();
  sh.getRange(FIRST_ROW, COL.qty, ROWS, 1).clearContent();
  sh.getRange(FIRST_ROW, COL.weight, ROWS, 1).setFormulas(range_(function (r) { return [weightFormula_(r)]; }));
}

// ======================= シートの作成・更新 =======================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('売上目標')
    .addItem('シートを作成・更新', 'setupSheet')
    .addToUi();
}

/**
 * 「売上目標」シートを最新のレイアウトで作り直す。
 * 入力済みの値（目標・対象月・曜日の重み・休業日・売上・本数）は引き継ぐ。
 * 旧形式（単価 × 目標本数）のシートからも移行できる。
 */
function setupSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEET_NAME);
    const keep = sh ? snapshot_(sh) : defaults_();
    if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
    sh.clear();
    sh.clearConditionalFormatRules();
    buildLayout_(sh);

    sh.getRange(CELL.goalSales).setValue(keep.goalSales);
    sh.getRange(CELL.goalQty).setValue(keep.goalQty);
    sh.getRange(CELL.month).setFormula('=DATE(' + keep.year + ',' + keep.month + ',1)');
    sh.getRange(CELL.weights).setValues([keep.weights]);
    sh.getRange(FIRST_ROW, COL.sales, ROWS, 1).setValues(keep.rows.map(function (x) { return [x.sales]; }));
    sh.getRange(FIRST_ROW, COL.qty, ROWS, 1).setValues(keep.rows.map(function (x) { return [x.qty]; }));
    keep.rows.forEach(function (x, i) { if (x.holiday) sh.getRange(FIRST_ROW + i, COL.weight).setValue(0); });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  try {
    SpreadsheetApp.getActive().toast('シートを最新の形式にしました（入力済みのデータは引き継いでいます）', '売上目標', 8);
  } catch (e) { /* エディタから実行したときは表示しない */ }
}

function defaults_() {
  const now = new Date();
  const rows = [];
  for (let i = 0; i < ROWS; i++) rows.push({ sales: '', qty: '', holiday: false });
  return { goalSales: 3000000, goalQty: 1000, year: now.getFullYear(), month: now.getMonth() + 1, weights: [1, 1, 1, 1, 1, 1, 1], rows: rows };
}

/** 今のシートから引き継ぐ値を読む（旧形式にも対応）。 */
function snapshot_(sh) {
  const d = defaults_();
  const label = sh.getRange('A3').getValue();
  if (label === '単価') {
    // 旧形式: B3 単価 / B4 目標本数 / B5 月間目標（=単価×本数）
    d.goalSales = Number(sh.getRange('B5').getValue()) || d.goalSales;
  } else if (label === '月間目標売上') {
    d.goalSales = Number(sh.getRange('B3').getValue()) || d.goalSales;
  }
  d.goalQty = Number(sh.getRange('B4').getValue()) || d.goalQty;

  const m = sh.getRange('B6').getValue();
  if (m instanceof Date) { d.year = m.getFullYear(); d.month = m.getMonth() + 1; }

  const w = sh.getRange(CELL.weights).getValues()[0].map(Number);
  if (w.length === 7 && w.every(function (x) { return isFinite(x) && x >= 0; }) && w.some(function (x) { return x > 0; })) d.weights = w;

  const headers = sh.getRange(HEADER_ROW, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const salesCol = headers.indexOf('売上実績') + 1;
  const qtyCol = headers.indexOf('本数実績') + 1;
  const dates = sh.getRange(FIRST_ROW, 1, ROWS, 1).getValues();
  const weightVals = sh.getRange(FIRST_ROW, COL.weight, ROWS, 1).getValues();
  const weightFormulas = sh.getRange(FIRST_ROW, COL.weight, ROWS, 1).getFormulas();
  const sales = salesCol ? sh.getRange(FIRST_ROW, salesCol, ROWS, 1).getValues() : null;
  const qty = qtyCol ? sh.getRange(FIRST_ROW, qtyCol, ROWS, 1).getValues() : null;
  d.rows = d.rows.map(function (_, i) {
    return {
      sales: sales && isNumber_(sales[i][0]) ? sales[i][0] : '',
      qty: qty && isNumber_(qty[i][0]) ? qty[i][0] : '',
      holiday: dates[i][0] instanceof Date && !weightFormulas[i][0] && weightVals[i][0] === 0,
    };
  });
  return d;
}

function buildLayout_(sh) {
  const A = rangeRef_(COL.date), C = rangeRef_(COL.weight), D = rangeRef_(COL.sales), G = rangeRef_(COL.qty);
  const yen = '#,##0', qty = '#,##0.0"本"';
  const input = '#FFF4C2', head = '#DDE7F5', note = '#666666';

  sh.getRange('A1').setValue('売上目標トラッカー').setFontWeight('bold').setFontSize(14);
  sh.getRange('A9').setValue('黄色のセルが入力欄です。Web アプリからも入力できます。').setFontColor(note);

  // 設定
  sh.getRange('A3:A6').setValues([['月間目標売上'], ['月間目標本数'], ['平均単価（目安）'], ['対象月（1日）']]).setFontWeight('bold');
  sh.getRange('B5').setFormula('=IFERROR(B3/B4,0)');
  sh.getRange('B3').setNumberFormat(yen);
  sh.getRange('B4').setNumberFormat('#,##0"本"');
  sh.getRange('B5').setNumberFormat(yen);
  sh.getRange('B6').setNumberFormat('yyyy/mm');
  sh.getRange('B3:B4').setBackground(input);
  sh.getRange('B6').setBackground(input);
  sh.getRange('B3:B6').setBorder(true, true, true, true, true, true, '#BBBBBB', null);

  // サマリー（売上 / 本数）
  sh.getRange('E2:F2').setValues([['売上', '本数']]).setFontWeight('bold').setBackground(head).setHorizontalAlignment('center');
  sh.getRange('D3:D7').setValues([['累計実績'], ['達成率'], ['残り目標'], ['残り営業日'], ['残り1日あたり']]).setFontWeight('bold');
  sh.getRange('E3:F7').setFormulas([
    ['=SUM(' + D + ')', '=SUM(' + G + ')'],
    ['=IFERROR(E3/$B$3,0)', '=IFERROR(F3/$B$4,0)'],
    ['=MAX(0,$B$3-E3)', '=MAX(0,$B$4-F3)'],
    // 「残り営業日」= 今日以降で未入力の日（過去の未入力日は 0 扱い）
    ['=COUNTIFS(' + A + ',">="&TODAY(),' + D + ',"",' + C + ',">0")', '=COUNTIFS(' + A + ',">="&TODAY(),' + G + ',"",' + C + ',">0")'],
    ['=IFERROR(E5/SUMIFS(' + C + ',' + D + ',"",' + A + ',">="&TODAY()),0)', '=IFERROR(F5/SUMIFS(' + C + ',' + G + ',"",' + A + ',">="&TODAY()),0)'],
  ]);
  sh.getRange('E3:E7').setNumberFormat(yen);
  sh.getRange('F3:F7').setNumberFormat(qty);
  sh.getRange('E4:F4').setNumberFormat('0.0%');
  sh.getRange('E6:F6').setNumberFormat('0"日"');
  sh.getRange('E7:F7').setFontWeight('bold').setFontColor('#C0392B');
  sh.getRange('E2:F7').setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  sh.getRange('D8').setValue('※重み1の日の目標').setFontColor(note).setFontSize(9);

  // 曜日別の重み
  sh.getRange('G2').setValue('曜日別の重み（売れる曜日は大きく）').setFontWeight('bold');
  sh.getRange('G3:M3').setValues([['月', '火', '水', '木', '金', '土', '日']]).setBackground(head).setHorizontalAlignment('center');
  sh.getRange(CELL.weights).setBackground(input).setHorizontalAlignment('center');
  sh.getRange('G3:M4').setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  sh.getRange('G5').setValue('※特定の日を休みにするときは、表の「重み」に 0 を入力').setFontColor(note).setFontSize(9);

  // 日別テーブル
  sh.getRange(HEADER_ROW, 1, 1, 9).setValues([['日付', '曜日', '重み', '売上実績', '売上目標', '売上差額', '本数実績', '本数目標', '本数差額']])
    .setFontWeight('bold').setBackground(head).setHorizontalAlignment('center');
  const f = range_(function (r) {
    const a = 'A' + r;
    return [
      r === FIRST_ROW ? '=$B$6' : '=IF(A' + (r - 1) + '="","",IF(MONTH(A' + (r - 1) + '+1)=MONTH($B$6),A' + (r - 1) + '+1,""))',
      '=IF(' + a + '="","",CHOOSE(WEEKDAY(' + a + '),"日","月","火","水","木","金","土"))',
      weightFormula_(r),
    ];
  });
  sh.getRange(FIRST_ROW, 1, ROWS, 3).setFormulas(f);
  sh.getRange(FIRST_ROW, COL.salesTarget, ROWS, 2).setFormulas(range_(function (r) { return targetFormulas_(r, 'D', '$B$3', '$E$5'); }));
  sh.getRange(FIRST_ROW, COL.qtyTarget, ROWS, 2).setFormulas(range_(function (r) { return targetFormulas_(r, 'G', '$B$4', '$F$5'); }));

  const table = sh.getRange(FIRST_ROW, 1, ROWS, 9);
  table.setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  sh.getRange(FIRST_ROW, 1, ROWS, 1).setNumberFormat('m/d');
  sh.getRange(FIRST_ROW, 2, ROWS, 2).setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, COL.sales, ROWS, 3).setNumberFormat(yen);
  sh.getRange(FIRST_ROW, COL.qty, ROWS, 3).setNumberFormat(qty);
  sh.getRange(FIRST_ROW, COL.sales, ROWS, 1).setBackground(input);
  sh.getRange(FIRST_ROW, COL.qty, ROWS, 1).setBackground(input);

  const diffRanges = [sh.getRange(FIRST_ROW, COL.salesDiff, ROWS, 1), sh.getRange(FIRST_ROW, COL.qtyDiff, ROWS, 1)];
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#C0392B').setRanges(diffRanges).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#1E8449').setRanges(diffRanges).build(),
  ]);

  [14, 8, 8, 14, 16, 14, 12, 12, 12, 6, 6, 6, 6].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
  sh.setFrozenRows(HEADER_ROW);
}

/**
 * 目標と差額の数式。
 * 実績入力済みの日・過ぎた日 → その日の朝時点の目標 / 今日以降の未入力日 → 今の残り目標を重みで配分。
 */
function targetFormulas_(r, col, goal, remaining) {
  const A = rangeRef_(COL.date), C = rangeRef_(COL.weight), X = '$' + col + '$' + FIRST_ROW + ':$' + col + '$' + LAST_ROW;
  const a = 'A' + r, c = 'C' + r, x = col + r, t = (col === 'D' ? 'E' : 'H') + r;
  return [
    '=IF(' + a + '="","",IF(N(' + c + ')=0,0,IF(OR(' + x + '<>"",' + a + '<TODAY()),' +
      'MAX(0,' + goal + '-SUMIFS(' + X + ',' + A + ',"<"&' + a + '))*' + c + '/SUMIFS(' + C + ',' + A + ',">="&' + a + '),' +
      'IFERROR(' + remaining + '*' + c + '/SUMIFS(' + C + ',' + X + ',"",' + A + ',">="&TODAY()),0))))',
    '=IF(OR(' + a + '="",' + x + '=""),"",' + x + '-' + t + ')',
  ];
}

// ======================= 補助 =======================

function passcode_() {
  return (PropertiesService.getScriptProperties().getProperty('PASSCODE') || '').trim();
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('「' + SHEET_NAME + '」シートが見つかりません。メニュー「売上目標 → シートを作成・更新」を実行してください');
  return sh;
}

function rowOf_(sh, date) {
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const dates = sh.getRange(FIRST_ROW, 1, ROWS, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    if (d instanceof Date && Utilities.formatDate(d, tz, 'yyyy-MM-dd') === date) return FIRST_ROW + i;
  }
  throw new Error(date + ' は対象月の日付ではありません');
}

function weightFormula_(row) {
  return '=IF(A' + row + '="","",INDEX($G$4:$M$4,WEEKDAY(A' + row + ',2)))';
}

function writeNumber_(cell, v, label) {
  if (v === undefined) return;
  if (v === null || v === '') { cell.clearContent(); return; }
  cell.setValue(nonNegative_(v, label, true));
}

function nonNegative_(v, label, allowZero) {
  const n = Number(v);
  if (!isFinite(n) || n < 0 || (!allowZero && n === 0)) throw new Error(label + 'が正しくありません');
  return n;
}

function numOrNull_(v) { return isNumber_(v) ? Number(v) : null; }
function isNumber_(v) { return v !== '' && v !== null && isFinite(Number(v)) && typeof v !== 'boolean'; }
function rangeRef_(col) { const L = String.fromCharCode(64 + col); return '$' + L + '$' + FIRST_ROW + ':$' + L + '$' + LAST_ROW; }
function range_(fn) { const out = []; for (let r = FIRST_ROW; r <= LAST_ROW; r++) out.push(fn(r)); return out; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
