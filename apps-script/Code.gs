/**
 * 売上目標トラッカー API（Google Apps Script）
 *
 * スプレッドシートの「売上目標」シートを読み書きして、Web アプリに JSON を返す。
 * セル位置は tools/make_sheet.py で作るシートに合わせてある。
 *
 * 必須: 「プロジェクトの設定 → スクリプト プロパティ」に PASSCODE を追加すること。
 * 閲覧も書き込みもパスコードが合っているときだけ応答する。未設定なら何も返さない。
 * パスコードは URL に載らないよう、すべて POST の本文で受け取る。
 */
const SHEET_NAME = '売上目標';
const FIRST_ROW = 12;
const LAST_ROW = 42;
const ROWS = LAST_ROW - FIRST_ROW + 1;

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

  if (req.action === 'get') return json_(readState_());

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    switch (req.action) {
      case 'setSales':
        setSales_(sh, req.date, req.sales);
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

// ---------- 読み込み ----------

function readState_() {
  const sh = sheet_();
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const s = sh.getRange('B3:B6').getValues();
  const weights = sh.getRange('G4:M4').getValues()[0].map(Number);
  const rows = sh.getRange(FIRST_ROW, 1, ROWS, 4).getValues();
  const formulas = sh.getRange(FIRST_ROW, 3, ROWS, 1).getFormulas();
  const days = [];
  rows.forEach(function (r, i) {
    if (!(r[0] instanceof Date)) return;
    days.push({
      date: Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'),
      weight: Number(r[2]) || 0,
      holiday: !formulas[i][0] && Number(r[2]) === 0, // 重みを手で 0 にした日
      sales: r[3] === '' ? null : Number(r[3]),
    });
  });
  return {
    unitPrice: Number(s[0][0]) || 0,
    targetQty: Number(s[1][0]) || 0,
    month: s[3][0] instanceof Date ? Utilities.formatDate(s[3][0], tz, 'yyyy-MM') : '',
    weights: weights,
    days: days,
    archives: sh.getParent().getSheets().map(function (x) { return x.getName(); })
      .filter(function (n) { return /^\d{4}-\d{2}$/.test(n); }),
    updatedAt: new Date().toISOString(),
  };
}

// ---------- 書き込み ----------

function setSales_(sh, date, sales) {
  const row = rowOf_(sh, date);
  const cell = sh.getRange(row, 4);
  if (sales === null || sales === '' || sales === undefined) {
    cell.clearContent();
  } else {
    const n = Number(sales);
    if (!isFinite(n) || n < 0) throw new Error('売上金額が正しくありません');
    cell.setValue(n);
  }
}

function setHoliday_(sh, date, holiday) {
  const row = rowOf_(sh, date);
  const cell = sh.getRange(row, 3);
  if (holiday) cell.setValue(0);
  else cell.setFormula(weightFormula_(row));
}

function setSettings_(sh, req) {
  if (req.unitPrice !== undefined) sh.getRange('B3').setValue(positive_(req.unitPrice, '単価'));
  if (req.targetQty !== undefined) sh.getRange('B4').setValue(positive_(req.targetQty, '目標本数'));
  if (req.weights !== undefined) {
    if (!Array.isArray(req.weights) || req.weights.length !== 7) throw new Error('曜日の重みが正しくありません');
    sh.getRange('G4:M4').setValues([req.weights.map(function (w) { return positive_(w, '重み', true); })]);
  }
}

/** 今のシートを「yyyy-MM」という名前で値だけ保存し、新しい月を始める。 */
function startMonth_(sh, month) {
  if (!/^\d{4}-\d{2}$/.test(String(month))) throw new Error('月の形式が正しくありません');
  const ss = sh.getParent();
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const current = sh.getRange('B6').getValue();
  if (current instanceof Date) {
    const name = Utilities.formatDate(current, tz, 'yyyy-MM');
    if (!ss.getSheetByName(name)) {
      const copy = sh.copyTo(ss).setName(name);
      const range = copy.getDataRange();
      range.copyTo(range, { contentsOnly: true });
    }
  }
  const parts = month.split('-');
  sh.getRange('B6').setFormula('=DATE(' + Number(parts[0]) + ',' + Number(parts[1]) + ',1)');
  sh.getRange(FIRST_ROW, 4, ROWS, 1).clearContent();
  const f = [];
  for (let r = FIRST_ROW; r <= LAST_ROW; r++) f.push([weightFormula_(r)]);
  sh.getRange(FIRST_ROW, 3, ROWS, 1).setFormulas(f);
}

// ---------- 補助 ----------

function passcode_() {
  return (PropertiesService.getScriptProperties().getProperty('PASSCODE') || '').trim();
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('「' + SHEET_NAME + '」シートが見つかりません');
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

function positive_(v, label, allowZero) {
  const n = Number(v);
  if (!isFinite(n) || n < 0 || (!allowZero && n === 0)) throw new Error(label + 'が正しくありません');
  return n;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
