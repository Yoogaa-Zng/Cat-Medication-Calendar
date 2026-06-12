// ── 貓咪餵藥記錄 - Google Apps Script ──────────────────────────────────────
// 貼到 Google Sheets 的 Apps Script 編輯器後，部署為「網頁應用程式」
// 執行身分：我，存取權：所有人

function doGet(e) {
  var action   = e.parameter.action;
  var callback = e.parameter.callback; // JSONP 支援（手機瀏覽器用）
  var data;

  try {
    if (action === 'getAll') data = getAllRecordsData();
    else if (action === 'set') data = setRecordData(
      e.parameter.date,
      e.parameter.period,
      e.parameter.value === '1'
    );
    else data = { error: 'Unknown action' };
  } catch (err) {
    data = { error: err.toString() };
  }

  var json = JSON.stringify(data);
  if (callback) {
    // JSONP：包在 callback 函式內，讓手機瀏覽器可跨域讀取
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// 各欄位對應欄號
function colForPeriod(period) {
  if (period === 'am')    return 2;
  if (period === 'pm')    return 3;
  if (period === 'vomit') return 4;
  if (period === 'pant')  return 5;
  return -1;
}

// 取得所有記錄（回傳資料物件）
// 重複列用 OR 合併：任何一列為 true 就算 true
function getAllRecordsData() {
  var sheet = getSheet();
  var rows  = sheet.getDataRange().getValues();
  var out   = {};

  for (var i = 1; i < rows.length; i++) {
    var date = formatDate(rows[i][0]);
    if (!date) continue;
    if (!out[date]) out[date] = { am: false, pm: false, vomit: false, pant: false };
    out[date].am    = out[date].am    || toBool(rows[i][1]);
    out[date].pm    = out[date].pm    || toBool(rows[i][2]);
    out[date].vomit = out[date].vomit || (rows[i].length > 3 ? toBool(rows[i][3]) : false);
    out[date].pant  = out[date].pant  || (rows[i].length > 4 ? toBool(rows[i][4]) : false);
  }
  return out;
}

// 寫入或更新一筆記錄（回傳資料物件）
// 更新所有相同日期的列，防止重複列造成讀寫不一致
function setRecordData(date, period, value) {
  if (!date || !period) return { error: 'Missing params' };
  var col = colForPeriod(period);
  if (col === -1) return { error: 'Unknown period: ' + period };

  var sheet      = getSheet();
  var rows       = sheet.getDataRange().getValues();
  var matchedIdx = [];

  for (var i = 1; i < rows.length; i++) {
    if (formatDate(rows[i][0]) === date) {
      matchedIdx.push(i + 1); // Sheets 從 1 起算
    }
  }

  if (matchedIdx.length === 0) {
    var newRow = [date, false, false, false, false];
    newRow[col - 1] = value;
    sheet.appendRow(newRow);
  } else {
    // 更新所有找到的列（處理重複列）
    for (var j = 0; j < matchedIdx.length; j++) {
      sheet.getRange(matchedIdx[j], col).setValue(value);
    }
  }

  return { ok: true };
}

// ── 工具函式 ──────────────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MedLog');
  if (!sheet) {
    sheet = ss.insertSheet('MedLog');
    sheet.appendRow(['date', 'am', 'pm', 'vomit', 'pant']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 70);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 70);
    sheet.setColumnWidth(5, 70);
  } else {
    // 舊 Sheet 自動補上 vomit / pant 標題欄
    var lastCol = sheet.getLastColumn();
    if (lastCol < 4) sheet.getRange(1, 4).setValue('vomit');
    if (lastCol < 5) sheet.getRange(1, 5).setValue('pant');
  }
  return sheet;
}

// Sheets 會把 YYYY-MM-DD 自動轉成 Date 物件，統一用試算表時區轉回字串
function formatDate(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val || '').trim();
  // 處理舊格式字串（如 "Sat May 22 2026 ..."）
  if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }
  return s;
}

function toBool(v) {
  return v === true || v === 'TRUE' || v === 1 || v === '1';
}
