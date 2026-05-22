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

// 取得所有記錄（回傳資料物件）
function getAllRecordsData() {
  var sheet = getSheet();
  var rows  = sheet.getDataRange().getValues();
  var out   = {};

  for (var i = 1; i < rows.length; i++) {
    var date = String(rows[i][0] || '').trim();
    if (!date) continue;
    out[date] = {
      am: toBool(rows[i][1]),
      pm: toBool(rows[i][2])
    };
  }
  return out;
}

// 寫入或更新一筆記錄（回傳資料物件）
function setRecordData(date, period, value) {
  if (!date || !period) return { error: 'Missing params' };

  var sheet  = getSheet();
  var rows   = sheet.getDataRange().getValues();
  var rowIdx = -1;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === date) {
      rowIdx = i + 1; // Sheets 從 1 起算
      break;
    }
  }

  if (rowIdx === -1) {
    var newRow = [date, false, false];
    if (period === 'am') newRow[1] = value;
    else                 newRow[2] = value;
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIdx, period === 'am' ? 2 : 3).setValue(value);
  }

  return { ok: true };
}

// ── 工具函式 ──────────────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MedLog');
  if (!sheet) {
    sheet = ss.insertSheet('MedLog');
    sheet.appendRow(['date', 'am', 'pm']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 80);
  }
  return sheet;
}

function toBool(v) {
  return v === true || v === 'TRUE' || v === 1 || v === '1';
}
