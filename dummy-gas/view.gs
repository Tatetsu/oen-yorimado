/**
 * ビューシート GAS - 共通エントリポイント + ユーティリティ
 *
 * シート構成:
 *   - 実績報告書         → jisseki-view.gs
 *   - 重度支援加算       → judo-view.gs
 *   - 送迎記録表         → sougei-view.gs
 *   - 献立カレンダー     → kondate-view.gs
 */

// === エントリポイント ===

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ビュー更新")
    .addItem("初期設定", "setupAll")
    .addSeparator()
    .addItem("実績報告書を更新", "updateJissekiView")
    .addItem("重度支援加算を更新", "updateJudoView")
    .addItem("送迎記録表を更新", "updateSougeiView")
    .addItem("献立カレンダーを更新", "updateKondateView")
    .addSeparator()
    .addItem("デバッグ：データ確認", "debugData")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheetName = e.range.getSheet().getName();
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (col !== 2 || row > 2) return;

  if (sheetName === VIEW_JISSEKI) {
    updateJissekiView();
  } else if (sheetName === VIEW_JUDO) {
    updateJudoView();
  } else if (sheetName === VIEW_SOUGEI) {
    updateSougeiView();
  } else if (sheetName === VIEW_KONDATE) {
    updateKondateView();
  }
}

function setupAll() {
  setupJissekiView_();
  setupJudoView_();
  setupSougeiView_();
  setupKondateView_();
  SpreadsheetApp.getUi().alert("初期設定完了");
}

// === 共通ユーティリティ ===

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/** 7行目以降のデータ行のみクリア（ヘッダー・スタイルは温存） */
function clearDataRows_(view) {
  var lastRow = Math.max(view.getLastRow(), 8);
  if (lastRow >= 7) {
    view.getRange(7, 1, lastRow - 6, 25).clearContent();
  }
}

function setTableHeader_(view, row, headers) {
  var range = view.getRange(row, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight("bold").setBackground("#4A86C8").setFontColor("#FFFFFF");
}

/** Date or String → "YYYY-MM" */
function toYm_(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ("0" + (val.getMonth() + 1)).slice(-2);
    return y + "-" + m;
  }
  var s = String(val).trim();
  var match = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (match) {
    return match[1] + "-" + ("0" + parseInt(match[2], 10)).slice(-2);
  }
  return "";
}

/** Date or String → "HH:mm"（1899/12/30問題を回避） */
function formatTime_(val) {
  if (val instanceof Date) {
    var h = ("0" + val.getHours()).slice(-2);
    var m = ("0" + val.getMinutes()).slice(-2);
    return h + ":" + m;
  }
  var s = String(val).trim();
  var match = s.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    return ("0" + match[1]).slice(-2) + ":" + match[2];
  }
  return s;
}

/** Date or String → "YYYY/MM/DD" */
function formatDate_(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ("0" + (val.getMonth() + 1)).slice(-2);
    var d = ("0" + val.getDate()).slice(-2);
    return y + "/" + m + "/" + d;
  }
  return String(val);
}

/** Date or String → "YYYY-MM-DD" */
function normDateKey_(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ("0" + (val.getMonth() + 1)).slice(-2);
    var d = ("0" + val.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
  }
  var s = String(val).trim();
  var match = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (match) {
    return match[1] + "-" + ("0" + parseInt(match[2], 10)).slice(-2) + "-" + ("0" + parseInt(match[3], 10)).slice(-2);
  }
  return "";
}

/** "2025年4月" → "2025-04"、Date対応 */
function parseYmLabel_(label) {
  if (label instanceof Date) {
    var y = label.getFullYear();
    var m = ("0" + (label.getMonth() + 1)).slice(-2);
    return y + "-" + m;
  }
  var m = String(label).match(/(\d{4})年(\d{1,2})月/);
  if (!m) return "";
  return m[1] + "-" + ("0" + parseInt(m[2], 10)).slice(-2);
}

/** データシートからユニーク児童名を取得 */
function getUniqueChildren_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var names = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][2]).trim();
    if (name) names[name] = true;
  }
  return Object.keys(names).sort();
}

/** データシートからユニーク年月ラベルを取得（"2025年4月" 形式） */
function getUniqueYearMonthLabels_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var yms = {};
  for (var i = 1; i < data.length; i++) {
    var ym = toYm_(data[i][0]);
    if (ym && ym.length === 7) yms[ym] = true;
  }

  return Object.keys(yms).sort().map(function(ym) {
    var p = ym.split("-");
    return p[0] + "年" + parseInt(p[1], 10) + "月";
  });
}

/** 児童名 + 年月でフィルタ */
function getFilteredData_(sheetName, childName, yearMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var filtered = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[2]).trim() === childName && toYm_(row[0]) === yearMonth) {
      filtered.push(row);
    }
  }
  return filtered;
}

/** 児童名のみでフィルタ（全期間） */
function getAllFilteredData_(sheetName, childName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var filtered = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[2]).trim() === childName) {
      filtered.push(row);
    }
  }
  return filtered;
}
