/**
 * ダミーデータ ビューシート GAS
 *
 * シート構成:
 *   - 実績報告書: CSVインポート済みデータ
 *   - 重度支援加算: CSVインポート済みデータ
 *   - ビュー（実績報告書）: 児童名・年月で実績報告書を表示
 *   - ビュー（重度支援加算）: 児童名・年月で重度支援加算を表示
 */

var SHEET_JISSEKI = "実績報告書";
var SHEET_JUDO = "重度支援加算";
var VIEW_JISSEKI = "ビュー（実績報告書）";
var VIEW_JUDO = "ビュー（重度支援加算）";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ダミーデータ")
    .addItem("初期設定", "setupAll")
    .addSeparator()
    .addItem("実績報告書を更新", "updateJissekiView")
    .addItem("重度支援加算を更新", "updateJudoView")
    .addSeparator()
    .addItem("デバッグ：データ確認", "debugData")
    .addToUi();
}

// === onEdit: 年月・児童名変更で自動更新 ===
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
  }
}

// === 初期設定（両ビュー作成） ===
function setupAll() {
  setupJissekiView_();
  setupJudoView_();
  SpreadsheetApp.getUi().alert("初期設定完了");
}

function setupJissekiView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_JISSEKI);
  view.clear();
  view.getRange("A1:Z1000").clearDataValidations();

  setupViewHeader_(view, SHEET_JISSEKI);
  updateJissekiView();
}

function setupJudoView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_JUDO);
  view.clear();
  view.getRange("A1:Z1000").clearDataValidations();

  setupViewHeader_(view, SHEET_JUDO);
  updateJudoView();
}

// === 共通ヘッダー設定 ===
function setupViewHeader_(view, dataSheetName) {
  view.getRange("A1").setValue("児童名：");
  view.getRange("A2").setValue("対象年月：");
  view.getRange("A1:A2").setFontWeight("bold");

  view.getRange("D1").setValue("児童名と年月を選んで利用実績を表示");
  view.getRange("D1").setFontColor("#0000FF");

  // 児童名ドロップダウン
  var children = getUniqueChildren_(dataSheetName);
  if (children.length > 0) {
    view.getRange("B1")
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(children, true).setAllowInvalid(false).build())
      .setValue(children[0]);
  }

  // 年月ドロップダウン（データから取得、"2025年4月" 形式）
  var ymLabels = getUniqueYearMonthLabels_(dataSheetName);
  if (ymLabels.length > 0) {
    view.getRange("B2")
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(ymLabels, true).setAllowInvalid(false).build())
      .setValue(ymLabels[ymLabels.length - 1]);
  }

  view.setColumnWidth(1, 130);
  view.setColumnWidth(2, 130);
}

// === 実績報告書ビュー更新 ===
function updateJissekiView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_JISSEKI);
  if (!view) return;

  var childName = String(view.getRange("B1").getValue()).trim();
  var ymLabel = String(view.getRange("B2").getValue()).trim();
  var yearMonth = parseYmLabel_(ymLabel);

  clearDataArea_(view);

  if (!childName || !yearMonth) {
    view.getRange("A4").setValue("児童名と年月を選択してください");
    return;
  }

  var data = getFilteredData_(SHEET_JISSEKI, childName, yearMonth);

  // サマリー
  view.getRange("A4").setValue("来館回数：");
  view.getRange("B4").setValue(data.length + "回");
  view.getRange("A4").setFontWeight("bold");

  // ヘッダー
  var headers = ["記録日", "体温", "夕食", "朝食", "昼食", "入浴", "便", "服薬(夜)", "服薬(朝)", "その他連絡事項"];
  setTableHeader_(view, 6, headers);

  if (data.length === 0) {
    view.getRange("A7").setValue("該当データなし");
    view.getRange("A8").setValue("デバッグ: 児童名=" + childName + " / 年月=" + yearMonth);
    view.getRange("A8").setFontColor("#999999");
    return;
  }

  // CSV: 日付(0), 受給者証番号(1), 児童名(2), 体温(3), 夕食(4), 朝食(5), 昼食(6), 入浴(7), 便(8), 服薬夜(9), 服薬朝(10), 連絡(11)
  var rows = data.map(function(r) {
    return [formatDate_(r[0]), r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]];
  });
  view.getRange(7, 1, rows.length, headers.length).setValues(rows);
  formatTable_(view, 6, rows.length, headers.length);

  view.setColumnWidth(1, 110);
  view.setColumnWidth(2, 50);
  for (var c = 3; c <= 9; c++) view.setColumnWidth(c, 60);
  view.setColumnWidth(10, 300);
}

// === 重度支援加算ビュー更新 ===
function updateJudoView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_JUDO);
  if (!view) return;

  var childName = String(view.getRange("B1").getValue()).trim();
  var ymLabel = String(view.getRange("B2").getValue()).trim();
  var yearMonth = parseYmLabel_(ymLabel);

  clearDataArea_(view);

  if (!childName || !yearMonth) {
    view.getRange("A4").setValue("児童名と年月を選択してください");
    return;
  }

  var dataSheet = ss.getSheetByName(SHEET_JUDO);
  if (!dataSheet) {
    view.getRange("A4").setValue("「重度支援加算」シートが見つかりません");
    return;
  }

  var allData = dataSheet.getDataRange().getValues();
  if (allData.length < 2) {
    view.getRange("A4").setValue("データなし");
    return;
  }

  var sheetHeaders = allData[0];
  var filtered = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[2]).trim() === childName && toYm_(row[0]) === yearMonth) {
      filtered.push(row);
    }
  }

  // サマリー
  view.getRange("A4").setValue("チェック日数：");
  view.getRange("B4").setValue(filtered.length + "日");
  view.getRange("A4").setFontWeight("bold");

  // ヘッダー（日付 + col3以降）
  var displayHeaders = ["記録日"];
  for (var h = 3; h < sheetHeaders.length; h++) {
    displayHeaders.push(String(sheetHeaders[h]));
  }
  setTableHeader_(view, 6, displayHeaders);

  if (filtered.length === 0) {
    view.getRange("A7").setValue("該当データなし");
    view.getRange("A8").setValue("デバッグ: 児童名=" + childName + " / 年月=" + yearMonth);
    view.getRange("A8").setFontColor("#999999");
    return;
  }

  var rows = filtered.map(function(row) {
    var r = [formatDate_(row[0])];
    for (var c = 3; c < row.length; c++) r.push(row[c]);
    return r;
  });
  view.getRange(7, 1, rows.length, displayHeaders.length).setValues(rows);
  formatTable_(view, 6, rows.length, displayHeaders.length);

  view.setColumnWidth(1, 110);
  for (var c = 2; c <= displayHeaders.length; c++) view.setColumnWidth(c, 50);
}

// === デバッグ：データシートの実際の値を確認 ===
function debugData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var messages = [];

  [SHEET_JISSEKI, SHEET_JUDO].forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      messages.push(sheetName + ": シートなし");
      return;
    }
    var data = sheet.getDataRange().getValues();
    messages.push(sheetName + ": " + (data.length - 1) + "行");

    if (data.length >= 2) {
      var row1 = data[1];
      var dateVal = row1[0];
      var dateType = typeof dateVal;
      if (dateVal instanceof Date) dateType = "Date";
      messages.push("  1行目: 日付=" + dateVal + " (型:" + dateType + ") / 児童名=" + row1[2]);
      messages.push("  toYm結果=" + toYm_(dateVal));

      // ユニーク年月
      var yms = {};
      var names = {};
      for (var i = 1; i < data.length; i++) {
        yms[toYm_(data[i][0])] = true;
        names[String(data[i][2]).trim()] = true;
      }
      messages.push("  年月一覧: " + Object.keys(yms).sort().join(", "));
      messages.push("  児童名一覧: " + Object.keys(names).sort().join(", "));
    }
  });

  SpreadsheetApp.getUi().alert(messages.join("\n"));
}

// === ユーティリティ ===

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function clearDataArea_(view) {
  var lastRow = Math.max(view.getLastRow(), 8);
  view.getRange(4, 1, lastRow - 3, 25).clear();
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
  // "YYYY-MM-DD" or "YYYY/MM/DD" 形式
  var match = s.match(/^(\d{4})[-\/](\d{1,2})/);
  if (match) {
    return match[1] + "-" + ("0" + parseInt(match[2], 10)).slice(-2);
  }
  return "";
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

/** "2025年4月" → "2025-04" */
function parseYmLabel_(label) {
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

function formatTable_(view, headerRow, dataRows, cols) {
  if (dataRows > 0) {
    view.getRange(headerRow + 1, 1, dataRows, cols)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }
  for (var r = 0; r < dataRows; r++) {
    view.getRange(headerRow + 1 + r, 1, 1, cols)
      .setBackground(r % 2 === 0 ? "#F8F9FA" : "#FFFFFF");
  }
  view.getRange(headerRow, 1, dataRows + 1, cols)
    .setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
}
