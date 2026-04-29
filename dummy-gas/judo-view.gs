/**
 * 重度支援加算ビュー
 *
 * データシート: 重度支援加算
 * ビューシート: 重度支援加算ビュー
 *
 * シート列構成（想定例・ヘッダー名駆動で動的に対応）:
 *   0:記録日 / 1:受給者証番号 / 2:児童名 / 3〜:入所日(or 入所日時)・入所時刻・退所日・退所時刻・各時刻の○×・睡眠
 *
 * 列の型判別はヘッダー文字列で動的に行う:
 *   - "日時" を含む  → 日時として表示（YYYY/MM/DD HH:mm）
 *   - "入所日"/"退所日"/"日付" → 日付として表示（YYYY/MM/DD）
 *   - "時刻"/"時間"を含む       → 時刻として表示（HH:mm）
 *   - それ以外（17:00 等の Date ヘッダー含む） → 値そのまま（○/×等）
 */

var SHEET_JUDO = "重度支援加算";
var VIEW_JUDO = "重度支援加算ビュー";

function setupJudoView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_JUDO);
  view.clear();
  view.getRange("A1:Z1000").clearDataValidations();

  view.getRange("A1").setValue("児童名：");
  view.getRange("A2").setValue("対象年：");
  view.getRange("A3").setValue("対象月：");
  view.getRange("A1:A3").setFontWeight("bold");
  view.getRange("D1").setValue("児童名を選んで利用実績を表示（年・月は任意）");
  view.getRange("D1").setFontColor("#0000FF");

  var children = getUniqueChildren_(SHEET_JUDO);
  if (children.length > 0) {
    view.getRange("B1")
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(children, true).setAllowInvalid(false).build())
      .setValue(children[0]);
  }

  var yearOptions = ["すべて"].concat(getUniqueYears_(SHEET_JUDO));
  view.getRange("B2")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(yearOptions, true).setAllowInvalid(true).build())
    .setValue("すべて");

  var monthOptions = ["すべて", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
  view.getRange("B3")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(monthOptions, true).setAllowInvalid(true).build())
    .setValue("すべて");

  view.getRange("A5").setValue("来館回数：").setFontWeight("bold");

  var dataSheet = ss.getSheetByName(SHEET_JUDO);
  if (dataSheet) {
    var sheetHeaders = dataSheet.getDataRange().getValues()[0];
    var displayHeaders = ["記録日"];
    for (var h = 3; h < sheetHeaders.length; h++) {
      var hVal = sheetHeaders[h];
      displayHeaders.push(hVal instanceof Date ? formatTime_(hVal) : String(hVal));
    }
    setTableHeader_(view, 7, displayHeaders);
    view.setColumnWidth(1, 110);
    for (var c = 2; c <= displayHeaders.length; c++) {
      var headerVal = sheetHeaders[c + 1];
      var hStr = (headerVal instanceof Date) ? "" : String(headerVal);
      if (hStr.indexOf("日時") >= 0) view.setColumnWidth(c, 130);
      else if (hStr === "入所日" || hStr === "退所日" || hStr.indexOf("日付") >= 0) view.setColumnWidth(c, 100);
      else if (hStr === "その他連絡事項" || hStr.indexOf("備考") >= 0) view.setColumnWidth(c, 200);
      else view.setColumnWidth(c, 50);
    }
  }

  updateJudoView();
}

function updateJudoView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_JUDO);
  if (!view) return;

  var childName = view.getRange("B1").getDisplayValue().trim();
  var filterYear = view.getRange("B2").getDisplayValue().trim();
  var filterMonth = view.getRange("B3").getDisplayValue().trim();

  clearDataRows_(view);

  if (!childName) {
    view.getRange("B5").setValue("―");
    return;
  }

  var dataSheet = ss.getSheetByName(SHEET_JUDO);
  if (!dataSheet) return;

  var allData = dataSheet.getDataRange().getValues();
  if (allData.length < 2) {
    view.getRange("B5").setValue("0日");
    return;
  }

  var sheetHeaders = allData[0];

  var filtered = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[2]).trim() === childName && matchYearMonth_(row[0], filterYear, filterMonth)) {
      filtered.push(row);
    }
  }

  view.getRange("B5").setValue(filtered.length + "日");

  if (filtered.length === 0) {
    view.getRange("A8").setValue("該当データなし");
    return;
  }

  var numCols = sheetHeaders.length - 3;
  var rows = filtered.map(function(row) {
    var r = [formatDate_(row[0])];
    for (var c = 3; c < row.length; c++) {
      r.push(formatJudoCell_(sheetHeaders[c], row[c]));
    }
    return r;
  });
  view.getRange(8, 1, rows.length, numCols + 1).setValues(rows);
}

/** ヘッダーの文字列に応じてセル値の表示形式を決定 */
function formatJudoCell_(headerVal, rowVal) {
  if (rowVal === null || rowVal === undefined || rowVal === "") return "";

  var headerStr = (headerVal instanceof Date) ? "" : String(headerVal);

  if (headerStr.indexOf("日時") >= 0) {
    return rowVal instanceof Date ? formatDateTime_(rowVal) : String(rowVal);
  }
  if (headerStr === "入所日" || headerStr === "退所日" || headerStr.indexOf("日付") >= 0) {
    return rowVal instanceof Date ? formatDate_(rowVal) : String(rowVal);
  }
  if (headerStr.indexOf("時刻") >= 0 || headerStr.indexOf("時間") >= 0) {
    return rowVal instanceof Date ? formatTime_(rowVal) : String(rowVal);
  }
  // 17:00 等の Date ヘッダー or 自由テキスト列
  if (rowVal instanceof Date) {
    return formatTime_(rowVal);
  }
  return rowVal;
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
      messages.push("  ヘッダー: " + data[0].join(" | "));
      messages.push("  1行目: " + row1.map(function(v) {
        return v instanceof Date ? "Date(" + v.toISOString() + ")" : String(v);
      }).join(" | "));
    }
  });

  SpreadsheetApp.getUi().alert(messages.join("\n"));
}
