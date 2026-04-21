/**
 * 重度支援加算ビュー
 *
 * データシート: 重度支援加算
 * ビューシート: 重度支援加算ビュー
 *
 * CSV列: 日付(0), 受給者証番号(1), 児童名(2), 時刻列(3〜)
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
    for (var c = 2; c <= displayHeaders.length; c++) view.setColumnWidth(c, 50);
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

  var numCols = allData[0].length - 3;
  var rows = filtered.map(function(row) {
    var r = [formatDate_(row[0])];
    for (var c = 3; c < row.length; c++) {
      r.push(row[c] instanceof Date ? formatTime_(row[c]) : row[c]);
    }
    return r;
  });
  view.getRange(8, 1, rows.length, numCols + 1).setValues(rows);
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

  var viewSheet = ss.getSheetByName(VIEW_JISSEKI);
  if (viewSheet) {
    var b1val = viewSheet.getRange("B1").getValue();
    var b1disp = viewSheet.getRange("B1").getDisplayValue();
    var b2val = viewSheet.getRange("B2").getValue();
    var b2disp = viewSheet.getRange("B2").getDisplayValue();
    messages.push("\n実績報告書ビュー:");
    messages.push("  B1: getValue=" + b1val + " (型:" + (typeof b1val) + ") / getDisplayValue=" + b1disp);
    messages.push("  B2: getValue=" + b2val + " (型:" + (typeof b2val) + (b2val instanceof Date ? "/Date" : "") + ") / getDisplayValue=" + b2disp);
    messages.push("  parseYmLabel結果=" + parseYmLabel_(String(b2val)));
  }

  SpreadsheetApp.getUi().alert(messages.join("\n"));
}
