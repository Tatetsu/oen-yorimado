/**
 * 実績報告書ビュー
 *
 * データシート: 実績報告書
 * ビューシート: 実績報告書ビュー
 *
 * CSV列: 日付(0), 受給者証番号(1), 児童名(2), 入所時間(3), 退所時間(4),
 *        体温(5), 夕食(6), 朝食(7), 昼食(8), 入浴(9), 便(10), 服薬夜(11), 服薬朝(12), 連絡(13)
 */

var SHEET_JISSEKI = "実績報告書";
var VIEW_JISSEKI = "実績報告書ビュー";

function setupJissekiView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_JISSEKI);
  view.clear();
  view.getRange("A1:Z1000").clearDataValidations();

  view.getRange("A1").setValue("児童名：");
  view.getRange("A2").setValue("対象年：");
  view.getRange("A3").setValue("対象月：");
  view.getRange("A1:A3").setFontWeight("bold");
  view.getRange("D1").setValue("児童名を選んで利用実績を表示（年・月は任意）");
  view.getRange("D1").setFontColor("#0000FF");

  var children = getUniqueChildren_(SHEET_JISSEKI);
  if (children.length > 0) {
    view.getRange("B1")
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(children, true).setAllowInvalid(false).build())
      .setValue(children[0]);
  }

  var yearOptions = ["すべて"].concat(getUniqueYears_(SHEET_JISSEKI));
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

  var headers = ["記録日", "入所時間", "退所時間", "体温", "夕食", "朝食", "昼食", "入浴", "便", "服薬(夜)", "服薬(朝)", "その他連絡事項"];
  setTableHeader_(view, 7, headers);
  view.setColumnWidth(1, 110);
  view.setColumnWidth(2, 80);
  view.setColumnWidth(3, 80);
  view.setColumnWidth(4, 50);
  for (var c = 5; c <= 11; c++) view.setColumnWidth(c, 60);
  view.setColumnWidth(12, 300);

  updateJissekiView();
}

function updateJissekiView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_JISSEKI);
  if (!view) return;

  var childName = view.getRange("B1").getDisplayValue().trim();
  var filterYear = view.getRange("B2").getDisplayValue().trim();
  var filterMonth = view.getRange("B3").getDisplayValue().trim();

  clearDataRows_(view);

  if (!childName) {
    view.getRange("B5").setValue("―");
    return;
  }

  var dataSheet = ss.getSheetByName(SHEET_JISSEKI);
  if (!dataSheet) { view.getRange("B5").setValue("―"); return; }

  var allData = dataSheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[2]).trim() === childName && matchYearMonth_(row[0], filterYear, filterMonth)) {
      data.push(row);
    }
  }

  view.getRange("B5").setValue(data.length + "回");

  if (data.length === 0) {
    view.getRange("A8").setValue("該当データなし");
    return;
  }

  var rows = data.map(function(r) {
    return [formatDate_(r[0]), formatTime_(r[3]), formatTime_(r[4]), r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12], r[13]];
  });
  view.getRange(8, 1, rows.length, 12).setValues(rows);
}
