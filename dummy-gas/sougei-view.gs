/**
 * 送迎記録表ビュー
 *
 * データシート: 送迎記録表
 * ビューシート: 送迎記録表ビュー
 *
 * CSV列: 日付(0), 曜日(1), 送迎区分(2), 送迎職員(3), 介助者(4), 送迎方法(5),
 *        車番(6), 点呼確認による異常(7), 利用者(8), 開始時刻(9), 終了時刻(10), 備考(11)
 */

var SHEET_SOUGEI = "送迎記録表";
var VIEW_SOUGEI = "送迎記録表ビュー";

function setupSougeiView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_SOUGEI);
  view.clear();
  view.getRange("A1:Z1000").clearDataValidations();

  view.getRange("A1").setValue("対象年月：");
  view.getRange("A2").setValue("送迎区分：");
  view.getRange("A1:A2").setFontWeight("bold");
  view.getRange("D1").setValue("年月・送迎区分を選んで送迎記録を表示");
  view.getRange("D1").setFontColor("#0000FF");

  var ymOptions = ["すべて"].concat(getUniqueYearMonthLabels_(SHEET_SOUGEI));
  view.getRange("B1")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(ymOptions, true).setAllowInvalid(true).build())
    .setValue(ymOptions.length > 1 ? ymOptions[ymOptions.length - 1] : "すべて");

  view.getRange("B2")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["すべて", "夕迎え", "朝送り", "昼送り"], true).setAllowInvalid(false).build())
    .setValue("すべて");

  view.getRange("A4").setValue("件数：").setFontWeight("bold");

  var headers = ["日付", "曜日", "送迎区分", "利用者", "開始時刻", "終了時刻", "備考"];
  setTableHeader_(view, 6, headers);
  view.setColumnWidth(1, 110);
  view.setColumnWidth(2, 50);
  view.setColumnWidth(3, 80);
  view.setColumnWidth(4, 500);
  view.setColumnWidth(5, 80);
  view.setColumnWidth(6, 80);
  view.setColumnWidth(7, 200);

  updateSougeiView();
}

function updateSougeiView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_SOUGEI);
  if (!view) return;

  var ymLabel = view.getRange("B1").getDisplayValue().trim();
  var kubun = view.getRange("B2").getDisplayValue().trim();

  clearDataRows_(view);

  var dataSheet = ss.getSheetByName(SHEET_SOUGEI);
  if (!dataSheet) {
    view.getRange("B4").setValue("―");
    return;
  }

  var allData = dataSheet.getDataRange().getValues();
  if (allData.length < 2) {
    view.getRange("B4").setValue("0件");
    return;
  }

  var isAllPeriod = (!ymLabel || ymLabel === "すべて");
  var yearMonth = isAllPeriod ? null : parseYmLabel_(ymLabel);
  if (!isAllPeriod && !yearMonth) isAllPeriod = true;

  var isAllKubun = (!kubun || kubun === "すべて");

  var filtered = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var ymMatch = isAllPeriod || toYm_(row[0]) === yearMonth;
    var kubunMatch = isAllKubun || String(row[2]).trim() === kubun;
    if (ymMatch && kubunMatch) filtered.push(row);
  }

  view.getRange("B4").setValue(filtered.length + "件");

  if (filtered.length === 0) {
    view.getRange("A7").setValue("該当データなし");
    return;
  }

  var rows = filtered.map(function(r) {
    return [
      formatDate_(r[0]),
      r[1],
      r[2],
      r[8],
      r[9] instanceof Date ? formatTime_(r[9]) : String(r[9]),
      r[10] instanceof Date ? formatTime_(r[10]) : String(r[10]),
      r[11]
    ];
  });
  view.getRange(7, 1, rows.length, 7).setValues(rows);
}
