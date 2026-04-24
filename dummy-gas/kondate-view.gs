/**
 * 献立カレンダービュー
 *
 * データシート: 献立カレンダー
 * ビューシート: 献立カレンダービュー
 *
 * CSV列: 日付(0), 曜日(1), メイン(2), 野菜(3), デザート(4), 飲み物(5), 栄養士(6)
 *
 * 年月選択時: 月間カレンダーグリッド（列=月〜日）
 *   1週 = 6行: 日付 / メイン / 野菜 / デザート・飲み物 / 栄養士 / セパレーター
 * 「すべて」選択時: 全件フラットテーブル
 */

var SHEET_KONDATE = "献立カレンダー";
var VIEW_KONDATE = "献立カレンダービュー";

var KONDATE_DAY_LABELS = ["月", "火", "水", "木", "金", "土", "日"];
var KONDATE_ROWS_PER_WEEK = 6;
var KONDATE_HEADER_ROW = 1;
var KONDATE_DATA_START_ROW = 2;

function setupKondateView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_KONDATE);
  view.clear();
  view.getRange("A1:H200").clearDataValidations();

  view.getRange("A1").setValue("年").setFontWeight("bold");
  view.getRange("A2").setValue("月").setFontWeight("bold");

  var now = new Date();
  var currentYear = String(now.getFullYear());
  var currentMonth = String(now.getMonth() + 1);

  var yearOptions = getUniqueYears_(SHEET_KONDATE);
  if (yearOptions.length === 0) yearOptions = [currentYear];
  view.getRange("B1")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(yearOptions, true).setAllowInvalid(true).build())
    .setValue(currentYear);

  var monthOptions = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
  view.getRange("B2")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(monthOptions, true).setAllowInvalid(true).build())
    .setValue(currentMonth);

  // 列幅の初期値（フラットテーブル用・D列スタート）
  view.setColumnWidth(4, 110);
  view.setColumnWidth(5, 50);
  view.setColumnWidth(6, 250);
  view.setColumnWidth(7, 200);
  view.setColumnWidth(8, 150);
  view.setColumnWidth(9, 80);
  view.setColumnWidth(10, 100);

  updateKondateView();
}

function updateKondateView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_KONDATE);
  if (!view) return;

  var filterYear = view.getRange("B1").getDisplayValue().trim();
  var filterMonth = view.getRange("B2").getDisplayValue().trim();

  var clearRows = Math.max(view.getLastRow() + 10, 500);
  // A〜C列（フィルタ行）は触れず、D列以降のみクリア
  view.getRange(1, 4, clearRows, 7).clearContent();

  var dataSheet = ss.getSheetByName(SHEET_KONDATE);
  if (!dataSheet) {
    view.getRange(KONDATE_DATA_START_ROW, 1).setValue("「" + SHEET_KONDATE + "」シートが見つかりません");
    return;
  }

  var allData = dataSheet.getDataRange().getValues();
  if (allData.length < 2) {
    view.getRange(KONDATE_DATA_START_ROW, 1).setValue("データがありません");
    return;
  }

  if (!filterYear || !filterMonth) return;

  var yearMonth = filterYear + "-" + ("0" + parseInt(filterMonth, 10)).slice(-2);
  var dateMap = {};
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (toYm_(row[0]) !== yearMonth) continue;
    var key = normDateKey_(row[0]);
    if (key) dateMap[key] = row;
  }
  drawKondateCalendar_(view, yearMonth, dateMap);
}

function drawKondateCalendar_(view, yearMonth, dateMap) {
  var parts = yearMonth.split("-");
  var year = parseInt(parts[0]);
  var month = parseInt(parts[1]);
  var daysInMonth = new Date(year, month, 0).getDate();

  var CAL_START_COL = 4; // D列スタート（A〜C列を汚染しない）

  // 曜日ヘッダー（D列から） 月〜金=無色 / 土=青系 / 日=赤系
  var headerBgs = [["#E8E8E8", "#E8E8E8", "#E8E8E8", "#E8E8E8", "#E8E8E8", "#5B9BD5", "#C94C4C"]];
  var headerFontColors = [["#333333", "#333333", "#333333", "#333333", "#333333", "#FFFFFF", "#FFFFFF"]];
  view.getRange(KONDATE_HEADER_ROW, CAL_START_COL, 1, 7)
    .setValues([KONDATE_DAY_LABELS])
    .setFontWeight("bold")
    .setBackgrounds(headerBgs)
    .setFontColors(headerFontColors)
    .setHorizontalAlignment("center");

  // カレンダー列幅（D〜J列）
  for (var c = CAL_START_COL; c <= CAL_START_COL + 6; c++) view.setColumnWidth(c, 130);

  var firstDow = new Date(year, month - 1, 1).getDay(); // 0=Sun
  var startCol = (firstDow === 0) ? 6 : firstDow - 1;   // 0=月, 6=日

  var currentRow = KONDATE_DATA_START_ROW;
  var dayInMonth = 1;

  while (dayInMonth <= daysInMonth) {
    var rowDate = new Array(7).fill("");
    var rowMain = new Array(7).fill("");
    var rowVeg = new Array(7).fill("");
    var rowDessert = new Array(7).fill("");
    var rowNutri = new Array(7).fill("");

    for (var col = startCol; col <= 6 && dayInMonth <= daysInMonth; col++, dayInMonth++) {
      var key = year + "-" + ("0" + month).slice(-2) + "-" + ("0" + dayInMonth).slice(-2);
      rowDate[col] = String(dayInMonth);
      var entry = dateMap[key];
      if (entry) {
        rowMain[col] = entry[2];
        rowVeg[col] = entry[3];
        rowDessert[col] = entry[4] + (entry[5] ? " / " + entry[5] : "");
        rowNutri[col] = entry[6];
      }
    }
    startCol = 0;

    // 日付行: 月〜金=無色 / 土=淡い青 / 日=淡い赤
    var dateBgs = [["#F0F0F0", "#F0F0F0", "#F0F0F0", "#F0F0F0", "#F0F0F0", "#D6E6F5", "#FCE4E4"]];
    var dateColors = [["#000000", "#000000", "#000000", "#000000", "#000000", "#1F4E79", "#B22222"]];
    view.getRange(currentRow, CAL_START_COL, 1, 7)
      .setNumberFormat("@")
      .setValues([rowDate])
      .setFontWeight("bold")
      .setBackgrounds(dateBgs)
      .setFontColors(dateColors)
      .setHorizontalAlignment("center");
    view.getRange(currentRow + 1, CAL_START_COL, 1, 7).setValues([rowMain]).setBackground("#FFFFFF").setFontWeight("normal").setHorizontalAlignment("left");
    view.getRange(currentRow + 2, CAL_START_COL, 1, 7).setValues([rowVeg]).setBackground("#FFFFFF").setHorizontalAlignment("left");
    view.getRange(currentRow + 3, CAL_START_COL, 1, 7).setValues([rowDessert]).setBackground("#FFFFFF").setHorizontalAlignment("left");
    view.getRange(currentRow + 4, CAL_START_COL, 1, 7).setValues([rowNutri]).setBackground("#F5F5F5").setFontColor("#888888").setHorizontalAlignment("left");
    view.getRange(currentRow + 5, CAL_START_COL, 1, 7).setBackground("#E0E0E0");

    for (var r = currentRow; r < currentRow + KONDATE_ROWS_PER_WEEK; r++) {
      view.setRowHeight(r, 20);
    }

    currentRow += KONDATE_ROWS_PER_WEEK;
  }
}

function showKondateFlatTable_(view, allData) {
  var FLAT_START_COL = 4; // D列スタート
  var headers = ["日付", "曜日", "メイン", "野菜", "デザート", "飲み物", "栄養士"];

  view.getRange(KONDATE_HEADER_ROW, FLAT_START_COL, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4A86C8")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  var rows = [];
  for (var i = 1; i < allData.length; i++) {
    var r = allData[i];
    rows.push([
      r[0] instanceof Date ? formatDate_(r[0]) : String(r[0]),
      r[1], r[2], r[3], r[4], r[5], r[6]
    ]);
  }
  if (rows.length > 0) {
    view.getRange(KONDATE_DATA_START_ROW, FLAT_START_COL, rows.length, 7)
      .setValues(rows)
      .setHorizontalAlignment("left");
  }
}
