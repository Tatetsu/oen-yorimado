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
var KONDATE_HEADER_ROW = 3;
var KONDATE_DATA_START_ROW = 4;

function setupKondateView_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = getOrCreateSheet_(ss, VIEW_KONDATE);
  view.clear();
  view.getRange("A1:H200").clearDataValidations();

  view.getRange("A1").setValue("対象年月：").setFontWeight("bold");
  view.getRange("D1").setValue("年月を選んで献立カレンダーを表示（「すべて」で一覧表示）");
  view.getRange("D1").setFontColor("#0000FF");

  var ymLabels = getUniqueYearMonthLabels_(SHEET_KONDATE);
  var ymOptions = ["すべて"].concat(ymLabels);
  view.getRange("B1")
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(ymOptions, true).setAllowInvalid(true).build())
    .setValue(ymLabels.length > 0 ? ymLabels[ymLabels.length - 1] : "すべて");

  updateKondateView();
}

function updateKondateView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var view = ss.getSheetByName(VIEW_KONDATE);
  if (!view) return;

  var ymLabel = view.getRange("B1").getDisplayValue().trim();

  // 3行目以降をクリア
  var lastRow = Math.max(view.getLastRow(), KONDATE_DATA_START_ROW + 50);
  view.getRange(KONDATE_HEADER_ROW, 1, lastRow - KONDATE_HEADER_ROW + 1, 8).clear();

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

  var isAllPeriod = (!ymLabel || ymLabel === "すべて");

  if (isAllPeriod) {
    showKondateFlatTable_(view, allData);
    return;
  }

  var yearMonth = parseYmLabel_(ymLabel);
  if (!yearMonth) {
    view.getRange(KONDATE_DATA_START_ROW, 1).setValue("年月の解析に失敗しました");
    return;
  }

  // 該当月のデータを日付キー → 行で管理
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

  // 曜日ヘッダー
  view.getRange(KONDATE_HEADER_ROW, 1, 1, 7)
    .setValues([KONDATE_DAY_LABELS])
    .setFontWeight("bold")
    .setBackground("#4A86C8")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  // 列幅（月〜日 = A〜G）
  for (var c = 1; c <= 7; c++) view.setColumnWidth(c, 130);

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
      rowDate[col] = dayInMonth;
      var entry = dateMap[key];
      if (entry) {
        rowMain[col] = entry[2];
        rowVeg[col] = entry[3];
        rowDessert[col] = entry[4] + (entry[5] ? " / " + entry[5] : "");
        rowNutri[col] = entry[6];
      }
    }
    startCol = 0;

    view.getRange(currentRow, 1, 1, 7)
      .setValues([rowDate])
      .setFontWeight("bold")
      .setBackground("#E8F0FE")
      .setHorizontalAlignment("center");
    view.getRange(currentRow + 1, 1, 1, 7).setValues([rowMain]).setBackground("#FFFFFF").setFontWeight("normal").setHorizontalAlignment("left");
    view.getRange(currentRow + 2, 1, 1, 7).setValues([rowVeg]).setBackground("#FFFFFF").setHorizontalAlignment("left");
    view.getRange(currentRow + 3, 1, 1, 7).setValues([rowDessert]).setBackground("#FFFFFF").setHorizontalAlignment("left");
    view.getRange(currentRow + 4, 1, 1, 7).setValues([rowNutri]).setBackground("#F5F5F5").setFontColor("#888888").setHorizontalAlignment("left");
    view.getRange(currentRow + 5, 1, 1, 7).setBackground("#E0E0E0");

    for (var r = currentRow; r < currentRow + KONDATE_ROWS_PER_WEEK; r++) {
      view.setRowHeight(r, 20);
    }

    currentRow += KONDATE_ROWS_PER_WEEK;
  }
}

function showKondateFlatTable_(view, allData) {
  var headers = ["日付", "曜日", "メイン", "野菜", "デザート", "飲み物", "栄養士"];
  setTableHeader_(view, KONDATE_HEADER_ROW, headers);

  view.setColumnWidth(1, 110);
  view.setColumnWidth(2, 50);
  view.setColumnWidth(3, 250);
  view.setColumnWidth(4, 200);
  view.setColumnWidth(5, 150);
  view.setColumnWidth(6, 80);
  view.setColumnWidth(7, 100);

  var rows = [];
  for (var i = 1; i < allData.length; i++) {
    var r = allData[i];
    rows.push([
      r[0] instanceof Date ? formatDate_(r[0]) : String(r[0]),
      r[1], r[2], r[3], r[4], r[5], r[6]
    ]);
  }
  if (rows.length > 0) {
    view.getRange(KONDATE_DATA_START_ROW, 1, rows.length, 7).setValues(rows);
  }
}
