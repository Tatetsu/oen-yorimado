/**
 * 実績報告書ビュー
 *
 * データシート: 実績報告書（フォーム回答）
 * ビューシート: 実績報告書ビュー
 *
 * シート列構成（0-indexed）:
 *   0:タイムスタンプ / 1:記録日 / 2:児童名 / 3:スタッフ1 / 4:スタッフ2
 *   5:入所日時 / 6:退所日時 / 7:体温 / 8:夕食 / 9:朝食 / 10:昼食
 *   11:入浴 / 12:睡眠 / 13:便 / 14:服薬(夜) / 15:服薬(朝) / 16:その他連絡事項
 *
 * 表示ロジック:
 *   1宿泊（入所日〜退所日）を本番ロジック（confirmed-visits.gs）と同じく日付展開し、
 *   入所日のみ→往=1, 退所日のみ→復=1, 中日（連泊2日目以降の中日）→両方1, 単日入退→両方1
 *   来館回数は展開後の日付件数。
 */

var SHEET_JISSEKI = "実績報告書";
var VIEW_JISSEKI = "実績報告書ビュー";

var JISSEKI_COL = {
  TIMESTAMP: 0,
  RECORD_DATE: 1,
  CHILD_NAME: 2,
  STAFF1: 3,
  STAFF2: 4,
  CHECK_IN: 5,
  CHECK_OUT: 6,
  TEMP: 7,
  DINNER: 8,
  BREAKFAST: 9,
  LUNCH: 10,
  BATH: 11,
  SLEEP: 12,
  BOWEL: 13,
  MED_NIGHT: 14,
  MED_MORNING: 15,
  NOTES: 16
};

var JISSEKI_VIEW_HEADERS = [
  "記録日", "スタッフ名", "入所時間", "退所時間", "往", "復",
  "体温", "夕食", "朝食", "昼食", "入浴", "睡眠", "便",
  "服薬(夜)", "服薬(朝)", "その他連絡事項"
];

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

  var children = getJissekiUniqueChildren_();
  if (children.length > 0) {
    view.getRange("B1")
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(children, true).setAllowInvalid(false).build())
      .setValue(children[0]);
  }

  var yearOptions = ["すべて"].concat(getJissekiUniqueYears_());
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

  setTableHeader_(view, 7, JISSEKI_VIEW_HEADERS);
  view.setColumnWidth(1, 110);  // 記録日
  view.setColumnWidth(2, 100);  // スタッフ名
  view.setColumnWidth(3, 70);   // 入所時間
  view.setColumnWidth(4, 70);   // 退所時間
  view.setColumnWidth(5, 40);   // 往
  view.setColumnWidth(6, 40);   // 復
  view.setColumnWidth(7, 50);   // 体温
  for (var c = 8; c <= 15; c++) view.setColumnWidth(c, 60);
  view.setColumnWidth(16, 300); // その他連絡事項

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
  var expanded = [];

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[JISSEKI_COL.CHILD_NAME]).trim() !== childName) continue;

    var checkIn = row[JISSEKI_COL.CHECK_IN];
    var checkOut = row[JISSEKI_COL.CHECK_OUT];
    if (!(checkIn instanceof Date) || !(checkOut instanceof Date)) continue;
    if (checkOut.getTime() < checkIn.getTime()) continue;

    var stayDates = expandStayDates_(checkIn, checkOut);
    var checkInKey = normDateKey_(checkIn);
    var checkOutKey = normDateKey_(checkOut);

    var staff1 = String(row[JISSEKI_COL.STAFF1] || "").trim();
    var staff2 = String(row[JISSEKI_COL.STAFF2] || "").trim();
    var staffName = staff1;
    if (staff2 && staff2 !== staff1) {
      staffName = staff1 ? (staff1 + " / " + staff2) : staff2;
    }

    var checkInTime = formatTime_(checkIn);
    var checkOutTime = formatTime_(checkOut);

    for (var j = 0; j < stayDates.length; j++) {
      var d = stayDates[j];
      if (!matchYearMonthByDate_(d, filterYear, filterMonth)) continue;

      var key = normDateKey_(d);
      var isInDay = (key === checkInKey);
      var isOutDay = (key === checkOutKey);
      var pickupOutbound, pickupReturn;
      if (!isInDay && !isOutDay) {
        pickupOutbound = 1;
        pickupReturn = 1;
      } else {
        pickupOutbound = isInDay ? 1 : "";
        pickupReturn = isOutDay ? 1 : "";
      }

      expanded.push([
        formatDate_(d),
        staffName,
        checkInTime,
        checkOutTime,
        pickupOutbound,
        pickupReturn,
        row[JISSEKI_COL.TEMP],
        row[JISSEKI_COL.DINNER],
        row[JISSEKI_COL.BREAKFAST],
        row[JISSEKI_COL.LUNCH],
        row[JISSEKI_COL.BATH],
        row[JISSEKI_COL.SLEEP],
        row[JISSEKI_COL.BOWEL],
        row[JISSEKI_COL.MED_NIGHT],
        row[JISSEKI_COL.MED_MORNING],
        row[JISSEKI_COL.NOTES]
      ]);
    }
  }

  expanded.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0]));
  });

  view.getRange("B5").setValue(expanded.length + "回");

  if (expanded.length === 0) {
    view.getRange("A8").setValue("該当データなし");
    return;
  }

  view.getRange(8, 1, expanded.length, JISSEKI_VIEW_HEADERS.length).setValues(expanded);
}

/** 実績報告書シートからユニーク児童名を取得 */
function getJissekiUniqueChildren_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_JISSEKI);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var names = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][JISSEKI_COL.CHILD_NAME]).trim();
    if (name) names[name] = true;
  }
  return Object.keys(names).sort();
}

/** 入所日〜退所日に出現する全年から年リストを生成 */
function getJissekiUniqueYears_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_JISSEKI);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var years = {};
  for (var i = 1; i < data.length; i++) {
    var ci = data[i][JISSEKI_COL.CHECK_IN];
    var co = data[i][JISSEKI_COL.CHECK_OUT];
    if (ci instanceof Date) years[ci.getFullYear()] = true;
    if (co instanceof Date) years[co.getFullYear()] = true;
  }
  return Object.keys(years).sort().reverse();
}
