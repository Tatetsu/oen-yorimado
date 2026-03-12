/**
 * F-04: 児童別ビュー更新
 * 選択された児童・年月の確定来館記録を児童別ビューに書き込む
 */

/**
 * 児童別ビューを更新する（ボタン実行）
 */
function updateChildView() {
  var sheet = getSheet(SHEET_NAMES.CHILD_VIEW);

  // 選択値を取得
  var childName = sheet.getRange('B1').getValue();
  var yearMonthStr = sheet.getRange('B2').getValue();

  if (!childName) {
    SpreadsheetApp.getUi().alert('児童名を選択してください');
    return;
  }
  if (!yearMonthStr) {
    SpreadsheetApp.getUi().alert('対象年月を選択してください');
    return;
  }

  var ym = parseYearMonth(yearMonthStr);

  // 基本情報を児童マスタから取得して書き込む
  writeChildBasicInfo_(sheet, childName, ym.year, ym.month);

  // 来館履歴を確定来館記録から取得して書き込む
  writeChildVisitHistory_(sheet, childName, ym.year, ym.month);

  Logger.log('児童別ビューを更新しました: ' + childName + ' ' + yearMonthStr);
}

/**
 * 児童の基本情報を書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {number} year 年
 * @param {number} month 月
 */
function writeChildBasicInfo_(sheet, childName, year, month) {
  // 児童マスタから該当児童を検索
  var masterData = getChildMasterData();
  var childRow = null;

  for (var i = 0; i < masterData.length; i++) {
    if (masterData[i][MASTER_COL.NAME - 1] === childName) {
      childRow = masterData[i];
      break;
    }
  }

  if (!childRow) {
    Logger.log('児童マスタに該当児童が見つかりません: ' + childName);
    SpreadsheetApp.getUi().alert('児童「' + childName + '」が児童マスタに見つかりません');
    return;
  }

  // 基本情報の書き込み
  sheet.getRange('B4').setValue(childRow[MASTER_COL.PARENT_NAME - 1]);
  sheet.getRange('B5').setValue(childRow[MASTER_COL.STAFF - 1]);
  sheet.getRange('B6').setValue(childRow[MASTER_COL.MONTHLY_QUOTA - 1]);
  sheet.getRange('B7').setValue(childRow[MASTER_COL.MEDICAL_TYPE - 1]);

  // 確定来館記録から来館回数を算出
  var quota = childRow[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
  var visitCount = countChildVisitsFromConfirmed_(childName, year, month);
  var remaining = quota - visitCount;
  var usageRate = quota > 0 ? Math.round((visitCount / quota) * 100) : 0;

  sheet.getRange('B8').setValue(visitCount + '回 / 残' + remaining + '枠 / ' + usageRate + '%');
}

/**
 * 児童の来館履歴を確定来館記録から書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {number} year 年
 * @param {number} month 月
 */
function writeChildVisitHistory_(sheet, childName, year, month) {
  // 既存の来館履歴データをクリア
  var lastRow = sheet.getLastRow();
  if (lastRow >= CHILD_VIEW_HISTORY_START_ROW) {
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, lastRow - CHILD_VIEW_HISTORY_START_ROW + 1, 11).clearContent();
  }

  // 確定来館記録から該当児童・年月のデータを取得
  var confirmedVisits = getConfirmedVisitsByMonth(year, month);
  var historyData = [];

  confirmedVisits.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) {
      historyData.push([
        row[CONFIRMED_COL.RECORD_DATE - 1],   // 記録日
        row[CONFIRMED_COL.STAFF_NAME - 1],     // スタッフ名
        row[CONFIRMED_COL.CHECK_IN - 1],       // 入所時間
        row[CONFIRMED_COL.CHECK_OUT - 1],      // 退所時間
        row[CONFIRMED_COL.TEMPERATURE - 1],    // 体温
        row[CONFIRMED_COL.MEAL - 1],           // 食事
        row[CONFIRMED_COL.BATH - 1],           // 入浴
        row[CONFIRMED_COL.SLEEP - 1],          // 睡眠
        row[CONFIRMED_COL.BOWEL - 1],          // 便
        row[CONFIRMED_COL.MEDICINE - 1],       // 服薬
        row[CONFIRMED_COL.NOTES - 1],          // その他連絡事項
      ]);
    }
  });

  if (historyData.length === 0) {
    Logger.log('該当する来館履歴がありません: ' + childName + ' ' + year + '年' + month + '月');
    return;
  }

  // 日付昇順でソート
  historyData.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });

  // 書き込み
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, historyData.length, 11).setValues(historyData);

  // 記録日列の表示形式
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, historyData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所時間・退所時間列の表示形式（列3, 4）
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 3, historyData.length, 2)
    .setNumberFormat('HH:mm');

  // 体温列の表示形式（列5）
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 5, historyData.length, 1)
    .setNumberFormat('0.0');

  Logger.log('来館履歴を書き込みました: ' + historyData.length + '件');
}

/**
 * 確定来館記録から児童の来館回数を算出する
 * @param {string} childName 児童名
 * @param {number} year 年
 * @param {number} month 月
 * @returns {number} 来館回数（実データ + 振り分けの合計）
 */
function countChildVisitsFromConfirmed_(childName, year, month) {
  var confirmedVisits = getConfirmedVisitsByMonth(year, month);
  var count = 0;

  confirmedVisits.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) {
      count++;
    }
  });

  return count;
}
