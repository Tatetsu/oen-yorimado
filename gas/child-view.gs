/**
 * F-04: 児童別ビュー更新
 * 選択された児童・年月の確定来館記録を児童別ビューに書き込む
 */

/**
 * 児童別ビューを更新する（ボタン実行）
 */
function updateChildView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    var sheet = getSheet(SHEET_NAMES.CHILD_VIEW);

    // 選択値を取得
    var childName = sheet.getRange('B1').getValue();
    var yearMonthStr = sheet.getRange('B2').getValue();

    if (!childName) {
      Logger.log('児童別ビュー: 児童名が未選択です');
      return;
    }

    ss.toast('児童別ビューを更新中...', '読み込み中', -1);

    var isAllPeriod = (!yearMonthStr || yearMonthStr === 'すべて');
    var yearOnly = parseYearOnly_(String(yearMonthStr).trim());

    if (isAllPeriod) {
      // 全期間表示
      writeChildBasicInfoAll_(sheet, childName);
      writeChildVisitHistoryAll_(sheet, childName);
    } else if (yearOnly !== null) {
      // 年間表示
      writeChildBasicInfoYear_(sheet, childName, yearOnly);
      writeChildVisitHistoryYear_(sheet, childName, yearOnly);
    } else {
      var ym = parseYearMonth(yearMonthStr);
      writeChildBasicInfo_(sheet, childName, ym.year, ym.month);
      writeChildVisitHistory_(sheet, childName, ym.year, ym.month);
    }

    ss.toast('児童別ビューの更新が完了しました', '完了', 3);
    var label = isAllPeriod ? '全期間' : (yearOnly !== null ? yearOnly + '年' : yearMonthStr);
    Logger.log('児童別ビューを更新しました: ' + childName + ' ' + label);
  } catch (error) {
    logError_('updateChildView', error);
    ss.toast('エラー: ' + error.message, 'エラー', 5);
  } finally {
    originalSheet.activate();
  }
}

/**
 * 児童の基本情報を書き込む（特定月）
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
    sheet.getRange('B4').setValue('（児童マスタに未登録）');
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
 * 児童の基本情報を書き込む（全期間）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 */
function writeChildBasicInfoAll_(sheet, childName) {
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
    sheet.getRange('B4').setValue('（児童マスタに未登録）');
    return;
  }

  sheet.getRange('B4').setValue(childRow[MASTER_COL.PARENT_NAME - 1]);
  sheet.getRange('B5').setValue(childRow[MASTER_COL.STAFF - 1]);
  sheet.getRange('B6').setValue(childRow[MASTER_COL.MONTHLY_QUOTA - 1]);
  sheet.getRange('B7').setValue(childRow[MASTER_COL.MEDICAL_TYPE - 1]);

  // 全期間の来館回数を算出
  var allVisits = getAllConfirmedVisits();
  var totalCount = 0;
  allVisits.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) {
      totalCount++;
    }
  });

  sheet.getRange('B8').setValue('全期間合計: ' + totalCount + '回');
}

/**
 * 児童の来館履歴を確定来館記録から書き込む（特定月）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {number} year 年
 * @param {number} month 月
 */
function writeChildVisitHistory_(sheet, childName, year, month) {
  // 既存の来館履歴データをクリア
  clearChildVisitHistory_(sheet);

  // 確定来館記録から該当児童・年月のデータを取得
  var confirmedVisits = getConfirmedVisitsByMonth(year, month);
  var historyData = extractChildHistory_(confirmedVisits, childName);

  if (historyData.length === 0) {
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setValue(year + '年' + month + '月の来館記録はありません');
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setFontColor('#999999');
    return;
  }

  writeHistoryToSheet_(sheet, historyData);
}

/**
 * 児童の来館履歴を確定来館記録から書き込む（全期間）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 */
function writeChildVisitHistoryAll_(sheet, childName) {
  clearChildVisitHistory_(sheet);

  var allVisits = getAllConfirmedVisits();
  var historyData = extractChildHistory_(allVisits, childName);

  if (historyData.length === 0) {
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setValue('来館記録はありません');
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setFontColor('#999999');
    return;
  }

  writeHistoryToSheet_(sheet, historyData);
}

/**
 * 児童の基本情報を書き込む（年間）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {number} year 年
 */
function writeChildBasicInfoYear_(sheet, childName, year) {
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
    sheet.getRange('B4').setValue('（児童マスタに未登録）');
    return;
  }

  sheet.getRange('B4').setValue(childRow[MASTER_COL.PARENT_NAME - 1]);
  sheet.getRange('B5').setValue(childRow[MASTER_COL.STAFF - 1]);
  sheet.getRange('B6').setValue(childRow[MASTER_COL.MONTHLY_QUOTA - 1]);
  sheet.getRange('B7').setValue(childRow[MASTER_COL.MEDICAL_TYPE - 1]);

  // 年間の来館回数を算出
  var yearVisits = getConfirmedVisitsByYear(year);
  var totalCount = 0;
  yearVisits.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) {
      totalCount++;
    }
  });

  sheet.getRange('B8').setValue(year + '年合計: ' + totalCount + '回');
}

/**
 * 児童の来館履歴を確定来館記録から書き込む（年間）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {number} year 年
 */
function writeChildVisitHistoryYear_(sheet, childName, year) {
  clearChildVisitHistory_(sheet);

  var yearVisits = getConfirmedVisitsByYear(year);
  var historyData = extractChildHistory_(yearVisits, childName);

  if (historyData.length === 0) {
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setValue(year + '年の来館記録はありません');
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setFontColor('#999999');
    return;
  }

  writeHistoryToSheet_(sheet, historyData);
}

/**
 * 来館履歴エリアをクリアする
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 */
function clearChildVisitHistory_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= CHILD_VIEW_HISTORY_START_ROW) {
    var range = sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, lastRow - CHILD_VIEW_HISTORY_START_ROW + 1, 11);
    range.clearContent();
    range.clearFormat();
  }
}

/**
 * 確定来館記録から指定児童のデータを抽出する
 * @param {Array<Array>} visits 確定来館記録データ
 * @param {string} childName 児童名
 * @returns {Array<Array>} 抽出された来館履歴
 */
function extractChildHistory_(visits, childName) {
  var historyData = [];
  visits.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) {
      historyData.push([
        row[CONFIRMED_COL.RECORD_DATE - 1],
        row[CONFIRMED_COL.STAFF_NAME - 1],
        row[CONFIRMED_COL.CHECK_IN - 1],
        row[CONFIRMED_COL.CHECK_OUT - 1],
        row[CONFIRMED_COL.TEMPERATURE - 1],
        row[CONFIRMED_COL.MEAL - 1],
        row[CONFIRMED_COL.BATH - 1],
        row[CONFIRMED_COL.SLEEP - 1],
        row[CONFIRMED_COL.BOWEL - 1],
        row[CONFIRMED_COL.MEDICINE - 1],
        row[CONFIRMED_COL.NOTES - 1],
      ]);
    }
  });

  // 日付昇順でソート
  historyData.sort(function(a, b) {
    return new Date(a[0]) - new Date(b[0]);
  });

  return historyData;
}

/**
 * 来館履歴データをシートに書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {Array<Array>} historyData 来館履歴データ
 */
function writeHistoryToSheet_(sheet, historyData) {
  var dataRange = sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, historyData.length, 11);
  dataRange.setValues(historyData);

  // 記録日列の表示形式
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, historyData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所時間・退所時間列の表示形式（列3, 4）
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 3, historyData.length, 2)
    .setNumberFormat('HH:mm');

  // 体温列の表示形式（列5）
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 5, historyData.length, 1)
    .setNumberFormat('0.0');

  // データ行の罫線（印刷向け）
  dataRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // データ行の中央寄せ（連絡事項以外）
  sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, historyData.length, 10)
    .setHorizontalAlignment('center');

  // 偶数行の背景色（ゼブラストライプ）
  for (var i = 0; i < historyData.length; i++) {
    if (i % 2 === 1) {
      sheet.getRange(CHILD_VIEW_HISTORY_START_ROW + i, 1, 1, 11)
        .setBackground('#F8F9FA');
    }
  }

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

/**
 * 児童別ビューを印刷用に整える（操作エリアを非表示にする）
 * メニューから実行する
 */
function prepareChildViewForPrint() {
  var sheet = getSheet(SHEET_NAMES.CHILD_VIEW);

  // 操作エリア（1〜3行目）を非表示
  sheet.hideRows(1, 3);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '印刷プレビューを確認してください。終わったら「印刷モード解除」を実行してください。',
    '印刷モード',
    -1
  );
}

/**
 * 児童別ビューを通常表示に戻す（印刷モード解除）
 * メニューから実行する
 */
function restoreChildViewFromPrint() {
  var sheet = getSheet(SHEET_NAMES.CHILD_VIEW);

  // 非表示にした行を再表示
  sheet.showRows(1, 3);

  SpreadsheetApp.getActiveSpreadsheet().toast('通常表示に戻しました', '完了', 3);
}
