/**
 * F-04: 児童別ビュー更新
 * 選択された児童・年月の確定来館記録を児童別ビューに書き込む
 */

/**
 * 児童別ビューを更新する（ボタン実行）
 * B1=児童名、B2=対象年、B3=対象月 を参照する
 */
function updateChildView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    var sheet = getSheet(SHEET_NAMES.CHILD_VIEW);

    var childName = sheet.getRange('B1').getValue();
    var yearStr = sheet.getRange('B2').getValue();
    var monthStr = sheet.getRange('B3').getValue();

    if (!childName) {
      Logger.log('児童別ビュー: 児童名が未選択です');
      return;
    }

    ss.toast('児童別ビューを更新中...', '読み込み中', -1);

    var scope = buildScope_(yearStr, monthStr);
    writeChildBasicInfo_(sheet, childName, scope);
    writeChildVisitHistory_(sheet, childName, scope);

    ss.toast('児童別ビューの更新が完了しました', '完了', 3);
    Logger.log('児童別ビューを更新しました: ' + childName + ' ' + describeScope_(scope));
  } catch (error) {
    logError_('updateChildView', error);
    ss.toast('エラー: ' + error.message, 'エラー', 5);
  } finally {
    originalSheet.activate();
  }
}

/**
 * 児童の基本情報を書き込む（期間スコープに応じて集計表示を切替）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {{type: string, year?: number, month?: number}} scope 期間スコープ
 */
function writeChildBasicInfo_(sheet, childName, scope) {
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

  // B8: 期間ごとの集計表示
  sheet.getRange('B8').setValue(buildChildSummaryText_(childRow, childName, scope));
}

/**
 * B8に表示する集計テキストを構築する
 */
function buildChildSummaryText_(childRow, childName, scope) {
  if (scope.type === 'month') {
    var quota = childRow[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
    var visitCount = countChildVisitsInRecords_(getConfirmedVisitsByScope(scope), childName);
    var remaining = quota - visitCount;
    var usageRate = quota > 0 ? Math.round((visitCount / quota) * 100) : 0;
    return visitCount + '回 / 残' + remaining + '枠 / ' + usageRate + '%';
  }

  var total = countChildVisitsInRecords_(getConfirmedVisitsByScope(scope), childName);
  if (scope.type === 'year') return scope.year + '年合計: ' + total + '回';
  if (scope.type === 'month_all_years') return '全期間の' + scope.month + '月合計: ' + total + '回';
  return '全期間合計: ' + total + '回';
}

/**
 * 確定来館記録配列から指定児童のレコード件数を数える
 */
function countChildVisitsInRecords_(records, childName) {
  var count = 0;
  records.forEach(function(row) {
    if (row[CONFIRMED_COL.CHILD_NAME - 1] === childName) count++;
  });
  return count;
}

/**
 * 児童の来館履歴を確定来館記録から書き込む（期間スコープに応じて対象データを切替）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 * @param {string} childName 児童名
 * @param {{type: string, year?: number, month?: number}} scope 期間スコープ
 */
function writeChildVisitHistory_(sheet, childName, scope) {
  clearChildVisitHistory_(sheet);

  var visits = getConfirmedVisitsByScope(scope);
  var emptyMessage = '来館記録はありません';
  if (scope.type === 'month') emptyMessage = scope.year + '年' + scope.month + '月の来館記録はありません';
  else if (scope.type === 'year') emptyMessage = scope.year + '年の来館記録はありません';
  else if (scope.type === 'month_all_years') emptyMessage = '全期間の' + scope.month + '月の来館記録はありません';

  var historyData = extractChildHistory_(visits, childName);

  if (historyData.length === 0) {
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setValue(emptyMessage);
    sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1).setFontColor('#999999');
    return;
  }

  writeHistoryToSheet_(sheet, historyData);
}

/**
 * 来館履歴エリアをクリアする
 * 列幅・フォントサイズは維持し、ストライプ背景・罫線・フォント色のみリセットする
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 児童別ビューシート
 */
function clearChildVisitHistory_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < CHILD_VIEW_HISTORY_START_ROW) return;
  var range = sheet.getRange(CHILD_VIEW_HISTORY_START_ROW, 1, lastRow - CHILD_VIEW_HISTORY_START_ROW + 1, 11);
  range.clearContent();
  range.setBackground(null);
  range.setFontColor(null);
  range.setBorder(false, false, false, false, false, false);
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
