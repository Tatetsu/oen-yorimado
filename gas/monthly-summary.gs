/**
 * F-02: 月別集計更新
 * フォームの回答（実記録）から集計し、月別集計シートに値を書き込む
 */

/**
 * 月別集計を更新する（メイン処理）
 * 月別集計シートのB1セル（対象年月）を参照して集計する
 */
function updateMonthlySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    var sheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);

    // 対象年月を取得
    var yearMonthStr = sheet.getRange('B1').getValue();
    if (!yearMonthStr) {
      Logger.log('対象年月が選択されていません');
      return;
    }

    var ym = parseYearMonth(yearMonthStr);

    ss.toast('月別集計を更新中...', '読み込み中', -1);

    // 児童マスタ取得
    var masterData = getChildMasterData();

    // フォームの回答（実記録）から該当月データ取得して集計
    var formResponses = getFormResponsesByMonth(ym.year, ym.month);
    var visitCounts = countVisitsByFormResponses_(formResponses);

    // データエリアをクリア（ヘッダーは残す）
    var lastRow = sheet.getLastRow();
    if (lastRow >= SUMMARY_DATA_START_ROW) {
      var lastCol = Math.max(sheet.getLastColumn(), 6);
      sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, lastCol).clearContent();
      sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, lastCol).clearFormat();
    }

    // ヘッダーを月別に復元
    var headers = ['No.', '児童名', '月間利用枠', '来館数', '残数', '利用率'];
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    var headerRange = sheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

    // 集計データを書き込み
    if (masterData.length === 0) {
      Logger.log('表示対象の児童がいません');
      return;
    }

    var outputData = masterData.map(function(row) {
      var childName = row[MASTER_COL.NAME - 1];
      var quota = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
      var visits = visitCounts[childName] || 0;
      var remaining = quota - visits;
      var usageRate = quota > 0 ? visits / quota : 0;

      return [
        row[MASTER_COL.NO - 1],
        childName,
        quota,
        visits,
        remaining,
        usageRate,
      ];
    });

    sheet.getRange(SUMMARY_DATA_START_ROW, 1, outputData.length, 6).setValues(outputData);

    // 利用率列の表示形式を%に設定
    sheet.getRange(SUMMARY_DATA_START_ROW, SUMMARY_COL.USAGE_RATE, outputData.length, 1)
      .setNumberFormat('0%');

    ss.toast('月別集計の更新が完了しました', '完了', 3);
    Logger.log('月別集計を更新しました: ' + yearMonthStr + ' (' + outputData.length + '名)');
  } catch (error) {
    logError_('updateMonthlySummary', error);
  } finally {
    originalSheet.activate();
  }
}

/**
 * フォームの回答データから児童名ごとの来館回数を集計する
 * （実記録のみカウント。振り分けは含まない）
 * @param {Array<Array>} formResponses フォームの回答データ
 * @returns {Object} {児童名: 回数}
 */
function countVisitsByFormResponses_(formResponses) {
  var counts = {};
  formResponses.forEach(function(row) {
    var childName = row[FORM_COL.CHILD_NAME - 1];
    if (childName) {
      counts[childName] = (counts[childName] || 0) + 1;
    }
  });
  return counts;
}
