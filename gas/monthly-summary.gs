/**
 * F-02: 月別集計更新
 * フォームの回答から集計し、月別集計シートに値を書き込む
 */

/**
 * 月別集計を更新する（メイン処理）
 * 月別集計シートのB1セル（対象年月）とB2セル（入所状況フィルタ）を参照して集計する
 */
function updateMonthlySummary() {
  var sheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);

  // 対象年月を取得
  var yearMonthStr = sheet.getRange('B1').getValue();
  if (!yearMonthStr) {
    Logger.log('対象年月が選択されていません');
    return;
  }
  var ym = parseYearMonth(yearMonthStr);

  // 児童マスタ取得
  var masterData = getChildMasterData();

  // フォームの回答から該当月データ取得
  var formResponses = getFormResponsesByMonth(ym.year, ym.month);

  // 児童名ごとの来館回数を集計
  var visitCounts = countVisitsByChildName_(formResponses);

  // データエリアをクリア（ヘッダーは残す）
  var lastRow = sheet.getLastRow();
  if (lastRow >= SUMMARY_DATA_START_ROW) {
    sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, 6).clearContent();
  }

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

  Logger.log('月別集計を更新しました: ' + yearMonthStr + ' (' + outputData.length + '名)');
}

/**
 * フォーム回答データから児童名ごとの来館回数を集計する
 * @param {Array<Array>} responses フォーム回答データ
 * @returns {Object} {児童名: 回数}
 */
function countVisitsByChildName_(responses) {
  var counts = {};
  responses.forEach(function(row) {
    var childName = row[FORM_COL.CHILD_NAME - 1];
    if (childName) {
      counts[childName] = (counts[childName] || 0) + 1;
    }
  });
  return counts;
}

