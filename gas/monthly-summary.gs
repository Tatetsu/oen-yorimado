/**
 * F-02: 月別集計更新
 * 確定来館記録から集計し、月別集計シートに値を書き込む
 */

/**
 * 月別集計を更新する（メイン処理）
 * 月別集計シートのB1セル（対象年月）を参照して集計する
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('月別集計を更新中...', '読み込み中', -1);

  // 児童マスタ取得
  var masterData = getChildMasterData();

  // 確定来館記録から該当月データ取得して集計
  var confirmedVisits = getConfirmedVisitsByMonth(ym.year, ym.month);
  var visitCounts = countVisitsByChildName_(confirmedVisits);

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

  ss.toast('月別集計の更新が完了しました', '完了', 3);
  Logger.log('月別集計を更新しました: ' + yearMonthStr + ' (' + outputData.length + '名)');
}

/**
 * 確定来館記録データから児童名ごとの来館回数を集計する
 * （実データ + 振り分けの両方をカウント）
 * @param {Array<Array>} records 確定来館記録データ
 * @returns {Object} {児童名: 回数}
 */
function countVisitsByChildName_(records) {
  var counts = {};
  records.forEach(function(row) {
    var childName = row[CONFIRMED_COL.CHILD_NAME - 1];
    if (childName) {
      counts[childName] = (counts[childName] || 0) + 1;
    }
  });
  return counts;
}
