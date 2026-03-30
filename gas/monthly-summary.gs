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

    ss.toast('月別集計を更新中...', '読み込み中', -1);

    // 年のみ指定（年間集計）の場合
    var yearOnly = parseYearOnly_(String(yearMonthStr).trim());
    if (yearOnly !== null) {
      updateAnnualSummary_(sheet, yearOnly);
      ss.toast('月別集計の更新が完了しました', '完了', 3);
      Logger.log('月別集計を更新しました（年間）: ' + yearMonthStr);
      return;
    }

    var ym = parseYearMonth(yearMonthStr);

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
 * 年間集計を月別集計シートに書き込む
 * 年間利用枠・年間来館数・残数・利用率を表示する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月別集計シート
 * @param {number} year 対象年
 */
function updateAnnualSummary_(sheet, year) {
  var masterData = getChildMasterData();

  // データエリアをクリア（ヘッダーは残す）
  var lastRow = sheet.getLastRow();
  if (lastRow >= SUMMARY_DATA_START_ROW) {
    var lastCol = Math.max(sheet.getLastColumn(), 6);
    sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, lastCol).clearContent();
    sheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, lastCol).clearFormat();
  }

  // ヘッダーを年間用に更新
  var headers = ['No.', '児童名', '年間利用枠', '来館数', '残数', '利用率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, headers.length)
    .setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  if (masterData.length === 0) {
    Logger.log('表示対象の児童がいません');
    return;
  }

  // 年間フォーム回答から来館数を集計
  var formResponses = getFormResponsesByYear(year);
  var visitCounts = countVisitsByFormResponses_(formResponses);

  var outputData = masterData.map(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[MASTER_COL.ANNUAL_QUOTA - 1] || 0;
    var visits = visitCounts[childName] || 0;
    var remaining = quota > 0 ? quota - visits : '';
    var usageRate = quota > 0 ? visits / quota : 0;

    return [
      row[MASTER_COL.NO - 1],
      childName,
      quota || '',
      visits,
      remaining,
      usageRate,
    ];
  });

  sheet.getRange(SUMMARY_DATA_START_ROW, 1, outputData.length, 6).setValues(outputData);
  sheet.getRange(SUMMARY_DATA_START_ROW, SUMMARY_COL.USAGE_RATE, outputData.length, 1)
    .setNumberFormat('0%');
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
    if (!childName) return;
    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];
    var days = calcStayDays_(checkIn, checkOut);
    counts[childName] = (counts[childName] || 0) + days;
  });
  return counts;
}
