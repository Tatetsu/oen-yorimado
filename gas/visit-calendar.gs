/**
 * 来館カレンダー更新
 * 日×児童のマトリクス形式で来館状況を表示する
 */

/**
 * 来館カレンダーを更新する（ボタン実行）
 * 来館カレンダーシートのB1セル（対象年月）を参照して対象月を決定する
 */
function updateVisitCalendar() {
  var sheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
  var yearMonthStr = sheet.getRange('B1').getValue();

  if (!yearMonthStr) {
    SpreadsheetApp.getUi().alert('対象年月を選択してください');
    return;
  }

  var ym = parseYearMonth(yearMonthStr);
  updateVisitCalendarByMonth_(sheet, ym.year, ym.month);
}

/**
 * 指定年月の来館カレンダーを生成する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 来館カレンダーシート
 * @param {number} year 年
 * @param {number} month 月（1-12）
 */
function updateVisitCalendarByMonth_(sheet, year, month) {
  // 稼働中の児童名リストを取得
  var children = getActiveChildren();
  var childNames = children.map(function(row) {
    return row[MASTER_COL.NAME - 1];
  });

  if (childNames.length === 0) {
    Logger.log('稼働中の児童がいません');
    return;
  }

  // 対象月の日数を取得
  var daysInMonth = new Date(year, month, 0).getDate();

  // フォームの回答から来館データを取得
  var formResponses = getFormResponsesByMonth(year, month);

  // 振り分け記録を取得
  var allocations = getAllocationsByMonth(year, month);

  // 来館マップを構築: { "YYYY/MM/DD_児童名": "実データ" or "振り分け" }
  var visitMap = buildVisitMap_(formResponses, allocations);

  // --- データエリアをクリア（ヘッダー行より下） ---
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow >= CALENDAR_LAYOUT.HEADER_ROW) {
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, lastRow - CALENDAR_LAYOUT.HEADER_ROW + 1, Math.max(lastCol, childNames.length + 3)).clearContent();
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, lastRow - CALENDAR_LAYOUT.HEADER_ROW + 1, Math.max(lastCol, childNames.length + 3)).clearFormat();
  }

  // --- ヘッダー行を書き込み ---
  var dailyTotalCol = CALENDAR_LAYOUT.CHILD_START_COL + childNames.length; // 日計列
  var headerRow = ['日付'];
  childNames.forEach(function(name) {
    headerRow.push(name);
  });
  headerRow.push('日計');

  sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, 1, headerRow.length).setValues([headerRow]);

  // ヘッダー書式
  var headerRange = sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, 1, headerRow.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // --- 日別データ行を書き込み ---
  var dataRows = [];
  for (var day = 1; day <= daysInMonth; day++) {
    var dateStr = Utilities.formatDate(new Date(year, month - 1, day), 'Asia/Tokyo', 'yyyy/MM/dd');
    var displayDate = month + '/' + day;
    var row = [displayDate];
    var dailyTotal = 0;

    childNames.forEach(function(name) {
      var key = dateStr + '_' + name;
      if (visitMap[key] === '実データ') {
        row.push('○');
        dailyTotal++;
      } else if (visitMap[key] === '振り分け') {
        row.push('△');
        dailyTotal++;
      } else {
        row.push('');
      }
    });

    row.push(dailyTotal > 0 ? dailyTotal : '');
    dataRows.push(row);
  }

  var dataStartRow = CALENDAR_LAYOUT.DATA_START_ROW;
  sheet.getRange(dataStartRow, 1, dataRows.length, headerRow.length).setValues(dataRows);

  // データ行のセンタリング（児童列＋日計列）
  sheet.getRange(dataStartRow, CALENDAR_LAYOUT.CHILD_START_COL, dataRows.length, childNames.length + 1)
    .setHorizontalAlignment('center');

  // 来館ありの日を薄い背景色でハイライト
  for (var r = 0; r < dataRows.length; r++) {
    for (var c = 1; c <= childNames.length; c++) {
      var cellValue = dataRows[r][c];
      if (cellValue === '○') {
        sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.CHILD_START_COL + c - 1)
          .setBackground('#E8F5E9');
      } else if (cellValue === '△') {
        sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.CHILD_START_COL + c - 1)
          .setBackground('#FFF3E0');
      }
    }
  }

  // --- サマリ行（月計・枠・残） ---
  var summaryStartRow = dataStartRow + daysInMonth + 1; // 1行空けてサマリ

  // 月別集計シートからデータを取得
  var summaryData = getMonthlySummaryData_(year, month, childNames);

  // 月計行
  var monthlyTotalRow = ['月計'];
  var grandTotal = 0;
  childNames.forEach(function(name) {
    var count = summaryData[name] ? summaryData[name].visits : 0;
    monthlyTotalRow.push(count);
    grandTotal += count;
  });
  monthlyTotalRow.push(grandTotal);

  // 枠行
  var quotaRow = ['枠'];
  var quotaTotal = 0;
  childNames.forEach(function(name) {
    var quota = summaryData[name] ? summaryData[name].quota : 0;
    quotaRow.push(quota);
    quotaTotal += quota;
  });
  quotaRow.push(quotaTotal);

  // 残行
  var remainingRow = ['残'];
  var remainingTotal = 0;
  childNames.forEach(function(name) {
    var remaining = summaryData[name] ? summaryData[name].remaining : 0;
    remainingRow.push(remaining);
    remainingTotal += remaining;
  });
  remainingRow.push(remainingTotal);

  sheet.getRange(summaryStartRow, 1, 3, headerRow.length)
    .setValues([monthlyTotalRow, quotaRow, remainingRow]);

  // サマリ行の書式
  var summaryRange = sheet.getRange(summaryStartRow, 1, 3, headerRow.length);
  summaryRange.setFontWeight('bold');
  summaryRange.setHorizontalAlignment('center');
  sheet.getRange(summaryStartRow, 1, 1, headerRow.length).setBackground('#E3F2FD');     // 月計: 青
  sheet.getRange(summaryStartRow + 1, 1, 1, headerRow.length).setBackground('#F3E5F5'); // 枠: 紫
  sheet.getRange(summaryStartRow + 2, 1, 1, headerRow.length).setBackground('#FFF8E1'); // 残: 黄

  // 列幅調整
  sheet.setColumnWidth(1, 60); // 日付列
  for (var ci = 0; ci < childNames.length; ci++) {
    sheet.setColumnWidth(CALENDAR_LAYOUT.CHILD_START_COL + ci, 80);
  }
  sheet.setColumnWidth(dailyTotalCol, 50); // 日計列

  // 行固定
  sheet.setFrozenRows(CALENDAR_LAYOUT.HEADER_ROW);

  Logger.log('来館カレンダーを更新しました: ' + year + '年' + month + '月 (' + childNames.length + '名)');
}

/**
 * フォーム回答と振り分け記録から来館マップを構築する
 * @param {Array<Array>} formResponses フォーム回答データ
 * @param {Array<Array>} allocations 振り分けデータ
 * @returns {Object} { "yyyy/MM/dd_児童名": "実データ" | "振り分け" }
 */
function buildVisitMap_(formResponses, allocations) {
  var map = {};

  formResponses.forEach(function(row) {
    var recordDate = new Date(row[FORM_COL.RECORD_DATE - 1]);
    var dateStr = Utilities.formatDate(recordDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    var childName = row[FORM_COL.CHILD_NAME - 1];
    var key = dateStr + '_' + childName;
    map[key] = '実データ';
  });

  allocations.forEach(function(row) {
    var allocDate = new Date(row[ALLOCATION_COL.ALLOCATION_DATE - 1]);
    var dateStr = Utilities.formatDate(allocDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    var childName = row[ALLOCATION_COL.CHILD_NAME - 1];
    var key = dateStr + '_' + childName;
    // 実データが既にある場合は上書きしない
    if (!map[key]) {
      map[key] = '振り分け';
    }
  });

  return map;
}

/**
 * 月別集計シートから児童ごとの枠・来館数・残数を取得する
 * @param {number} year 年
 * @param {number} month 月
 * @param {Array<string>} childNames 児童名リスト
 * @returns {Object} { 児童名: { quota, visits, remaining } }
 */
function getMonthlySummaryData_(year, month, childNames) {
  var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);

  // 月別集計の対象年月を確認
  var currentYm = summarySheet.getRange('B1').getValue();
  var ym = parseYearMonth(currentYm);

  // 対象年月が異なる場合は一時的に更新して取得
  var needsRestore = false;
  if (ym.year !== year || ym.month !== month) {
    summarySheet.getRange('B1').setValue(year + '年' + month + '月');
    updateMonthlySummary();
    needsRestore = true;
  }

  // 月別集計データを読み取り
  var lastRow = summarySheet.getLastRow();
  var result = {};

  if (lastRow >= SUMMARY_DATA_START_ROW) {
    var data = summarySheet.getRange(SUMMARY_DATA_START_ROW, 1, lastRow - SUMMARY_DATA_START_ROW + 1, 6).getValues();
    data.forEach(function(row) {
      var name = row[SUMMARY_COL.NAME - 1];
      if (name) {
        result[name] = {
          quota: row[SUMMARY_COL.QUOTA - 1] || 0,
          visits: row[SUMMARY_COL.VISITS - 1] || 0,
          remaining: row[SUMMARY_COL.REMAINING - 1] || 0,
        };
      }
    });
  }

  // 元の年月に戻す
  if (needsRestore) {
    summarySheet.getRange('B1').setValue(currentYm);
    updateMonthlySummary();
  }

  return result;
}
