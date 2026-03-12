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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('来館カレンダーを更新中...', '読み込み中', -1);

  var ym = parseYearMonth(yearMonthStr);
  updateVisitCalendarByMonth_(sheet, ym.year, ym.month);

  ss.toast('来館カレンダーの更新が完了しました', '完了', 3);
}

/**
 * 指定年月の来館カレンダーを生成する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 来館カレンダーシート
 * @param {number} year 年
 * @param {number} month 月（1-12）
 */
function updateVisitCalendarByMonth_(sheet, year, month) {
  // 対象月の日数を取得
  var daysInMonth = new Date(year, month, 0).getDate();

  // 確定来館記録から来館マップを構築
  var visitMap = buildVisitMapFromConfirmed_(year, month);

  // 来館マップからその月に記録がある児童名を抽出
  var visitedChildSet = {};
  Object.keys(visitMap).forEach(function(key) {
    var childName = key.split('_').slice(1).join('_'); // "yyyy/MM/dd_児童名" → 児童名
    visitedChildSet[childName] = true;
  });

  // 児童マスタの並び順を維持しつつ、来館記録がある児童のみフィルタ
  var allChildren = getChildMasterData();
  var childNames = allChildren.map(function(row) {
    return row[MASTER_COL.NAME - 1];
  }).filter(function(name) {
    return visitedChildSet[name];
  });

  // --- データエリアをクリア（ヘッダー行より下） ---
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow >= CALENDAR_LAYOUT.HEADER_ROW) {
    var clearCols = Math.max(lastCol, childNames.length + 3);
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, lastRow - CALENDAR_LAYOUT.HEADER_ROW + 1, clearCols).clearContent();
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, lastRow - CALENDAR_LAYOUT.HEADER_ROW + 1, clearCols).clearFormat();
  }

  if (childNames.length === 0) {
    // データがない月：メッセージを表示して終了
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1).setValue(year + '年' + month + '月の来館記録はありません');
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1).setFontColor('#999999');
    SpreadsheetApp.getUi().alert(year + '年' + month + '月の来館記録はありません。');
    return;
  }

  // --- 祝日マップを取得 ---
  var holidayMap = getJapaneseHolidays_(year, month);

  // --- ヘッダー行を書き込み ---
  var dailyTotalCol = CALENDAR_LAYOUT.CHILD_START_COL + childNames.length; // 日計列
  var headerRow = ['日付', '曜日'];
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
  var DOW_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
  var dataRows = [];
  var rowDowInfo = []; // 各行の曜日情報を保持（色付け用）
  for (var day = 1; day <= daysInMonth; day++) {
    var dateObj = new Date(year, month - 1, day);
    var dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
    var displayDate = month + '/' + day;
    var dow = dateObj.getDay(); // 0=日, 6=土
    var isHoliday = !!holidayMap[dateStr];
    var row = [displayDate, DOW_LABELS[dow]];
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
    rowDowInfo.push({ dow: dow, isHoliday: isHoliday });
  }

  var dataStartRow = CALENDAR_LAYOUT.DATA_START_ROW;
  sheet.getRange(dataStartRow, 1, dataRows.length, headerRow.length).setValues(dataRows);

  // データ行のセンタリング（曜日列＋児童列＋日計列）
  sheet.getRange(dataStartRow, CALENDAR_LAYOUT.DOW_COL, dataRows.length, childNames.length + 2)
    .setHorizontalAlignment('center');

  // 土日祝の行背景色と曜日文字色を設定
  for (var r = 0; r < dataRows.length; r++) {
    var info = rowDowInfo[r];
    var rowRange = sheet.getRange(dataStartRow + r, 1, 1, headerRow.length);

    if (info.dow === 0 || info.isHoliday) {
      // 日曜・祝日: 薄い赤背景、曜日文字を赤
      rowRange.setBackground('#FFEBEE');
      sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.DOW_COL).setFontColor('#D32F2F');
    } else if (info.dow === 6) {
      // 土曜: 薄い青背景、曜日文字を青
      rowRange.setBackground('#E3F2FD');
      sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.DOW_COL).setFontColor('#1565C0');
    }
  }

  // 来館ありの日を薄い背景色でハイライト（土日祝の行色より上書き）
  for (var r = 0; r < dataRows.length; r++) {
    for (var c = 0; c < childNames.length; c++) {
      var cellValue = dataRows[r][c + 2]; // 曜日列分オフセット
      if (cellValue === '○') {
        sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.CHILD_START_COL + c)
          .setBackground('#E8F5E9');
      } else if (cellValue === '△') {
        sheet.getRange(dataStartRow + r, CALENDAR_LAYOUT.CHILD_START_COL + c)
          .setBackground('#FFF3E0');
      }
    }
  }

  // --- サマリ行（月計・枠・残） ---
  var summaryStartRow = dataStartRow + daysInMonth + 1; // 1行空けてサマリ

  // 月別集計シートからデータを取得
  var summaryData = getMonthlySummaryData_(year, month, childNames);

  // 月計行
  var monthlyTotalRow = ['月計', ''];
  var grandTotal = 0;
  childNames.forEach(function(name) {
    var count = summaryData[name] ? summaryData[name].visits : 0;
    monthlyTotalRow.push(count);
    grandTotal += count;
  });
  monthlyTotalRow.push(grandTotal);

  // 枠行
  var quotaRow = ['枠', ''];
  var quotaTotal = 0;
  childNames.forEach(function(name) {
    var quota = summaryData[name] ? summaryData[name].quota : 0;
    quotaRow.push(quota);
    quotaTotal += quota;
  });
  quotaRow.push(quotaTotal);

  // 残行
  var remainingRow = ['残', ''];
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
  sheet.setColumnWidth(1, 60);  // 日付列
  sheet.setColumnWidth(CALENDAR_LAYOUT.DOW_COL, 40);  // 曜日列
  for (var ci = 0; ci < childNames.length; ci++) {
    sheet.setColumnWidth(CALENDAR_LAYOUT.CHILD_START_COL + ci, 80);
  }
  sheet.setColumnWidth(dailyTotalCol, 50); // 日計列

  // 行固定
  sheet.setFrozenRows(CALENDAR_LAYOUT.HEADER_ROW);

  Logger.log('来館カレンダーを更新しました: ' + year + '年' + month + '月 (' + childNames.length + '名)');
}

/**
 * 確定来館記録シートから指定月の来館マップを構築する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Object} { "yyyy/MM/dd_児童名": "実データ" | "振り分け" }
 */
function buildVisitMapFromConfirmed_(year, month) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var lastRow = sheet.getLastRow();
  var map = {};

  if (lastRow < CONFIRMED_DATA_START_ROW) {
    return map;
  }

  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, CONFIRMED_COL.DATA_TYPE).getValues();

  data.forEach(function(row) {
    var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
    if (recordDate.getFullYear() !== year || (recordDate.getMonth() + 1) !== month) {
      return;
    }
    var dateStr = Utilities.formatDate(recordDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    var childName = row[CONFIRMED_COL.CHILD_NAME - 1];
    var dataType = row[CONFIRMED_COL.DATA_TYPE - 1];
    var key = dateStr + '_' + childName;
    map[key] = dataType;
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

/**
 * Googleカレンダーから日本の祝日を取得する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Object} { "yyyy/MM/dd": 祝日名 }
 */
function getJapaneseHolidays_(year, month) {
  var map = {};
  try {
    var cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (!cal) return map;
    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0);
    var events = cal.getEvents(startDate, endDate);
    events.forEach(function(event) {
      var dateStr = Utilities.formatDate(event.getStartTime(), 'Asia/Tokyo', 'yyyy/MM/dd');
      map[dateStr] = event.getTitle();
    });
  } catch (e) {
    Logger.log('祝日カレンダーの取得に失敗: ' + e.message);
  }
  return map;
}
