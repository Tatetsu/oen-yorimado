/**
 * 来館カレンダー更新
 * 日×児童のマトリクス形式で来館状況を表示する
 * 年別（12ヶ月分）または月別で表示可能
 */

/**
 * 来館カレンダーを更新する（ボタン実行 / onEdit）
 */
function updateVisitCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    var sheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    var rawValue = sheet.getRange('B1').getValue();

    if (!rawValue) {
      SpreadsheetApp.getUi().alert('対象を選択してください');
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('来館カレンダーを更新中...', '読み込み中', -1);

    var selectionStr = String(rawValue).trim();
    var yearOnly = parseYearOnly_(selectionStr);

    if (yearOnly !== null) {
      updateVisitCalendarByYear_(sheet, yearOnly);
    } else {
      var ym = parseYearMonth(rawValue);
      updateVisitCalendarByMonth_(sheet, ym.year, ym.month);
    }

    ss.toast('来館カレンダーの更新が完了しました', '完了', 3);
  } catch (error) {
    logError_('updateVisitCalendar', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  } finally {
    originalSheet.activate();
  }
}
// parseYearOnly_ は utils.gs に移動済み

// ========================================
// 月別表示
// ========================================

/**
 * 指定年月の来館カレンダーを生成する
 * レイアウト: ヘッダー(3行目) → 集計(4-6行目) → 日別データ(7行目〜)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 来館カレンダーシート
 * @param {number} year 年
 * @param {number} month 月（1-12）
 */
function updateVisitCalendarByMonth_(sheet, year, month) {
  var daysInMonth = new Date(year, month, 0).getDate();
  var visitMap = buildVisitMapFromConfirmed_(year, month);

  var visitedChildSet = {};
  Object.keys(visitMap).forEach(function(key) {
    visitedChildSet[key.split('_').slice(1).join('_')] = true;
  });
  var allChildren = getChildMasterData();
  var childNames = allChildren.map(function(r) { return r[MASTER_COL.NAME - 1]; })
    .filter(function(n) { return visitedChildSet[n]; });

  // 前回の児童列数を保存（列幅維持の判定用）
  var prevChildCount = 0;
  if (sheet.getLastRow() >= CALENDAR_LAYOUT.HEADER_ROW && sheet.getLastColumn() > CALENDAR_LAYOUT.CHILD_START_COL) {
    prevChildCount = sheet.getLastColumn() - CALENDAR_LAYOUT.CHILD_START_COL;
  }

  clearCalendarArea_(sheet, childNames.length);

  if (childNames.length === 0) {
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1)
      .setValue(year + '年' + month + '月の来館記録はありません')
      .setFontColor('#999999');
    SpreadsheetApp.getUi().alert(year + '年' + month + '月の来館記録はありません。');
    return;
  }

  var holidayMap = getJapaneseHolidays_(year, month);
  var headerRow = buildHeaderRow_(childNames);
  var dailyTotalCol = CALENDAR_LAYOUT.CHILD_START_COL + childNames.length;

  // 凡例（A2）
  writeLegend_(sheet);

  // ヘッダー行
  writeHeaderRow_(sheet, headerRow);

  // 集計行（月間利用数/月間利用枠/残り利用枠）をヘッダー直下に配置
  var summaryStartRow = CALENDAR_LAYOUT.HEADER_ROW + 1;
  writeMonthlySummary_(sheet, summaryStartRow, childNames, visitMap, allChildren, headerRow.length, year, month);

  // 日別データ（集計行の下から開始）
  var dataStartRow = summaryStartRow + 3;
  var DOW_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
  var dataRows = [];
  var typeRows = [];
  var rowDowInfo = [];

  for (var day = 1; day <= daysInMonth; day++) {
    var dateObj = new Date(year, month - 1, day);
    var dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
    var dow = dateObj.getDay();
    var row = [month + '/' + day, DOW_LABELS[dow]];
    var typeRow = ['', ''];
    var dailyTotal = 0;

    childNames.forEach(function(name) {
      var val = visitMap[dateStr + '_' + name];
      if (val === '実データ') { row.push('○'); typeRow.push('実データ'); dailyTotal++; }
      else if (val === '振り分け') { row.push('○'); typeRow.push('振り分け'); dailyTotal++; }
      else { row.push(''); typeRow.push(''); }
    });
    row.push(dailyTotal > 0 ? dailyTotal : '');
    typeRow.push('');
    dataRows.push(row);
    typeRows.push(typeRow);
    rowDowInfo.push({ dow: dow, isHoliday: !!holidayMap[dateStr] });
  }

  sheet.getRange(dataStartRow, 1, dataRows.length, headerRow.length).setValues(dataRows);
  applyDataFormatting_(sheet, dataStartRow, dataRows, rowDowInfo, childNames, headerRow.length, typeRows);

  // 児童数が変わった場合のみ列幅を再設定
  if (childNames.length !== prevChildCount) {
    applyColumnWidths_(sheet, childNames.length, dailyTotalCol);
  }

  // ヘッダー＋集計行を固定
  sheet.setFrozenRows(dataStartRow - 1);

  Logger.log('来館カレンダーを更新しました: ' + year + '年' + month + '月 (' + childNames.length + '名)');
}

// ========================================
// 年別表示
// ========================================

/**
 * 指定年の来館カレンダーを生成する（12ヶ月分）
 * レイアウト: ヘッダー(3行目) → 集計(4-5行目) → 日別データ(6行目〜)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 来館カレンダーシート
 * @param {number} year 年
 */
function updateVisitCalendarByYear_(sheet, year) {
  var visitMap = buildVisitMapFromConfirmedYear_(year);

  var visitedChildSet = {};
  Object.keys(visitMap).forEach(function(key) {
    visitedChildSet[key.split('_').slice(1).join('_')] = true;
  });
  var allChildren = getChildMasterData();
  var childNames = allChildren.map(function(r) { return r[MASTER_COL.NAME - 1]; })
    .filter(function(n) { return visitedChildSet[n]; });

  // 前回の児童列数を保存（列幅維持の判定用）
  var prevChildCount = 0;
  if (sheet.getLastRow() >= CALENDAR_LAYOUT.HEADER_ROW && sheet.getLastColumn() > CALENDAR_LAYOUT.CHILD_START_COL) {
    prevChildCount = sheet.getLastColumn() - CALENDAR_LAYOUT.CHILD_START_COL;
  }

  clearCalendarArea_(sheet, childNames.length);

  if (childNames.length === 0) {
    sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1)
      .setValue(year + '年の来館記録はありません')
      .setFontColor('#999999');
    return;
  }

  var holidayMap = getJapaneseHolidaysYear_(year);
  var headerRow = buildHeaderRow_(childNames);
  var dailyTotalCol = CALENDAR_LAYOUT.CHILD_START_COL + childNames.length;

  // 凡例（A2）
  writeLegend_(sheet);

  // ヘッダー行
  writeHeaderRow_(sheet, headerRow);

  // 集計行（年計/月枠）をヘッダー直下に配置
  var summaryStartRow = CALENDAR_LAYOUT.HEADER_ROW + 1;
  writeYearlySummary_(sheet, summaryStartRow, childNames, allChildren, visitMap, headerRow.length, year);

  // 日別データ（集計行の下から開始）
  var dataStartRow = summaryStartRow + 3;
  var DOW_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
  var dataRows = [];
  var typeRows = [];
  var rowDowInfo = [];

  for (var m = 1; m <= 12; m++) {
    var daysInMonth = new Date(year, m, 0).getDate();
    for (var day = 1; day <= daysInMonth; day++) {
      var dateObj = new Date(year, m - 1, day);
      var dateStr = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
      var dow = dateObj.getDay();
      var row = [m + '/' + day, DOW_LABELS[dow]];
      var typeRow = ['', ''];
      var dailyTotal = 0;

      childNames.forEach(function(name) {
        var val = visitMap[dateStr + '_' + name];
        if (val === '実データ') { row.push('○'); typeRow.push('実データ'); dailyTotal++; }
        else if (val === '振り分け') { row.push('○'); typeRow.push('振り分け'); dailyTotal++; }
        else { row.push(''); typeRow.push(''); }
      });
      row.push(dailyTotal > 0 ? dailyTotal : '');
      typeRow.push('');
      dataRows.push(row);
      typeRows.push(typeRow);
      rowDowInfo.push({ dow: dow, isHoliday: !!holidayMap[dateStr] });
    }
  }

  sheet.getRange(dataStartRow, 1, dataRows.length, headerRow.length).setValues(dataRows);
  applyDataFormatting_(sheet, dataStartRow, dataRows, rowDowInfo, childNames, headerRow.length, typeRows);

  // 児童数が変わった場合のみ列幅を再設定
  if (childNames.length !== prevChildCount) {
    applyColumnWidths_(sheet, childNames.length, dailyTotalCol);
  }

  // ヘッダー＋集計行を固定
  sheet.setFrozenRows(dataStartRow - 1);

  Logger.log('来館カレンダーを更新しました: ' + year + '年 (' + childNames.length + '名)');
}

// ========================================
// 共通ヘルパー関数
// ========================================

/**
 * A2セルに凡例をリッチテキストで書き込む
 * 実データ=緑（#2E7D32）、振り分け=橙（#E65100）
 */
function writeLegend_(sheet) {
  var text = '凡例: 緑=実データ  橙=振り分け';
  // 文字位置（0始まり）: "凡例: " = 0-3, "緑=実データ" = 4-9, "  " = 10-11, "橙=振り分け" = 12-17
  var green  = SpreadsheetApp.newTextStyle().setForegroundColor('#2E7D32').build();
  var orange = SpreadsheetApp.newTextStyle().setForegroundColor('#E65100').build();
  var richText = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(4, 10, green)   // 緑=実データ
    .setTextStyle(12, 18, orange) // 橙=振り分け
    .build();
  var cell = sheet.getRange('A2');
  cell.setRichTextValue(richText);
  cell.setFontSize(9);
}

/**
 * ヘッダー行の配列を構築する
 */
function buildHeaderRow_(childNames) {
  var row = ['日付', '曜日'];
  childNames.forEach(function(name) { row.push(name); });
  row.push('日計');
  return row;
}

/**
 * カレンダーのデータエリアをクリアする
 */
function clearCalendarArea_(sheet, childCount) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow >= CALENDAR_LAYOUT.HEADER_ROW) {
    var cols = Math.max(lastCol, childCount + 4);
    var range = sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, lastRow - CALENDAR_LAYOUT.HEADER_ROW + 1, cols);
    range.clearContent();
    range.clearFormat();
  }
}

/**
 * ヘッダー行を書き込む
 */
function writeHeaderRow_(sheet, headerRow) {
  var range = sheet.getRange(CALENDAR_LAYOUT.HEADER_ROW, 1, 1, headerRow.length);
  range.setValues([headerRow]);
  range.setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * 月別集計行（月間利用数/月間利用枠/残り利用枠）をヘッダー直下に書き込む
 * visitMap から実データ・振り分け両方を含めて集計する
 */
function writeMonthlySummary_(sheet, startRow, childNames, visitMap, allChildren, totalCols, year, month) {
  // 月間利用数（実データ + 振り分け の合計）
  var visitCounts = {};
  childNames.forEach(function(n) { visitCounts[n] = 0; });
  Object.keys(visitMap).forEach(function(key) {
    var name = key.split('_').slice(1).join('_');
    if (visitCounts[name] !== undefined) visitCounts[name]++;
  });

  // 月間利用枠: 児童マスタの J列（MONTHLY_QUOTA）をそのまま使用
  var masterQuotaMap = {};
  allChildren.forEach(function(row) {
    var name = row[MASTER_COL.NAME - 1];
    masterQuotaMap[name] = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
  });

  var monthlyTotalRow = ['月間利用数', ''];
  childNames.forEach(function(name) {
    monthlyTotalRow.push(visitCounts[name] || 0);
  });
  monthlyTotalRow.push(''); // 日計列は空

  var quotaRow = ['月間利用枠', ''];
  childNames.forEach(function(name) {
    quotaRow.push(masterQuotaMap[name] || 0);
  });
  quotaRow.push(''); // 日計列は空

  var remainingRow = ['残り利用枠', ''];
  childNames.forEach(function(name) {
    var q = masterQuotaMap[name] || 0;
    var c = visitCounts[name] || 0;
    remainingRow.push(q - c);
  });
  remainingRow.push(''); // 日計列は空

  sheet.getRange(startRow, 1, 3, totalCols)
    .setValues([monthlyTotalRow, quotaRow, remainingRow]);

  var range = sheet.getRange(startRow, 1, 3, totalCols);
  range.setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(startRow, 1, 1, totalCols).setBackground('#E3F2FD');      // 月間利用数：薄青
  sheet.getRange(startRow + 1, 1, 1, totalCols).setBackground('#E8F5E9'); // 月間利用枠：薄緑
  sheet.getRange(startRow + 2, 1, 1, totalCols).setBackground('#F3E5F5'); // 残り利用枠：薄紫（振り分け色 #FFF3E0 と区別）
}

/**
 * 年別集計行（年計/年間枠/月間枠）をヘッダー直下に書き込む
 */
function writeYearlySummary_(sheet, startRow, childNames, allChildren, visitMap, totalCols, year) {
  var visitCounts = {};
  childNames.forEach(function(n) { visitCounts[n] = 0; });
  Object.keys(visitMap).forEach(function(key) {
    var name = key.split('_').slice(1).join('_');
    if (visitCounts[name] !== undefined) visitCounts[name]++;
  });

  var yearTotalRow = ['年計', ''];
  childNames.forEach(function(name) {
    yearTotalRow.push(visitCounts[name]);
  });
  yearTotalRow.push(''); // 日計列は空

  // 年間利用枠: 児童マスタの I列（ANNUAL_QUOTA）
  // 月間利用枠: 児童マスタの J列（MONTHLY_QUOTA）
  var annualQuotaMap = {};
  var monthlyQuotaMap = {};
  allChildren.forEach(function(row) {
    var name = row[MASTER_COL.NAME - 1];
    annualQuotaMap[name] = row[MASTER_COL.ANNUAL_QUOTA - 1] || 0;
    monthlyQuotaMap[name] = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
  });

  var annualQuotaRow = ['年間枠', ''];
  childNames.forEach(function(name) {
    annualQuotaRow.push(annualQuotaMap[name] || 0);
  });
  annualQuotaRow.push(''); // 日計列は空

  var monthlyQuotaRow = ['月間枠', ''];
  childNames.forEach(function(name) {
    monthlyQuotaRow.push(monthlyQuotaMap[name] || 0);
  });
  monthlyQuotaRow.push(''); // 日計列は空

  sheet.getRange(startRow, 1, 3, totalCols)
    .setValues([yearTotalRow, annualQuotaRow, monthlyQuotaRow]);

  var range = sheet.getRange(startRow, 1, 3, totalCols);
  range.setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(startRow, 1, 1, totalCols).setBackground('#E3F2FD');     // 年計：薄青
  sheet.getRange(startRow + 1, 1, 1, totalCols).setBackground('#E8F5E9'); // 年間枠：薄緑
  sheet.getRange(startRow + 2, 1, 1, totalCols).setBackground('#F3E5F5'); // 月間枠：薄紫
}

/**
 * 日別データ行の書式を一括適用する（バッチ処理で高速化）
 */
function applyDataFormatting_(sheet, dataStartRow, dataRows, rowDowInfo, childNames, totalCols, typeRows) {
  var numRows = dataRows.length;
  if (numRows === 0) return;

  // センタリング
  sheet.getRange(dataStartRow, CALENDAR_LAYOUT.DOW_COL, numRows, childNames.length + 2)
    .setHorizontalAlignment('center');

  // 背景色・文字色をバッチで構築
  var backgrounds = [];
  var fontColors = [];

  for (var r = 0; r < numRows; r++) {
    var bgRow = [];
    var fcRow = [];
    var info = rowDowInfo[r];
    var rowBg = null;
    var dowColor = null;

    if (info.dow === 0 || info.isHoliday) {
      rowBg = '#FFEBEE';
      dowColor = '#D32F2F';
    } else if (info.dow === 6) {
      rowBg = '#E3F2FD';
      dowColor = '#1565C0';
    }

    for (var c = 0; c < totalCols; c++) {
      var bg = rowBg;
      var fc = null;

      // 曜日列の文字色
      if (c === (CALENDAR_LAYOUT.DOW_COL - 1) && dowColor) {
        fc = dowColor;
      }

      // 児童セルのハイライト（土日祝の行色を上書き）
      var childIdx = c - (CALENDAR_LAYOUT.CHILD_START_COL - 1);
      if (childIdx >= 0 && childIdx < childNames.length) {
        var type = typeRows[r][c];
        if (type === '実データ') {
          bg = '#E8F5E9';
        } else if (type === '振り分け') {
          bg = '#FFF3E0';
        }
      }

      bgRow.push(bg);
      fcRow.push(fc);
    }
    backgrounds.push(bgRow);
    fontColors.push(fcRow);
  }

  var dataRange = sheet.getRange(dataStartRow, 1, numRows, totalCols);
  dataRange.setBackgrounds(backgrounds);
  dataRange.setFontColors(fontColors);
}

/**
 * 列幅を調整する
 */
function applyColumnWidths_(sheet, childCount, dailyTotalCol) {
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(CALENDAR_LAYOUT.DOW_COL, 40);
  for (var i = 0; i < childCount; i++) {
    sheet.setColumnWidth(CALENDAR_LAYOUT.CHILD_START_COL + i, 80);
  }
  sheet.setColumnWidth(dailyTotalCol, 50);
}

// ========================================
// データ取得
// ========================================

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
 * 確定来館記録シートから指定年の来館マップを構築する（年別表示用）
 * @param {number} year 年
 * @returns {Object} { "yyyy/MM/dd_児童名": "実データ" | "振り分け" }
 */
function buildVisitMapFromConfirmedYear_(year) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var lastRow = sheet.getLastRow();
  var map = {};

  if (lastRow < CONFIRMED_DATA_START_ROW) return map;

  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, CONFIRMED_COL.DATA_TYPE).getValues();

  data.forEach(function(row) {
    var recordDate = new Date(row[CONFIRMED_COL.RECORD_DATE - 1]);
    if (recordDate.getFullYear() !== year) return;
    var dateStr = Utilities.formatDate(recordDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    var childName = row[CONFIRMED_COL.CHILD_NAME - 1];
    var dataType = row[CONFIRMED_COL.DATA_TYPE - 1];
    map[dateStr + '_' + childName] = dataType;
  });

  return map;
}


/**
 * Googleカレンダーから日本の祝日を取得する（月単位）
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

/**
 * Googleカレンダーから日本の祝日を取得する（年単位）
 * @param {number} year 年
 * @returns {Object} { "yyyy/MM/dd": 祝日名 }
 */
function getJapaneseHolidaysYear_(year) {
  var map = {};
  try {
    var cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (!cal) return map;
    var startDate = new Date(year, 0, 1);
    var endDate = new Date(year, 11, 31, 23, 59, 59);
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
