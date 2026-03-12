/**
 * F-01: シート初期セットアップ
 * 月別集計・振り分け記録・来館カレンダー・児童別ビューのシート作成、ヘッダー・書式設定
 */

/**
 * 全シートの初期セットアップを実行する（手動実行・1回のみ）
 */
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupMonthlySummarySheet_(ss);
  setupAllocationSheet_(ss);
  setupConfirmedVisitsSheet_(ss);
  setupVisitCalendarSheet_(ss);
  setupChildViewSheet_(ss);

  Logger.log('全シートの初期セットアップが完了しました');
  SpreadsheetApp.getUi().alert('初期セットアップが完了しました');
}

/**
 * 月別集計シートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupMonthlySummarySheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.MONTHLY_SUMMARY);

  // 操作エリア（1行目）
  sheet.getRange('A1').setValue('対象年月:');

  // 年月ドロップダウン（B1）
  var yearMonthOptions = generateYearMonthOptions();
  var yearMonthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(yearMonthOptions, true)
    .build();
  sheet.getRange('B1').setDataValidation(yearMonthRule);
  // 当月をデフォルトに設定
  var now = new Date();
  sheet.getRange('B1').setValue(now.getFullYear() + '年' + (now.getMonth() + 1) + '月');

  // ヘッダー行（2行目）
  var headers = ['No.', '児童名', '月間利用枠', '来館数', '残数', '利用率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(2);

  // 列幅調整
  sheet.setColumnWidth(1, 40);   // No.
  sheet.setColumnWidth(2, 100);  // 児童名
  sheet.setColumnWidth(3, 90);   // 月間利用枠
  sheet.setColumnWidth(4, 80);   // 来館数
  sheet.setColumnWidth(5, 60);   // 残数
  sheet.setColumnWidth(6, 70);   // 利用率
  Logger.log('月別集計シートのセットアップ完了');
}

/**
 * 振り分け記録シートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupAllocationSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.ALLOCATION);

  // ヘッダー行
  var headers = ['対象年月', '児童名', '振り分け日', 'スタッフ名', '入所時間', '退所時間', '体温', '食事', '入浴', '睡眠', '便', '服薬', 'その他連絡事項', '実行日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(1);

  // 列幅調整
  sheet.setColumnWidth(1, 100);   // 対象年月
  sheet.setColumnWidth(2, 100);   // 児童名
  sheet.setColumnWidth(3, 100);   // 振り分け日
  sheet.setColumnWidth(4, 100);   // スタッフ名
  sheet.setColumnWidth(5, 80);    // 入所時間
  sheet.setColumnWidth(6, 80);    // 退所時間
  sheet.setColumnWidth(7, 60);    // 体温
  sheet.setColumnWidth(8, 50);    // 食事
  sheet.setColumnWidth(9, 50);    // 入浴
  sheet.setColumnWidth(10, 50);   // 睡眠
  sheet.setColumnWidth(11, 50);   // 便
  sheet.setColumnWidth(12, 50);   // 服薬
  sheet.setColumnWidth(13, 200);  // その他連絡事項
  sheet.setColumnWidth(14, 150);  // 実行日時

  Logger.log('振り分け記録シートのセットアップ完了');
}

/**
 * 確定来館記録シートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupConfirmedVisitsSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CONFIRMED_VISITS);

  // ヘッダー行
  var headers = ['記録日', '児童名', 'データ区分', 'スタッフ名', '入所時間', '退所時間', '体温', '食事', '入浴', '睡眠', '便', '服薬', 'その他連絡事項'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(1);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 記録日
  sheet.setColumnWidth(2, 100);  // 児童名
  sheet.setColumnWidth(3, 80);   // データ区分
  sheet.setColumnWidth(4, 100);  // スタッフ名
  sheet.setColumnWidth(5, 80);   // 入所時間
  sheet.setColumnWidth(6, 80);   // 退所時間
  sheet.setColumnWidth(7, 60);   // 体温
  sheet.setColumnWidth(8, 50);   // 食事
  sheet.setColumnWidth(9, 50);   // 入浴
  sheet.setColumnWidth(10, 50);  // 睡眠
  sheet.setColumnWidth(11, 50);  // 便
  sheet.setColumnWidth(12, 50);  // 服薬
  sheet.setColumnWidth(13, 200); // その他連絡事項

  Logger.log('確定来館記録シートのセットアップ完了');
}

/**
 * 来館カレンダーシートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupVisitCalendarSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.VISIT_CALENDAR);

  // 操作エリア（1行目）
  sheet.getRange('A1').setValue('対象年月:');
  sheet.getRange('A1').setFontWeight('bold');

  // 年月ドロップダウン（B1）- デフォルトは前月
  var yearMonthOptions = generateYearMonthOptions();
  var yearMonthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(yearMonthOptions, true)
    .build();
  sheet.getRange('B1').setDataValidation(yearMonthRule);

  var now = new Date();
  var prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  sheet.getRange('B1').setValue(prevMonth.getFullYear() + '年' + (prevMonth.getMonth() + 1) + '月');

  // 凡例（2行目）
  sheet.getRange('A2').setValue('凡例: ○=実データ  △=振り分け');
  sheet.getRange('A2').setFontSize(9);
  sheet.getRange('A2').setFontColor('#666666');

  Logger.log('来館カレンダーシートのセットアップ完了');
}

/**
 * 児童別ビューシートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupChildViewSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CHILD_VIEW);

  // 操作エリア（1〜2行目）
  sheet.getRange('A1').setValue('児童名:');
  sheet.getRange('A2').setValue('対象年月:');

  // 児童名ドロップダウン（B1）
  var childNames = getChildNameOptions();
  if (childNames.length > 0) {
    var childNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(childNames, true)
      .build();
    sheet.getRange('B1').setDataValidation(childNameRule);
  }

  // 年月ドロップダウン（B2）
  var yearMonthOptions = generateYearMonthOptions();
  var yearMonthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(yearMonthOptions, true)
    .build();
  sheet.getRange('B2').setDataValidation(yearMonthRule);
  var now = new Date();
  sheet.getRange('B2').setValue(now.getFullYear() + '年' + (now.getMonth() + 1) + '月');

  // 基本情報エリアのラベル（4〜8行目）
  sheet.getRange('A4').setValue('保護者名:');
  sheet.getRange('A5').setValue('担当スタッフ:');
  sheet.getRange('A6').setValue('月間利用枠:');
  sheet.getRange('A7').setValue('医療型の有無:');
  sheet.getRange('A8').setValue('来館回数 / 残枠 / 利用率:');

  // ラベル列を太字に
  sheet.getRange('A4:A8').setFontWeight('bold');

  // 来館履歴ヘッダー（9行目）
  var historyHeaders = ['記録日', 'スタッフ名', '入所時間', '退所時間', '体温', '食事', '入浴', '睡眠', '便', '服薬', 'その他連絡事項'];
  sheet.getRange(9, 1, 1, historyHeaders.length).setValues([historyHeaders]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(9, 1, 1, historyHeaders.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(9);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 記録日
  sheet.setColumnWidth(2, 100);  // スタッフ名
  sheet.setColumnWidth(3, 80);   // 入所時間
  sheet.setColumnWidth(4, 80);   // 退所時間
  sheet.setColumnWidth(5, 60);   // 体温
  sheet.setColumnWidth(6, 50);   // 食事
  sheet.setColumnWidth(7, 50);   // 入浴
  sheet.setColumnWidth(8, 50);   // 睡眠
  sheet.setColumnWidth(9, 50);   // 便
  sheet.setColumnWidth(10, 50);  // 服薬
  sheet.setColumnWidth(11, 200); // その他連絡事項

  // 印刷設定（A4横向き）
  sheet.getRange('A1:L1').activate();

  Logger.log('児童別ビューシートのセットアップ完了');
}

/**
 * シートを取得、存在しなければ作成する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('シート「' + sheetName + '」を新規作成しました');
  } else {
    Logger.log('シート「' + sheetName + '」は既に存在します');
  }
  return sheet;
}
