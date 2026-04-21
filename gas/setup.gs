/**
 * F-01: シート初期セットアップ
 * 月別集計・確定来館記録・来館カレンダー・児童別ビューのシート作成、ヘッダー・書式設定
 */

/**
 * 全シートの初期セットアップを実行する（手動実行・1回のみ）
 * シート構造を整えたあと、データが存在するビューは初期描画も行う
 */
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  setupMonthlySummarySheet_(ss);
  setupConfirmedVisitsSheet_(ss);
  setupVisitCalendarSheet_(ss);
  setupChildViewSheet_(ss);
  setupLogSheet_(ss);
  setupChildMasterValidations_();

  // 初期描画（フォーム回答がある場合のみ）
  populateViewsIfDataExists_();

  Logger.log('全シートの初期セットアップが完了しました');
  SpreadsheetApp.getUi().alert('初期セットアップが完了しました');
}

/**
 * フォーム回答データがある場合、ビューシートを初期描画する
 * 児童別ビューは児童名の選択が必要なため除外
 */
function populateViewsIfDataExists_() {
  try {
    var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FORM_RESPONSE);
    if (!formSheet || formSheet.getLastRow() < FORM_DATA_START_ROW) return;

    updateMonthlySummary();
    updateConfirmedVisits();
    filterConfirmedVisits_();
    updateVisitCalendar();
  } catch (e) {
    logError_('populateViewsIfDataExists_', e);
  }
}

/**
 * 児童マスタのドロップダウンを設定する（優先度 1〜5）
 * 既存シートへの追加適用にも使用できる
 */
function setupChildMasterValidations_() {
  var sheet = getSheet(SHEET_NAMES.CHILD_MASTER);
  var lastRow = Math.max(sheet.getLastRow(), 2);

  // 重度支援区分列（MASTER_COL.PRIORITY = 8列目）に 区分1〜区分5 のドロップダウンを設定
  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['区分1', '区分2', '区分3', '区分4', '区分5'], true)
    .build();
  sheet.getRange(2, MASTER_COL.PRIORITY, lastRow - 1, 1).setDataValidation(priorityRule);

  Logger.log('児童マスタのドロップダウン設定完了（重度支援区分 区分1〜区分5）');
}

/**
 * 児童マスタのドロップダウンを手動で再設定する（メニューから実行）
 */
function refreshChildMasterValidations() {
  setupChildMasterValidations_();
  SpreadsheetApp.getUi().alert('児童マスタのドロップダウンを更新しました（重度支援区分 区分1〜区分5）');
}

/**
 * 月別集計シートを作成・設定する
 * レイアウト: 1行目=対象年, 2行目=対象月, 3行目=ヘッダー, 4行目〜=データ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupMonthlySummarySheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.MONTHLY_SUMMARY);

  // 操作エリア（1〜2行目）
  sheet.getRange('A1').setValue('対象年:').setFontWeight('bold');
  sheet.getRange('A2').setValue('対象月:').setFontWeight('bold');

  // 年ドロップダウン（B1）
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateYearOptions(), true).build();
  sheet.getRange('B1').setDataValidation(yearRule);
  var years = collectYearsFromFormResponses_();
  sheet.getRange('B1').setValue(years[0] + '年');

  // 月ドロップダウン（B2）- デフォルトは「すべて」
  var monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateMonthOptions(), true).build();
  sheet.getRange('B2').setDataValidation(monthRule);
  sheet.getRange('B2').setValue('すべて');

  // ヘッダー行（3行目）
  var headers = ['No.', '児童名', '利用枠', '来館数', '残数', '利用率'];
  var headerRange = sheet.getRange(SUMMARY_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(SUMMARY_HEADER_ROW);

  // 列幅調整
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 70);
  Logger.log('月別集計シートのセットアップ完了');
}

/**
 * 確定来館記録シートを作成・設定する
 * レイアウト: 1行目=対象年, 2行目=対象月, 3行目=ヘッダー, 4行目〜=データ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupConfirmedVisitsSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CONFIRMED_VISITS);

  // 操作エリア（1〜2行目）
  sheet.getRange('A1').setValue('対象年:').setFontWeight('bold');
  sheet.getRange('A2').setValue('対象月:').setFontWeight('bold');

  // 年ドロップダウン（B1）- デフォルトは前月の年
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateYearOptions(), true).build();
  sheet.getRange('B1').setDataValidation(yearRule);

  // 月ドロップダウン（B2）- デフォルトは前月
  var monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateMonthOptions(), true).build();
  sheet.getRange('B2').setDataValidation(monthRule);

  // デフォルト値：現在日付の1ヶ月前（月初トリガーで前月データを集計するため）
  var lastMonth = new Date();
  lastMonth.setDate(1);
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  sheet.getRange('B1').setValue(lastMonth.getFullYear() + '年');
  sheet.getRange('B2').setValue((lastMonth.getMonth() + 1) + '月');

  // ヘッダー行（3行目）
  var headers = ['記録日', '児童名', 'データ区分', 'スタッフ1', 'スタッフ2', '入所日時', '退所日時', '体温', '食事', '入浴', '睡眠', '便', '服薬', 'その他連絡事項'];
  var headerRange = sheet.getRange(CONFIRMED_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  // 行固定（操作エリア+ヘッダー）
  sheet.setFrozenRows(CONFIRMED_HEADER_ROW);

  // 列幅調整
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 60);
  sheet.setColumnWidth(9, 50);
  sheet.setColumnWidth(10, 50);
  sheet.setColumnWidth(11, 50);
  sheet.setColumnWidth(12, 50);
  sheet.setColumnWidth(13, 50);
  sheet.setColumnWidth(14, 200);

  Logger.log('確定来館記録シートのセットアップ完了');
}

/**
 * 来館カレンダーシートを作成・設定する
 * B1=対象年、B2=対象月。月が「すべて」なら年別、具体値なら月別カレンダーを描画する
 * レイアウト: 1行目=対象年, 2行目=対象月（凡例は3行目以降にカレンダーが描かれる）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupVisitCalendarSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.VISIT_CALENDAR);

  // 操作エリア（1〜2行目）
  sheet.getRange('A1').setValue('対象年:').setFontWeight('bold');
  sheet.getRange('A2').setValue('対象月:').setFontWeight('bold');

  // 年ドロップダウン（B1）
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateYearOptions(), true).build();
  sheet.getRange('B1').setDataValidation(yearRule);
  var years = collectYearsFromFormResponses_();
  sheet.getRange('B1').setValue(years[0] + '年');

  // 月ドロップダウン（B2）- デフォルトは「すべて」（=年別カレンダー）
  var monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateMonthOptions(), true).build();
  sheet.getRange('B2').setDataValidation(monthRule);
  sheet.getRange('B2').setValue('すべて');

  Logger.log('来館カレンダーシートのセットアップ完了');
}

/**
 * 児童別ビューシートを作成・設定する
 * レイアウト: 1行目=児童名, 2行目=対象年, 3行目=対象月, 4〜8行目=基本情報, 9行目=履歴ヘッダー
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupChildViewSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CHILD_VIEW);

  // 操作エリア（1〜3行目）
  sheet.getRange('A1').setValue('児童名:');
  sheet.getRange('A2').setValue('対象年:');
  sheet.getRange('A3').setValue('対象月:');

  // 児童名ドロップダウン（B1）
  var childNames = getAllChildNameOptions();
  if (childNames.length > 0) {
    var childNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(childNames, true).build();
    sheet.getRange('B1').setDataValidation(childNameRule);
  }

  // 年ドロップダウン（B2）- デフォルトは「すべて」
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateYearOptions(), true).build();
  sheet.getRange('B2').setDataValidation(yearRule);
  sheet.getRange('B2').setValue('すべて');

  // 月ドロップダウン（B3）- デフォルトは「すべて」
  var monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(generateMonthOptions(), true).build();
  sheet.getRange('B3').setDataValidation(monthRule);
  sheet.getRange('B3').setValue('すべて');

  // 基本情報エリアのラベル（4〜8行目）
  sheet.getRange('A4').setValue('保護者名:');
  sheet.getRange('A5').setValue('担当スタッフ:');
  sheet.getRange('A6').setValue('月間利用枠:');
  sheet.getRange('A7').setValue('医療型の有無:');
  sheet.getRange('A8').setValue('来館回数 / 残枠 / 利用率:');

  // ラベル列を太字に
  sheet.getRange('A1:A8').setFontWeight('bold');

  // 基本情報エリアの罫線（A4:B8）
  var infoRange = sheet.getRange('A4:B8');
  infoRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // 来館履歴ヘッダー（9行目）
  var historyHeaders = ['記録日', 'スタッフ名', '入所時間', '退所時間', '体温', '食事', '入浴', '睡眠', '便', '服薬', 'その他連絡事項'];
  sheet.getRange(9, 1, 1, historyHeaders.length).setValues([historyHeaders]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(9, 1, 1, historyHeaders.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setBorder(true, true, true, true, true, true, '#333333', SpreadsheetApp.BorderStyle.SOLID);

  // ヘッダーテキストを中央寄せ
  headerRange.setHorizontalAlignment('center');

  // 行固定
  sheet.setFrozenRows(9);

  // 列幅調整（A4横向き 約1050px に収まるよう最適化）
  sheet.setColumnWidth(1, 90);   // 記録日
  sheet.setColumnWidth(2, 90);   // スタッフ名
  sheet.setColumnWidth(3, 65);   // 入所時間
  sheet.setColumnWidth(4, 65);   // 退所時間
  sheet.setColumnWidth(5, 50);   // 体温
  sheet.setColumnWidth(6, 45);   // 食事
  sheet.setColumnWidth(7, 45);   // 入浴
  sheet.setColumnWidth(8, 45);   // 睡眠
  sheet.setColumnWidth(9, 45);   // 便
  sheet.setColumnWidth(10, 45);  // 服薬
  sheet.setColumnWidth(11, 180); // その他連絡事項

  // フォントサイズを印刷向けに統一
  sheet.getRange('A1:K100').setFontSize(10);
  sheet.getRange('A1:A2').setFontWeight('bold');

  Logger.log('児童別ビューシートのセットアップ完了');
}

/**
 * ログシートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupLogSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.LOG);

  // ヘッダー行
  var headers = ['タイムスタンプ', '関数名', 'エラーメッセージ', 'スタックトレース'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー書式設定
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');

  // 行固定
  sheet.setFrozenRows(1);

  // 列幅調整
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 400);

  Logger.log('ログシートのセットアップ完了');
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
