/**
 * F-01: シート初期セットアップ
 * 月別集計・確定来館記録・来館カレンダー・児童別ビューのシート作成、ヘッダー・書式設定
 */

/**
 * 初期セットアップ（シート作成＋トリガー作成）
 * 既存シート・既存トリガーは削除せず、不足分だけ追加する。
 * 月別集計・来館カレンダーなどのデータ更新は「月次一括処理」側で実行する。
 */
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. シート作成・初期化
  setupMonthlySummarySheet_(ss);
  setupConfirmedVisitsSheet_(ss);
  setupVisitCalendarSheet_(ss);
  setupChildViewSheet_(ss);
  setupLogSheet_(ss);
  setupNotesMasterSheet_(ss);
  setupSettingsSheet_(ss);
  setupChildMasterValidations_();

  // 2. トリガー作成
  setupFormSyncTrigger();
  setupMonthlyProcessTrigger();
  setupEmailTrigger();
  setupBounceCheckTrigger();

  Logger.log('初期セットアップ完了（シート・トリガー）');
  SpreadsheetApp.getUi().alert(
    '初期セットアップが完了しました\n\n' +
    '【作成・初期化したシート】\n' +
    '・月別集計 / 来館カレンダー / 確定来館記録 / 児童別ビュー / ログ / 定型文マスタ / 設定\n\n' +
    '【設定したトリガー】\n' +
    '・毎日 AM1時: フォーム児童名・スタッフ同期\n' +
    '・毎月1日 AM3時: 月次一括処理\n' +
    '・毎朝 AM8時: 保護者メール送信\n' +
    '・毎日 AM9時: バウンスメール確認\n\n' +
    '次は「月次一括処理」からデータ集計を実行してください。'
  );
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

  // ヘッダー行（3行目）- CONFIRMED_COL と完全一致させる
  var headers = ['記録日', '児童名', 'データ区分', 'スタッフ1', 'スタッフ2', '入所日時', '退所日時', '体温', '夕食', '朝食', '昼食', '入浴', '睡眠', '便', '服薬(夜)', '服薬(朝)', 'その他連絡事項'];
  var headerRange = sheet.getRange(CONFIRMED_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  // 行固定（操作エリア+ヘッダー）
  sheet.setFrozenRows(CONFIRMED_HEADER_ROW);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 記録日
  sheet.setColumnWidth(2, 100);  // 児童名
  sheet.setColumnWidth(3, 80);   // データ区分
  sheet.setColumnWidth(4, 100);  // スタッフ1
  sheet.setColumnWidth(5, 100);  // スタッフ2
  sheet.setColumnWidth(6, 130);  // 入所日時
  sheet.setColumnWidth(7, 130);  // 退所日時
  sheet.setColumnWidth(8, 60);   // 体温
  sheet.setColumnWidth(9, 50);   // 夕食
  sheet.setColumnWidth(10, 50);  // 朝食
  sheet.setColumnWidth(11, 50);  // 昼食
  sheet.setColumnWidth(12, 50);  // 入浴
  sheet.setColumnWidth(13, 50);  // 睡眠
  sheet.setColumnWidth(14, 50);  // 便
  sheet.setColumnWidth(15, 60);  // 服薬(夜)
  sheet.setColumnWidth(16, 60);  // 服薬(朝)
  sheet.setColumnWidth(17, 200); // その他連絡事項

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

  // 来館履歴ヘッダー（9行目）- child-view.gs の履歴出力と一致させる
  var historyHeaders = ['記録日', 'スタッフ名', '入所時間', '退所時間', '体温', '夕食', '朝食', '昼食', '入浴', '睡眠', '便', '服薬(夜)', '服薬(朝)', 'その他連絡事項'];
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
  sheet.setColumnWidth(1, 85);   // 記録日
  sheet.setColumnWidth(2, 85);   // スタッフ名
  sheet.setColumnWidth(3, 60);   // 入所時間
  sheet.setColumnWidth(4, 60);   // 退所時間
  sheet.setColumnWidth(5, 45);   // 体温
  sheet.setColumnWidth(6, 40);   // 夕食
  sheet.setColumnWidth(7, 40);   // 朝食
  sheet.setColumnWidth(8, 40);   // 昼食
  sheet.setColumnWidth(9, 40);   // 入浴
  sheet.setColumnWidth(10, 40);  // 睡眠
  sheet.setColumnWidth(11, 40);  // 便
  sheet.setColumnWidth(12, 55);  // 服薬(夜)
  sheet.setColumnWidth(13, 55);  // 服薬(朝)
  sheet.setColumnWidth(14, 160); // その他連絡事項

  // フォントサイズを印刷向けに統一
  sheet.getRange('A1:N100').setFontSize(10);
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
 * 定型文マスタシートを作成・設定する
 * 振り分け時「その他連絡事項」のフォールバック用
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupNotesMasterSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.NOTES_MASTER);

  var header = ['定型文（その他連絡事項）'];
  var headerRange = sheet.getRange(1, 1, 1, 1);
  headerRange.setValues([header]);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 350);

  Logger.log('定型文マスタシートのセットアップ完了');
}

/**
 * 設定シートを作成・初期化する
 * SETTINGS_ROW の行順に沿って「設定項目 / デフォルト値 / 備考」を配置する。
 * 既存シートがある場合は上書きせず、シートが空の時だけ初期値を書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupSettingsSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.SETTINGS);

  // ヘッダー（1行目）
  var headers = ['設定項目', 'デフォルト値', '備考'];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 列幅
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 360);
  sheet.setColumnWidth(3, 240);

  // 設定項目と初期値の定義（SETTINGS_ROW と1:1で対応）
  var rows = [
    { row: SETTINGS_ROW.MAX_VISITS_PER_DAY, label: '1日最大来館数', value: DEFAULT_MAX_VISITS_PER_DAY, note: '1日あたりの最大来館数' },
    { row: SETTINGS_ROW.CHECK_IN,            label: '入所時間',     value: ALLOCATION_DEFAULTS.CHECK_IN,     note: 'HH:mm 形式' },
    { row: SETTINGS_ROW.CHECK_OUT,           label: '退所時間',     value: ALLOCATION_DEFAULTS.CHECK_OUT,    note: 'HH:mm 形式' },
    { row: SETTINGS_ROW.BUSINESS_DAYS,       label: '営業日',       value: '月曜日 火曜日 水曜日 木曜日 金曜日 土曜日 日曜日', note: '曜日をスペース・カンマ区切り' },
    { row: SETTINGS_ROW.TEMPERATURE,         label: '体温',         value: ALLOCATION_DEFAULTS.TEMPERATURE,  note: '' },
    { row: SETTINGS_ROW.MEAL_DINNER,         label: '夕食',         value: ALLOCATION_DEFAULTS.MEAL_DINNER,  note: '' },
    { row: SETTINGS_ROW.MEAL_BREAKFAST,      label: '朝食',         value: ALLOCATION_DEFAULTS.MEAL_BREAKFAST, note: '' },
    { row: SETTINGS_ROW.MEAL_LUNCH,          label: '昼食',         value: ALLOCATION_DEFAULTS.MEAL_LUNCH,   note: '他サービス併給時は − 推奨' },
    { row: SETTINGS_ROW.BATH,                label: '入浴',         value: ALLOCATION_DEFAULTS.BATH,         note: '' },
    { row: SETTINGS_ROW.SLEEP,               label: '睡眠',         value: ALLOCATION_DEFAULTS.SLEEP,        note: '' },
    { row: SETTINGS_ROW.BOWEL,               label: '便',           value: ALLOCATION_DEFAULTS.BOWEL,        note: '' },
    { row: SETTINGS_ROW.MEDICINE_MORNING,    label: '服薬（朝）',   value: ALLOCATION_DEFAULTS.MEDICINE_MORNING, note: '' },
    { row: SETTINGS_ROW.MEDICINE_NIGHT,      label: '服薬（夜）',   value: ALLOCATION_DEFAULTS.MEDICINE_NIGHT,   note: '' },
    { row: SETTINGS_ROW.NOTES,               label: '連絡事項',     value: ALLOCATION_DEFAULTS.NOTES,        note: '複数候補は「定型文マスタ」シートに1行1件で登録' },
    { row: SETTINGS_ROW.DUMMY_STAFF_NAME,    label: '固定スタッフ', value: '溝口母',                         note: '振り分け・スタッフ2補完用（7人満枠日に補完）' },
    { row: SETTINGS_ROW.ERROR_EMAIL,         label: 'エラー通知先メール', value: '',                         note: '複数はカンマ区切り' },
    { row: SETTINGS_ROW.EMAIL_SUBJECT,       label: 'メール件名',   value: DEFAULT_EMAIL_SUBJECT,            note: '保護者向け来館報告メールの件名' },
    { row: SETTINGS_ROW.EMAIL_BODY,          label: 'メール本文',   value: DEFAULT_EMAIL_TEMPLATE,           note: '{保護者名}{日付}{児童名}{入所時間}{退所時間}{体温}{夕食}{朝食}{昼食}{入浴}{睡眠}{便}{服薬(夜)}{服薬(朝)}{連絡事項} が置換される' },
  ];

  // A列（項目名）は未記入行のみ書き込む。B列（値）は空の場合のみ初期値を書き込む。
  // ユーザーが変更済みの値は保持する。
  rows.forEach(function(r) {
    var labelCell = sheet.getRange(r.row, 1);
    if (!labelCell.getValue()) labelCell.setValue(r.label);
    var valueCell = sheet.getRange(r.row, 2);
    if (valueCell.getValue() === '' || valueCell.getValue() === null) {
      valueCell.setValue(r.value);
    }
    var noteCell = sheet.getRange(r.row, 3);
    if (!noteCell.getValue() && r.note) noteCell.setValue(r.note);
  });

  // メール本文セルは折り返し表示・上揃えにし、行高さを確保
  sheet.getRange(SETTINGS_ROW.EMAIL_BODY, 2).setWrap(true).setVerticalAlignment('top');
  sheet.setRowHeight(SETTINGS_ROW.EMAIL_BODY, 300);

  Logger.log('設定シートのセットアップ完了');
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
