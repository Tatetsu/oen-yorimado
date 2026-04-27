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
  setupAllowedUsersSheet_(ss);
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
    '・月別集計 / 来館カレンダー / 確定来館記録 / 児童別ビュー / ログ / 定型文マスタ / 設定 / 許可ユーザー\n\n' +
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
  setListValidation_(
    sheet.getRange(2, MASTER_COL.PRIORITY, lastRow - 1, 1),
    ['区分1', '区分2', '区分3', '区分4', '区分5']
  );

  Logger.log('児童マスタのドロップダウン設定完了（重度支援区分 区分1〜区分5）');
}


/**
 * 月別集計シートを作成・設定する
 * レイアウト: 1行目=対象年, 2行目=対象月, 3行目=ヘッダー, 4行目〜=データ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupMonthlySummarySheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.MONTHLY_SUMMARY);
  var pristine = isPristineSheet_(sheet);

  // 操作エリア（1〜2行目）- ラベル文言は常に維持、太字は初回のみ
  sheet.getRange('A1').setValue('対象年:');
  sheet.getRange('A2').setValue('対象月:');
  if (pristine) {
    sheet.getRange('A1:A2').setFontWeight('bold');
  }

  // 年ドロップダウン（B1）
  setListValidation_(sheet.getRange('B1'), generateYearOptions());
  if (pristine) {
    var years = collectYearsFromFormResponses_();
    sheet.getRange('B1').setValue(years[0] + '年');
  }

  // 月ドロップダウン（B2）- デフォルトは「すべて」
  setListValidation_(sheet.getRange('B2'), generateMonthOptions());
  if (pristine) {
    sheet.getRange('B2').setValue('すべて');
  }

  // ヘッダー行（3行目）
  var headers = ['No.', '児童名', '利用枠', '来館数', '残数', '利用率'];
  var headerRange = sheet.getRange(SUMMARY_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  if (pristine) {
    styleSheetHeader_(headerRange, SUMMARY_HEADER_ROW);
    setColumnWidths_(sheet, { 1: 40, 2: 100, 3: 90, 4: 80, 5: 60, 6: 70 });
  }
  Logger.log('月別集計シートのセットアップ完了');
}

/**
 * 確定来館記録シートを作成・設定する
 * レイアウト: 1行目=対象年, 2行目=対象月, 3行目=ヘッダー, 4行目〜=データ
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupConfirmedVisitsSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CONFIRMED_VISITS);
  var pristine = isPristineSheet_(sheet);

  // 操作エリア（1〜2行目）- ラベル文言は維持、太字は初回のみ
  sheet.getRange('A1').setValue('対象年:');
  sheet.getRange('A2').setValue('対象月:');
  if (pristine) {
    sheet.getRange('A1:A2').setFontWeight('bold');
  }

  // 年ドロップダウン（B1）- デフォルトは前月の年
  setListValidation_(sheet.getRange('B1'), generateYearOptions());

  // 月ドロップダウン（B2）- デフォルトは前月
  setListValidation_(sheet.getRange('B2'), generateMonthOptions());

  // デフォルト値：現在日付の1ヶ月前（初回のみ。再実行時にユーザー選択を上書きしない）
  if (pristine) {
    var lastMonth = new Date();
    lastMonth.setDate(1);
    lastMonth.setMonth(lastMonth.getMonth() - 1);
    sheet.getRange('B1').setValue(lastMonth.getFullYear() + '年');
    sheet.getRange('B2').setValue((lastMonth.getMonth() + 1) + '月');
  }

  // ヘッダー行（3行目）- CONFIRMED_COL と完全一致させる
  // 8〜9列目「往」「復」は記録日が入所日/退所予定日と一致するときに 1 が立つ集計用列
  // 21列目「宿泊PK」は児童名+入所日時のユニークキーで、デフォルト非表示
  var headers = ['データ区分', '記録日', 'スタッフ1', 'スタッフ2', '児童名', '入所日時', '退所予定日時', '往', '復', '体温', '夕食', '朝食', '昼食', '入浴', '睡眠', '便', '服薬(夜)', '服薬(朝)', 'その他連絡事項', '連泊', '宿泊PK'];
  var headerRange = sheet.getRange(CONFIRMED_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  if (pristine) {
    styleSheetHeader_(headerRange, CONFIRMED_HEADER_ROW);
    setColumnWidths_(sheet, {
      1: 80,   // データ区分
      2: 100,  // 記録日
      3: 100,  // スタッフ1
      4: 100,  // スタッフ2
      5: 100,  // 児童名
      6: 130,  // 入所日時
      7: 130,  // 退所予定日時
      8: 40,   // 往
      9: 40,   // 復
      10: 60,  // 体温
      11: 50,  // 夕食
      12: 50,  // 朝食
      13: 50,  // 昼食
      14: 50,  // 入浴
      15: 50,  // 睡眠
      16: 50,  // 便
      17: 60,  // 服薬(夜)
      18: 60,  // 服薬(朝)
      19: 200, // その他連絡事項
      20: 60,  // 連泊
      21: 200, // 宿泊PK
    });
  }

  // 宿泊PK列はデフォルト非表示（運用上は不要、デバッグ・整合性確認用）
  sheet.hideColumns(CONFIRMED_COL.STAY_PK);

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
  var pristine = isPristineSheet_(sheet);

  // 操作エリア（1〜2行目）- ラベル文言は維持、太字は初回のみ
  sheet.getRange('A1').setValue('対象年:');
  sheet.getRange('A2').setValue('対象月:');
  if (pristine) {
    sheet.getRange('A1:A2').setFontWeight('bold');
  }

  // 年ドロップダウン（B1）
  setListValidation_(sheet.getRange('B1'), generateYearOptions());
  if (pristine) {
    var years = collectYearsFromFormResponses_();
    sheet.getRange('B1').setValue(years[0] + '年');
  }

  // 月ドロップダウン（B2）- デフォルトは「すべて」（=年別カレンダー）
  setListValidation_(sheet.getRange('B2'), generateMonthOptions());
  if (pristine) {
    sheet.getRange('B2').setValue('すべて');
  }

  Logger.log('来館カレンダーシートのセットアップ完了');
}

/**
 * 児童別ビューシートを作成・設定する
 * レイアウト: 1行目=児童名, 2行目=対象年, 3行目=対象月, 4〜8行目=基本情報, 9行目=履歴ヘッダー
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupChildViewSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.CHILD_VIEW);
  var pristine = isPristineSheet_(sheet);

  // 操作エリア（1〜3行目）- ラベル文言は維持
  sheet.getRange('A1').setValue('児童名:');
  sheet.getRange('A2').setValue('対象年:');
  sheet.getRange('A3').setValue('対象月:');

  // 児童名ドロップダウン（B1）
  var childNames = getAllChildNameOptions();
  if (childNames.length > 0) {
    setListValidation_(sheet.getRange('B1'), childNames);
  }

  // 年ドロップダウン（B2）- デフォルトは「すべて」
  setListValidation_(sheet.getRange('B2'), generateYearOptions());
  if (pristine) sheet.getRange('B2').setValue('すべて');

  // 月ドロップダウン（B3）- デフォルトは「すべて」
  setListValidation_(sheet.getRange('B3'), generateMonthOptions());
  if (pristine) sheet.getRange('B3').setValue('すべて');

  // 基本情報エリアのラベル（4〜8行目）
  sheet.getRange('A4').setValue('保護者名:');
  sheet.getRange('A5').setValue('担当スタッフ:');
  sheet.getRange('A6').setValue('月間利用枠:');
  sheet.getRange('A7').setValue('医療型の有無:');
  sheet.getRange('A8').setValue('来館回数 / 残枠 / 利用率:');

  // 来館履歴ヘッダー（9行目）- child-view.gs の履歴出力と一致させる
  var historyHeaders = ['記録日', 'スタッフ名', '入所時間', '退所時間', '往', '復', '体温', '夕食', '朝食', '昼食', '入浴', '睡眠', '便', '服薬(夜)', '服薬(朝)', 'その他連絡事項'];
  sheet.getRange(9, 1, 1, historyHeaders.length).setValues([historyHeaders]);

  if (pristine) {
    // ラベル列を太字に
    sheet.getRange('A1:A8').setFontWeight('bold');

    // 基本情報エリアの罫線（A4:B8）
    sheet.getRange('A4:B8').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    // ヘッダー書式設定（配置中央寄せ・行固定）
    var headerRange = sheet.getRange(9, 1, 1, historyHeaders.length);
    styleSheetHeader_(headerRange, 9, { horizontalAlignment: 'center' });
    headerRange.setBorder(true, true, true, true, true, true, '#333333', SpreadsheetApp.BorderStyle.SOLID);

    // 列幅調整（A4横向き 約1050px に収まるよう最適化）
    setColumnWidths_(sheet, {
      1: 85,   // 記録日
      2: 85,   // スタッフ名
      3: 60,   // 入所時間
      4: 60,   // 退所時間
      5: 35,   // 往
      6: 35,   // 復
      7: 45,   // 体温
      8: 40,   // 夕食
      9: 40,   // 朝食
      10: 40,  // 昼食
      11: 40,  // 入浴
      12: 40,  // 睡眠
      13: 40,  // 便
      14: 55,  // 服薬(夜)
      15: 55,  // 服薬(朝)
      16: 140, // その他連絡事項
    });

    // フォントサイズを印刷向けに統一
    sheet.getRange('A1:P100').setFontSize(10);
  }

  Logger.log('児童別ビューシートのセットアップ完了');
}

/**
 * ログシートを作成・設定する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupLogSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.LOG);
  var pristine = isPristineSheet_(sheet);

  // ヘッダー行
  var headers = ['タイムスタンプ', '関数名', 'エラーメッセージ', 'スタックトレース'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (pristine) {
    styleSheetHeader_(sheet.getRange(1, 1, 1, headers.length), 1);
    setColumnWidths_(sheet, { 1: 160, 2: 200, 3: 400, 4: 400 });
  }

  Logger.log('ログシートのセットアップ完了');
}

/**
 * 定型文マスタシートを作成・設定する
 * 振り分け時「その他連絡事項」のフォールバック用
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupNotesMasterSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.NOTES_MASTER);
  var pristine = isPristineSheet_(sheet);

  var header = ['定型文（その他連絡事項）'];
  var headerRange = sheet.getRange(1, 1, 1, 1);
  headerRange.setValues([header]);
  if (pristine) {
    styleSheetHeader_(headerRange, 1);
    setColumnWidths_(sheet, { 1: 350 });
  }

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
  var pristine = isPristineSheet_(sheet);

  // ヘッダー（1行目）
  var headers = ['設定項目', 'デフォルト値', '備考'];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  if (pristine) {
    styleSheetHeader_(headerRange, 1);
    setColumnWidths_(sheet, { 1: 180, 2: 360, 3: 240 });
  }

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
    { row: SETTINGS_ROW.DUMMY_STAFF_NAME,    label: '固定スタッフ', value: '溝口照代',                       note: '振り分け・スタッフ2補完用（7人満枠日に補完）' },
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

  // メール本文セルは折り返し表示・上揃えにし、行高さを確保（初回のみ）
  if (pristine) {
    sheet.getRange(SETTINGS_ROW.EMAIL_BODY, 2).setWrap(true).setVerticalAlignment('top');
    sheet.setRowHeight(SETTINGS_ROW.EMAIL_BODY, 300);
  }

  Logger.log('設定シートのセットアップ完了');
}


/**
 * 許可ユーザーシートを作成・初期化する
 * WebView のアクセストークン検証用。
 * レイアウト: Row 1〜5=マニュアル/編集画面URL(B2)/TOKEN/ランダム文字列等の管理者編集領域 /
 *             Row 6=ヘッダー / Row 7〜=データ
 * D列(URL)はARRAYFORMULAで自動生成される（B2のURLとC列のトークンを結合）
 * Row 1〜5 はユーザー管理領域のため触らない。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupAllowedUsersSheet_(ss) {
  var sheet = getOrCreateSheet_(ss, SHEET_NAMES.ALLOWED_USERS);
  var pristine = isPristineSheet_(sheet);

  // Row 6: ヘッダー
  var headers = ['メールアドレス', '氏名', 'トークン', 'URL', '有効', '備考'];
  var headerRange = sheet.getRange(ALLOWED_USERS_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  if (pristine) {
    styleSheetHeader_(headerRange, ALLOWED_USERS_HEADER_ROW);
    var allowedUsersWidths = {};
    allowedUsersWidths[ALLOWED_USERS_COL.EMAIL] = 240;
    allowedUsersWidths[ALLOWED_USERS_COL.NAME] = 140;
    allowedUsersWidths[ALLOWED_USERS_COL.TOKEN] = 200;
    allowedUsersWidths[ALLOWED_USERS_COL.URL] = 360;
    allowedUsersWidths[ALLOWED_USERS_COL.ACTIVE] = 80;
    allowedUsersWidths[ALLOWED_USERS_COL.NOTE] = 240;
    setColumnWidths_(sheet, allowedUsersWidths);
  }

  // D列(URL): ARRAYFORMULA をデータ開始行に投入（既に何か入っていれば上書きしない）
  var dStart = sheet.getRange(ALLOWED_USERS_DATA_START_ROW, ALLOWED_USERS_COL.URL);
  if (!dStart.getFormula() && !dStart.getValue()) {
    var startRow = ALLOWED_USERS_DATA_START_ROW;
    var formula = '=ARRAYFORMULA(IF(($' + ALLOWED_USERS_BASE_URL_CELL + '<>"")*(A' + startRow + ':A<>"")*(C' + startRow + ':C<>""), $' + ALLOWED_USERS_BASE_URL_CELL + '&"?t="&C' + startRow + ':C, ""))';
    dStart.setFormula(formula);
  }

  // 「有効」列にチェックボックス（データ開始行以降）
  var maxRow = Math.max(sheet.getMaxRows(), 100);
  var checkboxRows = maxRow - ALLOWED_USERS_DATA_START_ROW + 1;
  if (checkboxRows > 0) {
    sheet.getRange(ALLOWED_USERS_DATA_START_ROW, ALLOWED_USERS_COL.ACTIVE, checkboxRows, 1)
      .insertCheckboxes();
  }

  Logger.log('許可ユーザーシートのセットアップ完了');
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

/**
 * シートが「初回扱い」（新規作成 or 中身が空）かを判定する
 * 書式の初回適用ガードに使用する。これが false の時は背景色・列幅・罫線等の
 * 装飾系呼び出しをスキップして、ユーザーの手動編集を尊重する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {boolean}
 */
function isPristineSheet_(sheet) {
  // getLastRow/getLastColumn は値が入っていない（書式のみ）の場合 0 を返す
  return sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;
}
