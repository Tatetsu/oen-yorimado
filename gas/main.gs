/**
 * エントリポイント・トリガー管理・ボタン実行関数
 */

/**
 * スプレッドシートを開いた時にカスタムメニューを追加する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('来館管理')
    .addItem('初期セットアップ', 'setupAllSheets')
    .addItem('月次一括処理', 'runMonthlyProcess')
    .addSeparator()
    .addItem('確定来館記録を手動更新', 'updateConfirmedVisitsAndCalendar')
    .addItem('来館報告メール手動送信', 'sendVisitReportsManual')
    .addItem('ドロップダウンを更新', 'refreshDropdowns')
    .addItem('児童マスタ ドロップダウン設定', 'refreshChildMasterValidations')
    .addSeparator()
    .addItem('バウンスメールを確認', 'collectBounceEmailsManual')
    .addToUi();
}

/**
 * 月次一括処理: 月別集計 → 振り分け → 確定来館記録 → 来館カレンダーを一括実行する
 * 選択肢は前月までに制限（未来月は非表示）
 */
function runMonthlyProcess() {
  var ui = SpreadsheetApp.getUi();

  try {
    var options = generateProcessableMonthOptions_();
    if (options.length === 0) {
      ui.alert('処理可能な月がありません（前月以前のデータがありません）');
      return;
    }
    var year = parseYearMonth(options[0]).year;

    // 現在の月別集計 B1(年)+B2(月) をデフォルト表示
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    var curY = summarySheet.getRange('B1').getDisplayValue();
    var curM = summarySheet.getRange('B2').getDisplayValue();
    var currentLabel = (parseYearOption_(curY) && parseMonthOption_(curM))
      ? curY + curM
      : options[options.length - 1];

    var prompt = '対象年月を選択してください（番号を入力）:\n\n';
    prompt += '0. ' + year + '年（年次一括処理）\n';
    for (var i = 0; i < options.length; i++) {
      var marker = (options[i] === currentLabel) ? ' ← 現在' : '';
      prompt += (i + 1) + '. ' + options[i] + marker + '\n';
    }

    var response = ui.prompt('月次一括処理', prompt, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;

    var inputNum = parseInt(response.getResponseText().trim(), 10);
    if (inputNum === 0) {
      runAnnualProcess();
      return;
    }

    var selectedIndex = inputNum - 1;
    if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= options.length) {
      ui.alert('無効な番号です。0〜' + options.length + 'の番号を入力してください。');
      return;
    }

    var ym = parseYearMonth(options[selectedIndex]);

    // 月別集計のドロップダウンを反映
    summarySheet.getRange('B1').setValue(ym.year + '年');
    summarySheet.getRange('B2').setValue(ym.month + '月');

    // 1. 振り分け実行（内部で確定来館記録・月別集計も更新される）
    allocateRemainingPoints_(ym.year, ym.month);

    // 2. 来館カレンダーを年単位で更新
    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setValue(ym.year + '年');
    updateVisitCalendarByYear_(calendarSheet, ym.year);

    ui.alert(
      ym.year + '年' + ym.month + '月の月次処理が完了しました\n' +
      '・確定来館記録（振り分け含む）\n・月別集計\n・来館カレンダー'
    );
  } catch (error) {
    logError_('runMonthlyProcess', error);
    ui.alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * 月次一括処理（トリガー用）: 前日の属する月を対象として全処理を自動実行する
 * 月初（例: 4/1）に実行すると前日（3/31）の月 = 前月が対象になる
 */
function runMonthlyProcessAutomatic() {
  var now = new Date();
  var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  var year = yesterday.getFullYear();
  var month = yesterday.getMonth() + 1;

  try {
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    summarySheet.getRange('B1').setValue(year + '年');
    summarySheet.getRange('B2').setValue(month + '月');

    allocateRemainingPoints_(year, month);

    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setValue(year + '年');
    updateVisitCalendarByYear_(calendarSheet, year);

    Logger.log('月次自動処理完了: ' + year + '年' + month + '月');
  } catch (error) {
    logError_('runMonthlyProcessAutomatic', error);
  }
}

/**
 * セル編集時のトリガー関数
 * 特定シートのB1セルが変更されたときに対応する更新処理を自動実行する
 * @param {Object} e 編集イベントオブジェクト
 */
function onEdit(e) {
  try {
    var sheetName = e.range.getSheet().getName();
    var cell = e.range.getA1Notation();

    if (sheetName === SHEET_NAMES.CHILD_VIEW && (cell === 'B1' || cell === 'B2' || cell === 'B3')) {
      updateChildView();
    } else if (sheetName === SHEET_NAMES.VISIT_CALENDAR && cell === 'B1') {
      updateVisitCalendar();
    } else if (sheetName === SHEET_NAMES.MONTHLY_SUMMARY && (cell === 'B1' || cell === 'B2')) {
      updateMonthlySummary();
    } else if (sheetName === SHEET_NAMES.CONFIRMED_VISITS && (cell === 'B1' || cell === 'B2')) {
      filterConfirmedVisits_();
    }
  } catch (error) {
    logError_('onEdit', error);
  }
}

/**
 * 児童別ビュー・月別集計のドロップダウン、およびGoogleフォームの児童名プルダウンを
 * 最新の児童マスタで更新する
 */
function refreshDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    // 児童名リスト（在籍中のみ・フォーム同期用）
    var childNames = getChildNameOptions();
    if (childNames.length === 0) {
      Logger.log('入所中の児童がいません');
      return;
    }

    // ビューシートのドロップダウン設定（シート名・セル・選択肢）
    var yearOpts = generateYearOptions();
    var monthOpts = generateMonthOptions();
    var dropdownConfigs = [
      { sheet: SHEET_NAMES.CHILD_VIEW,       cell: 'B1', options: getAllChildNameOptions() },
      { sheet: SHEET_NAMES.CHILD_VIEW,       cell: 'B2', options: yearOpts },
      { sheet: SHEET_NAMES.CHILD_VIEW,       cell: 'B3', options: monthOpts },
      { sheet: SHEET_NAMES.MONTHLY_SUMMARY,  cell: 'B1', options: yearOpts },
      { sheet: SHEET_NAMES.MONTHLY_SUMMARY,  cell: 'B2', options: monthOpts },
      { sheet: SHEET_NAMES.VISIT_CALENDAR,   cell: 'B1', options: yearOpts },
      { sheet: SHEET_NAMES.CONFIRMED_VISITS, cell: 'B1', options: yearOpts },
      { sheet: SHEET_NAMES.CONFIRMED_VISITS, cell: 'B2', options: monthOpts },
    ];

    dropdownConfigs.forEach(function(cfg) {
      var sheet = getSheet(cfg.sheet);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(cfg.options, true)
        .build();
      sheet.getRange(cfg.cell).setDataValidation(rule);
    });

    // Googleフォームの児童名プルダウン更新
    updateFormChildNameDropdown_(childNames);

    Logger.log('ドロップダウンを更新しました');
  } catch (error) {
    logError_('refreshDropdowns', error);
  } finally {
    originalSheet.activate();
  }
}

/**
 * Googleフォームの児童名プルダウンを更新する
 * スクリプトプロパティ FORM_ID と FORM_CHILD_NAME_QUESTION にフォームIDと質問タイトルを設定すること
 * @param {Array<string>} childNames 児童名の配列
 */
function updateFormChildNameDropdown_(childNames) {
  var props = PropertiesService.getScriptProperties();
  var formId = props.getProperty('FORM_ID');
  var questionTitle = props.getProperty('FORM_CHILD_NAME_QUESTION') || '児童名';

  if (!formId) {
    Logger.log('フォームプルダウン更新スキップ: スクリプトプロパティ FORM_ID が未設定');
    return;
  }

  var form = FormApp.openById(formId);
  var items = form.getItems();

  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === questionTitle) {
      items[i].asListItem().setChoiceValues(childNames);
      Logger.log('フォームプルダウン更新完了: ' + childNames.length + '件の選択肢を設定');
      return;
    }
  }

  Logger.log('フォームの質問「' + questionTitle + '」が見つかりませんでした');
}

/**
 * 年次一括処理: 対象年の1月〜前月までを一括で処理する
 * 振り分け → 確定来館記録 → 月別集計を対象月数分実行し、最後に年別カレンダーを更新する
 * 前月より後の月は処理しない。さらに、既に書き込まれていた振り分け行があれば自動クリアする
 */
function runAnnualProcess() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
  var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);

  // 対象年は前月年以下に制限し、前月年なら1月〜前月、それ以前なら1月〜12月を処理
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastY = lastMonth.getFullYear();
  var lastM = lastMonth.getMonth() + 1;
  var formYear = getTargetYearFromFormResponses_();
  var year = Math.min(formYear, lastY);
  var maxMonth = (year === lastY) ? lastM : 12;

  if (maxMonth <= 0) {
    ui.alert('処理可能な月がありません（前月以前のデータがありません）');
    return;
  }

  try {
    ss.toast(year + '年の年次一括処理を開始します...', '処理中', -1);

    for (var month = 1; month <= maxMonth; month++) {
      ss.toast(year + '年' + month + '月 を処理中...', '処理中', -1);
      summarySheet.getRange('B1').setValue(year + '年');
      summarySheet.getRange('B2').setValue(month + '月');
      allocateRemainingPoints_(year, month);
    }

    // 前月より後の月に残っていた振り分け行をクリア（未来月の幽霊データ除去）
    for (var fm = maxMonth + 1; fm <= 12; fm++) {
      clearAllocationsForMonth_(year, fm);
    }

    // 来館カレンダーを年全体で更新
    calendarSheet.getRange('B1').setValue(year + '年');
    updateVisitCalendarByYear_(calendarSheet, year);

    // 月別集計の表示を最終処理月に合わせる
    summarySheet.getRange('B1').setValue(year + '年');
    summarySheet.getRange('B2').setValue(maxMonth + '月');
    updateMonthlySummary();

    ss.toast(year + '年の年次一括処理が完了しました', '完了', 5);
    ui.alert(year + '年の年次一括処理が完了しました\n処理範囲: 1月〜' + maxMonth + '月\n・振り分け・確定来館記録\n・来館カレンダー（年別）\n・月別集計');
  } catch (error) {
    logError_('runAnnualProcess', error);
    ui.alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * 月次一括処理の自動実行トリガーを設定する（手動で1回実行）
 * 毎月1日の午前3時に前日（前月末日）の属する月を対象として全処理を実行する
 */
function setupMonthlyProcessTrigger() {
  // 既存のrunMonthlyProcessAutomaticトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runMonthlyProcessAutomatic') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎月1日の午前3時に実行
  ScriptApp.newTrigger('runMonthlyProcessAutomatic')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  Logger.log('月次一括処理トリガーを設定しました');
  SpreadsheetApp.getUi().alert('月次一括処理トリガーを設定しました（毎月1日 午前3時）');
}

/**
 * 確定来館記録シートの年月フィルタを適用する
 * B1=年, B2=月 の組合せでスコープを決定し、行の表示/非表示を切り替える
 */
function filterConfirmedVisits_() {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var scope = buildScope_(sheet.getRange('B1').getValue(), sheet.getRange('B2').getValue());
  var lastRow = sheet.getLastRow();

  if (lastRow < CONFIRMED_DATA_START_ROW) return;

  var numDataRows = lastRow - CONFIRMED_DATA_START_ROW + 1;

  // まず全データ行を表示
  sheet.showRows(CONFIRMED_DATA_START_ROW, numDataRows);

  if (scope.type === 'all') return;

  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, numDataRows, 1).getValues();

  // 連続する非表示行をバッチで処理（高速化）
  var hideStart = -1;
  var hideCount = 0;

  for (var i = 0; i <= data.length; i++) {
    var shouldHide = false;
    if (i < data.length && data[i][0]) {
      shouldHide = !matchesScope_(new Date(data[i][0]), scope);
    }

    if (shouldHide) {
      if (hideStart === -1) hideStart = i;
      hideCount++;
    } else {
      if (hideCount > 0) {
        sheet.hideRows(CONFIRMED_DATA_START_ROW + hideStart, hideCount);
        hideStart = -1;
        hideCount = 0;
      }
    }
  }
}

/**
 * 確定来館記録を手動更新し、来館カレンダーも連動更新する
 * メニュー「確定来館記録を手動更新」から実行される
 */
function updateConfirmedVisitsAndCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    ss.toast('確定来館記録を更新中...', '処理中', -1);

    updateConfirmedVisits();

    ss.toast('来館カレンダーを更新中...', '処理中', -1);

    updateVisitCalendar();

    ss.toast('確定来館記録と来館カレンダーの更新が完了しました', '完了', 3);
  } catch (error) {
    logError_('updateConfirmedVisitsAndCalendar', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  } finally {
    originalSheet.activate();
  }
}

/**
 * 月次一括処理で選択可能な「YYYY年M月」ラベルを前月まで生成する
 * - 対象年はフォーム回答の最新年と前月の年のうち、前月年以下に制限
 * - 対象年が前月と同じ場合は1月〜前月、それ以前の年なら1月〜12月
 * @returns {Array<string>} 処理可能な年月ラベル（例: ["2026年1月", "2026年2月", "2026年3月"]）
 */
function generateProcessableMonthOptions_() {
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastY = lastMonth.getFullYear();
  var lastM = lastMonth.getMonth() + 1;

  var formYear = getTargetYearFromFormResponses_();
  var targetYear = Math.min(formYear, lastY);
  var maxMonth = (targetYear === lastY) ? lastM : 12;

  var options = [];
  for (var m = 1; m <= maxMonth; m++) {
    options.push(targetYear + '年' + m + '月');
  }
  return options;
}
