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
    .addSeparator()
    .addItem('月別集計を更新', 'updateMonthlySummary')
    .addItem('確定来館記録を更新', 'updateConfirmedVisits')
    .addItem('来館カレンダーを更新', 'updateVisitCalendar')
    .addItem('児童別ビューを更新', 'updateChildView')
    .addSeparator()
    .addItem('余りポイント振り分け実行', 'runAllocationManual')
    .addSeparator()
    .addItem('来館報告メール手動送信', 'sendVisitReportsManual')
    .addItem('メール送信トリガー設定', 'setupEmailTrigger')
    .addSeparator()
    .addItem('ドロップダウンを更新', 'refreshDropdowns')
    .addToUi();
}

/**
 * フォーム送信時のトリガー関数
 * フォームの回答が追加された際に月別集計と来館カレンダーを自動更新する
 * @param {Object} e フォーム送信イベントオブジェクト
 */
function onFormSubmit(e) {
  try {
    updateMonthlySummary();
    updateConfirmedVisits();
    updateVisitCalendar();
    Logger.log('フォーム送信トリガー: 月別集計・確定来館記録・来館カレンダーを更新しました');
  } catch (error) {
    Logger.log('フォーム送信トリガーでエラーが発生しました: ' + error.message);
  }
}

/**
 * フォーム送信トリガーを設定する（初回のみ手動実行）
 */
function setupFormSubmitTrigger() {
  // 既存のonFormSubmitトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新規トリガーを作成
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log('フォーム送信トリガーを設定しました');
  SpreadsheetApp.getUi().alert('フォーム送信トリガーを設定しました');
}

/**
 * 児童別ビュー・月別集計のドロップダウン、およびGoogleフォームの児童名プルダウンを
 * 最新の児童マスタで更新する
 */
function refreshDropdowns() {
  try {
    // 児童名リスト
    var childNames = getChildNameOptions();
    if (childNames.length === 0) {
      Logger.log('入所中の児童がいません');
      return;
    }

    // 年月リスト
    var yearMonthOptions = generateYearMonthOptions();

    // 児童別ビューの児童名ドロップダウン更新
    var childViewSheet = getSheet(SHEET_NAMES.CHILD_VIEW);
    var childNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(childNames, true)
      .build();
    childViewSheet.getRange('B1').setDataValidation(childNameRule);

    // 児童別ビューの年月ドロップダウン更新
    var yearMonthRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(yearMonthOptions, true)
      .build();
    childViewSheet.getRange('B2').setDataValidation(yearMonthRule);

    // 月別集計の年月ドロップダウン更新
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    summarySheet.getRange('B1').setDataValidation(yearMonthRule);

    // 来館カレンダーの年月ドロップダウン更新
    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setDataValidation(yearMonthRule);

    // Googleフォームの児童名プルダウン更新
    updateFormChildNameDropdown_(childNames);

    Logger.log('ドロップダウンを更新しました');
  } catch (error) {
    Logger.log('ドロップダウン更新でエラー: ' + error.message);
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
