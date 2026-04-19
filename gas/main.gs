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
    .addItem('確定来館記録（Webビュー）を開く', 'openWebView')
    .addSeparator()
    .addItem('バウンスメールを確認', 'collectBounceEmailsManual')
    .addToUi();
}

/**
 * 月次一括処理: 月別集計 → 振り分け → 確定来館記録 → 来館カレンダーを一括実行する
 * 月別集計シートのB1セル（対象年月）を参照して全処理を実行する
 */
function runMonthlyProcess() {
  var ui = SpreadsheetApp.getUi();

  try {
    // 年月選択ダイアログを表示
    var options = generateYearMonthOptions();
    if (options.length === 0) {
      ui.alert('選択可能な年月がありません');
      return;
    }

    // 現在の月別集計B1の値をデフォルトとして表示
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    var currentValue = summarySheet.getRange('B1').getDisplayValue() || options[options.length - 1];

    var year = getTargetYearFromFormResponses_();
    var prompt = '対象年月を選択してください（番号を入力）:\n\n';
    prompt += '0. ' + year + '年（年次一括処理・全12ヶ月）\n';
    for (var i = 0; i < options.length; i++) {
      var marker = (options[i] === currentValue) ? ' ← 現在' : '';
      prompt += (i + 1) + '. ' + options[i] + marker + '\n';
    }

    var response = ui.prompt('月次一括処理', prompt, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    var input = response.getResponseText().trim();
    var inputNum = parseInt(input, 10);

    // 0 = 年次一括処理
    if (inputNum === 0) {
      runAnnualProcess();
      return;
    }

    var selectedIndex = inputNum - 1;
    if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= options.length) {
      ui.alert('無効な番号です。0〜' + options.length + 'の番号を入力してください。');
      return;
    }

    var yearMonthStr = options[selectedIndex];
    var ym = parseYearMonth(yearMonthStr);

    // 月別集計B1を選択した年月に更新
    summarySheet.getRange('B1').setValue(yearMonthStr);

    // 1. 振り分け実行（内部で確定来館記録・月別集計も更新される）
    allocateRemainingPoints_(ym.year, ym.month);

    // 2. 来館カレンダーを更新
    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setValue(yearMonthStr);
    updateVisitCalendarByMonth_(calendarSheet, ym.year, ym.month);

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
  var yearMonthStr = year + '年' + month + '月';

  try {
    // 月別集計B1を更新
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    summarySheet.getRange('B1').setValue(yearMonthStr);

    // 1. 振り分け実行（内部で確定来館記録・月別集計も更新される）
    allocateRemainingPoints_(year, month);

    // 2. 来館カレンダーを更新
    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setValue(yearMonthStr);
    updateVisitCalendarByMonth_(calendarSheet, year, month);

    Logger.log('月次自動処理完了: ' + yearMonthStr);
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

    if (sheetName === SHEET_NAMES.CHILD_VIEW && (cell === 'B1' || cell === 'B2')) {
      updateChildView();
    } else if (cell === 'B1') {
      if (sheetName === SHEET_NAMES.VISIT_CALENDAR) {
        updateVisitCalendar();
      } else if (sheetName === SHEET_NAMES.MONTHLY_SUMMARY) {
        updateMonthlySummary();
      } else if (sheetName === SHEET_NAMES.CONFIRMED_VISITS) {
        filterConfirmedVisits_();
      }
    }
  } catch (error) {
    logError_('onEdit', error);
  }
}

/**
 * フォーム送信時のトリガー関数
 * フォームの回答が追加された際に月別集計と来館カレンダーを自動更新する
 * @param {Object} e フォーム送信イベントオブジェクト
 */
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    updateMonthlySummary();
    updateConfirmedVisits();
    updateVisitCalendar();
    Logger.log('フォーム送信トリガー: 月別集計・確定来館記録・来館カレンダーを更新しました');
  } catch (error) {
    logError_('onFormSubmit', error);
  } finally {
    originalSheet.activate();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
    // 児童名リスト
    var childNames = getChildNameOptions();
    if (childNames.length === 0) {
      Logger.log('入所中の児童がいません');
      return;
    }

    // 年月リスト
    var yearMonthOptions = generateYearMonthOptions();

    // 児童別ビューの児童名ドロップダウン更新（全児童: 在籍中→退所済みの順）
    var childViewSheet = getSheet(SHEET_NAMES.CHILD_VIEW);
    var allChildNames = getAllChildNameOptions();
    var childViewNameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allChildNames, true)
      .build();
    childViewSheet.getRange('B1').setDataValidation(childViewNameRule);

    // 児童別ビューの年月ドロップダウン更新（すべて・年・月の選択肢を含む）
    var childViewOptions = generateChildViewOptions();
    var childViewYmRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(childViewOptions, true)
      .build();
    childViewSheet.getRange('B2').setDataValidation(childViewYmRule);

    // 月別集計の年月ドロップダウン更新（年全体オプション付き）
    var summaryOptions = generateMonthlySummaryOptions();
    var summaryYmRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(summaryOptions, true)
      .build();
    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    summarySheet.getRange('B1').setDataValidation(summaryYmRule);

    // 来館カレンダーの年月ドロップダウン更新（年オプション付き）
    var calendarOptions = generateCalendarOptions();
    var calendarRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(calendarOptions, true)
      .build();
    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setDataValidation(calendarRule);

    // 確定来館記録の年月ドロップダウン更新
    var confirmedOptions = generateConfirmedVisitsOptions();
    var confirmedRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(confirmedOptions, true)
      .build();
    var confirmedSheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
    confirmedSheet.getRange('B1').setDataValidation(confirmedRule);

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
 * 年次一括処理: 対象年の全12ヶ月を一括で処理する（月次ダイアログから呼び出し）
 * 振り分け → 確定来館記録 → 月別集計を全月分実行し、最後に年別カレンダーを更新する
 */
function runAnnualProcess() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = getTargetYearFromFormResponses_();
  var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
  var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);

  try {
    ss.toast(year + '年の年次一括処理を開始します...', '処理中', -1);

    for (var month = 1; month <= 12; month++) {
      var yearMonthStr = year + '年' + month + '月';
      ss.toast(yearMonthStr + ' を処理中...', '処理中', -1);
      summarySheet.getRange('B1').setValue(yearMonthStr);
      allocateRemainingPoints_(year, month);
    }

    // 来館カレンダーを年全体で更新
    calendarSheet.getRange('B1').setValue(year + '年');
    updateVisitCalendarByYear_(calendarSheet, year);

    // 月別集計の表示を最終月に合わせて戻す（12月）
    summarySheet.getRange('B1').setValue(year + '年12月');
    updateMonthlySummary();

    ss.toast(year + '年の年次一括処理が完了しました', '完了', 5);
    ui.alert(year + '年の年次一括処理が完了しました\n・振り分け・確定来館記録（全12ヶ月）\n・来館カレンダー（年別）\n・月別集計（12月）');
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
 * B1セルの値に基づいて行の表示/非表示を切り替える
 */
function filterConfirmedVisits_() {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var value = sheet.getRange('B1').getValue();
  var lastRow = sheet.getLastRow();

  if (lastRow < CONFIRMED_DATA_START_ROW) return;

  var numDataRows = lastRow - CONFIRMED_DATA_START_ROW + 1;

  // まず全データ行を表示
  sheet.showRows(CONFIRMED_DATA_START_ROW, numDataRows);

  if (!value || value === 'すべて') return;

  var yearOnly = parseYearOnly_(String(value).trim());
  var ym = yearOnly ? null : parseYearMonth(value);

  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, numDataRows, 1).getValues();

  // 連続する非表示行をバッチで処理（高速化）
  var hideStart = -1;
  var hideCount = 0;

  for (var i = 0; i <= data.length; i++) {
    var shouldHide = false;
    if (i < data.length && data[i][0]) {
      var recordDate = new Date(data[i][0]);
      if (yearOnly) {
        shouldHide = recordDate.getFullYear() !== yearOnly;
      } else if (ym) {
        shouldHide = recordDate.getFullYear() !== ym.year || (recordDate.getMonth() + 1) !== ym.month;
      }
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
 * 確定来館記録のWebビューをブラウザで開く
 * 事前にWebアプリとしてデプロイが必要
 */
function openWebView() {
  var ui = SpreadsheetApp.getUi();
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) {
      ui.alert('WebビューのURLが取得できません。\n\n「デプロイ > デプロイを管理」からWebアプリとしてデプロイしてください。');
      return;
    }
    var html = HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;font-size:14px">以下のURLをブラウザで開いてください：</p>'
      + '<p><a href="' + url + '" target="_blank" style="font-size:13px;word-break:break-all">' + url + '</a></p>'
      + '<p style="font-size:12px;color:#888;margin-top:12px">このウィンドウは閉じて構いません。</p>'
    ).setWidth(480).setHeight(120);
    ui.showModelessDialog(html, '確定来館記録 Webビュー');
  } catch (e) {
    ui.alert('エラー: ' + e.message + '\n\nWebアプリとしてデプロイされているか確認してください。');
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
