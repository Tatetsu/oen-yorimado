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
    .addSeparator()
    .addItem('バウンスメールを確認', 'collectBounceEmailsManual')
    .addToUi();
}

/**
 * 月次一括処理: 月別集計 → 振り分け → 確定来館記録 → 来館カレンダーを一括実行する
 * 選択肢は当月までに制限（未来月は非表示）
 */
function runMonthlyProcess() {
  var ui = SpreadsheetApp.getUi();

  try {
    var years = generateProcessableYears_();
    if (years.length === 0) {
      ui.alert('処理可能な月がありません（当月以前のデータがありません）');
      return;
    }

    var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);

    // Stage 1: 年選択
    var yearPrompt = '対象年を選択してください（番号を入力）:\n\n';
    for (var i = 0; i < years.length; i++) {
      yearPrompt += i + '. ' + years[i] + '年\n';
    }
    var yearResponse = ui.prompt('月次一括処理 - 年選択', yearPrompt, ui.ButtonSet.OK_CANCEL);
    if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

    var yearIdx = parseInt(yearResponse.getResponseText().trim(), 10);
    if (isNaN(yearIdx) || yearIdx < 0 || yearIdx >= years.length) {
      ui.alert('無効な番号です。0〜' + (years.length - 1) + 'の番号を入力してください。');
      return;
    }
    var selectedYear = years[yearIdx];

    // Stage 2: 月選択（年次一括処理を先頭に含む）
    var monthOptions = generateMonthOptionsForYear_(selectedYear);
    var monthPrompt = '対象月を選択してください（番号を入力）:\n\n';
    for (var j = 0; j < monthOptions.length; j++) {
      monthPrompt += j + '. ' + monthOptions[j].label + '\n';
    }
    var monthResponse = ui.prompt('月次一括処理 - ' + selectedYear + '年', monthPrompt, ui.ButtonSet.OK_CANCEL);
    if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

    var monthIdx = parseInt(monthResponse.getResponseText().trim(), 10);
    if (isNaN(monthIdx) || monthIdx < 0 || monthIdx >= monthOptions.length) {
      ui.alert('無効な番号です。0〜' + (monthOptions.length - 1) + 'の番号を入力してください。');
      return;
    }

    var selected = monthOptions[monthIdx];
    if (selected.kind === 'annual') {
      runAnnualProcess(selected.year);
      return;
    }

    var ym = { year: selected.year, month: selected.month };

    // 月別集計のドロップダウンを反映
    summarySheet.getRange('B1').setValue(ym.year + '年');
    summarySheet.getRange('B2').setValue(ym.month + '月');

    // 0. フォーム回答のバリデーション（入所/退所日時の欠落・時系列逆転・期間重複）
    var responseIssues = validateOvernightRecords(ym.year, ym.month);
    if (responseIssues.length > 0) {
      var lines = responseIssues.slice(0, 20).map(function(it) {
        return '・' + formatDateYMD_(it.recordDate) + ' ' + String(it.childName || '(児童名なし)') + ': ' + it.issues.join(' / ');
      });
      var more = responseIssues.length > 20 ? '\n（他 ' + (responseIssues.length - 20) + ' 件）' : '';
      var proceed = ui.alert(
        'フォーム回答に問題が ' + responseIssues.length + ' 件あります',
        lines.join('\n') + more + '\n\nこのまま処理を続行しますか？',
        ui.ButtonSet.YES_NO
      );
      if (proceed !== ui.Button.YES) return;
    }

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
    } else if (sheetName === SHEET_NAMES.VISIT_CALENDAR && (cell === 'B1' || cell === 'B2')) {
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
 * 各ビューシートのドロップダウン、および Google フォームのドロップダウン
 * （児童マスタの児童名・スタッフマスタのスタッフ1/2）を一括で最新化する
 */
function refreshDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  try {
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
      { sheet: SHEET_NAMES.VISIT_CALENDAR,   cell: 'B2', options: monthOpts },
      { sheet: SHEET_NAMES.CONFIRMED_VISITS, cell: 'B1', options: yearOpts },
      { sheet: SHEET_NAMES.CONFIRMED_VISITS, cell: 'B2', options: monthOpts },
    ];

    dropdownConfigs.forEach(function(cfg) {
      var sheet = getSheet(cfg.sheet);
      setListValidation_(sheet.getRange(cfg.cell), cfg.options);
    });

    // 児童マスタのドロップダウン（重度支援区分）を再適用
    setupChildMasterValidations_();

    // Google フォームの児童名・スタッフ1・スタッフ2を同期
    syncFormDropdowns();

    Logger.log('ドロップダウンを更新しました（シート + フォーム 児童マスタ・スタッフマスタ）');
    SpreadsheetApp.getUi().alert('ドロップダウンを更新しました\n\n・各ビューシート（年/月/児童名）\n・児童マスタ（重度支援区分）\n・フォーム（児童名・スタッフ1・スタッフ2）');
  } catch (error) {
    logError_('refreshDropdowns', error);
    SpreadsheetApp.getUi().alert('ドロップダウン更新中にエラーが発生しました: ' + error.message);
  } finally {
    originalSheet.activate();
  }
}

/**
 * 年次一括処理: 対象年の1月〜前月までを一括で処理する
 * 振り分け → 確定来館記録 → 月別集計を対象月数分実行し、最後に年別カレンダーを更新する
 * 前月より後の月は処理しない。さらに、既に書き込まれていた振り分け行があれば自動クリアする
 */
function runAnnualProcess(targetYear) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
  var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);

  // 対象年は前月年以下に制限し、前月年なら1月〜前月、それ以前なら1月〜12月を処理
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastY = lastMonth.getFullYear();
  var lastM = lastMonth.getMonth() + 1;
  var baseYear = (typeof targetYear === 'number' && targetYear > 0)
    ? targetYear
    : getTargetYearFromFormResponses_();
  var year = Math.min(baseYear, lastY);
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
 * 翌月1日の午前6時に前日（前月末日）の属する月を対象として全処理を実行する
 */
function setupMonthlyProcessTrigger() {
  setupTimeTrigger_('runMonthlyProcessAutomatic', { onMonthDay: 1, atHour: 6 });
  Logger.log('月次一括処理トリガーを設定しました（翌月1日 午前6時）');
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
 * 月次一括処理と同じ 2段プロンプト（年→月）で対象スコープを選ばせ、
 * 月選択時はその月、「年次一括処理」選択時はその年の実データを洗い替える
 */
function updateConfirmedVisitsAndCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var originalSheet = ss.getActiveSheet();
  try {
    // Stage 1: 年選択
    var years = generateProcessableYears_();
    if (years.length === 0) {
      ui.alert('処理可能な月がありません（当月以前のデータがありません）');
      return;
    }
    var yearPrompt = '対象年を選択してください（番号を入力）:\n\n';
    for (var i = 0; i < years.length; i++) {
      yearPrompt += i + '. ' + years[i] + '年\n';
    }
    var yearResponse = ui.prompt('確定来館記録を手動更新 - 年選択', yearPrompt, ui.ButtonSet.OK_CANCEL);
    if (yearResponse.getSelectedButton() !== ui.Button.OK) return;
    var yearIdx = parseInt(yearResponse.getResponseText().trim(), 10);
    if (isNaN(yearIdx) || yearIdx < 0 || yearIdx >= years.length) {
      ui.alert('無効な番号です。0〜' + (years.length - 1) + 'の番号を入力してください。');
      return;
    }
    var selectedYear = years[yearIdx];

    // Stage 2: 月選択（先頭は年次一括）
    var monthOptions = generateMonthOptionsForYear_(selectedYear);
    var monthPrompt = '対象月を選択してください（番号を入力）:\n\n';
    for (var j = 0; j < monthOptions.length; j++) {
      monthPrompt += j + '. ' + monthOptions[j].label + '\n';
    }
    var monthResponse = ui.prompt('確定来館記録を手動更新 - ' + selectedYear + '年', monthPrompt, ui.ButtonSet.OK_CANCEL);
    if (monthResponse.getSelectedButton() !== ui.Button.OK) return;
    var monthIdx = parseInt(monthResponse.getResponseText().trim(), 10);
    if (isNaN(monthIdx) || monthIdx < 0 || monthIdx >= monthOptions.length) {
      ui.alert('無効な番号です。0〜' + (monthOptions.length - 1) + 'の番号を入力してください。');
      return;
    }
    var selected = monthOptions[monthIdx];
    var isAnnual = (selected.kind === 'annual');
    var targetMonth = isAnnual ? null : selected.month;

    // 振り分け済みチェック（選択スコープ内に限定）
    var allocatedInScope = listAllocatedMonths_().filter(function(ym) {
      if (ym.year !== selectedYear) return false;
      return isAnnual ? true : (ym.month === targetMonth);
    });
    if (allocatedInScope.length > 0) {
      var shown = allocatedInScope.map(function(ym) {
        return '・' + ym.year + '年' + ym.month + '月';
      });
      var proceed = ui.alert(
        '振り分け済みの月があります',
        '以下の月には既に振り分け（申請用データ）が登録されています:\n\n' +
        shown.join('\n') + '\n\n' +
        '実データを更新すると、振り分け行はそのまま残りますが\n' +
        '実来館との整合性が崩れる可能性があります。\n\n' +
        'このまま更新しますか？',
        ui.ButtonSet.YES_NO
      );
      if (proceed !== ui.Button.YES) return;
    }

    var scopeLabel = isAnnual ? (selectedYear + '年') : (selectedYear + '年' + targetMonth + '月');
    ss.toast(scopeLabel + ' の確定来館記録を更新中...', '処理中', -1);

    if (isAnnual) {
      updateConfirmedVisits(selectedYear);
    } else {
      updateConfirmedVisits(selectedYear, targetMonth);
    }

    ss.toast('来館カレンダーを更新中...', '処理中', -1);

    var calendarSheet = getSheet(SHEET_NAMES.VISIT_CALENDAR);
    calendarSheet.getRange('B1').setValue(selectedYear + '年');
    calendarSheet.getRange('B2').setValue(isAnnual ? 'すべて' : (targetMonth + '月'));
    updateVisitCalendar();

    ss.toast(scopeLabel + ' の確定来館記録と来館カレンダーを更新しました', '完了', 3);
  } catch (error) {
    logError_('updateConfirmedVisitsAndCalendar', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  } finally {
    originalSheet.activate();
  }
}

/**
 * 月次一括処理で選択可能な年リストを降順で返す
 * フォーム回答に存在する年のうち、当月以前のデータが対象になり得る年のみ
 * @returns {Array<number>}
 */
function generateProcessableYears_() {
  var now = new Date();
  var currentY = now.getFullYear();
  return collectYearsFromFormResponses_().filter(function(y) { return y <= currentY; });
}

/**
 * 指定年で選択可能な月オプションを生成する（先頭に年次一括処理を含む）
 * 当年は1〜当月、それ以前の年は1〜12月を対象
 * @param {number} year 対象年
 * @returns {Array<{kind:string, year:number, month:number=, label:string}>}
 */
function generateMonthOptionsForYear_(year) {
  var now = new Date();
  var currentY = now.getFullYear();
  var currentM = now.getMonth() + 1;
  var maxMonth = (year === currentY) ? currentM : 12;
  var options = [{ kind: 'annual', year: year, label: '年次一括処理' }];
  for (var m = 1; m <= maxMonth; m++) {
    options.push({ kind: 'month', year: year, month: m, label: m + '月' });
  }
  return options;
}
