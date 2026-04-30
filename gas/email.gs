/**
 * F-07: 保護者向け来館報告メール送信
 *
 * TODO: 同日に兄弟（同一保護者メールアドレス）が来館した場合、
 *       個別メールではなく1通にまとめるかどうか検討・実装する
 */

/**
 * 過去24時間以内に送信されたフォーム回答のうち、未送信の保護者宛メールを送る
 * 自動トリガー（毎日AM8時）用。「タイムスタンプ ≧ now-24h」で対象抽出する。
 */
function sendDailyVisitReports() {
  try {
    sendVisitReportsRecent_();
  } catch (error) {
    logError_('sendDailyVisitReports', error);
  }
}

/**
 * 手動実行用: HTMLダイアログで対象日付を選択してメール送信する
 * フロー: Step1(日付選択) → Step2(件数プレビュー) → Step3(送信完了)
 */
function sendVisitReportsManual() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var defaultDate = formatDateYMD_(yesterday, 'yyyy-MM-dd');

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '  body { font-family: "Google Sans", sans-serif; padding: 16px; margin: 0; }' +
    '  h3 { margin: 0 0 16px; font-size: 16px; }' +
    '  input[type="date"] { font-size: 16px; padding: 8px 12px; border: 1px solid #dadce0; border-radius: 4px; width: 100%; box-sizing: border-box; }' +
    '  .buttons { margin-top: 20px; text-align: right; }' +
    '  button { font-size: 14px; padding: 8px 24px; border: none; border-radius: 4px; cursor: pointer; margin-left: 8px; }' +
    '  button:disabled { opacity: 0.6; cursor: default; }' +
    '  .cancel { background: #f1f3f4; color: #5f6368; }' +
    '  .submit { background: #1a73e8; color: #fff; }' +
    '  .submit:hover:not(:disabled) { background: #1765cc; }' +
    '  .summary { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 4px; padding: 12px 16px; font-size: 14px; line-height: 1.8; }' +
    '  .summary .row { display: flex; justify-content: space-between; }' +
    '  .summary .label { color: #5f6368; }' +
    '  .summary .value { font-weight: 500; }' +
    '  .summary .value.sendable { color: #1a73e8; }' +
    '  .target-date { font-size: 13px; color: #5f6368; margin-bottom: 8px; }' +
    '  .result-main { font-size: 18px; font-weight: 500; text-align: center; margin: 16px 0 12px; }' +
    '  .result-sub { font-size: 13px; color: #5f6368; text-align: center; }' +
    '  .hidden { display: none; }' +
    '</style>' +

    '<div id="step1">' +
    '  <h3>対象日付を選択してください</h3>' +
    '  <input type="date" id="targetDate" value="' + defaultDate + '">' +
    '  <div class="buttons">' +
    '    <button class="cancel" onclick="google.script.host.close()">キャンセル</button>' +
    '    <button class="submit" id="confirmBtn" onclick="goConfirm()">確認</button>' +
    '  </div>' +
    '</div>' +

    '<div id="step2" class="hidden">' +
    '  <h3>送信内容の確認</h3>' +
    '  <div class="target-date" id="step2Date"></div>' +
    '  <div class="summary">' +
    '    <div class="row"><span class="label">送信予定</span><span class="value sendable" id="cntSendable">-</span></div>' +
    '  </div>' +
    '  <div class="buttons">' +
    '    <button class="cancel" onclick="goBack()">戻る</button>' +
    '    <button class="submit" id="sendBtn" onclick="goSend()">送信</button>' +
    '  </div>' +
    '</div>' +

    '<div id="step3" class="hidden">' +
    '  <h3>送信完了</h3>' +
    '  <div class="target-date" id="step3Date"></div>' +
    '  <div class="result-main" id="resultMain"></div>' +
    '  <div class="result-sub" id="resultSub"></div>' +
    '  <div class="buttons">' +
    '    <button class="submit" onclick="google.script.host.close()">閉じる</button>' +
    '  </div>' +
    '</div>' +

    '<script>' +
    '  var currentDate = "";' +
    '  function show(stepId) {' +
    '    ["step1","step2","step3"].forEach(function(id){' +
    '      document.getElementById(id).classList.toggle("hidden", id !== stepId);' +
    '    });' +
    '  }' +
    '  function goBack() { show("step1"); }' +
    '  function goConfirm() {' +
    '    var date = document.getElementById("targetDate").value;' +
    '    if (!date) { alert("日付を選択してください"); return; }' +
    '    currentDate = date;' +
    '    var btn = document.getElementById("confirmBtn");' +
    '    btn.disabled = true; btn.textContent = "確認中...";' +
    '    google.script.run' +
    '      .withSuccessHandler(function(r) {' +
    '        btn.disabled = false; btn.textContent = "確認";' +
    '        document.getElementById("step2Date").textContent = "対象日: " + r.date;' +
    '        document.getElementById("cntSendable").textContent = r.sendable + "件";' +
    '        var sendBtn = document.getElementById("sendBtn");' +
    '        sendBtn.disabled = (r.sendable === 0);' +
    '        sendBtn.textContent = (r.sendable === 0) ? "送信対象なし" : "送信";' +
    '        show("step2");' +
    '      })' +
    '      .withFailureHandler(function(e) {' +
    '        btn.disabled = false; btn.textContent = "確認";' +
    '        alert("エラー: " + e.message);' +
    '      })' +
    '      .countVisitReportsByDate(date);' +
    '  }' +
    '  function goSend() {' +
    '    var btn = document.getElementById("sendBtn");' +
    '    btn.disabled = true; btn.textContent = "送信中...";' +
    '    google.script.run' +
    '      .withSuccessHandler(function(r) {' +
    '        document.getElementById("step3Date").textContent = "対象日: " + r.date;' +
    '        document.getElementById("resultMain").textContent = r.sent + "/" + r.sendable + "件 送信完了";' +
    '        var subParts = [];' +
    '        if (r.skipped > 0) subParts.push("スキップ: " + r.skipped + "件");' +
    '        if (r.errors > 0) subParts.push("エラー: " + r.errors + "件");' +
    '        document.getElementById("resultSub").textContent = subParts.join(" / ");' +
    '        show("step3");' +
    '      })' +
    '      .withFailureHandler(function(e) {' +
    '        btn.disabled = false; btn.textContent = "送信";' +
    '        alert("エラー: " + e.message);' +
    '      })' +
    '      .sendVisitReportsByDateFromDialog(currentDate);' +
    '  }' +
    '</script>'
  )
  .setWidth(360)
  .setHeight(240);

  SpreadsheetApp.getUi().showModalDialog(html, '来館報告メール送信');
}

/**
 * HTMLダイアログから呼ばれる: 指定日の送信予定件数を集計する（送信は行わない）
 * 条件: タイムスタンプの日付 = 対象日 / メール送信済が空
 * @param {string} dateStr yyyy-MM-dd形式の日付文字列
 * @returns {{date: string, sendable: number}}
 */
function countVisitReportsByDate(dateStr) {
  var targetDate = parseDateInput_(dateStr);
  if (!targetDate) {
    throw new Error('日付の形式が不正です: ' + dateStr);
  }
  var tz = Session.getScriptTimeZone();
  var targetDateStr = formatDateYMD_(targetDate, 'yyyy/MM/dd', tz);

  var records = getFormResponsesByDate_(targetDate);
  var sendable = 0;
  for (var i = 0; i < records.length; i++) {
    var row = records[i].data;
    var timestamp = row[FORM_COL.TIMESTAMP - 1];
    var sentFlag = row[FORM_COL.EMAIL_SENT - 1];
    if (!timestamp) continue;
    if (sentFlag && String(sentFlag).trim() !== '') continue;
    sendable++;
  }
  return { date: targetDateStr, sendable: sendable };
}

/**
 * HTMLダイアログから呼ばれるメール送信処理
 * @param {string} dateStr yyyy-MM-dd形式の日付文字列
 * @returns {{date: string, sent: number, sendable: number, skipped: number, errors: number}}
 */
function sendVisitReportsByDateFromDialog(dateStr) {
  var targetDate = parseDateInput_(dateStr);
  if (!targetDate) {
    throw new Error('日付の形式が不正です: ' + dateStr);
  }
  return sendVisitReportsByDate_(targetDate);
}

/**
 * 指定日の来館記録を保護者にメール送信する（手動ダイアログ用）
 * @param {Date} targetDate 対象日付
 * @returns {{date: string, sent: number, sendable: number, skipped: number, errors: number}} 処理結果
 */
function sendVisitReportsByDate_(targetDate) {
  var tz = Session.getScriptTimeZone();
  var targetDateStr = formatDateYMD_(targetDate, 'yyyy/MM/dd', tz);
  Logger.log('来館報告メール送信開始: 対象日 = ' + targetDateStr);

  // フォームの回答から対象日のレコードを抽出
  var records = getFormResponsesByDate_(targetDate);
  if (records.length === 0) {
    Logger.log('対象日の来館記録がありません: ' + targetDateStr);
    return { date: targetDateStr, sent: 0, sendable: 0, skipped: 0, errors: 0 };
  }

  var result = processEmailSend_(records, targetDate);
  return {
    date: targetDateStr,
    sent: result.sent,
    sendable: result.sendable,
    skipped: result.skipped,
    errors: result.errors,
  };
}

/**
 * フォーム回答レコード群を保護者宛メール送信する共通処理
 * @param {Array<{rowIndex:number, data:Array}>} records
 * @param {Date|null} fixedTargetDate 全件で使う対象日付。null の場合は各レコードのタイムスタンプ日付を使う。
 * @returns {{sent:number, sendable:number, skipped:number, errors:number}}
 */
function processEmailSend_(records, fixedTargetDate) {
  var tz = Session.getScriptTimeZone();

  // スタッフ用_回答シートへの参照（送信済フラグ書き込み用）
  var staffSs = SpreadsheetApp.openById(STAFF_SHEET_ID);
  var formSheet = staffSs.getSheetByName(SHEET_NAMES.FORM_RESPONSE);

  // ヘッダーにメール送信済列がなければ設定
  var headerValue = formSheet.getRange(1, FORM_COL.EMAIL_SENT).getValue();
  if (!headerValue) {
    formSheet.getRange(1, FORM_COL.EMAIL_SENT).setValue('メール送信済');
  }

  // 児童マスタを取得してマップ化（児童名 → 行データ）
  var childMasterMap = buildChildNameToRowMap_();

  // スクリプトプロパティから施設名を取得
  var props = PropertiesService.getScriptProperties();
  var facilityName = props.getProperty('FACILITY_NAME') || '施設';
  var senderName = props.getProperty('EMAIL_SENDER_NAME') || facilityName;

  var sentCount = 0;
  var skipCount = 0;
  var errorCount = 0;
  var sendableCount = 0;

  for (var i = 0; i < records.length; i++) {
    var record = records[i].data;
    var rowIndex = records[i].rowIndex;
    var childName = record[FORM_COL.CHILD_NAME - 1];

    // 送信済チェック
    var sentFlag = formSheet.getRange(rowIndex, FORM_COL.EMAIL_SENT).getValue();
    if (sentFlag && String(sentFlag).indexOf('送信済') !== -1) {
      Logger.log('送信済みのためスキップ: ' + childName + ' (行' + rowIndex + ')');
      skipCount++;
      continue;
    }

    var masterRow = childMasterMap[childName];

    if (!masterRow) {
      Logger.log('児童マスタに該当なし（スキップ）: ' + childName);
      skipCount++;
      continue;
    }

    var parentEmail = masterRow[MASTER_COL.PARENT_EMAIL - 1];
    if (!parentEmail || String(parentEmail).trim() === '') {
      Logger.log('保護者メールアドレスが未設定（スキップ）: ' + childName);
      skipCount++;
      continue;
    }

    sendableCount++;
    try {
      var ts2 = record[FORM_COL.TIMESTAMP - 1];
      // メール本文の対象日: 引数で固定日が指定されていればそれを、無ければレコードの利用日を使う
      var perRecordTargetDate = fixedTargetDate || getRowRecordDate_(record) || (ts2 instanceof Date ? new Date(ts2.getFullYear(), ts2.getMonth(), ts2.getDate()) : new Date());
      var emailData = buildEmailData_(record, masterRow, perRecordTargetDate, facilityName);
      MailApp.sendEmail({
        to: String(parentEmail).trim(),
        subject: emailData.subject,
        body: emailData.body,
        name: senderName,
      });

      // 送信済フラグを書き込み
      var sentTimestamp = formatDateYMD_(new Date(), 'yyyy/MM/dd HH:mm', tz);
      formSheet.getRange(rowIndex, FORM_COL.EMAIL_SENT).setValue('送信済 ' + sentTimestamp);

      Logger.log('メール送信成功: ' + childName + ' → ' + parentEmail);
      sentCount++;
    } catch (error) {
      Logger.log('メール送信エラー: ' + childName + ' → ' + parentEmail + ' / ' + error.message);
      errorCount++;
    }
  }

  Logger.log('来館報告メール送信完了: 送信=' + sentCount + '件, スキップ=' + skipCount + '件, エラー=' + errorCount + '件');
  return {
    sent: sentCount,
    sendable: sendableCount,
    skipped: skipCount,
    errors: errorCount,
  };
}

/**
 * フォームの回答から指定日のデータを取得する（行番号付き）
 * フォームから「記録日」項目を廃止したため、タイムスタンプの日付を基準にフィルタする。
 * @param {Date} targetDate 対象日付
 * @returns {Array<{rowIndex: number, data: Array}>} 該当日のフォーム回答データ（rowIndexはシート上の行番号、1始まり）
 */
function getFormResponsesByDate_(targetDate) {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var data = sheet.getDataRange().getValues();
  var responses = data.slice(1);
  var tz = Session.getScriptTimeZone();
  var targetStr = formatDateYMD_(targetDate, 'yyyy-MM-dd', tz);

  var results = [];
  for (var i = 0; i < responses.length; i++) {
    var ts = responses[i][FORM_COL.TIMESTAMP - 1];
    if (!isValidDate_(ts)) continue;
    var dateStr = formatDateYMD_(ts, 'yyyy-MM-dd', tz);
    if (dateStr === targetStr) {
      results.push({
        rowIndex: i + 2,
        data: responses[i],
      });
    }
  }
  return results;
}

/**
 * 過去24時間以内のフォーム回答のうち、メール未送信のものに保護者宛メールを送る
 * - 抽出条件: タイムスタンプ ≧ now - 24時間 / メール送信済が空
 * - 自動トリガー（毎日AM8時）から呼ばれる
 * @returns {{from: string, sent: number, sendable: number, skipped: number, errors: number}}
 */
function sendVisitReportsRecent_() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var since = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  Logger.log('来館報告メール(過去24時間)送信開始: ' + formatDateYMD_(since, 'yyyy/MM/dd HH:mm', tz) + ' 〜 ' + formatDateYMD_(now, 'yyyy/MM/dd HH:mm', tz));

  var records = getFormResponsesSince_(since);
  if (records.length === 0) {
    Logger.log('過去24時間に新規回答はありません');
    return { from: formatDateYMD_(since, 'yyyy/MM/dd HH:mm', tz), sent: 0, sendable: 0, skipped: 0, errors: 0 };
  }

  var result = processEmailSend_(records, null /* targetDate: 各レコードのタイムスタンプ日付を本文に使う */);
  return {
    from: formatDateYMD_(since, 'yyyy/MM/dd HH:mm', tz),
    sent: result.sent,
    sendable: result.sendable,
    skipped: result.skipped,
    errors: result.errors,
  };
}

/**
 * フォーム回答シートからタイムスタンプが指定時刻以降のレコードを取得する
 * @param {Date} since
 * @returns {Array<{rowIndex: number, data: Array}>}
 */
function getFormResponsesSince_(since) {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var data = sheet.getDataRange().getValues();
  var responses = data.slice(1);
  var results = [];
  var sinceTime = since.getTime();
  for (var i = 0; i < responses.length; i++) {
    var ts = responses[i][FORM_COL.TIMESTAMP - 1];
    if (!isValidDate_(ts)) continue;
    if (ts.getTime() < sinceTime) continue;
    results.push({ rowIndex: i + 2, data: responses[i] });
  }
  return results;
}

/**
 * メールの件名と本文を組み立てる
 * @param {Array} record フォーム回答の1行
 * @param {Array} masterRow 児童マスタの1行
 * @param {Date} targetDate 対象日付
 * @param {string} facilityName 施設名
 * @param {Object} [pairedStay] 連泊ペアリング後の論理1宿泊（連泊時の入退所表示に使用）
 * @returns {{subject: string, body: string}}
 */
function buildEmailData_(record, masterRow, targetDate, facilityName) {
  var dateStr = formatDateYMD_(targetDate, 'M月d日');
  var childName = record[FORM_COL.CHILD_NAME - 1];
  var parentName = masterRow[MASTER_COL.PARENT_NAME - 1] || '';

  var subject = getEmailSubjectFromSettings_();
  var template = getEmailBodyFromSettings_();

  // フォーム1行=1宿泊運用のため、レコードの入退所をそのまま採用する
  var displayCheckIn = record[FORM_COL.CHECK_IN - 1];
  var displayCheckOut = record[FORM_COL.CHECK_OUT - 1];

  var body = template
    .replace(/{保護者名}/g, parentName)
    .replace(/{日付}/g, dateStr)
    .replace(/{児童名}/g, childName)
    .replace(/{入所時間}/g, formatTime_(displayCheckIn))
    .replace(/{退所時間}/g, formatTime_(displayCheckOut))
    .replace(/{体温}/g, record[FORM_COL.TEMPERATURE - 1] || '')
    .replace(/{夕食}/g, record[FORM_COL.MEAL_DINNER - 1] || '')
    .replace(/{朝食}/g, record[FORM_COL.MEAL_BREAKFAST - 1] || '')
    .replace(/{昼食}/g, record[FORM_COL.MEAL_LUNCH - 1] || '')
    .replace(/{入浴}/g, record[FORM_COL.BATH - 1] || '')
    .replace(/{入眠時刻}/g, record[FORM_COL.SLEEP_ONSET - 1] || '')
    .replace(/{朝4時チェック}/g, record[FORM_COL.SLEEP_CHECK_4AM - 1] || '')
    .replace(/{起床時刻}/g, record[FORM_COL.WAKE_UP - 1] || '')
    .replace(/{便}/g, record[FORM_COL.BOWEL - 1] || '')
    .replace(/{服薬\(夜\)}/g, record[FORM_COL.MEDICINE_NIGHT - 1] || '')
    .replace(/{服薬\(朝\)}/g, record[FORM_COL.MEDICINE_MORNING - 1] || '')
    .replace(/{連絡事項}/g, record[FORM_COL.NOTES - 1] || '特になし');

  return { subject: subject, body: body };
}

/**
 * 時刻値をHH:mm形式の文字列にフォーマットする
 * @param {*} value 時刻値（Date型またはテキスト）
 * @returns {string} フォーマット済み時刻文字列
 */
function formatTime_(value) {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return formatDateYMD_(value, 'HH:mm');
  }
  return String(value);
}

/**
 * 日付入力文字列をパースする（yyyy/MM/dd または yyyy-MM-dd）
 * @param {string} input 日付文字列
 * @returns {Date|null} パース結果（不正な場合はnull）
 */
function parseDateInput_(input) {
  var match = String(input).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!match) {
    return null;
  }
  var date = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
  if (isNaN(date.getTime())) {
    return null;
  }
  return date;
}

/**
 * ログシートに新しいエラーが追加されたときにメール通知する（トリガー用）
 * 送信先: GAS実行者（固定） + スクリプトプロパティ ERROR_NOTIFY_EMAILS（任意・カンマ区切り）
 */
function notifyErrorLog() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.LOG);
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    // 最新のエラー行を取得
    var lastRowData = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
    var timestamp = lastRowData[0];
    var functionName = lastRowData[1];
    var errorMessage = lastRowData[2];
    var stackTrace = lastRowData[3];

    // 送信先を構築（実行者 + 追加通知先）
    var recipients = getErrorNotifyRecipients_();

    var props = PropertiesService.getScriptProperties();
    var facilityName = props.getProperty('FACILITY_NAME') || '施設';

    var subject = '【' + facilityName + '】エラー通知: ' + functionName;
    var body = 'エラーが発生しました。\n\n'
      + '■ 発生日時: ' + timestamp + '\n'
      + '■ 関数名: ' + functionName + '\n'
      + '■ エラーメッセージ:\n' + errorMessage + '\n\n'
      + '■ スタックトレース:\n' + (stackTrace || 'なし') + '\n\n'
      + '---\n'
      + 'スプレッドシート: ' + ss.getUrl() + '\n';

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body,
      name: facilityName + ' システム通知',
    });

    Logger.log('エラー通知メール送信完了: ' + recipients.join(', '));
  } catch (e) {
    Logger.log('エラー通知メール送信に失敗: ' + e.message);
  }
}

/**
 * エラー通知の送信先メールアドレスを取得する
 * GAS実行者（固定） + 設定シートのエラー通知先メール（カンマ区切り）
 *   + スクリプトプロパティ ERROR_NOTIFY_EMAILS（後方互換・任意）
 * @returns {Array<string>} メールアドレスの配列（重複排除済み）
 */
function getErrorNotifyRecipients_() {
  // GAS実行者は必ず含める
  var ownerEmail = Session.getEffectiveUser().getEmail();
  var recipients = [ownerEmail];

  // 設定シート優先
  var sheetEmails = getErrorEmailsFromSettings_();

  // スクリプトプロパティ（後方互換）
  var props = PropertiesService.getScriptProperties();
  var extraProp = props.getProperty('ERROR_NOTIFY_EMAILS') || '';
  var propEmails = extraProp.trim()
    ? extraProp.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return !!e; })
    : [];

  sheetEmails.concat(propEmails).forEach(function(email) {
    if (email && recipients.indexOf(email) === -1) recipients.push(email);
  });

  return recipients;
}

/**
 * メール送信の時間トリガーを設定する（毎朝8時）
 */
function setupEmailTrigger() {
  setupTimeTrigger_('sendDailyVisitReports', { everyDays: 1, atHour: 8 });
  Logger.log('メール送信トリガーを設定しました（毎朝8時）');
}
