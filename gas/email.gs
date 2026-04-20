/**
 * F-07: 保護者向け来館報告メール送信
 *
 * TODO: 同日に兄弟（同一保護者メールアドレス）が来館した場合、
 *       個別メールではなく1通にまとめるかどうか検討・実装する
 */

/**
 * 前日の来館記録を保護者にメール送信する（自動トリガー用）
 */
function sendDailyVisitReports() {
  try {
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    sendVisitReportsByDate_(yesterday);
  } catch (error) {
    logError_('sendDailyVisitReports', error);
  }
}

/**
 * 手動実行用: HTMLダイアログで対象日付を選択してメール送信する
 */
function sendVisitReportsManual() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var defaultDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '  body { font-family: "Google Sans", sans-serif; padding: 16px; }' +
    '  h3 { margin: 0 0 16px; font-size: 16px; }' +
    '  input[type="date"] { font-size: 16px; padding: 8px 12px; border: 1px solid #dadce0; border-radius: 4px; width: 100%; box-sizing: border-box; }' +
    '  .buttons { margin-top: 20px; text-align: right; }' +
    '  button { font-size: 14px; padding: 8px 24px; border: none; border-radius: 4px; cursor: pointer; margin-left: 8px; }' +
    '  .cancel { background: #f1f3f4; color: #5f6368; }' +
    '  .submit { background: #1a73e8; color: #fff; }' +
    '  .submit:hover { background: #1765cc; }' +
    '</style>' +
    '<h3>対象日付を選択してください</h3>' +
    '<input type="date" id="targetDate" value="' + defaultDate + '">' +
    '<div class="buttons">' +
    '  <button class="cancel" onclick="google.script.host.close()">キャンセル</button>' +
    '  <button class="submit" onclick="submitDate()">送信</button>' +
    '</div>' +
    '<script>' +
    '  function submitDate() {' +
    '    var date = document.getElementById("targetDate").value;' +
    '    if (!date) { alert("日付を選択してください"); return; }' +
    '    document.querySelector(".submit").disabled = true;' +
    '    document.querySelector(".submit").textContent = "送信中...";' +
    '    google.script.run' +
    '      .withSuccessHandler(function(result) {' +
    '        google.script.host.close();' +
    '        if (result) { google.script.run.showResultAlert(result); }' +
    '      })' +
    '      .withFailureHandler(function(e) { alert("エラー: " + e.message); document.querySelector(".submit").disabled = false; document.querySelector(".submit").textContent = "送信"; })' +
    '      .sendVisitReportsByDateFromDialog(date);' +
    '  }' +
    '</script>'
  )
  .setWidth(320)
  .setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(html, '来館報告メール送信');
}

/**
 * HTMLダイアログから呼ばれるメール送信処理
 * @param {string} dateStr yyyy-MM-dd形式の日付文字列
 */
function sendVisitReportsByDateFromDialog(dateStr) {
  var targetDate = parseDateInput_(dateStr);
  if (!targetDate) {
    throw new Error('日付の形式が不正です: ' + dateStr);
  }
  return sendVisitReportsByDate_(targetDate);
}

/**
 * 処理結果をアラートダイアログで表示する（HTMLダイアログから呼び出し用）
 * @param {string} message 表示するメッセージ
 */
function showResultAlert(message) {
  SpreadsheetApp.getUi().alert('来館報告メール送信結果', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 指定日の来館記録を保護者にメール送信する
 * @param {Date} targetDate 対象日付
 * @returns {string} 処理結果メッセージ
 */
function sendVisitReportsByDate_(targetDate) {
  var tz = Session.getScriptTimeZone();
  var targetDateStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
  Logger.log('来館報告メール送信開始: 対象日 = ' + targetDateStr);

  // フォームの回答から対象日のレコードを抽出
  var records = getFormResponsesByDate_(targetDate);
  if (records.length === 0) {
    Logger.log('対象日の来館記録がありません: ' + targetDateStr);
    return '対象日（' + targetDateStr + '）の来館記録がありません。';
  }

  // スタッフ用_回答シートへの参照（送信済フラグ書き込み用）
  var staffSs = SpreadsheetApp.openById(STAFF_SHEET_ID);
  var formSheet = staffSs.getSheetByName(SHEET_NAMES.FORM_RESPONSE);

  // ヘッダーにメール送信済列がなければ設定
  var headerValue = formSheet.getRange(1, FORM_COL.EMAIL_SENT).getValue();
  if (!headerValue) {
    formSheet.getRange(1, FORM_COL.EMAIL_SENT).setValue('メール送信済');
  }

  // 児童マスタを取得してマップ化（児童名 → 行データ）
  var childMasterMap = buildChildMasterMap_();

  // スクリプトプロパティから施設名を取得
  var props = PropertiesService.getScriptProperties();
  var facilityName = props.getProperty('FACILITY_NAME') || '施設';
  var senderName = props.getProperty('EMAIL_SENDER_NAME') || facilityName;

  var sentCount = 0;
  var skipCount = 0;
  var errorCount = 0;

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

    try {
      var emailData = buildEmailData_(record, masterRow, targetDate, facilityName);
      MailApp.sendEmail({
        to: String(parentEmail).trim(),
        subject: emailData.subject,
        body: emailData.body,
        name: senderName,
      });

      // 送信済フラグを書き込み
      var sentTimestamp = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
      formSheet.getRange(rowIndex, FORM_COL.EMAIL_SENT).setValue('送信済 ' + sentTimestamp);

      Logger.log('メール送信成功: ' + childName + ' → ' + parentEmail);
      sentCount++;
    } catch (error) {
      Logger.log('メール送信エラー: ' + childName + ' → ' + parentEmail + ' / ' + error.message);
      errorCount++;
    }
  }

  var resultMessage = '来館報告メール送信完了（' + targetDateStr + '）\n\n'
    + '送信: ' + sentCount + '件\n'
    + 'スキップ: ' + skipCount + '件\n'
    + 'エラー: ' + errorCount + '件';
  Logger.log('来館報告メール送信完了: 送信=' + sentCount + '件, スキップ=' + skipCount + '件, エラー=' + errorCount + '件');
  return resultMessage;
}

/**
 * フォームの回答から指定日のデータを取得する（行番号付き）
 * @param {Date} targetDate 対象日付
 * @returns {Array<{rowIndex: number, data: Array}>} 該当日のフォーム回答データ（rowIndexはシート上の行番号、1始まり）
 */
function getFormResponsesByDate_(targetDate) {
  var sheet = getSheet(SHEET_NAMES.FORM_RESPONSE);
  var data = sheet.getDataRange().getValues();
  var responses = data.slice(1);
  var tz = Session.getScriptTimeZone();
  var targetStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');

  var results = [];
  for (var i = 0; i < responses.length; i++) {
    var recordDate = responses[i][FORM_COL.RECORD_DATE - 1];
    if (!(recordDate instanceof Date)) continue;
    var rowStr = Utilities.formatDate(recordDate, tz, 'yyyy-MM-dd');
    if (rowStr === targetStr) {
      results.push({
        rowIndex: i + 2, // ヘッダー行(1行目)の次から
        data: responses[i],
      });
    }
  }
  return results;
}

/**
 * 児童マスタデータを児童名→行データのマップに変換する
 * @returns {Object} 児童名をキーとするマップ
 */
function buildChildMasterMap_() {
  var masterData = getChildMasterData();
  var map = {};
  for (var i = 0; i < masterData.length; i++) {
    var name = masterData[i][MASTER_COL.NAME - 1];
    map[name] = masterData[i];
  }
  return map;
}

/**
 * メールの件名と本文を組み立てる
 * @param {Array} record フォーム回答の1行
 * @param {Array} masterRow 児童マスタの1行
 * @param {Date} targetDate 対象日付
 * @param {string} facilityName 施設名
 * @returns {{subject: string, body: string}}
 */
function buildEmailData_(record, masterRow, targetDate, facilityName) {
  var tz = Session.getScriptTimeZone();
  var dateStr = Utilities.formatDate(targetDate, tz, 'M月d日');
  var childName = record[FORM_COL.CHILD_NAME - 1];
  var parentName = masterRow[MASTER_COL.PARENT_NAME - 1] || '';

  var subject = '【テスト施設　来館記録のお知らせ】';

  var body = EMAIL_TEMPLATE
    .replace(/{保護者名}/g, parentName)
    .replace(/{日付}/g, dateStr)
    .replace(/{児童名}/g, childName)
    .replace(/{入所時間}/g, formatTime_(record[FORM_COL.CHECK_IN - 1]))
    .replace(/{退所時間}/g, formatTime_(record[FORM_COL.CHECK_OUT - 1]))
    .replace(/{体温}/g, record[FORM_COL.TEMPERATURE - 1] || '')
    .replace(/{食事}/g, record[FORM_COL.MEAL - 1] || '')
    .replace(/{入浴}/g, record[FORM_COL.BATH - 1] || '')
    .replace(/{睡眠}/g, record[FORM_COL.SLEEP - 1] || '')
    .replace(/{便}/g, record[FORM_COL.BOWEL - 1] || '')
    .replace(/{服薬}/g, record[FORM_COL.MEDICINE - 1] || '')
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
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
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
 * GAS実行者（固定） + スクリプトプロパティ ERROR_NOTIFY_EMAILS（カンマ区切り・任意）
 * @returns {Array<string>} メールアドレスの配列（重複排除済み）
 */
function getErrorNotifyRecipients_() {
  // GAS実行者は必ず含める
  var ownerEmail = Session.getEffectiveUser().getEmail();
  var recipients = [ownerEmail];

  // 追加通知先があれば追加
  var props = PropertiesService.getScriptProperties();
  var extraEmails = props.getProperty('ERROR_NOTIFY_EMAILS') || '';
  if (extraEmails.trim()) {
    var extras = extraEmails.split(',').map(function(email) {
      return email.trim();
    }).filter(function(email) {
      return email !== '' && email !== ownerEmail;
    });
    recipients = recipients.concat(extras);
  }

  return recipients;
}

/**
 * メール送信の時間トリガーを設定する（毎朝8時）
 */
function setupEmailTrigger() {
  // 既存のsendDailyVisitReportsトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendDailyVisitReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝8時のトリガーを作成
  ScriptApp.newTrigger('sendDailyVisitReports')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('メール送信トリガーを設定しました（毎朝8時）');
  SpreadsheetApp.getUi().alert('メール送信トリガーを設定しました（毎朝8時）');
}
