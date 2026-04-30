/**
 * F-08: バウンスメール（送信失敗返送）収集
 *
 * GmailのNDR（Non-Delivery Report）を定期検索し、
 * 送信失敗した保護者メールアドレスをバウンスログシートに記録する。
 *
 * 検知できるケース:
 *   - 宛先アドレスが存在しない（550 User unknown 等）
 *   - メールボックスが満杯
 *   - ドメインが存在しない
 * 検知できないケース:
 *   - 受信側がサイレントに破棄（スパムフォルダ含む）
 */

/**
 * Gmail受信箱のNDRメールを検索してバウンスログシートに記録する（自動トリガー用）
 */
function collectBounceEmails() {
  try {
    var lastCheckDate = getLastBounceCheckDate_();
    var afterStr = formatDateYMD_(lastCheckDate);

    // mailer-daemon または postmaster からのメールを検索
    var query = 'from:(mailer-daemon OR postmaster) after:' + afterStr;
    var threads = GmailApp.search(query, 0, 100);

    if (threads.length === 0) {
      Logger.log('バウンスメールなし（検索期間: ' + afterStr + ' 以降）');
      saveLastBounceCheckDate_(new Date());
      return;
    }

    // 保護者メールアドレス → 児童名 のマップを構築
    var emailToChildMap = buildEmailToChildMap_();

    // バウンスログシートを取得（なければ作成）
    var sheet = setupBounceLogSheet_();

    var newCount = 0;

    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        var msg = messages[j];
        var msgDate = msg.getDate();

        // 前回チェック日時より古いメッセージはスキップ
        if (msgDate < lastCheckDate) continue;

        var body = msg.getPlainBody();
        var subject = msg.getSubject();
        var recipient = extractOriginalRecipient_(body, emailToChildMap);

        if (!recipient) continue;

        var childName = emailToChildMap[recipient.toLowerCase()] || '（マスタ未登録）';
        var detectedAt = formatDateYMD_(new Date(), 'yyyy/MM/dd HH:mm');
        var bouncedAt = formatDateYMD_(msgDate, 'yyyy/MM/dd HH:mm');

        sheet.appendRow([detectedAt, bouncedAt, recipient, childName, subject, '未対応']);
        newCount++;
      }
    }

    saveLastBounceCheckDate_(new Date());
    Logger.log('バウンスメール収集完了: ' + newCount + '件を記録');

    if (newCount > 0) {
      notifyBounceDetected_(newCount);
    }
  } catch (error) {
    logError_('collectBounceEmails', error);
  }
}

/**
 * 手動実行用: バウンスメールを即時確認する
 */
function collectBounceEmailsManual() {
  collectBounceEmails();
  SpreadsheetApp.getUi().alert(
    'バウンスメール確認完了',
    '確認が完了しました。\n送信失敗が検出された場合は「バウンスログ」シートに記録されています。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * バウンスログシートを取得する（存在しない場合は作成して返す）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function setupBounceLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.BOUNCE_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BOUNCE_LOG);
    var headers = ['検出日時', 'バウンス発生日時', '送信先メールアドレス', '児童名', 'バウンスメール件名', '対応状況'];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    // バウンスログはエラー系シートのため赤ヘッダー
    styleSheetHeader_(headerRange, 1, { bgColor: '#E53935' });
    setColumnWidths_(sheet, { 1: 150, 2: 150, 3: 220, 4: 120, 5: 300, 6: 100 });
  }

  return sheet;
}

/**
 * NDRメール本文から元の送信先アドレスを抽出する
 *
 * 優先順位:
 *   1. "Final-Recipient: rfc822; address" （SMTP DSN 標準ヘッダー）
 *   2. 既知の保護者メールアドレスとの部分一致（フォールバック）
 *
 * @param {string} body メール本文
 * @param {Object} emailToChildMap メールアドレス → 児童名 のマップ
 * @returns {string|null} 抽出したメールアドレス（見つからなければnull）
 */
function extractOriginalRecipient_(body, emailToChildMap) {
  // 1. SMTP DSN の標準パターン
  var match = body.match(/Final-Recipient:\s*rfc822;\s*([^\s\r\n<>]+)/i);
  if (match) {
    return match[1].toLowerCase();
  }

  // 2. 既知の保護者メールアドレスがメール本文中に含まれていないか確認
  if (emailToChildMap) {
    var knownEmails = Object.keys(emailToChildMap);
    var lowerBody = body.toLowerCase();
    for (var i = 0; i < knownEmails.length; i++) {
      if (lowerBody.indexOf(knownEmails[i]) !== -1) {
        return knownEmails[i];
      }
    }
  }

  return null;
}

/**
 * 児童マスタから 保護者メールアドレス（小文字） → 児童名 のマップを構築する
 * @returns {Object}
 */
function buildEmailToChildMap_() {
  var masterData = getChildMasterData();
  var map = {};
  for (var i = 0; i < masterData.length; i++) {
    var email = masterData[i][MASTER_COL.PARENT_EMAIL - 1];
    var name = masterData[i][MASTER_COL.NAME - 1];
    if (email && String(email).trim()) {
      map[String(email).trim().toLowerCase()] = name;
    }
  }
  return map;
}

/**
 * バウンスが検出されたことを管理者にメール通知する
 * @param {number} count 検出件数
 */
function notifyBounceDetected_(count) {
  try {
    var recipients = getErrorNotifyRecipients_();
    var props = PropertiesService.getScriptProperties();
    var facilityName = props.getProperty('FACILITY_NAME') || '施設';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var subject = '【' + facilityName + '】バウンスメール検出: ' + count + '件';
    var body = count + '件の送信失敗メール（バウンス）が検出されました。\n\n'
      + '保護者のメールアドレスが誤っている可能性があります。\n'
      + '以下のスプレッドシートの「バウンスログ」シートを確認してください。\n\n'
      + ss.getUrl() + '\n\n'
      + '対応完了後は「対応状況」列を「対応済」に更新してください。';

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body,
      name: facilityName + ' システム通知',
    });

    Logger.log('バウンス通知メール送信: ' + recipients.join(', '));
  } catch (e) {
    Logger.log('バウンス通知メール送信に失敗: ' + e.message);
  }
}

/**
 * 前回のバウンスチェック日時をPropertiesServiceから取得する
 * 初回は7日前を返す
 * @returns {Date}
 */
function getLastBounceCheckDate_() {
  var props = PropertiesService.getScriptProperties();
  var lastCheck = props.getProperty('BOUNCE_CHECK_LAST_RUN');
  if (lastCheck) {
    return new Date(parseInt(lastCheck, 10));
  }
  var date = new Date();
  date.setDate(date.getDate() - 7);
  return date;
}

/**
 * バウンスチェック日時をPropertiesServiceに保存する
 * @param {Date} date
 */
function saveLastBounceCheckDate_(date) {
  PropertiesService.getScriptProperties().setProperty(
    'BOUNCE_CHECK_LAST_RUN',
    String(date.getTime())
  );
}

/**
 * フォーム回答のバリデーションを実行する
 *
 * 検知ルール:
 *   1. 入退所いずれかが空欄
 *   2. 入所日時 > 退所日時（時系列が逆転している）
 *   3. 同一児童で重複期間の宿泊がある（既存の宿泊期間に新しい宿泊期間が重なる）
 *
 * @param {number} [year] 対象年（省略時は全期間）
 * @param {number} [month] 対象月 1-12（省略時は全期間）
 * @returns {Array<{childName: string, recordDate: Date, issues: Array<string>}>}
 */
function validateFormResponses_(year, month) {
  var allResponses = getFormResponsesAll_();
  var filterByMonth = (year !== undefined && month !== undefined);
  var monthStart = filterByMonth ? new Date(year, month - 1, 1) : null;
  var monthEnd = filterByMonth ? new Date(year, month, 0) : null;

  // フォーム回答を1行=1宿泊として正規化
  var stays = allResponses.map(function(row) {
    var checkIn = row[FORM_COL.CHECK_IN - 1];
    var checkOut = row[FORM_COL.CHECK_OUT - 1];
    var hasIn = (checkIn instanceof Date) && checkIn.getFullYear() >= 1900;
    var hasOut = (checkOut instanceof Date) && checkOut.getFullYear() >= 1900;
    return {
      childName: row[FORM_COL.CHILD_NAME - 1],
      checkIn: hasIn ? checkIn : null,
      checkOut: hasOut ? checkOut : null,
      recordDate: getRowRecordDate_(row),
      issues: [],
    };
  }).filter(function(s) { return !!s.childName; });

  // 1. 各レコード単体のチェック（空欄・時系列逆転）
  stays.forEach(function(stay) {
    if (!stay.checkIn || !stay.checkOut) {
      stay.issues.push('入所日時または退所日時が空欄');
    } else if (stay.checkIn.getTime() > stay.checkOut.getTime()) {
      stay.issues.push('入所日時が退所日時より後になっている');
    }
  });

  // 2. 児童ごとに宿泊期間の重複をチェック
  var byChild = {};
  stays.forEach(function(stay) {
    if (!stay.childName || !stay.checkIn || !stay.checkOut) return;
    var name = String(stay.childName);
    if (!byChild[name]) byChild[name] = [];
    byChild[name].push(stay);
  });
  Object.keys(byChild).forEach(function(name) {
    var list = byChild[name].slice().sort(function(a, b) {
      return a.checkIn.getTime() - b.checkIn.getTime();
    });
    for (var i = 0; i < list.length - 1; i++) {
      var cur = list[i];
      var nxt = list[i + 1];
      if (cur.checkOut.getTime() > nxt.checkIn.getTime()) {
        cur.issues.push('別の宿泊（' + formatDateYMD_(nxt.checkIn) + '〜）と期間が重複');
      }
    }
  });

  // 3. 月フィルタ + 結果整形
  var issues = [];
  stays.forEach(function(stay) {
    if (!stay.issues || stay.issues.length === 0) return;
    if (filterByMonth) {
      var spanStart = stay.checkIn || stay.recordDate;
      var spanEnd = stay.checkOut || stay.recordDate;
      if (!(spanStart instanceof Date) || !(spanEnd instanceof Date)) return;
      if (spanEnd < monthStart || spanStart > monthEnd) return;
    }
    issues.push({
      childName: stay.childName,
      recordDate: stay.recordDate,
      issues: stay.issues.slice(),
    });
  });

  return issues;
}

/**
 * 旧API互換: validateFormResponses_ のラッパー
 * @deprecated validateFormResponses_ を直接呼ぶこと
 */
function validateOvernightRecords(year, month) {
  return validateFormResponses_(year, month);
}

/**
 * フォーム回答バリデーション結果をログに書き出して通知する（手動実行用）
 */
function runFormResponseValidationManual() {
  return runOvernightValidationManualImpl_();
}

/**
 * @deprecated 旧API互換: 既存メニューバインディングを壊さないため残置
 */
function runOvernightValidationManual() {
  return runOvernightValidationManualImpl_();
}

function runOvernightValidationManualImpl_() {
  var issues = validateFormResponses_();
  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert('フォーム回答検証', '問題は検出されませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var lines = issues.map(function(it) {
    return '・' + formatDateYMD_(it.recordDate) + ' ' + String(it.childName || '(児童名なし)') + ': ' + it.issues.join(' / ');
  });
  SpreadsheetApp.getUi().alert('フォーム回答検証（' + issues.length + '件）', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * バウンスチェックの時間トリガーを設定する（毎日9時）
 */
function setupBounceCheckTrigger() {
  setupTimeTrigger_('collectBounceEmails', { everyDays: 1, atHour: 9 });
  Logger.log('バウンスチェックトリガーを設定しました（毎日9時）');
}
