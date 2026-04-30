/**
 * フォームのドロップダウンをマスタデータに同期する
 */

// フォームの質問タイトル（実際の質問文と異なる場合は修正すること）
var FORM_TITLE_STAFF1 = 'スタッフ1';
var FORM_TITLE_STAFF2 = 'スタッフ2';
var FORM_TITLE_CHILD = '児童名';
var FORM_TITLE_CHECK_IN = '入所日時';
var FORM_TITLE_CHECK_OUT_OLD = '退所日時';
var FORM_TITLE_CHECK_OUT = '退所予定日時';

// 廃止項目（同期実行時にフォームから削除する）
var FORM_TITLE_RECORD_DATE = '記録日';

// 旧「睡眠」→新3項目（入眠時刻 / 朝4時チェック / 起床時刻）
var FORM_TITLE_SLEEP_OLD = '睡眠';
var FORM_TITLE_BATH = '入浴';
var FORM_TITLE_SLEEP_ONSET = '入眠時刻';
var FORM_TITLE_SLEEP_CHECK_4AM = '朝4時チェック';
var FORM_TITLE_WAKE_UP = '起床時刻';

var SLEEP_ONSET_CHOICES = ['20:30', '20:40', '20:50', '21:00', '21:10', '21:20', '21:30', '21:40', '21:50', '22:00'];
var SLEEP_CHECK_4AM_CHOICES = ['睡眠', '覚醒確認後に付き添い'];
var WAKE_UP_CHOICES = ['6:00', '6:10', '6:20', '6:30', '6:40', '6:50', '7:00', '7:10', '7:20', '7:30'];

/**
 * フォームのドロップダウンをマスタデータで更新する
 * 手動実行・トリガーから呼び出す
 */
function syncFormDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formUrl = ss.getFormUrl();
  if (!formUrl) {
    Logger.log('リンクされたフォームが見つかりません');
    return;
  }

  var form = FormApp.openByUrl(formUrl);
  var staffNames = getStaffNamesFromMaster_(ss);
  var childNames = getChildNamesFromMaster_(ss);

  if (!staffNames.length || !childNames.length) {
    Logger.log('マスタデータが空のため同期をスキップしました');
    return;
  }

  form.getItems().forEach(function(item) {
    var title = item.getTitle();
    if (title === FORM_TITLE_STAFF1 || title === FORM_TITLE_STAFF2) {
      setListChoices_(item, staffNames);
    } else if (title === FORM_TITLE_CHILD) {
      setListChoices_(item, childNames);
    } else if (title === FORM_TITLE_CHECK_OUT_OLD) {
      // 旧名「退所日時」→「退所予定日時」へリネーム（一度だけ実行される）
      item.setTitle(FORM_TITLE_CHECK_OUT);
    }
  });

  // 「睡眠」項目を入眠時刻 / 朝4時チェック / 起床時刻 の3項目に置き換える
  ensureSleepThreeItems_(form);

  // 「記録日」項目は廃止（タイムスタンプの日付を利用日として扱う）
  removeRecordDateField_(form);

  Logger.log('同期完了: スタッフ ' + staffNames.length + '名 / 児童 ' + childNames.length + '名');
}

/**
 * フォームから「記録日」項目を削除する。既に削除済みなら何もしない。
 * @param {GoogleAppsScript.Forms.Form} form
 */
function removeRecordDateField_(form) {
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === FORM_TITLE_RECORD_DATE) {
      form.deleteItem(items[i]);
      Logger.log('「記録日」項目をフォームから削除しました');
      return;
    }
  }
}

/**
 * フォームの「睡眠」項目を3項目（入眠時刻 / 朝4時チェック / 起床時刻）に置き換える。
 * - 既に3項目がそろっていれば選択肢の整合のみ取って終了（再実行で上書きしない）
 * - 旧「睡眠」項目があれば「入浴」の直後の位置に新3項目を挿入してから旧項目を削除する
 * - 新3項目はすべてプルダウン（LIST）・必須
 * @param {GoogleAppsScript.Forms.Form} form
 */
function ensureSleepThreeItems_(form) {
  var items = form.getItems();
  var existing = {};
  var oldSleepIndex = -1;
  var bathIndex = -1;

  items.forEach(function(item, idx) {
    var t = item.getTitle();
    if (t === FORM_TITLE_SLEEP_ONSET) existing.onset = item;
    else if (t === FORM_TITLE_SLEEP_CHECK_4AM) existing.check4am = item;
    else if (t === FORM_TITLE_WAKE_UP) existing.wakeUp = item;
    else if (t === FORM_TITLE_SLEEP_OLD) oldSleepIndex = idx;
    if (t === FORM_TITLE_BATH) bathIndex = idx;
  });

  // 既に3項目すべてある場合は選択肢・必須・タイプのみ整える
  if (existing.onset && existing.check4am && existing.wakeUp) {
    applySleepListItem_(existing.onset, SLEEP_ONSET_CHOICES);
    applySleepListItem_(existing.check4am, SLEEP_CHECK_4AM_CHOICES);
    applySleepListItem_(existing.wakeUp, WAKE_UP_CHOICES);
    return;
  }

  // 挿入位置: 旧「睡眠」があればその位置、無ければ「入浴」の直後、それも無ければ末尾
  var insertIndex;
  if (oldSleepIndex >= 0) {
    insertIndex = oldSleepIndex;
  } else if (bathIndex >= 0) {
    insertIndex = bathIndex + 1;
  } else {
    insertIndex = form.getItems().length;
  }

  // 不足分を追加（既存があればスキップ）
  if (!existing.onset) {
    var onset = form.addListItem();
    onset.setTitle(FORM_TITLE_SLEEP_ONSET).setRequired(true).setChoiceValues(SLEEP_ONSET_CHOICES);
    form.moveItem(onset.getIndex(), insertIndex);
    insertIndex++;
  }
  if (!existing.check4am) {
    var c4 = form.addListItem();
    c4.setTitle(FORM_TITLE_SLEEP_CHECK_4AM).setRequired(true).setChoiceValues(SLEEP_CHECK_4AM_CHOICES);
    form.moveItem(c4.getIndex(), insertIndex);
    insertIndex++;
  }
  if (!existing.wakeUp) {
    var wu = form.addListItem();
    wu.setTitle(FORM_TITLE_WAKE_UP).setRequired(true).setChoiceValues(WAKE_UP_CHOICES);
    form.moveItem(wu.getIndex(), insertIndex);
    insertIndex++;
  }

  // 旧「睡眠」を削除（最後に削除することで挿入位置がずれない）
  if (oldSleepIndex >= 0) {
    var refreshed = form.getItems();
    for (var i = 0; i < refreshed.length; i++) {
      if (refreshed[i].getTitle() === FORM_TITLE_SLEEP_OLD) {
        form.deleteItem(refreshed[i]);
        break;
      }
    }
  }
}

/**
 * 既存のフォーム項目をプルダウン・必須・指定選択肢に整える。
 * 項目タイプがLIST以外の場合は警告ログのみ（GASでは型変更不可のため手動修正を促す）。
 */
function applySleepListItem_(item, choices) {
  var type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(choices).setRequired(true);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(choices).setRequired(true);
  } else {
    Logger.log('警告: 「' + item.getTitle() + '」の型がLIST/MULTIPLE_CHOICEではありません。手動でプルダウンに変更してください');
  }
}

/**
 * スタッフマスタからスタッフ名一覧を取得する（B列・重複除去）
 */
function getStaffNamesFromMaster_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.STAFF_MASTER);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAMES.STAFF_MASTER);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    if (name && !seen[name]) {
      seen[name] = true;
      names.push(name);
    }
  }
  return names;
}

/**
 * 児童マスタから児童名一覧を取得する（稼働・休止のみ・重複除去）
 */
function getChildNamesFromMaster_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.CHILD_MASTER);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + SHEET_NAMES.CHILD_MASTER);
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var seen = {};
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][MASTER_COL.NAME - 1] || '').trim();
    var status = String(data[i][MASTER_COL.ENROLLMENT - 1] || '').trim();
    if (name && (status === '稼働' || status === '休止') && !seen[name]) {
      seen[name] = true;
      names.push(name);
    }
  }
  return names;
}

/**
 * フォームアイテムの選択肢を更新する（リスト・ラジオ・チェックボックス対応）
 */
function setListChoices_(item, choices) {
  var type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(choices);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(choices);
  } else if (type === FormApp.ItemType.CHECKBOX) {
    item.asCheckboxItem().setChoiceValues(choices);
  }
}

/**
 * 時間ベーストリガーを設定する（毎日AM1時に自動同期）
 * 一度だけ手動で実行してください
 */
function setupFormSyncTrigger() {
  setupTimeTrigger_('syncFormDropdowns', { everyDays: 1, atHour: 1 });
  Logger.log('フォーム同期トリガー設定完了（毎日AM1時）');
}
