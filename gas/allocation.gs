/**
 * F-05 / F-06: 余りポイント自動振り分け
 * 前月の来館記録から余りポイントを算出し、未来館日に自動振り分けする
 */

/**
 * 振り分けを手動実行する（F-06）
 * 月別集計シートのB1セル（対象年月）を参照して振り分けを実行する
 */
function runAllocationManual() {
  var sheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
  var yearMonthStr = sheet.getRange('B1').getValue();
  if (!yearMonthStr) {
    SpreadsheetApp.getUi().alert('月別集計シートの対象年月を選択してください');
    return;
  }
  var ym = parseYearMonth(yearMonthStr);
  allocateRemainingPoints_(ym.year, ym.month);
  SpreadsheetApp.getUi().alert(ym.year + '年' + ym.month + '月の振り分けが完了しました');
}

/**
 * 振り分けを月初自動実行する（F-05）
 * 前月を対象として振り分けを実行する
 */
function runAllocationAutomatic() {
  var now = new Date();
  var prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var year = prevMonth.getFullYear();
  var month = prevMonth.getMonth() + 1;

  try {
    allocateRemainingPoints_(year, month);
    Logger.log('月初自動振り分け完了: ' + year + '年' + month + '月');
  } catch (error) {
    Logger.log('月初自動振り分けでエラー: ' + error.message);
  }
}

/**
 * 月初自動振り分けトリガーを設定する（手動で1回実行）
 */
function setupAllocationTrigger() {
  // 既存のrunAllocationAutomaticトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runAllocationAutomatic') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎月1日の午前2時に実行
  ScriptApp.newTrigger('runAllocationAutomatic')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();

  Logger.log('月初自動振り分けトリガーを設定しました');
  SpreadsheetApp.getUi().alert('月初自動振り分けトリガーを設定しました（毎月1日 午前2時）');
}

/**
 * 余りポイント振り分けのメインロジック
 * @param {number} year 対象年
 * @param {number} month 対象月（1-12）
 */
function allocateRemainingPoints_(year, month) {
  // 1. 入所状況「稼働」の児童を取得
  var masterData = getChildMasterData();
  var activeChildren = masterData.filter(function(row) {
    return row[MASTER_COL.ENROLLMENT - 1] === '稼働';
  });

  if (activeChildren.length === 0) {
    Logger.log('振り分け対象の児童がいません');
    return;
  }

  // 2. フォーム回答から対象月の実来館データ取得
  var formResponses = getFormResponsesByMonth(year, month);

  // 3. 児童名ごとの実来館回数と来館日マップを作成
  var visitCountMap = {};
  var visitDateMap = {};  // {児童名: {日付文字列: true}}
  formResponses.forEach(function(row) {
    var childName = row[FORM_COL.CHILD_NAME - 1];
    if (!childName) return;
    visitCountMap[childName] = (visitCountMap[childName] || 0) + 1;
    if (!visitDateMap[childName]) visitDateMap[childName] = {};
    var dateKey = formatDateKey_(new Date(row[FORM_COL.RECORD_DATE - 1]));
    visitDateMap[childName][dateKey] = true;
  });

  // 4. 残枠がある児童を抽出し、優先度順にソート
  var childrenWithRemaining = activeChildren.filter(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
    if (quota <= 0) return false;
    var visits = visitCountMap[childName] || 0;
    return (quota - visits) > 0;
  });

  childrenWithRemaining.sort(function(a, b) {
    var priorityA = a[MASTER_COL.PRIORITY - 1] || 9999;
    var priorityB = b[MASTER_COL.PRIORITY - 1] || 9999;
    return priorityA - priorityB;
  });

  if (childrenWithRemaining.length === 0) {
    Logger.log('残枠のある児童がいません。振り分け不要です。');
    clearAllocationsForMonth_(year, month);
    updateConfirmedVisits();
    updateMonthlySummary();
    return;
  }

  // 5. 対象月の全日付を取得
  var allDates = getAllDatesInMonth_(year, month);

  // 6. 各日付の既存来館数マップを作成（実データのみ。振り分けは洗い替えするため含めない）
  var dailyVisitCounts = {};
  allDates.forEach(function(date) {
    dailyVisitCounts[formatDateKey_(date)] = 0;
  });
  formResponses.forEach(function(row) {
    var dateKey = formatDateKey_(new Date(row[FORM_COL.RECORD_DATE - 1]));
    if (dailyVisitCounts[dateKey] !== undefined) {
      dailyVisitCounts[dateKey]++;
    }
  });

  // 7. 振り分け記録シートから対象月の既存データを削除（洗い替え）
  clearAllocationsForMonth_(year, month);

  // 8. 各児童の補完データを事前計算
  var childDefaultsMap = {};
  activeChildren.forEach(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    childDefaultsMap[childName] = computeChildDefaults_(childName, row, formResponses);
  });

  // 9. 児童ごとに振り分けを実行
  var allocationResults = [];
  var executedAt = new Date();

  childrenWithRemaining.forEach(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
    var actualVisits = visitCountMap[childName] || 0;
    var remaining = quota - actualVisits;
    var childVisitDates = visitDateMap[childName] || {};
    var visitDayStr = row[MASTER_COL.VISIT_DAYS - 1];

    // 来館曜日を数値に変換
    var visitDayNumbers = parseVisitDays_(visitDayStr);

    // 候補日を作成: 来館済み・同一児童重複を除外
    var preferredDates = [];   // 来館曜日に該当する候補日
    var otherDates = [];       // その他の候補日

    allDates.forEach(function(date) {
      var dateKey = formatDateKey_(date);
      // 既に来館済みの日は除外
      if (childVisitDates[dateKey]) return;
      // 7名満枠の日は除外
      if (dailyVisitCounts[dateKey] >= MAX_VISITS_PER_DAY) return;

      if (visitDayNumbers.length > 0 && visitDayNumbers.indexOf(date.getDay()) !== -1) {
        preferredDates.push(date);
      } else {
        otherDates.push(date);
      }
    });

    // 来館曜日優先 → その他で足りない分を補う
    var allocated = 0;
    var candidatePools = [preferredDates, otherDates];

    for (var poolIdx = 0; poolIdx < candidatePools.length && allocated < remaining; poolIdx++) {
      var pool = candidatePools[poolIdx];

      while (allocated < remaining && pool.length > 0) {
        // 来館数が最も少ない日を選択（均等分散）
        pool.sort(function(a, b) {
          return dailyVisitCounts[formatDateKey_(a)] - dailyVisitCounts[formatDateKey_(b)];
        });

        var selectedDate = pool.shift();
        var selectedKey = formatDateKey_(selectedDate);

        // 再度上限チェック（他の児童の振り分けで埋まっている可能性）
        if (dailyVisitCounts[selectedKey] >= MAX_VISITS_PER_DAY) continue;

        // 振り分け確定
        var defaults = childDefaultsMap[childName];
        allocationResults.push([
          new Date(year, month - 1, 1),  // 対象年月
          childName,                      // 児童名
          selectedDate,                   // 振り分け日
          defaults.staffName,             // スタッフ名
          defaults.checkIn,               // 入所時間
          defaults.checkOut,              // 退所時間
          defaults.temperature,           // 体温
          defaults.meal,                  // 食事
          defaults.bath,                  // 入浴
          defaults.sleep,                 // 睡眠
          defaults.bowel,                 // 便
          defaults.medicine,              // 服薬
          defaults.notes,                 // その他連絡事項
          executedAt,                     // 実行日時
        ]);

        dailyVisitCounts[selectedKey]++;
        allocated++;
      }
    }

    if (allocated < remaining) {
      Logger.log('警告: ' + childName + ' の振り分けが不完全です（残り' + (remaining - allocated) + '枠分の空きがありません）');
    }
  });

  // 10. 振り分け記録シートに書き込み
  if (allocationResults.length > 0) {
    writeAllocations_(allocationResults);
    Logger.log('振り分け完了: ' + allocationResults.length + '件');
  } else {
    Logger.log('振り分け結果: 0件（振り分け先がありません）');
  }

  // 11. 確定来館記録を再生成
  updateConfirmedVisits();

  // 12. 月別集計を更新
  updateMonthlySummary();
}

/**
 * 対象月の振り分け記録を削除する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 */
function clearAllocationsForMonth_(year, month) {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.ALLOCATION);
  } catch (e) {
    return;
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // 下の行から削除（行番号のずれを防ぐ）
  for (var i = data.length - 1; i >= 1; i--) {
    var targetMonth = new Date(data[i][ALLOCATION_COL.TARGET_MONTH - 1]);
    if (targetMonth.getFullYear() === year && (targetMonth.getMonth() + 1) === month) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * 振り分け結果を振り分け記録シートに書き込む
 * @param {Array<Array>} results 振り分け結果の2次元配列（14列）
 */
function writeAllocations_(results) {
  var sheet = getSheet(SHEET_NAMES.ALLOCATION);
  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1;

  sheet.getRange(startRow, 1, results.length, ALLOCATION_COL_COUNT).setValues(results);

  // 日付列の表示形式
  sheet.getRange(startRow, ALLOCATION_COL.TARGET_MONTH, results.length, 1)
    .setNumberFormat('yyyy/mm');
  sheet.getRange(startRow, ALLOCATION_COL.ALLOCATION_DATE, results.length, 1)
    .setNumberFormat('yyyy/mm/dd');
  // 入所時間・退所時間
  sheet.getRange(startRow, ALLOCATION_COL.CHECK_IN, results.length, 2)
    .setNumberFormat('H:mm');
  // 体温
  sheet.getRange(startRow, ALLOCATION_COL.TEMPERATURE, results.length, 1)
    .setNumberFormat('0.0');
  // 実行日時
  sheet.getRange(startRow, ALLOCATION_COL.EXECUTED_AT, results.length, 1)
    .setNumberFormat('yyyy/mm/dd H:mm:ss');
}

/**
 * 指定月の全日付を配列で返す
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {Array<Date>} 日付の配列
 */
function getAllDatesInMonth_(year, month) {
  var dates = [];
  var daysInMonth = new Date(year, month, 0).getDate();
  for (var d = 1; d <= daysInMonth; d++) {
    dates.push(new Date(year, month - 1, d));
  }
  return dates;
}

/**
 * 来館曜日文字列をDay番号の配列に変換する
 * @param {string} visitDayStr 来館曜日（例: "月" または "月,水,金"）
 * @returns {Array<number>} Day番号の配列（0=日, 1=月, ... 6=土）
 */
function parseVisitDays_(visitDayStr) {
  if (!visitDayStr) return [];
  var days = String(visitDayStr).split(',');
  var result = [];
  days.forEach(function(day) {
    var trimmed = day.trim();
    if (DAY_OF_WEEK_MAP[trimmed] !== undefined) {
      result.push(DAY_OF_WEEK_MAP[trimmed]);
    }
  });
  return result;
}

/**
 * 日付をYYYY-MM-DD形式の文字列に変換する（比較用キー）
 * @param {Date} date 日付
 * @returns {string} YYYY-MM-DD形式
 */
function formatDateKey_(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * 児童の振り分け補完データを実データから算出する
 * @param {string} childName 児童名
 * @param {Array} masterRow 児童マスタの行データ
 * @param {Array<Array>} formResponses 同月のフォーム回答データ
 * @returns {Object} 補完データ
 */
function computeChildDefaults_(childName, masterRow, formResponses) {
  // スタッフ名は児童マスタから取得
  var staffName = masterRow[MASTER_COL.STAFF - 1] || '';

  // 同じ児童の実データを抽出
  var childData = formResponses.filter(function(row) {
    return row[FORM_COL.CHILD_NAME - 1] === childName;
  });

  // 実データがある場合は最頻値を算出
  if (childData.length > 0) {
    return {
      staffName: staffName,
      checkIn: getModeValue_(childData, FORM_COL.CHECK_IN - 1, ALLOCATION_DEFAULTS.CHECK_IN),
      checkOut: getModeValue_(childData, FORM_COL.CHECK_OUT - 1, ALLOCATION_DEFAULTS.CHECK_OUT),
      temperature: getModeNumeric_(childData, FORM_COL.TEMPERATURE - 1, ALLOCATION_DEFAULTS.TEMPERATURE),
      meal: getModeValue_(childData, FORM_COL.MEAL - 1, ALLOCATION_DEFAULTS.MEAL),
      bath: getModeValue_(childData, FORM_COL.BATH - 1, ALLOCATION_DEFAULTS.BATH),
      sleep: getModeValue_(childData, FORM_COL.SLEEP - 1, ALLOCATION_DEFAULTS.SLEEP),
      bowel: getModeValue_(childData, FORM_COL.BOWEL - 1, ALLOCATION_DEFAULTS.BOWEL),
      medicine: getModeValue_(childData, FORM_COL.MEDICINE - 1, ALLOCATION_DEFAULTS.MEDICINE),
      notes: pickRandomNote_(childName, childData, formResponses),
    };
  }

  // 実データがない場合はデフォルト値を使用
  return {
    staffName: staffName,
    checkIn: ALLOCATION_DEFAULTS.CHECK_IN,
    checkOut: ALLOCATION_DEFAULTS.CHECK_OUT,
    temperature: ALLOCATION_DEFAULTS.TEMPERATURE,
    meal: ALLOCATION_DEFAULTS.MEAL,
    bath: ALLOCATION_DEFAULTS.BATH,
    sleep: ALLOCATION_DEFAULTS.SLEEP,
    bowel: ALLOCATION_DEFAULTS.BOWEL,
    medicine: ALLOCATION_DEFAULTS.MEDICINE,
    notes: pickRandomNote_(childName, [], formResponses),
  };
}

/**
 * 配列の指定列から最頻値を返す
 * @param {Array<Array>} data データ配列
 * @param {number} colIndex 列インデックス（0始まり）
 * @param {*} defaultValue デフォルト値
 * @returns {*} 最頻値
 */
function getModeValue_(data, colIndex, defaultValue) {
  var counts = {};
  data.forEach(function(row) {
    var val = row[colIndex];
    if (val === '' || val === null || val === undefined) return;
    // 時刻型の場合は文字列キーに変換
    var key = (val instanceof Date) ? formatTimeKey_(val) : String(val);
    counts[key] = (counts[key] || 0) + 1;
  });

  var maxCount = 0;
  var modeKey = null;
  Object.keys(counts).forEach(function(key) {
    if (counts[key] > maxCount) {
      maxCount = counts[key];
      modeKey = key;
    }
  });

  if (modeKey === null) return defaultValue;

  // 元の値を返す（時刻型の最頻値キーから最初に一致する元の値を返す）
  for (var i = 0; i < data.length; i++) {
    var val = data[i][colIndex];
    if (val === '' || val === null || val === undefined) continue;
    var key = (val instanceof Date) ? formatTimeKey_(val) : String(val);
    if (key === modeKey) return val;
  }
  return defaultValue;
}

/**
 * 数値列の最頻値を返す（小数1桁に丸めて集計）
 * @param {Array<Array>} data データ配列
 * @param {number} colIndex 列インデックス（0始まり）
 * @param {number} defaultValue デフォルト値
 * @returns {number} 最頻値
 */
function getModeNumeric_(data, colIndex, defaultValue) {
  var counts = {};
  data.forEach(function(row) {
    var val = row[colIndex];
    if (val === '' || val === null || val === undefined || isNaN(val)) return;
    // 小数1桁に丸めてキーにする
    var key = (Math.round(val * 10) / 10).toFixed(1);
    counts[key] = (counts[key] || 0) + 1;
  });

  var maxCount = 0;
  var modeKey = null;
  Object.keys(counts).forEach(function(key) {
    if (counts[key] > maxCount) {
      maxCount = counts[key];
      modeKey = key;
    }
  });

  return modeKey !== null ? parseFloat(modeKey) : defaultValue;
}

/**
 * 連絡事項をランダムに選択する
 * 優先順: 自児童のノート → 全児童のノート → デフォルト
 * @param {string} childName 児童名
 * @param {Array<Array>} childData 自児童の実データ
 * @param {Array<Array>} allData 全児童の実データ
 * @returns {string} 連絡事項テキスト
 */
function pickRandomNote_(childName, childData, allData) {
  // 自児童のノートを収集（空でないもの）
  var childNotes = collectNotes_(childData);
  if (childNotes.length > 0) {
    return childNotes[Math.floor(Math.random() * childNotes.length)];
  }

  // 他児童のノートを収集
  var allNotes = collectNotes_(allData);
  if (allNotes.length > 0) {
    return allNotes[Math.floor(Math.random() * allNotes.length)];
  }

  return ALLOCATION_DEFAULTS.NOTES;
}

/**
 * データ配列から空でない連絡事項を収集する
 * @param {Array<Array>} data データ配列
 * @returns {Array<string>} ノートの配列
 */
function collectNotes_(data) {
  var notes = [];
  data.forEach(function(row) {
    var note = row[FORM_COL.NOTES - 1];
    if (note && String(note).trim() !== '') {
      notes.push(String(note).trim());
    }
  });
  return notes;
}

/**
 * Date型の時刻をHH:mm形式の文字列に変換する
 * @param {Date} date 日時
 * @returns {string} HH:mm形式
 */
function formatTimeKey_(date) {
  var h = ('0' + date.getHours()).slice(-2);
  var m = ('0' + date.getMinutes()).slice(-2);
  return h + ':' + m;
}
