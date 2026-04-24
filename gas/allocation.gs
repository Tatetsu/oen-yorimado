/**
 * F-05 / F-06: 余りポイント自動振り分け
 * 前月の来館記録から余りポイントを算出し、未来館日に自動振り分けする
 * 振り分け結果は確定来館記録シートに直接書き込む（データ区分=「振り分け」）
 */

/**
 * 振り分けを手動実行する（F-06）
 * 月別集計シートのB1=対象年、B2=対象月を参照して振り分けを実行する
 */
function runAllocationManual() {
  var ui = SpreadsheetApp.getUi();

  try {
    var sheet = getSheet(SHEET_NAMES.MONTHLY_SUMMARY);
    var year = parseYearOption_(sheet.getRange('B1').getValue());
    var month = parseMonthOption_(sheet.getRange('B2').getValue());
    if (year === null || month === null) {
      ui.alert('月別集計シートの対象年・対象月を具体的に選択してください（「すべて」は不可）');
      return;
    }
    var ym = { year: year, month: month };

    // 既に振り分け済みかチェック
    if (hasAllocationsForMonth_(ym.year, ym.month)) {
      var response = ui.alert(
        '確認',
        ym.year + '年' + ym.month + '月は既に振り分け済みです。\n' +
        '再実行すると手動修正を含む既存の振り分けデータが全て上書きされます。\n\n' +
        '本当に再実行しますか？',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        return;
      }
    }

    allocateRemainingPoints_(ym.year, ym.month);
    ui.alert(ym.year + '年' + ym.month + '月の振り分けが完了しました');
  } catch (error) {
    logError_('runAllocationManual', error);
    ui.alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * 振り分けを月初自動実行する（F-05）
 * 前月を対象として振り分けを実行する
 */
function runAllocationAutomatic() {
  var now = new Date();
  var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  var year = yesterday.getFullYear();
  var month = yesterday.getMonth() + 1;

  try {
    allocateRemainingPoints_(year, month);
    Logger.log('月初自動振り分け完了: ' + year + '年' + month + '月');
  } catch (error) {
    logError_('runAllocationAutomatic', error);
  }
}

/**
 * 月初自動振り分けトリガーを設定する（手動で1回実行）
 */
function setupAllocationTrigger() {
  setupTimeTrigger_('runAllocationAutomatic', { onMonthDay: 1, atHour: 2 });
  Logger.log('月初自動振り分けトリガーを設定しました');
  SpreadsheetApp.getUi().alert('月初自動振り分けトリガーを設定しました（毎月1日 午前2時）');
}

/**
 * 余りポイント振り分けのメインロジック
 * @param {number} year 対象年
 * @param {number} month 対象月（1-12）
 */
function allocateRemainingPoints_(year, month) {
  // 0. 確定来館記録を対象月の実データで最新化（他の月はそのまま保持）
  updateConfirmedVisits(year, month);

  // 1. 振り分け対象の児童を取得
  //    - 稼働: 通常の振り分け対象
  //    - 退所（別施設移動無）: 退所月の残枠を振り分け対象にする
  //    - 退所（別施設移動）: 振り分け対象外
  var masterData = getChildMasterData();
  var activeChildren = masterData.filter(function(row) {
    var enrollment = row[MASTER_COL.ENROLLMENT - 1];
    if (enrollment === '稼働') return true;
    if (enrollment === '退所') {
      var departureStatus = row[MASTER_COL.DEPARTURE_STATUS - 1];
      return departureStatus === '別施設移動無';
    }
    return false;
  });

  if (activeChildren.length === 0) {
    Logger.log('振り分け対象の児童がいません');
    return;
  }

  // 2. フォーム回答から対象月の実来館データ取得（連泊ペアリング後）
  var formResponses = getFormResponsesByMonth(year, month);
  var stays = pairOvernightRecords_(formResponses);

  // 3. 児童名ごとの実来館回数と来館日マップを作成（対象月の日数のみカウント・月またぎ対応）
  var visitCountMap = {};
  var visitDateMap = {};  // {児童名: {日付文字列: true}}
  stays.forEach(function(stay) {
    var childName = stay.childName;
    if (!childName) return;
    if (!visitDateMap[childName]) visitDateMap[childName] = {};
    expandStayToDates_(stay.recordDate, stay.checkIn, stay.checkOut).forEach(function(d) {
      // 対象月の日付のみカウント
      if (d.getFullYear() === year && (d.getMonth() + 1) === month) {
        var dateKey = formatDateKey_(d);
        if (!visitDateMap[childName][dateKey]) {
          visitDateMap[childName][dateKey] = true;
          visitCountMap[childName] = (visitCountMap[childName] || 0) + 1;
        }
      }
    });
  });

  // 4. 残枠がある児童を抽出し、優先度順にソート
  var childrenWithRemaining = activeChildren.filter(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
    if (quota <= 0) return false;
    var actualVisits = visitCountMap[childName] || 0;
    // 実来館2日未満は振り分け対象外（退所フラグ付け忘れによる誤振り分けを防止）
    if (actualVisits < 2) {
      Logger.log(childName + ': 対象月実来館' + actualVisits + '日（2日未満のため振り分け対象外）');
      return false;
    }
    return (quota - actualVisits) > 0;
  });

  childrenWithRemaining.sort(function(a, b) {
    var priorityA = parsePriority_(a[MASTER_COL.PRIORITY - 1]);
    var priorityB = parsePriority_(b[MASTER_COL.PRIORITY - 1]);
    return priorityA - priorityB;
  });

  // 5. 確定来館記録から対象月の既存振り分け行を削除（洗い替え）
  clearAllocationsForMonth_(year, month);

  if (childrenWithRemaining.length === 0) {
    Logger.log('残枠のある児童がいません。振り分け不要です。');
    updateMonthlySummary();
    return;
  }

  // 6. 対象月の全日付と1日最大来館数を取得
  var allDates = getAllDatesInMonth_(year, month);
  var maxVisitsPerDay = getMaxVisitsPerDay_();

  // 7. 各日付の新着者数マップを作成（満枠判定用）
  //    満枠=1日最大来館数は「その日の新着者（入所日=その日の宿泊）」を対象とする。
  //    連泊2日目（持ち越し）は新着者ではないためカウントしない。
  //    （月間利用枠の消費は別途 visitCountMap で連泊展開済み）
  var dailyVisitCounts = {};
  allDates.forEach(function(date) {
    dailyVisitCounts[formatDateKey_(date)] = 0;
  });
  stays.forEach(function(stay) {
    var arrival = stay.checkIn instanceof Date ? stay.checkIn : (stay.recordDate instanceof Date ? stay.recordDate : null);
    if (!arrival || arrival.getFullYear() < 1900) return;
    if (arrival.getFullYear() !== year || (arrival.getMonth() + 1) !== month) return;
    var dateKey = formatDateKey_(new Date(arrival.getFullYear(), arrival.getMonth(), arrival.getDate()));
    if (dailyVisitCounts[dateKey] !== undefined) {
      dailyVisitCounts[dateKey]++;
    }
  });

  // 8. 各児童の補完データを事前計算
  var childDefaultsMap = {};
  activeChildren.forEach(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    childDefaultsMap[childName] = computeChildDefaults_(childName, row, formResponses);
  });

  // 年間利用枠チェック用: 当年の確定来館記録（実データ+振り分け）を集計
  var ytdVisitMap = buildYtdVisitMap_(year);

  // 9. 児童ごとに振り分けを実行
  var allocationResults = [];

  childrenWithRemaining.forEach(function(row) {
    var childName = row[MASTER_COL.NAME - 1];
    var quota = row[MASTER_COL.MONTHLY_QUOTA - 1] || 0;
    var actualVisits = visitCountMap[childName] || 0;

    // 月間利用枠に ±1 のランダム幅を加える（年間累計超過は後段の annualRemaining チェックで制御）
    var delta = Math.floor(Math.random() * 3) - 1; // -1, 0, +1
    var effectiveQuota = Math.max(0, quota + delta);
    var remaining = effectiveQuota - actualVisits;

    // 年間利用枠による上限チェック
    var annualQuota = row[MASTER_COL.ANNUAL_QUOTA - 1];
    if (annualQuota && annualQuota > 0) {
      var ytdVisits = ytdVisitMap[childName] || 0;
      var annualRemaining = annualQuota - ytdVisits;
      if (annualRemaining <= 0) {
        Logger.log(childName + ': 年間利用枠（' + annualQuota + '日）に達しているため振り分けスキップ');
        return;
      }
      // 月次上限: 月間利用枠±1 と 年間枠の平日按分 の小さい方を採用
      var monthlyMax = calcMonthlyQuota_(annualQuota, year, month);
      var effectiveRemaining = Math.min(effectiveQuota, monthlyMax) - actualVisits;
      if (effectiveRemaining <= 0) {
        Logger.log(childName + ': 月次上限（月間枠' + effectiveQuota + '日 / 年間按分' + monthlyMax + '日）に達しているため振り分けスキップ');
        return;
      }
      remaining = Math.min(effectiveRemaining, annualRemaining);
    }
    var childVisitDates = visitDateMap[childName] || {};
    var visitDayStr = row[MASTER_COL.VISIT_DAYS - 1];
    var nonVisitDayStr = row[MASTER_COL.NON_VISIT_DAYS - 1];

    // 来館曜日・非来館曜日を数値に変換
    var visitDayNumbers = parseVisitDays_(visitDayStr);
    var nonVisitDayNumbers = parseVisitDays_(nonVisitDayStr);

    // 候補日を作成: 来館済み・同一児童重複を除外
    var preferredDates = [];   // 来館曜日に該当する候補日
    var otherDates = [];       // その他の候補日
    var businessDayNumbers = getBusinessDays();

    allDates.forEach(function(date) {
      var dateKey = formatDateKey_(date);
      // 既に来館済みの日は除外
      if (childVisitDates[dateKey]) return;
      // その日の新着者数が 1日最大来館数 に達していたらスキップ
      if (dailyVisitCounts[dateKey] >= maxVisitsPerDay) return;
      // 営業日以外は除外（設定シートに営業日が設定されている場合）
      if (businessDayNumbers.length > 0 && businessDayNumbers.indexOf(date.getDay()) === -1) return;
      // 非来館曜日は除外
      if (nonVisitDayNumbers.length > 0 && nonVisitDayNumbers.indexOf(date.getDay()) !== -1) return;

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
        if (dailyVisitCounts[selectedKey] >= maxVisitsPerDay) continue;

        // 振り分け確定 → 確定来館記録の形式で追加
        var defaults = childDefaultsMap[childName];
        var checkInDT = toDateTimeOnDate_(selectedDate, defaults.checkIn);
        var checkOutDT = toDateTimeOnDate_(selectedDate, defaults.checkOut);
        // 退所時刻が入所時刻以前なら翌日扱い（例: 17:00入所 → 翌8:00退所）
        if (checkOutDT.getTime() <= checkInDT.getTime()) {
          checkOutDT = new Date(checkOutDT.getTime() + 24 * 60 * 60 * 1000);
        }

        allocationResults.push([
          selectedDate,             // 記録日
          childName,                // 児童名
          '振り分け',                // データ区分
          defaults.staffName,       // スタッフ1（固定スタッフ）
          defaults.staffName2,      // スタッフ2（固定スタッフ）
          checkInDT,                // 入所日時（selectedDate + 時刻）
          checkOutDT,               // 退所日時（selectedDate + 時刻、必要に応じて翌日）
          defaults.temperature,     // 体温
          defaults.mealDinner,      // 夕食
          defaults.mealBreakfast,   // 朝食
          defaults.mealLunch,       // 昼食
          defaults.bath,            // 入浴
          defaults.sleep,           // 睡眠
          defaults.bowel,           // 便
          defaults.medicineNight,   // 服薬(夜)
          defaults.medicineMorning, // 服薬(朝)
          defaults.notes,           // その他連絡事項
        ]);

        dailyVisitCounts[selectedKey]++;
        allocated++;
      }
    }

    if (allocated < remaining) {
      Logger.log('警告: ' + childName + ' の振り分けが不完全です（残り' + (remaining - allocated) + '枠分の空きがありません）');
    }
  });

  // 10. 確定来館記録シートに振り分け行を追加
  if (allocationResults.length > 0) {
    writeAllocationsToConfirmed_(allocationResults);
    Logger.log('振り分け完了: ' + allocationResults.length + '件');
  } else {
    Logger.log('振り分け結果: 0件（振り分け先がありません）');
  }

  // 11. 7人満枠の日でスタッフ2が空欄の行に固定スタッフを補完
  fillStaff2ForFullDays_(year, month);

  // 12. 月別集計を更新
  updateMonthlySummary();
}

/**
 * 対象月の振り分けが確定来館記録に存在するかチェックする
 * @param {number} year 年
 * @param {number} month 月（1-12）
 * @returns {boolean} 振り分け行が存在する場合true
 */
function hasAllocationsForMonth_(year, month) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIRMED_DATA_START_ROW) return false;

  var targetYM = year + '/' + ('0' + month).slice(-2);
  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    var dataType = data[i][CONFIRMED_COL.DATA_TYPE - 1];
    if (dataType !== '振り分け') continue;
    var rawDate = data[i][CONFIRMED_COL.RECORD_DATE - 1];
    if (!rawDate) continue;
    var dateYM = formatDateYMD_(new Date(rawDate), 'yyyy/MM');
    if (dateYM === targetYM) return true;
  }
  return false;
}

/**
 * 振り分け行が存在する年月の一覧を返す（年月昇順、重複排除）
 * @returns {Array<{year:number, month:number}>}
 */
function listAllocatedMonths_() {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  } catch (e) {
    return [];
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIRMED_DATA_START_ROW) return [];

  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, 3).getValues();
  var seen = {};
  data.forEach(function(row) {
    if (row[CONFIRMED_COL.DATA_TYPE - 1] !== '振り分け') return;
    var rawDate = row[CONFIRMED_COL.RECORD_DATE - 1];
    if (!rawDate) return;
    var d = new Date(rawDate);
    if (isNaN(d.getTime()) || d.getFullYear() < 1900) return;
    var key = d.getFullYear() + '-' + (d.getMonth() + 1);
    seen[key] = { year: d.getFullYear(), month: d.getMonth() + 1 };
  });
  return Object.keys(seen).map(function(k) { return seen[k]; }).sort(function(a, b) {
    return a.year !== b.year ? a.year - b.year : a.month - b.month;
  });
}

/**
 * 確定来館記録から対象月の振り分け行を削除する
 * @param {number} year 年
 * @param {number} month 月（1-12）
 */
function clearAllocationsForMonth_(year, month) {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  } catch (e) {
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIRMED_DATA_START_ROW) return;

  var colCount = CONFIRMED_COL.NOTES; // 列数=17
  var rowCount = lastRow - CONFIRMED_DATA_START_ROW + 1;
  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, rowCount, colCount).getValues();

  // 削除対象（=対象年月の振り分け行）以外を残し、値だけ上に詰める。
  // 行物理削除は使わない（書式が崩れるため）。
  var kept = data.filter(function(row) {
    var rawDate = row[CONFIRMED_COL.RECORD_DATE - 1];
    var dataType = row[CONFIRMED_COL.DATA_TYPE - 1];
    if (!rawDate) return false; // 元から空行は除外
    if (dataType !== '振り分け') return true;
    var d = new Date(rawDate);
    return !(d.getFullYear() === year && (d.getMonth() + 1) === month);
  });

  if (kept.length === rowCount) return; // 削除対象なし

  // 既存範囲の値だけクリア → 残す行を先頭から書き戻し（書式は維持）
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, rowCount, colCount).clearContent();
  if (kept.length > 0) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, kept.length, colCount).setValues(kept);
  }
}

/**
 * 振り分け結果を確定来館記録シートに追加書き込みする
 * @param {Array<Array>} results 振り分け結果の2次元配列（17列: CONFIRMED_COL と同じ形式）
 */
function writeAllocationsToConfirmed_(results) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var colCount = CONFIRMED_COL.NOTES; // 列の末尾=総数（NOTES=17）

  // 既存データと振り分け結果をマージして日付順にソートし直す
  var lastRow = sheet.getLastRow();
  var existingData = [];
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    existingData = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).getValues();
  }

  var allData = existingData.concat(results);

  // 日付昇順 → 児童名昇順でソート
  allData.sort(function(a, b) {
    var dateCompare = new Date(a[0]) - new Date(b[0]);
    if (dateCompare !== 0) return dateCompare;
    return String(a[1]).localeCompare(String(b[1]));
  });

  // 既存データをクリアして書き直し
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).clearContent();
  }

  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, colCount).setValues(allData);

  // 記録日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, 1)
    .setNumberFormat('yyyy/mm/dd');

  // 入所日時・退所日時列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.CHECK_IN, allData.length, 2)
    .setNumberFormat('yyyy/mm/dd H:mm');
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
 * 重度支援区分をソート用の数値に変換する
 * @param {*} val 区分値（例: "区分1", "1", null）
 * @returns {number} ソート用数値（未設定は9999）
 */
function parsePriority_(val) {
  if (!val) return 9999;
  var match = String(val).trim().match(/区分(\d+)/);
  if (match) return parseInt(match[1], 10);
  var num = parseInt(String(val), 10);
  return isNaN(num) ? 9999 : num;
}

/**
 * 来館曜日文字列をDay番号の配列に変換する
 * @param {string} visitDayStr 来館曜日（例: "月" または "月,水,金"）
 * @returns {Array<number>} Day番号の配列（0=日, 1=月, ... 6=土）
 */
function parseVisitDays_(visitDayStr) {
  if (!visitDayStr) return [];
  // 半角カンマ、全角カンマ（読点）、全角コンマ、スペース、改行で分割
  var days = String(visitDayStr).split(/[,、，\s]+/);
  var result = [];
  days.forEach(function(day) {
    var trimmed = day.trim();
    if (!trimmed) return;
    // "月曜日" → "月"、"月曜" → "月" のように先頭1文字を抽出してマッチ
    if (DAY_OF_WEEK_MAP[trimmed] !== undefined) {
      result.push(DAY_OF_WEEK_MAP[trimmed]);
    } else {
      var firstChar = trimmed.charAt(0);
      if (DAY_OF_WEEK_MAP[firstChar] !== undefined) {
        result.push(DAY_OF_WEEK_MAP[firstChar]);
      }
    }
  });
  return result;
}

/**
 * 基準日と時刻値から datetime の Date オブジェクトを組み立てる
 * - timeVal が Date の場合: その時:分を基準日に適用
 * - timeVal が "HH:mm" 文字列の場合: パースして基準日に適用
 * @param {Date} date 基準日
 * @param {Date|string} timeVal 時刻値
 * @returns {Date} 基準日 + 時刻 の Date
 */
function toDateTimeOnDate_(date, timeVal) {
  var result = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
  if (timeVal instanceof Date) {
    result.setHours(timeVal.getHours(), timeVal.getMinutes(), 0, 0);
    return result;
  }
  var parts = String(timeVal || '').split(':');
  var hh = parseInt(parts[0], 10);
  var mm = parseInt(parts[1], 10);
  result.setHours(isNaN(hh) ? 0 : hh, isNaN(mm) ? 0 : mm, 0, 0);
  return result;
}

/**
 * 児童の振り分け補完データを実データから算出する
 * @param {string} childName 児童名
 * @param {Array} masterRow 児童マスタの行データ
 * @param {Array<Array>} formResponses 同月のフォーム回答データ
 * @returns {Object} 補完データ
 */
function computeChildDefaults_(childName, masterRow, formResponses) {
  // スタッフ1は設定シートの固定スタッフ名。スタッフ2は空欄（7人満枠時のみ fillStaff2ForFullDays_ が補完）
  var staffName = getDummyStaffName_();
  var staffName2 = '';

  // 設定シートの補完値を取得（未設定項目は ALLOCATION_DEFAULTS にフォールバック）
  var settings = getAllocationDefaultsFromSettings_();

  // 同じ児童の実データを抽出
  var childData = formResponses.filter(function(row) {
    return row[FORM_COL.CHILD_NAME - 1] === childName;
  });

  // 実データがある場合は最頻値を算出、なければ設定シート値を使用
  if (childData.length > 0) {
    return {
      staffName: staffName,
      staffName2: staffName2,
      checkIn: getModeValue_(childData, FORM_COL.CHECK_IN - 1, settings.CHECK_IN),
      checkOut: getModeValue_(childData, FORM_COL.CHECK_OUT - 1, settings.CHECK_OUT),
      temperature: getModeNumeric_(childData, FORM_COL.TEMPERATURE - 1, settings.TEMPERATURE),
      mealDinner: getModeValue_(childData, FORM_COL.MEAL_DINNER - 1, settings.MEAL_DINNER),
      mealBreakfast: getModeValue_(childData, FORM_COL.MEAL_BREAKFAST - 1, settings.MEAL_BREAKFAST),
      mealLunch: getModeValue_(childData, FORM_COL.MEAL_LUNCH - 1, settings.MEAL_LUNCH),
      bath: getModeValue_(childData, FORM_COL.BATH - 1, settings.BATH),
      sleep: getModeValue_(childData, FORM_COL.SLEEP - 1, settings.SLEEP),
      bowel: getModeValue_(childData, FORM_COL.BOWEL - 1, settings.BOWEL),
      medicineNight: getModeValue_(childData, FORM_COL.MEDICINE_NIGHT - 1, settings.MEDICINE_NIGHT),
      medicineMorning: getModeValue_(childData, FORM_COL.MEDICINE_MORNING - 1, settings.MEDICINE_MORNING),
      notes: pickRandomNote_(childName, childData, formResponses),
    };
  }

  return {
    staffName: staffName,
    staffName2: staffName2,
    checkIn: settings.CHECK_IN,
    checkOut: settings.CHECK_OUT,
    temperature: settings.TEMPERATURE,
    mealDinner: settings.MEAL_DINNER,
    mealBreakfast: settings.MEAL_BREAKFAST,
    mealLunch: settings.MEAL_LUNCH,
    bath: settings.BATH,
    sleep: settings.SLEEP,
    bowel: settings.BOWEL,
    medicineNight: settings.MEDICINE_NIGHT,
    medicineMorning: settings.MEDICINE_MORNING,
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

  // 定型文マスタからランダム取得
  var masterNotes = getNotesMasterData_();
  if (masterNotes.length > 0) {
    return masterNotes[Math.floor(Math.random() * masterNotes.length)];
  }

  // 設定シートの「連絡事項」値にフォールバック
  var settingNote = getSettingValue_(SETTINGS_ROW.NOTES);
  return settingNote ? String(settingNote) : ALLOCATION_DEFAULTS.NOTES;
}

/**
 * 指定年の確定来館記録から児童ごとの来館数マップを構築する（年間利用枠チェック用）
 * @param {number} year 対象年
 * @returns {Object} { 児童名: 来館数 }
 */
function buildYtdVisitMap_(year) {
  var records = getConfirmedVisitsByYear(year);
  var map = {};
  records.forEach(function(row) {
    var name = row[CONFIRMED_COL.CHILD_NAME - 1];
    if (name) {
      map[name] = (map[name] || 0) + 1;
    }
  });
  return map;
}

/**
 * 満枠日でスタッフ2が空欄の行に固定スタッフ名を補完する
 * 対象月の確定来館記録を走査し、来館数が1日最大来館数以上の日の
 * スタッフ2列が空欄のレコードに設定シートの固定スタッフ名を書き込む
 * @param {number} year 対象年
 * @param {number} month 対象月（1-12）
 */
function fillStaff2ForFullDays_(year, month) {
  var sheet;
  try {
    sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  } catch (e) {
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIRMED_DATA_START_ROW) return;

  var numRows = lastRow - CONFIRMED_DATA_START_ROW + 1;
  var colCount = CONFIRMED_COL.NOTES; // 17
  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, numRows, colCount).getValues();

  // 対象月の行インデックスを収集。
  // 満枠判定は「その日の新着者（記録日=入所日のプライマリ行）」のみでカウント。
  // 連泊2日目（持ち越し行）は新着者にカウントしないが、満枠成立時の補完対象には含める。
  var dailyCountMap = {};
  var rowsByDate = {};

  for (var i = 0; i < data.length; i++) {
    var recordDate = new Date(data[i][CONFIRMED_COL.RECORD_DATE - 1]);
    if (recordDate.getFullYear() !== year || (recordDate.getMonth() + 1) !== month) continue;
    var dateKey = formatDateKey_(recordDate);

    // 新着者判定: 記録日と入所日（日付部分）が一致する行のみ満枠カウントに加算
    var checkInRaw = data[i][CONFIRMED_COL.CHECK_IN - 1];
    var isNewArrival = true;
    if (checkInRaw instanceof Date && !isNaN(checkInRaw.getTime())) {
      var checkInDateStr = formatDateKey_(new Date(checkInRaw.getFullYear(), checkInRaw.getMonth(), checkInRaw.getDate()));
      isNewArrival = (checkInDateStr === dateKey);
    }
    if (isNewArrival) {
      dailyCountMap[dateKey] = (dailyCountMap[dateKey] || 0) + 1;
    }
    if (!rowsByDate[dateKey]) rowsByDate[dateKey] = [];
    rowsByDate[dateKey].push(i);
  }

  // 満枠日でスタッフ2が空欄の行に固定スタッフを補完
  var dummyStaff = getDummyStaffName_();
  if (!dummyStaff) {
    Logger.log('固定スタッフが設定シートに未設定のためスタッフ2補完をスキップ');
    return;
  }

  var maxVisitsPerDay = getMaxVisitsPerDay_();
  var updated = false;
  Object.keys(dailyCountMap).forEach(function(dateKey) {
    if (dailyCountMap[dateKey] < maxVisitsPerDay) return;
    rowsByDate[dateKey].forEach(function(rowIdx) {
      if (!data[rowIdx][CONFIRMED_COL.STAFF_NAME_2 - 1]) {
        data[rowIdx][CONFIRMED_COL.STAFF_NAME_2 - 1] = dummyStaff;
        updated = true;
      }
    });
  });

  if (!updated) return;

  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, numRows, colCount).setValues(data);
  Logger.log('スタッフ2補完完了: ' + year + '年' + month + '月（' + maxVisitsPerDay + '人満枠日対象）');
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
