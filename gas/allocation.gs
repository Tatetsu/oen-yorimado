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
  var stays = pairStayRecords_(formResponses);

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

  // スタッフ稼働曜日マップ（曜日ごとの稼働スタッフ）と日ごとの担当スタッフキャッシュ
  var staffByWeekday = getActiveStaffByWeekday_();
  var fallbackStaff = getDummyStaffName_();
  var dailyStaffMap = {}; // {dateKey: staffName} 同日同一スタッフ保証
  var pickStaffForDate_ = function(date) {
    var key = formatDateKey_(date);
    if (dailyStaffMap[key]) return dailyStaffMap[key];
    var pool = staffByWeekday[date.getDay()] || [];
    var name = pool.length > 0 ? pool[Math.floor(Math.random() * pool.length)] : fallbackStaff;
    dailyStaffMap[key] = name;
    return name;
  };

  // フォーム選択肢ベースの加重ランダムジェネレータ（行ごとに値生成）
  var randomGen = getAllocationRandomGenerators_();

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

    // 同一児童の連続日抑制用: 既存の実来館日 + これまで振り分けた日 を記録
    var childOccupiedKeys = {};
    Object.keys(childVisitDates).forEach(function(k) { childOccupiedKeys[k] = true; });
    // 隣接日ペナルティ: dailyVisitCounts の取り得る最大値より十分大きく、
    //   かつ「空きが本当にない月」では連続日も許容するためのソフト制約
    var ADJACENCY_PENALTY = 100;
    var scoreDate_ = function(date) {
      var key = formatDateKey_(date);
      var s = dailyVisitCounts[key] || 0;
      var prevKey = formatDateKey_(new Date(date.getFullYear(), date.getMonth(), date.getDate() - 1));
      var nextKey = formatDateKey_(new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1));
      if (childOccupiedKeys[prevKey]) s += ADJACENCY_PENALTY;
      if (childOccupiedKeys[nextKey]) s += ADJACENCY_PENALTY;
      return s;
    };

    // 来館曜日優先 → その他で足りない分を補う
    var allocated = 0;
    var candidatePools = [preferredDates, otherDates];

    for (var poolIdx = 0; poolIdx < candidatePools.length && allocated < remaining; poolIdx++) {
      var pool = candidatePools[poolIdx];

      while (allocated < remaining && pool.length > 0) {
        // 来館数が最も少ない日を選択（均等分散）+ 同一児童の連続日にはペナルティ
        pool.sort(function(a, b) {
          return scoreDate_(a) - scoreDate_(b);
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
        var checkOutDate = new Date(checkOutDT.getFullYear(), checkOutDT.getMonth(), checkOutDT.getDate());
        var isOvernight = (selectedDate.getTime() !== checkOutDate.getTime());
        // 退所日が対象月外（=月末またぎ）の場合は退所日行は作らない（実データの月別フィルタと同じ挙動）
        var checkOutInScope = (checkOutDate.getFullYear() === year && (checkOutDate.getMonth() + 1) === month);
        var generateCheckoutRow = isOvernight && checkOutInScope;
        // 月間利用枠を「日数」で消費するため、連泊の2日消費でオーバーシュートする場合はこの候補をスキップ
        if (generateCheckoutRow && allocated + 2 > remaining) continue;
        var stayPk = isOvernight ? buildStayPk_(childName, checkInDT) : '';

        // 1宿泊=フォーム1レコード相当の値を1度だけ生成し、入所日行/退所日行で複製する
        // （実データはフォーム1回送信→確定来館記録で複製の流れ。同じ内容で記録日と往/復のみ変わる）
        var stayStaff = pickStaffForDate_(selectedDate);
        var stayValues = {
          temperature: randomGen.temperature(),
          mealDinner: randomGen.mealDinner(),
          mealBreakfast: randomGen.mealBreakfast(),
          mealLunch: randomGen.mealLunch(),
          bath: randomGen.bath(),
          sleepOnset: randomGen.sleepOnset(),
          sleepCheck4am: randomGen.sleepCheck4am(),
          wakeUp: randomGen.wakeUp(),
          bowel: randomGen.bowel(),
          medicineNight: randomGen.medicineNight(),
          medicineMorning: randomGen.medicineMorning(),
          notes: pickRandomNote_(),
        };

        var buildRow_ = function(recordDate, pickupOutbound, pickupReturn) {
          return [
            '振り分け',
            recordDate,
            stayStaff,
            '',
            childName,
            checkInDT,
            checkOutDT,
            pickupOutbound,
            pickupReturn,
            stayValues.temperature,
            stayValues.mealDinner,
            stayValues.mealBreakfast,
            stayValues.mealLunch,
            stayValues.bath,
            stayValues.sleepOnset,
            stayValues.sleepCheck4am,
            stayValues.wakeUp,
            stayValues.bowel,
            stayValues.medicineNight,
            stayValues.medicineMorning,
            stayValues.notes,
            isOvernight,
            stayPk,
          ];
        };

        // 入所日行（記録日=入所日、往=1、復は退所日行で立つので空）
        allocationResults.push(buildRow_(selectedDate, 1, ''));

        // 退所日行（対象月内に退所が収まる場合のみ複製。月またぎは欠落）
        if (generateCheckoutRow) {
          allocationResults.push(buildRow_(checkOutDate, '', 1));
        }

        dailyVisitCounts[selectedKey]++;
        childOccupiedKeys[selectedKey] = true;
        if (isOvernight && checkOutInScope) {
          // 退所日も同児童の連続日抑制対象に追加（実データ集計の通し日付として扱う）
          childOccupiedKeys[formatDateKey_(checkOutDate)] = true;
        }
        // 月間利用枠は「日数」で消費する。連泊で対象月内に退所する場合は2日分を消費
        allocated += generateCheckoutRow ? 2 : 1;
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
  // データ区分(1)・記録日(2)を読むので最低でも RECORD_DATE まで取得
  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, CONFIRMED_COL.RECORD_DATE).getValues();
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

  // データ区分(1)・記録日(2)を読むので最低でも RECORD_DATE まで取得
  var data = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, CONFIRMED_COL.RECORD_DATE).getValues();
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

  var colCount = CONFIRMED_COL.STAY_PK; // 列数=末尾(STAY_PK=23)
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
 * @param {Array<Array>} results 振り分け結果の2次元配列（21列: CONFIRMED_COL と同じ形式）
 */
function writeAllocationsToConfirmed_(results) {
  var sheet = getSheet(SHEET_NAMES.CONFIRMED_VISITS);
  var colCount = CONFIRMED_COL.STAY_PK; // 列の末尾=総数（STAY_PK=23）

  // 既存データと振り分け結果をマージして日付順にソートし直す
  var lastRow = sheet.getLastRow();
  var existingData = [];
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    existingData = sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).getValues();
  }

  var allData = existingData.concat(results);

  // 日付昇順 → 児童名昇順でソート
  allData.sort(function(a, b) {
    var dateCompare = new Date(a[CONFIRMED_COL.RECORD_DATE - 1]) - new Date(b[CONFIRMED_COL.RECORD_DATE - 1]);
    if (dateCompare !== 0) return dateCompare;
    return String(a[CONFIRMED_COL.CHILD_NAME - 1]).localeCompare(String(b[CONFIRMED_COL.CHILD_NAME - 1]));
  });

  // 既存データをクリアして書き直し
  if (lastRow >= CONFIRMED_DATA_START_ROW) {
    sheet.getRange(CONFIRMED_DATA_START_ROW, 1, lastRow - CONFIRMED_DATA_START_ROW + 1, colCount).clearContent();
  }

  sheet.getRange(CONFIRMED_DATA_START_ROW, 1, allData.length, colCount).setValues(allData);

  // 記録日列の表示形式
  sheet.getRange(CONFIRMED_DATA_START_ROW, CONFIRMED_COL.RECORD_DATE, allData.length, 1)
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
 * 体温・食事(夕/朝)・入浴・入眠時刻・朝4時チェック・起床時刻・便・服薬は行ごとにランダム抽選するため、
 * ここでは値そのものではなく抽選用の候補配列(プール)を返す
 * 同児童のフォーム回答にその列の値があれば優先、なければ全児童の回答から、
 * それでもなければ設定シート値1件をプールに入れる
 * 入所/退所時刻・昼食は固定運用のため最頻値(なければ設定値)で確定する
 * @param {string} childName 児童名
 * @param {Array} masterRow 児童マスタの行データ
 * @param {Array<Array>} formResponses 同月のフォーム回答データ（全児童分）
 * @returns {Object} 補完データ（固定値 + 抽選プール）
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
  var hasChildData = childData.length > 0;

  return {
    staffName: staffName,
    staffName2: staffName2,
    // 入所/退所時刻は固定運用：最頻値、なければ設定値
    checkIn: hasChildData ? getModeValue_(childData, FORM_COL.CHECK_IN - 1, settings.CHECK_IN) : settings.CHECK_IN,
    checkOut: hasChildData ? getModeValue_(childData, FORM_COL.CHECK_OUT - 1, settings.CHECK_OUT) : settings.CHECK_OUT,
    // 昼食は固定運用：最頻値、なければ設定値（通常 "-"）
    mealLunch: hasChildData ? getModeValue_(childData, FORM_COL.MEAL_LUNCH - 1, settings.MEAL_LUNCH) : settings.MEAL_LUNCH,
    // 行ごとランダム抽選するための候補プール（同児童 → 全児童 → 設定値の順でフォールバック）
    temperaturePool: buildValuePool_(childData, formResponses, FORM_COL.TEMPERATURE - 1, settings.TEMPERATURE),
    mealDinnerPool: buildValuePool_(childData, formResponses, FORM_COL.MEAL_DINNER - 1, settings.MEAL_DINNER),
    mealBreakfastPool: buildValuePool_(childData, formResponses, FORM_COL.MEAL_BREAKFAST - 1, settings.MEAL_BREAKFAST),
    bathPool: buildValuePool_(childData, formResponses, FORM_COL.BATH - 1, settings.BATH),
    sleepOnsetPool: buildValuePool_(childData, formResponses, FORM_COL.SLEEP_ONSET - 1, settings.SLEEP_ONSET),
    sleepCheck4amPool: buildValuePool_(childData, formResponses, FORM_COL.SLEEP_CHECK_4AM - 1, settings.SLEEP_CHECK_4AM),
    wakeUpPool: buildValuePool_(childData, formResponses, FORM_COL.WAKE_UP - 1, settings.WAKE_UP),
    bowelPool: buildValuePool_(childData, formResponses, FORM_COL.BOWEL - 1, settings.BOWEL),
    medicineNightPool: buildValuePool_(childData, formResponses, FORM_COL.MEDICINE_NIGHT - 1, settings.MEDICINE_NIGHT),
    medicineMorningPool: buildValuePool_(childData, formResponses, FORM_COL.MEDICINE_MORNING - 1, settings.MEDICINE_MORNING),
  };
}

/**
 * 抽選候補プールを構築する
 * 優先順: 同児童データ > 全件データ > 設定値1件
 * 出現回数の重みはそのまま（よく出る値が選ばれやすい）
 */
function buildValuePool_(childData, allData, colIndex, defaultValue) {
  var collect = function(rows) {
    return rows
      .map(function(row) { return row[colIndex]; })
      .filter(function(v) { return v !== '' && v !== null && v !== undefined; });
  };
  var pool = collect(childData);
  if (pool.length > 0) return pool;
  pool = collect(allData);
  if (pool.length > 0) return pool;
  return [defaultValue];
}

/**
 * 配列からランダムに1要素を返す
 */
function pickRandomFromPool_(pool) {
  if (!pool || pool.length === 0) return '';
  return pool[Math.floor(Math.random() * pool.length)];
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
 * 連絡事項を定型文マスタからランダムに選択する
 * 振り分け行は実データを参照せず、定型文マスタのみを参照する
 * @returns {string} 連絡事項テキスト
 */
function pickRandomNote_() {
  var masterNotes = getNotesMasterData_();
  if (masterNotes.length > 0) {
    return masterNotes[Math.floor(Math.random() * masterNotes.length)];
  }

  // 定型文マスタが空の場合は設定シート値→デフォルトにフォールバック
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
  var colCount = CONFIRMED_COL.STAY_PK; // 23
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
 * Date型の時刻をHH:mm形式の文字列に変換する
 * @param {Date} date 日時
 * @returns {string} HH:mm形式
 */
function formatTimeKey_(date) {
  var h = ('0' + date.getHours()).slice(-2);
  var m = ('0' + date.getMinutes()).slice(-2);
  return h + ':' + m;
}
