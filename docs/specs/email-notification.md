# メール通知 仕様書

## 概要

3種類のメール通知がある。保護者向け来館報告メール（F-07）、エラー通知メール（F-10）、バウンスメール検出（bounce-checker.gs）。

---

## 保護者向け来館報告メール（F-07）

実装: `gas/email.gs`

### 実行タイミング

| 方法 | 関数名 | タイミング |
|------|-------|---------|
| 自動（時間トリガー） | `sendDailyVisitReports` | 毎朝8時（前日の来館記録が対象） |
| 手動（メニュー） | `sendVisitReportsManual` | HTMLダイアログで対象日付を選択 |

### 対象レコードの抽出

- フォームの回答シートから指定日の記録を取得
- N列（`EMAIL_SENT`）が空欄のレコードのみ送信対象（送信済みは重複送信しない）

### 送信後の処理

N列に `「送信済 yyyy/MM/dd HH:mm」` を書き込む。書き込み先は**スタッフ用_回答シート**（`STAFF_SHEET_ID`）。

### メール本文テンプレート

`utils.gs` の `DEFAULT_EMAIL_TEMPLATE` 定数（設定シート row19「メール本文」で上書き可能）。
2026-04 で食事(夕/朝/昼)・服薬(夜/朝)の分離に対応。

```
{保護者名} 様

いつもお世話になっております。
テスト施設です。

{日付}の{児童名}さんの来館記録をお知らせいたします。

■ 来館記録
・入所時間: {入所時間}
・退所時間: {退所時間}
・体温: {体温}℃
・夕食: {夕食}
・朝食: {朝食}
・昼食: {昼食}
・入浴: {入浴}
・睡眠: {睡眠}
・便: {便}
・服薬(夜): {服薬(夜)}
・服薬(朝): {服薬(朝)}

 ■ 連絡事項
・{連絡事項}
```

#### プレースホルダー

| プレースホルダー | 対応列 | GAS定数 |
|--------------|-------|--------|
| `{保護者名}` | 児童マスタ C列 | MASTER_COL.PARENT_NAME |
| `{日付}` | フォーム回答 B列（記録日） | FORM_COL.RECORD_DATE |
| `{児童名}` | フォーム回答 E列 | FORM_COL.CHILD_NAME |
| `{入所時間}` | フォーム回答 F列 | FORM_COL.CHECK_IN |
| `{退所時間}` | フォーム回答 G列 | FORM_COL.CHECK_OUT |
| `{体温}` | フォーム回答 H列 | FORM_COL.TEMPERATURE |
| `{夕食}` | フォーム回答 I列 | FORM_COL.MEAL_DINNER |
| `{朝食}` | フォーム回答 J列 | FORM_COL.MEAL_BREAKFAST |
| `{昼食}` | フォーム回答 K列 | FORM_COL.MEAL_LUNCH |
| `{入浴}` | フォーム回答 L列 | FORM_COL.BATH |
| `{睡眠}` | フォーム回答 M列 | FORM_COL.SLEEP |
| `{便}` | フォーム回答 N列 | FORM_COL.BOWEL |
| `{服薬(夜)}` | フォーム回答 O列 | FORM_COL.MEDICINE_NIGHT |
| `{服薬(朝)}` | フォーム回答 P列 | FORM_COL.MEDICINE_MORNING |
| `{連絡事項}` | フォーム回答 Q列 | FORM_COL.NOTES |

### メール件名

設定シート row18（`SETTINGS_ROW.EMAIL_SUBJECT`）の値を使用。未設定時のフォールバックは `DEFAULT_EMAIL_SUBJECT`。

### 宛先

児童マスタ D列（保護者メールアドレス）。

### 未実装事項（TODO）

同日に兄弟（同一保護者メールアドレス）が来館した場合、現在は児童ごとに個別送信。1通にまとめるかは未実装・未決定（`email.gs` コメント参照）。

---

## エラー通知メール（F-10）

実装: `gas/email.gs` / `utils.gs`

### 仕組み

GAS の各関数の catch ブロックで `logError_('関数名', error)` を呼び出す。

`logError_` は:
1. ログシートにタイムスタンプ・関数名・エラーメッセージ・スタックトレースを書き込む
2. ログシートが存在しない場合は自動作成する

### エラー通知先

設定シート row17（`SETTINGS_ROW.ERROR_EMAIL`）にカンマ区切りでメールアドレスを設定する。GAS実行者のメールは常に含まれる（`getErrorNotifyRecipients_`）。

---

## 設定シートのメール関連行

| 行 | 設定項目名 | 内容 |
|----|----------|------|
| 17 | エラー通知先メール | カンマ区切り複数可 |
| 18 | メール件名 | 保護者向けメールの件名テンプレート |
| 19 | メール本文 | 保護者向けメールの本文テンプレート |

---

## バウンスメール検出（bounce-checker.gs）

Gmail の NDR（`mailer-daemon` / `postmaster` 発）を定期検索し、バウンスログシートに記録 → 検出時は `ERROR_NOTIFY_EMAILS` 宛に通知メールを送信する。

### 実行タイミング

| 方法 | 関数名 | タイミング |
|------|-------|---------|
| 自動（時間トリガー） | `collectBounceEmails` | 毎日 AM9時（`setupBounceCheckTrigger()` で登録） |
| 手動（メニュー） | `collectBounceEmailsManual` | 「来館管理」メニュー → 「バウンスメールを確認」 |

初期セットアップ（`setupAllSheets`）で自動トリガーが作成される。検索対象は前回チェック日時以降のメール（PropertiesService `BOUNCE_CHECK_LAST_RUN` で管理）。
