# docs/ インデックス

## 最優先参照

| ファイル | 内容 | 最終更新 |
|---------|------|---------|
| [requirements.md](requirements.md) | 要件定義書（シート仕様・GAS機能一覧・振り分けロジック概要） | 2026-04-20 |
| [architecture.md](architecture.md) | システム構成図・GAS処理フロー・ファイル構成 | 2026-04-20 |
| [er-diagram.md](er-diagram.md) | シート間のデータ関係・各シートの列定義 | 2026-03-30 |

---

## 仕様書（specs/）

個別のロジック・機能の詳細設計。GASの実装コードから起こしたもの。

| ファイル | 内容 |
|---------|------|
| [specs/google-form.md](specs/google-form.md) | フォームの質問項目・列定義・プルダウン同期の仕組み・スクリプトプロパティ |
| [specs/allocation-logic.md](specs/allocation-logic.md) | 振り分けロジック詳細（優先度・±1ランダム幅・年間枠按分・補完値算出・均等分散） |
| [specs/overnight-logic.md](specs/overnight-logic.md) | 連泊展開ロジック・月またぎカウントルール |
| [specs/aggregation-logic.md](specs/aggregation-logic.md) | スコープ4種の挙動・月別集計/児童別ビュー/カレンダーのデータソース |
| [specs/email-notification.md](specs/email-notification.md) | 保護者メール・メール本文テンプレート・バウンス検出・エラー通知 |
| [specs/triggers.md](specs/triggers.md) | 全トリガー一覧（時間・onEdit・メニュー）とトリガー設定手順 |
| [specs/child-special-rules.md](specs/child-special-rules.md) | 個別児童の特殊対応ルール（送迎順序・迎え不要曜日） |

---

## 手順書・運用

| ファイル | 内容 | 最終更新 |
|---------|------|---------|
| [setup-guide.md](setup-guide.md) | 初期セットアップ手順（施設スタッフ向け） | 2026-04-21 |
| [manual_form-edit.md](manual_form-edit.md) | Webビューでの来館記録の編集・削除手順 | 2026-04-20 |
| [送迎記録表_仕様.md](送迎記録表_仕様.md) | 送迎記録表CSV生成スクリプトの仕様 | 2026-04-10 |

---

## 参考（作成時点の情報。現状と乖離している可能性あり）

| ファイル | 内容 |
|---------|------|
| [progress-report.md](progress-report.md) | 開発進捗レポート（2026-03-12時点） |
| [todo.md](todo.md) | TODOリスト（2026-03-20時点） |
| [skills.md](skills.md) | スキル・技術メモ（2026-03-09時点） |
