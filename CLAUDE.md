# CLAUDE.md - oen-yorimado

## プロジェクト概要

児童施設の来館記録管理システム。Google Sheets + GAS + Google Forms + CLASP で構成。

## CLASP運用ルール（必須）

- GASファイルを編集した場合、**必ずユーザーに確認を取ってから** `clasp push` を実行する
- 勝手にプッシュしない。確認なしのデプロイは禁止
- `.clasp.json` でscriptIdとrootDir（./gas）は設定済み

## 参照優先順位

1. `docs/requirements.md` — 要件定義（最優先）
2. `rules/project.md` — プロジェクト固有ルール
3. `rules/coding.md` — コーディング規約
4. `rules/ai-collab.md` — AI連携規約
5. `docs/architecture.md` — システム構成図
6. `docs/er-diagram.md` — データ設計
