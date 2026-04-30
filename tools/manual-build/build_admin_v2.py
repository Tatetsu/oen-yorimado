#!/usr/bin/env python3
"""
Yorimado_来館記録管理マニュアル.pptx を v2 へ更新する。

仕様変更点:
- 「記録日」→「利用日」リネーム
- 宿泊展開ロジック追記
- 児童マスタ・設定シートの最新カラムへ書換
- 振り分けロジック説明刷新
- 来館カレンダーのマーク説明修正(○のみ・■廃止)
- メールプレースホルダ拡充(6→17)
- フォーム回答シートのIMPORTRANGE仕様追記

各 <a:t> をインデックスで指定して置換する。インデックスは tools/manual-build/dump_indexed.py で確認したもの。
"""
import os
import shutil
import zipfile
from pathlib import Path
from lxml import etree

SRC = Path("/Users/masa/choi-dx/projects/oen_yorimado/docs/manual/Yorimado_来館記録管理マニュアル.pptx")
DST = Path("/Users/masa/choi-dx/projects/oen_yorimado/docs/manual/Yorimado_来館記録管理マニュアル_v2.pptx")
WORK = Path("/tmp/pptx_admin_build")

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A_T = f"{{{NS_A}}}t"

# slideN.xml -> { index: new_text }
REPLACE: dict[str, dict[int, str]] = {
    # === slide6: フォーム回答を確認する 手順 ===
    "slide6.xml": {
        24: "タイムスタンプ（送信日）・スタッフ名・児童名・入退所日時などが自動入力されます。本シートはスタッフ用_回答シートからIMPORTRANGEで参照する読取専用シートです。",
        32: "送信タイミングと修正方法",
        33: "送信は退所した翌朝に1件のみ。フォーム送信日のタイムスタンプが利用日として扱われ、毎朝8時に過去24時間以内のフォーム送信を対象に保護者メールが自動送信されます。誤送信はスプレッドシートを直接書き換えず、Webビュー（修正ツール）で修正・削除してください。",
    },

    # === slide11: 確定来館記録を見る 手順 ===
    "slide11.xml": {
        32: "利用日・スタッフ・入退所日時・体温・食事・入浴・入眠時刻・朝4時チェック・起床時刻・便・服薬など。フォーム1行（入所〜退所）は利用日ごとに1行ずつ展開して書込まれます（例:3/1入所→3/3退所＝3行）。",
    },

    # === slide12: 実データと振り分けの違い ===
    "slide12.xml": {
        18: "月間利用枠の残数を、児童マスタの",
        19: "「来館曜日／非来館曜日／重度支援区分／年間利用枠」と設定シートの「営業日」を考慮して自動振り分けしたシステム予測。",
        21: "重度支援区分の高い児童から優先して埋め、月間枠±1のランダム幅と年間枠の平日按分上限の小さい方を上限とします。",
        32: "複数日宿泊なのに1日しかカウントされない",
        33: "入所日時・退所予定日時が同日になっている可能性。Webビューから修正",
    },

    # === slide14: 来館カレンダー マークの見方 ===
    # 旧: ○=実データ(緑), △=振り分け(薄緑), 空欄=来館なし, ■=予定だったが記録なし
    # 新: ○=実データ(緑) / ○=振り分け(薄緑) / 空欄=来館なし / ■行は廃止
    "slide14.xml": {
        32: "来館あり（実データ）",
        35: "○",
        43: "—",
        44: "（廃止）現仕様では使用しません",
        45: "—",
    },

    # === slide18: 児童を追加する 入力項目テーブル ===
    # 旧8行 → 新マスタ15列のうち代表項目8行
    "slide18.xml": {
        # Row1: 児童名 / 例：山田太郎  (no change)
        # Row2: よみがな / やまだたろう → 保護者名 / 山田花子
        32: "保護者名",
        33: "山田花子",
        # Row3: 保護者名／保護者メール / 報告メールの送信先 → 保護者メールアドレス / 報告メール送信先
        34: "保護者メールアドレス",
        35: "報告メール送信先",
        # Row4: 月間利用枠 / 20 など → 担当スタッフ / プルダウン
        36: "担当スタッフ",
        37: "プルダウン（スタッフマスタから）",
        # Row5: 医療型 / あり／なし → 医療型の有無 / あり / なし
        38: "医療型の有無",
        39: "あり / なし",
        # Row6: 曜日1〜 / 月曜日・水曜日 → 重度支援(区分) / あり/なし＋区分1〜5
        40: "重度支援 / 重度支援区分",
        41: "あり/なし ＋ 区分1〜区分5",
        # Row7: 担当スタッフ1／2 / プルダウン選択 → 年間/月間利用枠 / 数値
        42: "年間利用枠 / 月間利用枠",
        43: "数値（年=180、月=20など。空欄=上限なし）",
        # Row8: 区分 / プルダウン選択 → 来館曜日/非来館曜日/入所状況/退所状況 / 設定例
        44: "来館曜日 / 非来館曜日 / 入所状況 / 退所状況",
        45: "月,水,金 形式 / 稼働・休止・退所 / 別施設移動・別施設移動無",
        # 注意書きの2行
        50: "既存の児童の行を並び替え・削除しない。「来館曜日」はカンマ区切り（例:月,水,金）で入力してください。",
        51: "「入所状況」=稼働/休止/退所、「退所状況」=別施設移動/別施設移動無 を正しく選んでください。",
    },

    # === slide27: 設定 主な設定項目 ===
    "slide27.xml": {
        23: "入所時間 / 退所時間",
        25: "1日最大来館数",
        27: "食事 / 入浴 / 便 / 服薬",
        28: "○",
        29: "入眠時刻 / 朝4時チェック / 起床時刻",
        30: "○",
        31: "メール件名 / メール本文",
        32: "○ 文言のみ",
        33: "営業日 / 固定スタッフ名 / エラー通知先",
        34: "△ 慎重に",
    },

    # === slide28: メール本文のプレースホルダ(全17個) ===
    "slide28.xml": {
        8: "送信時に以下の17個が自動で置き換わります。名前は絶対に変更しないでください。",
        # Row1: {保護者名} / その児童の保護者名 (no change)
        # Row2: {日付} / 来館日 → {日付} / 利用日(フォーム送信日)
        16: "利用日（フォーム送信日のタイムスタンプ）",
        # Row3: {児童名} / 児童の名前 (no change)
        # Row4: {入所時間} / 入所時刻 → {入所時間}{退所時間}{体温} / 入退所時刻・体温
        23: "{入所時間} {退所時間} {体温}",
        24: "入退所時刻・体温",
        # Row5: {退所時間} / 退所時刻 → {夕食}{朝食}{昼食}{入浴}{便} / 食事3種・入浴・便
        27: "{夕食} {朝食} {昼食} {入浴} {便}",
        28: "食事3種・入浴・便",
        # Row6: {体温} / 体温 → {入眠時刻}{朝4時チェック}{起床時刻}{服薬(夜)}{服薬(朝)}{連絡事項} / 睡眠3項目・服薬・連絡事項
        31: "{入眠時刻} {朝4時チェック} {起床時刻} {服薬(夜)} {服薬(朝)} {連絡事項}",
        32: "睡眠3項目・服薬2種・連絡事項",
    },
}


def apply_replacements(slide_xml: Path, idx_map: dict[int, str]) -> int:
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(slide_xml), parser)
    root = tree.getroot()
    ts = list(root.iter(A_T))
    applied = 0
    for idx, new_text in idx_map.items():
        if idx >= len(ts):
            print(f"  [WARN] {slide_xml.name}: index {idx} out of range ({len(ts)})")
            continue
        old = ts[idx].text or ""
        ts[idx].text = new_text
        applied += 1
        print(f"  [{idx}] {old[:30]!r} -> {new_text[:30]!r}")
    tree.write(str(slide_xml), xml_declaration=True, encoding="UTF-8", standalone=True)
    return applied


def main():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    with zipfile.ZipFile(SRC, "r") as zf:
        zf.extractall(WORK)

    total = 0
    for slide_name, idx_map in REPLACE.items():
        path = WORK / "ppt" / "slides" / slide_name
        if not path.exists():
            print(f"[ERROR] not found: {slide_name}")
            continue
        print(f"\n=== {slide_name} ===")
        total += apply_replacements(path, idx_map)

    print(f"\n[TOTAL] {total} replacements applied")

    if DST.exists():
        DST.unlink()
    with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zf:
        for root_dir, _, files in os.walk(WORK):
            for fn in files:
                full = Path(root_dir) / fn
                rel = full.relative_to(WORK)
                zf.write(full, rel.as_posix())

    print(f"\n=> {DST}")


if __name__ == "__main__":
    main()
