#!/usr/bin/env python3
"""
Yorimado_フォーム入力マニュアル.pptx を v2 へ全面再構成する。

主要変更:
- PART 0(GAS認証) を新設し 5ページ追加
- 旧 STEP1(記録日)/STEP25(連泊チェック)/連泊早見表 を削除/再利用
- 旧 STEP20-21(入浴・睡眠) を 入浴のみ に縮小
- 便/入眠時刻/朝4時チェック/起床時刻 の4ステップを新規追加
- 全STEPを18ステップに再採番
- 入退所日時の入力例ページを新設（旧連泊早見表を再利用）
- 各種文言更新

スライド配置:
- 既存slide7 → 認証STEP A1 (画像差替: image.png)
- 既存slide16 → STEP10 入浴 (睡眠部分を縮小)
- 既存slide19 → 認証STEP A2 (画像差替: image copy.png)
- 既存slide20 → 入退所日時の入力例 (テキスト全書換)
- 新規slide34: PART0扉（slide5複製）
- 新規slide35: 認証A3（slide7複製、image copy 2.png）
- 新規slide36: 認証A4（slide7複製、image copy 3.png）
- 新規slide37: 認証A5（slide7複製、image copy 4.png）
- 新規slide38: STEP11 便（slide17複製、画像はプレースホルダ）
- 新規slide39: STEP12 入眠時刻（slide17複製）
- 新規slide40: STEP13 朝4時チェック（slide17複製）
- 新規slide41: STEP14 起床時刻（slide17複製）

最後に slideIdLst を再順序づけ。
"""
import os
import re
import shutil
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

PROJ = Path("/Users/masa/choi-dx/projects/oen_yorimado")
SRC = PROJ / "docs/manual/Yorimado_フォーム入力マニュアル.pptx"
DST = PROJ / "docs/manual/Yorimado_フォーム入力マニュアル_v2.pptx"
WORK = Path("/tmp/pptx_form_build")

# 認証スクショ + 新ステップ用スクショ
AUTH_IMAGES = {
    "image.png": PROJ / "image.png",
    "imagecopy.png": PROJ / "image copy.png",
    "imagecopy2.png": PROJ / "image copy 2.png",
    "imagecopy3.png": PROJ / "image copy 3.png",
    "imagecopy4.png": PROJ / "image copy 4.png",
    # 新ステップ画像 (入眠時刻 / 朝4時チェック / 起床時刻)
    # 便は服薬と同じラジオボタン形式のため流用（=画像差替なし）
    "imagecopy5.png": PROJ / "image copy 5.png",  # 入眠時刻
    "imagecopy6.png": PROJ / "image copy 6.png",  # 朝4時チェック
    "imagecopy7.png": PROJ / "image copy 7.png",  # 起床時刻
}

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

A_T = f"{{{NS_A}}}t"


# --------------------------------------------------------------
# テキスト置換マップ: slideN.xml -> { index: new_text }
# --------------------------------------------------------------
EXISTING_REPLACE: dict[str, dict[int, str]] = {
    # ---- slide2: はじめに ----
    "slide2.xml": {
        4: "もし間違って送ってしまっても、Webビュー（修正ツール）で後から直せます。本マニュアルは、Webビューの初回認証から日々の送信、修正までを1冊にまとめています。",
        19: "1回の来館（単泊・宿泊いずれも）をフォーム1行で正しく送信できる",
        22: "2 / 41",
    },

    # ---- slide3: 目次 (全面書換) ----
    "slide3.xml": {
        3: "PART 0   WebViewの初回認証",
        4: "00",
        5: "認証の流れ（5ステップ）",
        6: "PART 1   フォームを入力する",
        7: "STEP 1：スタッフ1を選ぶ",
        8: "01",
        9: "全体の流れ",
        10: "02",
        11: "STEP 2：スタッフ2（任意）",
        12: "03",
        13: "STEP 3：児童名を選ぶ",
        14: "04",
        15: "STEP 4：入所日時を入れる",
        16: "05",
        17: "STEP 5：退所予定日時を入れる",
        18: "06",
        19: "入退所日時の入力例",
        20: "07",
        21: "STEP 6：体温を入れる",
        22: "08",
        23: "STEP 7-8：食事（夕・朝）",
        24: "09",
        25: "STEP 9：昼食（他サービス併給時は−）",
        26: "10",
        27: "STEP 10：入浴を記録する",
        28: "11",
        29: "STEP 11：便を記録する",
        30: "12",
        31: "STEP 12：入眠時刻",
        32: "PART 2   フォームを修正する",
        33: "13",
        34: "STEP 13：朝4時チェック",
        35: "14",
        36: "STEP 14：起床時刻",
        37: "15",
        38: "STEP 15-16：服薬（夜・朝）",
        39: "16",
        40: "STEP 17：その他連絡事項",
        41: "17",
        42: "STEP 18：送信ボタン",
        43: "18",
        44: "PART 1 まとめ：よくあるトラブル",
        45: "19",
        46: "PART 2：Webビューで修正する",
        47: "20",
        48: "PART 2 まとめ：大事なお願い",
        49: "21",
        50: "ご活用ください",
        51: "→ 初回はPART 0で認証、日々の送信はPART 1、修正はPART 2を見てください。",
        53: "3 / 41",
    },

    # ---- slide4: 全体の流れ ----
    "slide4.xml": {
        7: "毎朝8時に過去24時間の",
        8: "送信分を保護者へ自動メール",
    },

    # ---- slide5: PART 1 扉 ----
    "slide5.xml": {
        4: "全18ステップを 18 ページにまとめています。",
    },

    # ---- slide6: 全体の流れ表 (連泊チェック削除・再採番) ----
    "slide6.xml": {
        1: "PART 1：18ステップで完結します",
        3: "Googleフォームは上から順に入力していくだけ。1回の来館（単泊・宿泊いずれも）をフォーム1行で記録します。",
        4: "① 基本情報",
        5: "スタッフ1 (1)",
        6: "スタッフ2 (2)",
        7: "児童名 (3)",
        8: "入所日時 (4)",
        9: "② 入退所・健康",
        10: "退所予定日時 (5)",
        11: "体温 (6)",
        12: "③ 食事・入浴・便",
        13: "夕食/朝食/昼食 (7-9)",
        14: "入浴 (10)",
        15: "便 (11)",
        16: "服薬 夜・朝 (15-16)",
        17: "④ 睡眠・連絡・送信",
        18: "入眠/朝4時/起床 (12-14)",
        19: "その他連絡事項 (17)",
        20: "送信ボタン (18)",
        22: "送信は退所した翌朝に1件のみ。複数日宿泊でもフォーム1行で完結します。",
        24: "6 / 41",
    },

    # ---- slide7: (旧STEP1記録日 → 認証STEP A1) ----
    "slide7.xml": {
        0: "00",
        1: "認証 A1：「REVIEW PERMISSIONS」を押す",
        2: "A1 / A5",
        3: "認証 STEP A1",
        4: "「Authorization needed」画面で青ボタンを押す",
        6: "施設オーナーから共有されたWebViewのURLをブラウザで開きます。",
        8: "「Google Apps Script / prd-oen-yorimado (Unverified)」と書かれた認証画面が表示されます。",
        10: "中央の青い「REVIEW PERMISSIONS」ボタンをクリックします。",
        12: "一度この認証を通せば次回以降は表示されません。組織内専用のアプリなので安全です。",
        13: "認証 STEP A1：REVIEW PERMISSIONS",
        15: "7 / 41",
    },

    # ---- slide8: (旧STEP2-3 スタッフ1) ----
    "slide8.xml": {
        1: "STEP 1：スタッフ1を選ぶ（必須）",
        2: "1 / 18",
        3: "STEP 1",
        13: "STEP 1：スタッフ1のプルダウン",
        15: "13 / 41",
    },

    # ---- slide9: (旧STEP4-5 スタッフ2) ----
    "slide9.xml": {
        1: "STEP 2：スタッフ2を選ぶ（任意）",
        2: "2 / 18",
        3: "STEP 2",
        13: "STEP 2：スタッフ2のプルダウン",
        15: "14 / 41",
    },

    # ---- slide10: (旧STEP8-9 児童名) ----
    "slide10.xml": {
        1: "STEP 3：児童名を選ぶ（必須）",
        2: "3 / 18",
        3: "STEP 3",
        13: "STEP 3：児童名のプルダウン",
        15: "15 / 41",
    },

    # ---- slide11: (旧STEP10-12 入所日時) - 連泊文言削除 ----
    "slide11.xml": {
        1: "STEP 4：入所日時を入れる（必須）",
        2: "4 / 18",
        3: "STEP 4",
        4: "「入所日時」の日付・時・分の3か所を入力（複数日宿泊でも入力する）",
        11: "[ ポイント ]",
        12: "単泊・1泊2日・連泊いずれでも 必ず入所日時を入力 します。同じフォーム1行に退所予定日時も入れて、入所〜退所を1行で記録します。",
        13: "STEP 4：日付を選ぶ",
        14: "STEP 4：時を入れる",
        15: "16 / 41",
    },

    # ---- slide12: (旧STEP13-15 退所予定日時) - 連泊文言削除 ----
    "slide12.xml": {
        1: "STEP 5：退所予定日時を入れる（必須）",
        2: "5 / 18",
        3: "STEP 5",
        4: "「退所予定日時」の日付・時・分を入力（複数日宿泊でも入力する）",
        11: "[ ポイント ]",
        12: "退所予定日時も全来館で必須。複数日宿泊の場合は最終日の日時を入力（例: 1泊2日 = 翌日 08:00）。次のページで入力例を確認してください。",
        13: "STEP 5：日付を選ぶ",
        14: "STEP 5：時を入れる",
        15: "17 / 41",
    },

    # ---- slide13: (旧STEP16 体温) ----
    "slide13.xml": {
        1: "STEP 6：体温を入れる",
        2: "6 / 18",
        3: "STEP 6",
        13: "STEP 6：体温の入力",
        15: "19 / 41",
    },

    # ---- slide14: (旧STEP17-18 食事 夕朝) ----
    "slide14.xml": {
        1: "STEP 7-8：食事の記録（夕食・朝食）",
        2: "7-8 / 18",
        3: "STEP 7-8",
        13: "STEP 7：夕食（提供なし＝−）",
        14: "STEP 8：朝食（完食を選択）",
        15: "20 / 41",
    },

    # ---- slide15: (旧STEP19 昼食) ----
    "slide15.xml": {
        1: "STEP 9：昼食について",
        2: "9 / 18",
        3: "STEP 9",
        13: "STEP 9：昼食欄（−を選択）",
        15: "21 / 41",
    },

    # ---- slide16: (旧STEP20-21 入浴+睡眠 → 入浴のみ) ----
    "slide16.xml": {
        1: "STEP 10：入浴を記録する",
        2: "10 / 18",
        3: "STEP 10",
        4: "「入浴」を ○ / × から選ぶ",
        8: "入浴できなかった日は「×」を選び、理由はその他連絡事項に書きます。",
        10: "体調不良で入浴を見送った場合も同様に「×」を選択します。",
        12: "睡眠は別の3項目（入眠時刻・朝4時チェック・起床時刻）として STEP 12〜14 に分かれました。",
        13: "STEP 10：入浴のプルダウン",
        14: "※睡眠は STEP 12〜14 に分離されました",
        15: "22 / 41",
    },

    # ---- slide17: (旧STEP22-23 服薬) ----
    "slide17.xml": {
        1: "STEP 15-16：服薬（夜・朝）を記録する",
        2: "15-16 / 18",
        3: "STEP 15-16",
        13: "STEP 15：服薬（夜）",
        14: "STEP 16：服薬（朝）",
        15: "27 / 41",
    },

    # ---- slide18: (旧STEP24 連絡事項) ----
    "slide18.xml": {
        1: "STEP 17：その他連絡事項を入れる",
        2: "17 / 18",
        3: "STEP 17",
        13: "STEP 17：その他連絡事項",
        15: "28 / 41",
    },

    # ---- slide19: (旧STEP25 連泊チェック → 認証STEP A2) ----
    "slide19.xml": {
        0: "00",
        1: "認証 A2：「詳細」をクリック",
        2: "A2 / A5",
        3: "認証 STEP A2",
        4: "「このアプリは Google で確認されていません」で詳細を開く",
        6: "次の画面で「⚠ このアプリは Google で確認されていません」という赤い警告が表示されます。",
        8: "怖い見た目ですが、組織内専用の社内アプリなので問題ありません。",
        10: "画面左下の「詳細」リンクをクリックします。",
        11: "[ 補足 ]",
        12: "この警告は「Google公式審査を通していない社内アプリ」の標準警告です。デベロッパーは tatetsum@choi-dx.com と表示されます。",
        13: "認証 STEP A2：警告画面の「詳細」リンク",
        15: "8 / 41",
    },

    # ---- slide20: (旧連泊早見表 → 入退所日時の入力例) ----
    "slide20.xml": {
        1: "入退所日時の入力例",
        2: "STEP 4-5 補足",
        3: "入所日時・退所予定日時はどちらも必須。1回の来館（単泊・宿泊いずれも）をフォーム1行で記録します。送信は退所した翌朝に1件のみ。",
        4: "来館形態",
        5: "入所日時",
        6: "退所予定日時",
        7: "利用日カウント",
        8: "ひとことメモ",
        9: "単泊（同日帰宅）",
        10: "例: 4/1",
        11: "4/1 17:00",
        12: "4/1 21:00",
        13: "1日",
        14: "通常の来館",
        15: "1泊2日",
        16: "例: 4/1〜4/2",
        17: "4/1 17:00",
        18: "4/2 08:00",
        19: "2日",
        20: "月別集計に2日カウント",
        21: "2泊3日",
        22: "例: 4/1〜4/3",
        23: "4/1 17:00",
        24: "4/3 08:00",
        25: "3日",
        26: "入所日〜退所予定日の各日にカウント",
        27: "3泊4日（月またぎ）",
        28: "例: 4/30〜5/3",
        29: "4/30 17:00",
        30: "5/3 08:00",
        31: "4月=1, 5月=3",
        32: "月またぎは両方の月にカウント",
        33: "[ ポイント — 2つだけ ]",
        34: "①入所日時・退所予定日時は **すべての来館で必須入力**。",
        35: "②フォーム送信は **退所した翌朝に1件のみ**。複数日宿泊でも分けて送らない。",
        36: "※連泊チェック項目は廃止されました。",
        38: "9 / 41",
    },

    # ---- slide21: (旧STEP26 送信) ----
    "slide21.xml": {
        1: "STEP 18：送信ボタンを押す",
        2: "18 / 18",
        3: "STEP 18",
        10: "翌朝8時に、過去24時間以内の送信分を対象に保護者へ自動メールが送信されます。",
        12: "ボタンを押しても反応がないときは、必須欄（赤い*）が未入力の可能性。スクロールして確認してください。",
        13: "STEP 18：送信ボタン",
        15: "29 / 41",
    },

    # ---- slide22: (PART 1 まとめ - トラブル文言更新) ----
    "slide22.xml": {
        6: "複数日宿泊なのに月別集計で1日しかカウントされない",
        7: "入所日時・退所予定日時が同じ日になっている可能性。",
        21: "30 / 41",
    },

    # ---- slide23: PART 2 扉 ----
    "slide23.xml": {},

    # ---- slide24: Webビューってなに ----
    "slide24.xml": {
        21: "32 / 41",
    },
    "slide25.xml": {15: "33 / 41"},
    "slide26.xml": {15: "34 / 41"},
    "slide27.xml": {15: "35 / 41"},
    "slide28.xml": {15: "36 / 41"},
    "slide29.xml": {
        6: "夕食・朝食・昼食 はプルダウンから選び直します。",
        8: "服薬（夜・朝）・入眠時刻・朝4時チェック・起床時刻 もプルダウンで選び直します。",
        15: "37 / 41",
    },
    "slide30.xml": {15: "38 / 41"},
    "slide31.xml": {19: "39 / 41"},
    "slide32.xml": {},  # PART 2 まとめ - 番号は別位置
    # slide33 表紙裏 / ご活用ください
    "slide33.xml": {4: "Yorimado / 2026.04 版（v2: 1宿泊1行・睡眠3項目化対応）"},
}

# slide32 のページ番号は最後のa:tに含まれる
# slide5 と slide23 は表紙系でページ番号なし

# 新規追加スライドの定義
# (テンプレート元slide, 新slide番号, ページ全体での挿入位置(0-indexed), タイトル一覧)
NEW_SLIDES = [
    # PART 0 扉 (slide5を複製)
    {
        "id": "slide34",
        "template": "slide5.xml",
        "image_target": None,
        "replacements": {
            0: "00",
            1: "PART 0",
            2: "WebViewの初回認証",
            3: "修正ツール（WebView）を初めて開くときに表示される認証画面の通し方です。",
            4: "全5ステップ。一度認証すれば次回以降は不要です。",
        },
    },
    # 認証 A3 (slide7を複製、画像はimage copy 2.png)
    {
        "id": "slide35",
        "template": "slide7.xml",
        "image_target": "imagecopy2.png",
        "replacements": {
            0: "00",
            1: "認証 A3：「prd-oen-yorimado に移動」",
            2: "A3 / A5",
            3: "認証 STEP A3",
            4: "詳細展開後の青リンクをクリックする",
            6: "詳細を展開すると、デベロッパー情報の下にもう少し説明文が表示されます。",
            8: "下部に表示される青字リンク「prd-oen-yorimado(安全ではないページ)に移動」をクリックします。",
            10: "Googleの警告文が並びますが、社内アプリなので「移動」して問題ありません。",
            12: "「安全ではないページ」と書かれていますが、社内専用アプリへの遷移リンクです。",
            13: "認証 STEP A3：詳細展開後の青リンク",
            15: "10 / 41",
        },
    },
    # 認証 A4 (slide7を複製、image copy 3.png)
    {
        "id": "slide36",
        "template": "slide7.xml",
        "image_target": "imagecopy3.png",
        "replacements": {
            0: "00",
            1: "認証 A4：Googleアカウントを選んで「次へ」",
            2: "A4 / A5",
            3: "認証 STEP A4",
            4: "業務用Googleアカウントでログイン",
            6: "「prd-oen-yorimado にログイン」画面が表示されます。",
            8: "業務用Googleアカウント（自分が使うアドレス）が選択されていることを確認します。",
            10: "右下の「次へ」ボタンをクリックします。",
            12: "違うアカウントが選ばれていたら、上部のプルダウンから業務用アドレスに切り替えてください。",
            13: "認証 STEP A4：アカウント確認画面",
            15: "11 / 41",
        },
    },
    # 認証 A5 (slide7を複製、image copy 4.png)
    {
        "id": "slide37",
        "template": "slide7.xml",
        "image_target": "imagecopy4.png",
        "replacements": {
            0: "00",
            1: "認証 A5：権限内容を確認して「続行」",
            2: "A5 / A5",
            3: "認証 STEP A5",
            4: "権限の説明をスクロールして「続行」を押す",
            6: "権限の項目（フォーム表示・メール送信・アプリ実行 等）が一覧表示されます。",
            8: "画面を下までスクロールして、「続行」ボタンを押します。",
            10: "認証完了後、WebView画面が開きます。次回以降この画面は出ません。",
            12: "認証完了後はPART 2「フォームを修正する」の手順に進めます。",
            13: "認証 STEP A5：権限同意画面",
            15: "12 / 41",
        },
    },
    # STEP 11 便 (slide17服薬を複製) - 便はラジオボタンのため服薬画像を流用
    {
        "id": "slide38",
        "template": "slide17.xml",
        "image_target": None,  # 服薬画像を流用（便はラジオボタン形式・服薬と同じUI）
        "replacements": {
            1: "STEP 11：便を記録する",
            2: "11 / 18",
            3: "STEP 11",
            4: "「便」欄を ○ / × から選ぶ",
            6: "排便あり＝○、なし＝× をラジオボタンで選びます。",
            8: "観察できなかった場合は連絡事項に補足を記載してください。",
            10: "便の状態に異変があった場合は、その他連絡事項に詳細を記載してください。",
            12: "排便の有無のみ記録します。詳細（量・状態など）は連絡事項へ。",
            13: "STEP 11：便のラジオボタン",
            14: "（服薬と同じ ○ / × の選択UIです）",
            15: "23 / 41",
        },
    },
    # STEP 12 入眠時刻 (image copy 5.png)
    {
        "id": "slide39",
        "template": "slide17.xml",
        "image_target": "imagecopy5.png",
        "replacements": {
            1: "STEP 12：入眠時刻（必須プルダウン）",
            2: "12 / 18",
            3: "STEP 12",
            4: "「入眠時刻」を 20:30〜22:00 から10分刻みで選ぶ",
            6: "児童が入眠した時刻に最も近い10分刻みを選びます（20:30 / 20:40 / 20:50 / 21:00 / 21:10 / 21:20 / 21:30 / 21:40 / 21:50 / 22:00）。",
            8: "毎回必須の入力です。空欄では送信できません。",
            10: "特例：林夏渚さんは入眠時刻 22:00 を選んでください（運用ルール）。",
            12: "睡眠記録は3項目（入眠時刻・朝4時チェック・起床時刻）に分かれました。",
            13: "STEP 12：入眠時刻のプルダウン",
            14: "20:30〜22:00 / 10分刻み",
            15: "24 / 41",
        },
    },
    # STEP 13 朝4時チェック (image copy 6.png)
    {
        "id": "slide40",
        "template": "slide17.xml",
        "image_target": "imagecopy6.png",
        "replacements": {
            1: "STEP 13：朝4時チェック（必須プルダウン）",
            2: "13 / 18",
            3: "STEP 13",
            4: "朝4時の見回り時の状態を選ぶ",
            6: "「睡眠」または「覚醒確認後に付き添い」のどちらかを選びます。",
            8: "通常通り眠っていれば「睡眠」、目覚めていて付き添った場合は「覚醒確認後に付き添い」。",
            10: "毎回必須の入力です。",
            12: "夜間巡視時の様子を記録する項目です。",
            13: "STEP 13：朝4時チェックのプルダウン",
            14: "睡眠 / 覚醒確認後に付き添い",
            15: "25 / 41",
        },
    },
    # STEP 14 起床時刻 (image copy 7.png)
    {
        "id": "slide41",
        "template": "slide17.xml",
        "image_target": "imagecopy7.png",
        "replacements": {
            1: "STEP 14：起床時刻（必須プルダウン）",
            2: "14 / 18",
            3: "STEP 14",
            4: "「起床時刻」を 6:00〜7:30 から10分刻みで選ぶ",
            6: "児童が起きた時刻に最も近い10分刻みを選びます（6:00 / 6:10 / 6:20 / 6:30 / 6:40 / 6:50 / 7:00 / 7:10 / 7:20 / 7:30）。",
            8: "毎回必須の入力です。空欄では送信できません。",
            10: "起床時刻が範囲外（5:50以前や7:40以降）になった場合は連絡事項に補足を。",
            12: "睡眠3項目の最後（入眠→朝4時→起床）の項目です。",
            13: "STEP 14：起床時刻のプルダウン",
            14: "6:00〜7:30 / 10分刻み",
            15: "26 / 41",
        },
    },
]

# 新スライド順 (slideNN.xml の名前で指定)
# 旧 slide1-33 + 新 slide34-41 から、最終の表示順を作る
NEW_ORDER = [
    "slide1.xml",   # 表紙
    "slide2.xml",   # はじめに
    "slide3.xml",   # 目次
    "slide4.xml",   # 全体像
    "slide34.xml",  # PART 0 扉
    "slide7.xml",   # 認証A1
    "slide19.xml",  # 認証A2
    "slide35.xml",  # 認証A3
    "slide36.xml",  # 認証A4
    "slide37.xml",  # 認証A5
    "slide5.xml",   # PART 1 扉
    "slide6.xml",   # PART 1 全体の流れ
    "slide8.xml",   # STEP 1 スタッフ1
    "slide9.xml",   # STEP 2 スタッフ2
    "slide10.xml",  # STEP 3 児童名
    "slide11.xml",  # STEP 4 入所日時
    "slide12.xml",  # STEP 5 退所予定日時
    "slide20.xml",  # 入退所日時の入力例
    "slide13.xml",  # STEP 6 体温
    "slide14.xml",  # STEP 7-8 食事
    "slide15.xml",  # STEP 9 昼食
    "slide16.xml",  # STEP 10 入浴
    "slide38.xml",  # STEP 11 便
    "slide39.xml",  # STEP 12 入眠時刻
    "slide40.xml",  # STEP 13 朝4時チェック
    "slide41.xml",  # STEP 14 起床時刻
    "slide17.xml",  # STEP 15-16 服薬
    "slide18.xml",  # STEP 17 連絡事項
    "slide21.xml",  # STEP 18 送信
    "slide22.xml",  # PART 1 まとめ
    "slide23.xml",  # PART 2 扉
    "slide24.xml",  # Webビューってなに
    "slide25.xml",  # PART 2 STEP 1-2
    "slide26.xml",  # PART 2 STEP 3-4
    "slide27.xml",  # PART 2 STEP 5
    "slide28.xml",  # PART 2 STEP 6
    "slide29.xml",  # PART 2 STEP 7-10
    "slide30.xml",  # PART 2 STEP 11
    "slide31.xml",  # 削除したいとき
    "slide32.xml",  # PART 2 まとめ
    "slide33.xml",  # ご活用ください
]


def apply_text_replace(slide_xml: Path, idx_map: dict[int, str]) -> int:
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(slide_xml), parser)
    ts = list(tree.getroot().iter(A_T))
    applied = 0
    for idx, new_text in idx_map.items():
        if idx >= len(ts):
            print(f"  [WARN] {slide_xml.name}: index {idx} out of range ({len(ts)})")
            continue
        ts[idx].text = new_text
        applied += 1
    tree.write(str(slide_xml), xml_declaration=True, encoding="UTF-8", standalone=True)
    return applied


def main():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    with zipfile.ZipFile(SRC, "r") as zf:
        zf.extractall(WORK)

    media_dir = WORK / "ppt" / "media"
    slides_dir = WORK / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    # ---- 1. 既存スライドのテキスト置換 ----
    print("\n=== Step 1: existing slide text replacements ===")
    total = 0
    for slide_name, idx_map in EXISTING_REPLACE.items():
        if not idx_map:
            continue
        path = slides_dir / slide_name
        if not path.exists():
            print(f"  [ERROR] missing {slide_name}")
            continue
        applied = apply_text_replace(path, idx_map)
        print(f"  {slide_name}: {applied} replacements")
        total += applied
    print(f"  total: {total}")

    # ---- 2. 認証用画像をmediaへコピー、image-7-1とimage-19-1を上書き ----
    print("\n=== Step 2: image swaps for repurposed slides ===")
    # slide7のimage-7-1.png → image.png 内容で上書き
    shutil.copy(AUTH_IMAGES["image.png"], media_dir / "image-7-1.png")
    print(f"  image-7-1.png <- image.png")
    # slide19のimage-19-1.png → image copy.png 内容で上書き
    shutil.copy(AUTH_IMAGES["imagecopy.png"], media_dir / "image-19-1.png")
    print(f"  image-19-1.png <- image copy.png")
    # 認証A3-A5用の新規画像
    shutil.copy(AUTH_IMAGES["imagecopy2.png"], media_dir / "image-auth-a3.png")
    shutil.copy(AUTH_IMAGES["imagecopy3.png"], media_dir / "image-auth-a4.png")
    shutil.copy(AUTH_IMAGES["imagecopy4.png"], media_dir / "image-auth-a5.png")
    print(f"  +image-auth-a3/a4/a5.png")
    # 新ステップ画像 (入眠 / 朝4時チェック / 起床)
    shutil.copy(AUTH_IMAGES["imagecopy5.png"], media_dir / "image-step-sleep.png")
    shutil.copy(AUTH_IMAGES["imagecopy6.png"], media_dir / "image-step-4am.png")
    shutil.copy(AUTH_IMAGES["imagecopy7.png"], media_dir / "image-step-wakeup.png")
    print(f"  +image-step-sleep/4am/wakeup.png")

    # ---- 3. 新規スライド作成 ----
    print("\n=== Step 3: new slide creation ===")
    for spec in NEW_SLIDES:
        sid = spec["id"]
        template_name = spec["template"]
        # スライドXMLを複製
        new_slide_path = slides_dir / f"{sid}.xml"
        shutil.copy(slides_dir / template_name, new_slide_path)
        # 元templateのrelsも複製
        template_rels = rels_dir / f"{template_name}.rels"
        new_rels = rels_dir / f"{sid}.xml.rels"
        if template_rels.exists():
            shutil.copy(template_rels, new_rels)
        # 画像差替が必要な場合
        if spec["image_target"]:
            # rels内のTarget="../media/image-XXX.png" を新画像に向ける
            target_map = {
                "imagecopy2.png": "image-auth-a3.png",
                "imagecopy3.png": "image-auth-a4.png",
                "imagecopy4.png": "image-auth-a5.png",
                "imagecopy5.png": "image-step-sleep.png",
                "imagecopy6.png": "image-step-4am.png",
                "imagecopy7.png": "image-step-wakeup.png",
            }
            new_img_filename = target_map[spec["image_target"]]
            rels_text = new_rels.read_text(encoding="utf-8")
            # 元templateの画像参照を新画像に書換 (slide7なら image-7-1.png → image-auth-a3.png 等)
            # template画像名を取得
            tmpl_match = re.search(r'Target="\.\./media/([^"]+)"', rels_text)
            if tmpl_match:
                tmpl_img = tmpl_match.group(1)
                rels_text = rels_text.replace(
                    f'Target="../media/{tmpl_img}"',
                    f'Target="../media/{new_img_filename}"',
                    1,
                )
                new_rels.write_text(rels_text, encoding="utf-8")
        # notesSlide参照は削除（新規notesSlideを作らないため）
        if new_rels.exists():
            rels_text = new_rels.read_text(encoding="utf-8")
            rels_text = re.sub(
                r'<Relationship[^/]*Type="[^"]*notesSlide"[^/]*/>',
                '',
                rels_text,
            )
            new_rels.write_text(rels_text, encoding="utf-8")
        # テキスト置換
        applied = apply_text_replace(new_slide_path, spec["replacements"])
        print(f"  {sid}.xml (from {template_name}): {applied} replacements")

    # ---- 4. Content_Types.xml に新スライドを登録 ----
    print("\n=== Step 4: update Content_Types.xml ===")
    ct_path = WORK / "[Content_Types].xml"
    ct_text = ct_path.read_text(encoding="utf-8")
    new_overrides = []
    for spec in NEW_SLIDES:
        sid = spec["id"]
        new_overrides.append(
            f'<Override PartName="/ppt/slides/{sid}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    # 既存の最後の </Types> 直前に挿入
    ct_text = ct_text.replace("</Types>", "".join(new_overrides) + "</Types>")
    ct_path.write_text(ct_text, encoding="utf-8")
    print(f"  added {len(new_overrides)} overrides")

    # ---- 5. presentation.xml.rels に新Relationshipを追加 ----
    print("\n=== Step 5: update presentation.xml.rels ===")
    pres_rels_path = WORK / "ppt" / "_rels" / "presentation.xml.rels"
    pres_rels_text = pres_rels_path.read_text(encoding="utf-8")
    # 既存の最大rIdを取得
    existing_ids = re.findall(r'Id="rId(\d+)"', pres_rels_text)
    next_id = max(int(x) for x in existing_ids) + 1
    slide_to_rid: dict[str, str] = {}
    # 既存slideのrIdマップ
    for m in re.finditer(
        r'<Relationship Id="(rId\d+)"[^>]*Type="[^"]*/slide"[^>]*Target="slides/(slide\d+\.xml)"',
        pres_rels_text,
    ):
        slide_to_rid[m.group(2)] = m.group(1)
    # 新スライド用のrIdを追加
    new_rels_xml = []
    for spec in NEW_SLIDES:
        sid = f"{spec['id']}.xml"
        rid = f"rId{next_id}"
        slide_to_rid[sid] = rid
        new_rels_xml.append(
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/{sid}"/>'
        )
        next_id += 1
    pres_rels_text = pres_rels_text.replace(
        "</Relationships>", "".join(new_rels_xml) + "</Relationships>"
    )
    pres_rels_path.write_text(pres_rels_text, encoding="utf-8")
    print(f"  added {len(new_rels_xml)} new relationships, slide_to_rid for {len(slide_to_rid)} slides")

    # ---- 6. presentation.xml の sldIdLst を再構築 ----
    print("\n=== Step 6: update presentation.xml sldIdLst ===")
    pres_path = WORK / "ppt" / "presentation.xml"
    pres_text = pres_path.read_text(encoding="utf-8")
    new_sldids = []
    sid_counter = 256
    for slide_name in NEW_ORDER:
        if slide_name not in slide_to_rid:
            print(f"  [ERROR] no rId for {slide_name}")
            continue
        rid = slide_to_rid[slide_name]
        new_sldids.append(f'<p:sldId id="{sid_counter}" r:id="{rid}"/>')
        sid_counter += 1
    new_sldidlst = f'<p:sldIdLst>{"".join(new_sldids)}</p:sldIdLst>'
    pres_text = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        new_sldidlst,
        pres_text,
        count=1,
        flags=re.S,
    )
    pres_path.write_text(pres_text, encoding="utf-8")
    print(f"  rebuilt sldIdLst with {len(new_sldids)} slides")

    # ---- 7. zip作成 ----
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
