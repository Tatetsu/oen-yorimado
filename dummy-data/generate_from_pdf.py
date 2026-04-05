#!/usr/bin/env python3
"""
実績PDFからダミーデータを生成するスクリプト

書類1: 実績報告書（全児童）
書類2: 重度支援加算チェック記録（対象児童のみ）

構成:
  児童別/実績報告書/  — 児童ごとの実績報告書CSV
  児童別/重度支援加算/ — 児童ごとの重度支援加算CSV
  実績報告書_全児童.csv — 実績報告書の統合ファイル
  重度支援加算_全児童.csv — 重度支援加算の統合ファイル

再申請対応:
  同一年月の申請が複数PDFに含まれる場合がある（修正再申請）。
  PDFはファイル名でソートし古い順に処理。同一児童・同一月は後のPDFで上書き。
  マージ時も (受給者証番号, 日付) をキーに新しいファイルが優先。

ルール:
  - 日付: YYYY-MM-DD形式（令和→西暦変換）
  - 曜日: なし
  - 他サービス併給の日 → 昼食は「−」
  - 重度支援加算: 17:00〜翌8:00毎時 → 12:30 → 睡眠チェック
  - 重度12:30: 他サービス併給→「−」、それ以外→「○」
  - 睡眠チェック: 基本21:00（林夏渚のみ22:00）、4:00、7:00
"""

import fitz
import re
import csv
import os
import glob
import random

random.seed(42)

# === 設定 ===
PDF_DIR = "../Gmail"
PDF_PATTERN = "*実績.pdf"
JISSEKI_DIR = "児童別/実績報告書"
JUDO_DIR = "児童別/重度支援加算"
MERGED_JISSEKI = "実績報告書_全児童.csv"
MERGED_JUDO = "重度支援加算_全児童.csv"

# 重度支援加算対象者
SEVERE_CHILDREN = {
    "9200364314": "溝口一花",
    "9200373703": "小野凌来",
    "9200809896": "林夏渚",
    "9200539980": "香村快",
    "9200934983": "香村慧",
}

# 林夏渚の受給者証番号（睡眠チェック22時特例）
HAYASHI_KANA_CERT = "9200809896"

# === CSV ヘッダー ===
JISSEKI_HEADER = [
    "日付", "受給者証番号", "児童名", "体温",
    "夕食", "朝食", "昼食", "入浴", "便",
    "服薬(夜)", "服薬(朝)", "その他連絡事項",
]

# 重度支援加算: 17:00〜翌8:00毎時(16列) + 12:30 + 睡眠チェック3列
JUDO_TIME_COLS = []
for h in range(17, 24):
    JUDO_TIME_COLS.append(f"{h}:00")
for h in range(0, 9):
    JUDO_TIME_COLS.append(f"{h}:00")
JUDO_TIME_COLS.append("12:30")

# === ダミーデータ生成用定義 ===
FOOD_CHOICES = ["完食", "半分", "食べなかった", "−"]
FOOD_WEIGHTS = [0.55, 0.30, 0.10, 0.05]

BATH_WEIGHTS = [0.9, 0.1]
STOOL_WEIGHTS = [0.7, 0.3]
MED_WEIGHTS = [0.85, 0.15]

CONTACT_NOTES = [
    "ユーチューブを見ていました。",
    "デイでのお出かけの話をしてくれました。",
    "トランプをして遊んでいました。",
    "お風呂に進んで入られてました。",
    "新しいお友達に優しくしてくれてました。",
    "シール交換をしていました。",
    "お料理のお手伝いをしてくれていました。",
    "お菓子パーティーを楽しみました。",
    "他のお友達の面倒を見てくれていました。",
    "お食事をお代わりされていました。",
    "夜間しっかり眠れたと話していました。",
    "トランポリンで遊んでいました。",
    "ピアノで遊んでいました。",
    "プラレールで遊んでいました。",
    "お食事と美味しいともりもり食べていました。",
    "楽しそうに過ごされていました。",
    "お友達とお風呂に入っていました。",
    "テレビの順番をお友達に譲ることができました。",
    "シーツを敷くのを手伝ってくれていました。",
    "車で遊んでいました。",
    "積み木で遊んでいました。",
    "色塗りしていました。",
    "ダイエットマシーンで遊んでいました。",
    "机を拭いてくれていました。",
    "絵本で遊んでいました。",
    "磁石で遊んでいました。",
    "マグネットで遊んでいました。",
]

WEEKDAY_SET = {"月", "火", "水", "木", "金", "土", "日"}


def generate_temp():
    temp = random.gauss(36.5, 0.3)
    temp = max(36.0, min(37.5, temp))
    return f"{temp:.1f}"


def weighted_choice(choices, weights):
    return random.choices(choices, weights=weights, k=1)[0]


def parse_year_month(page):
    """ページのヘッダー領域（y=45-60）から令和年月を抽出"""
    blocks = page.get_text("dict")["blocks"]
    header_texts = []
    for b in blocks:
        if "lines" in b:
            for line in b["lines"]:
                for s in line["spans"]:
                    if 45 < s["bbox"][1] < 60:
                        header_texts.append((s["bbox"][0], s["text"]))
    header_texts.sort(key=lambda x: x[0])
    header = "".join([t[1] for t in header_texts])
    m = re.search(r"令和\s*(\d+)\s*年\s*(\d+)\s*月", header)
    if m:
        reiwa = int(m.group(1))
        month = int(m.group(2))
        western = 2018 + reiwa
        return western, month
    return None, None


def parse_page(page, page_num, pdf_name):
    """1ページ分のデータを抽出"""
    western_year, month = parse_year_month(page)
    if not western_year:
        print(f"  WARNING: {pdf_name} Page {page_num} - 年月抽出失敗")
        return None

    text = page.get_text()
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    cert_no = None
    parent_name = None
    child_name = None

    for j, line in enumerate(lines):
        m = re.match(r"^(9200\d+)$", line)
        if m:
            cert_no = m.group(1)
            # j+1=親名, j+2=児童名 が基本だが、
            # 一部PDFでは j+1=児童名, j+2=事業所番号 のパターンがある
            candidates = []
            for offset in range(1, 4):
                if j + offset < len(lines):
                    candidates.append(lines[j + offset])
            # 事業所番号(2712...)を除外して親名・児童名を特定
            names = [c for c in candidates if not re.match(r"^\d+$", c)
                     and "短期入所" not in c and "事業" not in c]
            if len(names) >= 2:
                parent_name = names[0]
                child_name = names[1]
            elif len(names) == 1:
                child_name = names[0]
                parent_name = ""
            break

    if not cert_no or not child_name:
        print(f"  WARNING: {pdf_name} Page {page_num} - 児童情報抽出失敗")
        return None

    # 利用日パース
    day_data = []
    for j, line in enumerate(lines):
        if line in WEEKDAY_SET and j > 0:
            if re.match(r"^\d{1,2}$", lines[j - 1]):
                day_num = int(lines[j - 1])
                has_other = False
                for k in range(j + 1, min(j + 15, len(lines))):
                    if lines[k] == "他サービス併給":
                        has_other = True
                        break
                    if lines[k] == "合計":
                        break
                    if (
                        re.match(r"^\d{1,2}$", lines[k])
                        and k + 1 < len(lines)
                        and lines[k + 1] in WEEKDAY_SET
                    ):
                        break
                day_data.append({"day": day_num, "other_service": has_other})

    clean_name = child_name.replace("　", "").replace(" ", "")

    return {
        "page": page_num,
        "year": western_year,
        "month": month,
        "cert_no": cert_no,
        "parent": parent_name,
        "child": clean_name,
        "days": day_data,
        "total": len(day_data),
        "is_severe": cert_no in SEVERE_CHILDREN,
    }


# === 書類1: 実績報告書 ===

def generate_jisseki_rows(child):
    """1児童・1ヶ月分の実績報告書データを生成"""
    rows = []
    for day in sorted(child["days"], key=lambda x: x["day"]):
        date_str = f"{child['year']}-{child['month']:02d}-{day['day']:02d}"
        temp = generate_temp()
        dinner = weighted_choice(FOOD_CHOICES, FOOD_WEIGHTS)
        breakfast = weighted_choice(FOOD_CHOICES, FOOD_WEIGHTS)

        if day["other_service"]:
            lunch = "−"
        else:
            lunch = weighted_choice(FOOD_CHOICES, FOOD_WEIGHTS)

        bath = weighted_choice(["○", "×"], BATH_WEIGHTS)
        stool = weighted_choice(["○", "×"], STOOL_WEIGHTS)
        med_night = weighted_choice(["○", "×"], MED_WEIGHTS)
        med_morning = weighted_choice(["○", "×"], MED_WEIGHTS)
        note = random.choice(CONTACT_NOTES)

        rows.append([
            date_str, child["cert_no"], child["child"], temp,
            dinner, breakfast, lunch, bath, stool,
            med_night, med_morning, note,
        ])
    return rows


# === 書類2: 重度支援加算チェック記録 ===

def generate_judo_rows(child):
    """1児童・1ヶ月分の重度支援加算データを生成
    列順: 17:00〜翌8:00(16列) → 12:30(1列) → 睡眠チェック(3列)
    """
    is_hayashi = child["cert_no"] == HAYASHI_KANA_CERT
    rows = []

    for day in sorted(child["days"], key=lambda x: x["day"]):
        date_str = f"{child['year']}-{child['month']:02d}-{day['day']:02d}"
        row = [date_str, child["cert_no"], child["child"]]

        # 17:00〜翌8:00 毎時チェック（16列）: 基本○、2%で×
        for _ in range(16):
            if random.random() < 0.02:
                row.append("×")
            else:
                row.append("○")

        # 12:30: 他サービス併給 → 「−」、それ以外 → 「○」
        if day["other_service"]:
            row.append("−")
        else:
            row.append("○")

        # 睡眠チェック1: 21:00（林夏渚は22:00）
        row.append("○" if random.random() < 0.95 else "×")

        # 睡眠チェック2: 4:00（時々×）
        row.append("○" if random.random() < 0.75 else "×")

        # 睡眠チェック3: 7:00
        row.append("○" if random.random() < 0.95 else "×")

        rows.append(row)
    return rows, is_hayashi


# === CSV書き込み・マージ ===

def write_csv(filepath, header, rows):
    """CSVファイルを書き込み"""
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    return filepath


def make_judo_header(is_hayashi):
    """重度支援加算のヘッダーを生成"""
    sleep1 = "睡眠(22:00)" if is_hayashi else "睡眠(21:00)"
    return ["日付", "受給者証番号", "児童名"] + JUDO_TIME_COLS + [
        sleep1, "睡眠(4:00)", "睡眠(7:00)",
    ]


def merge_csvs(individual_dir, output_path, header):
    """個別CSVをマージ。同一キーはファイル更新日時が新しい方を採用。"""
    csv_files = sorted(glob.glob(os.path.join(individual_dir, "*.csv")))
    csv_files.sort(key=lambda f: os.path.getmtime(f))

    merged = {}
    for csv_file in csv_files:
        with open(csv_file, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            next(reader)  # skip header
            for row in reader:
                if len(row) >= 2:
                    key = (row[1], row[0])  # (受給者証番号, 日付)
                    merged[key] = row

    all_rows = sorted(merged.values(), key=lambda r: (r[2], r[0]))

    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(all_rows)

    return len(all_rows)


def find_pdfs(pdf_dir, pattern):
    """PDFファイルをファイル名でソート（古い順）して返す"""
    pdf_files = sorted(glob.glob(os.path.join(pdf_dir, pattern)))
    return pdf_files


def main():
    print("=== ダミーデータ生成（全PDF） ===\n")

    pdf_files = find_pdfs(PDF_DIR, PDF_PATTERN)
    print(f"対象PDF: {len(pdf_files)}ファイル\n")

    os.makedirs(JISSEKI_DIR, exist_ok=True)
    os.makedirs(JUDO_DIR, exist_ok=True)

    all_children = []

    # 全PDF処理（古い順 → 新しいもので上書き）
    for pdf_path in pdf_files:
        pdf_name = os.path.basename(pdf_path)
        doc = fitz.open(pdf_path)
        print(f"[{pdf_name}] ({doc.page_count}ページ)")

        for i in range(doc.page_count):
            child = parse_page(doc[i], i + 1, pdf_name)
            if child:
                all_children.append(child)

        doc.close()

    print(f"\n合計: {len(all_children)}エントリ\n")

    # --- 書類1: 実績報告書 ---
    print("--- 書類1: 実績報告書 ---")
    jisseki_count = 0
    for child in all_children:
        rows = generate_jisseki_rows(child)
        filename = f"{child['child']}_{child['cert_no']}_{child['year']}-{child['month']:02d}.csv"
        write_csv(os.path.join(JISSEKI_DIR, filename), JISSEKI_HEADER, rows)
        jisseki_count += 1
        severe_mark = " [重度]" if child["is_severe"] else ""
        print(f"  {child['child']} - {child['year']}-{child['month']:02d} "
              f"({child['total']}日){severe_mark}")

    total_jisseki = merge_csvs(JISSEKI_DIR, MERGED_JISSEKI, JISSEKI_HEADER)
    print(f"\n  個別: {jisseki_count}ファイル → {JISSEKI_DIR}/")
    print(f"  統合: {MERGED_JISSEKI} ({total_jisseki}行)\n")

    # --- 書類2: 重度支援加算チェック記録 ---
    print("--- 書類2: 重度支援加算チェック記録 ---")
    judo_count = 0
    for child in all_children:
        if not child["is_severe"]:
            continue
        rows, is_hayashi = generate_judo_rows(child)
        header = make_judo_header(is_hayashi)
        filename = f"{child['child']}_{child['cert_no']}_{child['year']}-{child['month']:02d}.csv"
        write_csv(os.path.join(JUDO_DIR, filename), header, rows)
        judo_count += 1
        sleep_note = " (睡眠22時)" if is_hayashi else ""
        print(f"  {child['child']} - {child['year']}-{child['month']:02d} "
              f"({child['total']}日){sleep_note}")

    merged_judo_header = make_judo_header(False)
    total_judo = merge_csvs(JUDO_DIR, MERGED_JUDO, merged_judo_header)
    print(f"\n  個別: {judo_count}ファイル → {JUDO_DIR}/")
    print(f"  統合: {MERGED_JUDO} ({total_judo}行)\n")

    print("=== 完了 ===")


if __name__ == "__main__":
    main()
