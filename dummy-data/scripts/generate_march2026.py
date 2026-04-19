"""
2026年3月分データ生成スクリプト
PDFから抽出した19名の実績データをベースに各種CSVを生成
"""

import csv
import random
from pathlib import Path
from datetime import date

random.seed(2026_03)

BASE_DIR = Path("dummy-data")
CHILD_REPORT_DIR = BASE_DIR / "児童別" / "実績報告書"
CHILD_JUDO_DIR = BASE_DIR / "児童別" / "重度支援加算"
MONTH = "2026-03"
YEAR, MON = 2026, 3

# === PDFから抽出した全19名の出席日 ===
PDF_CHILDREN = {
    "溝口一花": {
        "cert": "9200364314",
        "days": [3,4,5,6,7,8,9,10,11,12,13,14,16,17,18,19,20,21,23,24],
        "judo": True,
    },
    "小野凌来": {
        "cert": "9200373703",  # 既存データに合わせる（PDF上は9200377703）
        "days": [2,3,4,5,6,7,9,10,11,12,16,17,18,19,23,24,25,26,30,31],
        "judo": True,
    },
    "藤川結梨": {
        "cert": "9200750009",
        "days": [1,2,3,8,9,26,27],
        "judo": False,
    },
    "寺田菫": {
        "cert": "9200771880",
        "days": [2,3,4,5,6,7,9,10,11,12,13,14,16,17,19,20,23,24,28,29],
        "judo": False,
    },
    "四宮鈴": {
        "cert": "9200935337",
        "days": [1,2,3,4,5,6,8,9,10,11,12,13,17,18,20,21,24,25,26,27,29,30,31],
        "judo": False,
    },
    "石原晴海": {
        "cert": "9200185800",  # 新規・障害児氏名空欄
        "days": [1,2,3,7,8,12,13,14,15,19,20,21,22,28,29],
        "judo": False,
    },
    "松岡蘭": {
        "cert": "9200616903",
        "days": [1,2,3,4,5,6,8,9,10,11,12,13,17,18,19,20,23,24,25,26,27,28,29,30,31],
        "judo": False,
    },
    "園田葵": {
        "cert": "9200701069",
        "days": [3,4,5,6,7,8,10,11,14,15,17,18,21,22,24,25,28,29,30,31],
        "judo": False,
    },
    "園田柚真": {
        "cert": "9200701655",
        "days": [7,8,12,13,14,15,21,22,28,29],
        "judo": False,
    },
    "富永日向": {
        "cert": "9200757988",
        "days": [1,2,3,4,5,6,8,9,10,11,12,13,14,17,18,19,20,24,25,26,27,28,29,30,31],
        "judo": False,
    },
    "西原早彩": {
        "cert": "9200886449",  # 新規
        "days": [7,8,9,10,14,15,16,17,21,22,23,24,28,29,30],
        "judo": False,
    },
    "向井瑠華": {
        "cert": "9200948090",
        "days": [1,7,8,14,15,17,18,19,20,21,22,26,27,28,29],
        "judo": False,
    },
    "向井永愛": {
        "cert": "9200948256",
        "days": [1,7,8,14,15,17,18,19,20,21,22,24,25,28,29],
        "judo": False,
    },
    "深津": {
        "cert": "9200758051",  # 新規・障害児名が姓のみ
        "days": [6,7,20,21],
        "judo": False,
    },
    "藤崎柊威": {
        "cert": "9200892116",
        "days": [1,2,3,9,10,16,17],
        "judo": False,
    },
    "辻村優樹": {
        "cert": "9200945153",
        "days": [1,2,4,5,11,12,18,19,25,26],
        "judo": False,
    },
    "香村快": {
        "cert": "9200539980",
        "days": [4,5,6,7,11,12,13,14,18,19,20,21,25,26,27,28],
        "judo": True,
    },
    "香村慧": {
        "cert": "9200934983",
        "days": [4,5,6,7,11,12,13,14,18,19,20,21,25,26,27,28],
        "judo": True,
    },
    "笠原英二次": {
        "cert": "9200664275",
        "days": [4,5,18,19],
        "judo": False,
    },
}

# 曜日名
WEEKDAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]

# ランダムデータ生成用
MEAL_CHOICES = ["完食", "完食", "完食", "半分", "半分", "食べなかった"]
MEAL_CHOICES_LUNCH = ["完食", "完食", "半分", "半分"]
BATH_CHOICES = ["○", "○", "○", "○", "×"]
BOWEL_CHOICES = ["○", "○", "○", "×", "×"]
MED_CHOICES = ["○", "○", "○", "×"]
REMARKS = [
    "楽しそうに過ごされていました。",
    "シール交換をしていました。",
    "ユーチューブを見ていました。",
    "お食事をお代わりされていました。",
    "積み木で遊んでいました。",
    "絵本で遊んでいました。",
    "お菓子パーティーを楽しみました。",
    "テレビの順番をお友達に譲ることができました。",
    "お友達とお風呂に入っていました。",
    "デイでのお出かけの話をしてくれました。",
    "お料理のお手伝いをしてくれていました。",
    "色塗りしていました。",
    "お風呂に進んで入られてました。",
    "トランプをして遊んでいました。",
    "マグネットで遊んでいました。",
    "車で遊んでいました。",
    "ピアノで遊んでいました。",
    "折り紙を折っていました。",
    "パズルで遊んでいました。",
    "ブロックで遊んでいました。",
    "お絵かきをしていました。",
    "おままごとをしていました。",
    "DVDを見て過ごしていました。",
    "お友達と仲良く遊んでいました。",
    "歌を歌って楽しんでいました。",
]

# 重度支援加算の×コメント
JUDO_X_COMMENTS = [
    "×（悪態をついていたが、声かけで落ち着きました）",
    "×（大声を出していたが、しばらくして落ち着きました）",
    "×（興奮気味でしたが、クールダウンして安定しました）",
    "×（不穏な様子でしたが、個別対応で安定しました）",
    "×（物を投げようとしたが、制止して落ち着きました）",
]


def gen_temp():
    """体温をランダム生成（36.0〜37.1）"""
    return f"{random.choice([36.0, 36.1, 36.2, 36.3, 36.4, 36.5, 36.6, 36.7, 36.8, 36.9, 37.0, 37.1])}"


def gen_meal_row(d):
    """食事・入浴等のランダムデータを生成"""
    weekday = d.weekday()
    # 土日は昼食あり（ランダム）、平日は基本なし
    has_lunch = weekday >= 5 and random.random() < 0.5
    # 夕食: たまに食べないパターンあり
    dinner = random.choice(MEAL_CHOICES) if random.random() > 0.1 else "−"
    breakfast = random.choice(MEAL_CHOICES)
    lunch = random.choice(MEAL_CHOICES_LUNCH) if has_lunch else "−"

    return {
        "体温": gen_temp(),
        "夕食": dinner,
        "朝食": breakfast,
        "昼食": lunch,
        "入浴": random.choice(BATH_CHOICES),
        "便": random.choice(BOWEL_CHOICES),
        "服薬(夜)": random.choice(MED_CHOICES),
        "服薬(朝)": random.choice(MED_CHOICES),
        "その他連絡事項": random.choice(REMARKS),
    }


def gen_judo_row(d, name, cert):
    """重度支援加算の1行を生成"""
    hours = ["17:00","18:00","19:00","20:00","21:00","22:00","23:00",
             "0:00","1:00","2:00","3:00","4:00","5:00","6:00","7:00","8:00","12:30"]
    row = {"日付": d.strftime("%Y-%m-%d"), "受給者証番号": cert, "児童名": name}

    for h in hours:
        if h == "12:30":
            row[h] = "−"
        elif random.random() < 0.05:
            row[h] = random.choice(JUDO_X_COMMENTS)
        else:
            row[h] = "○"

    row["睡眠(21:00)"] = "○"
    row["睡眠(4:00)"] = random.choice(["○", "○", "○", "×"])
    row["睡眠(7:00)"] = random.choice(["○", "○", "○", "×"])

    return row


def main():
    print("=== 2026年3月分データ生成（PDF19名） ===\n")

    march_attendance = {}
    for name, info in PDF_CHILDREN.items():
        march_attendance[name] = {
            "cert": info["cert"],
            "days": info["days"],
            "judo": info["judo"],
        }
        flags = []
        if info["judo"]:
            flags.append("重度")
        print(f"  {name} ({info['cert']}): {len(info['days'])}日" + (f" [{','.join(flags)}]" if flags else ""))

    total_days = sum(len(v["days"]) for v in march_attendance.values())
    judo_count = sum(1 for v in march_attendance.values() if v["judo"])
    print(f"\n  合計: {len(march_attendance)}名, 延べ{total_days}日, 重度対象{judo_count}名")

    # === 1. 児童別 実績報告書CSV生成 ===
    print(f"\n--- 児童別 実績報告書CSV生成 ---")
    child_report_header = ["日付","受給者証番号","児童名","体温","夕食","朝食","昼食","入浴","便","服薬(夜)","服薬(朝)","その他連絡事項"]

    for name, info in sorted(march_attendance.items()):
        cert = info["cert"]
        filename = CHILD_REPORT_DIR / f"{name}_{cert}_{MONTH}.csv"
        rows = []
        for day in sorted(info["days"]):
            d = date(YEAR, MON, day)
            meal = gen_meal_row(d)
            row = {"日付": d.strftime("%Y-%m-%d"), "受給者証番号": cert, "児童名": name}
            row.update(meal)
            rows.append(row)

        with open(filename, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=child_report_header)
            writer.writeheader()
            writer.writerows(rows)
        print(f"  {filename.name}: {len(rows)}行")

    # === 2. 児童別 重度支援加算CSV生成 ===
    print(f"\n--- 児童別 重度支援加算CSV生成 ---")
    judo_header = ["日付","受給者証番号","児童名",
                   "17:00","18:00","19:00","20:00","21:00","22:00","23:00",
                   "0:00","1:00","2:00","3:00","4:00","5:00","6:00","7:00","8:00","12:30",
                   "睡眠(21:00)","睡眠(4:00)","睡眠(7:00)"]

    for name, info in sorted(march_attendance.items()):
        if not info["judo"]:
            continue
        cert = info["cert"]
        filename = CHILD_JUDO_DIR / f"{name}_{cert}_{MONTH}.csv"
        rows = []
        for day in sorted(info["days"]):
            d = date(YEAR, MON, day)
            rows.append(gen_judo_row(d, name, cert))

        with open(filename, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=judo_header)
            writer.writeheader()
            writer.writerows(rows)
        print(f"  {filename.name}: {len(rows)}行")

    # === 3. 統合 実績報告書CSV に追記 ===
    print(f"\n--- 統合 実績報告書CSV追記 ---")
    combined_report = BASE_DIR / "実績報告書.csv"
    combined_header = ["日付","受給者証番号","児童名","入所時間","退所時間","体温","夕食","朝食","昼食","入浴","便","服薬(夜)","服薬(朝)","その他連絡事項"]

    # 乱数シードをリセットして児童別と統合で同じデータにする
    random.seed(2026_03_99)

    all_combined = []
    for name, info in march_attendance.items():
        cert = info["cert"]
        for day in sorted(info["days"]):
            d = date(YEAR, MON, day)
            meal = gen_meal_row(d)
            row = {
                "日付": d.strftime("%Y-%m-%d"),
                "受給者証番号": cert,
                "児童名": name,
                "入所時間": "17:00",
                "退所時間": "8:00",
            }
            row.update(meal)
            all_combined.append(row)

    all_combined.sort(key=lambda r: (r["日付"], r["児童名"]))

    # 末尾に改行があるか確認して追記
    with open(combined_report, "rb") as f:
        f.seek(-1, 2)
        last_byte = f.read(1)
    needs_newline = last_byte not in (b'\n', b'\r')

    with open(combined_report, "a", encoding="utf-8-sig", newline="") as f:
        if needs_newline:
            f.write("\r\n")
        writer = csv.DictWriter(f, fieldnames=combined_header)
        writer.writerows(all_combined)
    print(f"  追記: {len(all_combined)}行")

    # === 4. 統合 重度支援加算CSV に追記 ===
    print(f"\n--- 統合 重度支援加算CSV追記 ---")
    combined_judo = BASE_DIR / "重度支援加算_全児童.csv"

    random.seed(2026_03_88)

    all_judo = []
    for name, info in march_attendance.items():
        if not info["judo"]:
            continue
        cert = info["cert"]
        for day in sorted(info["days"]):
            d = date(YEAR, MON, day)
            row = gen_judo_row(d, name, cert)
            all_judo.append(row)

    all_judo.sort(key=lambda r: (r["日付"], r["児童名"]))

    with open(combined_judo, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        for row in all_judo:
            writer.writerow([
                row["日付"], row["受給者証番号"], row["児童名"],
                *[row.get(h, "") for h in ["17:00","18:00","19:00","20:00","21:00","22:00","23:00",
                                            "0:00","1:00","2:00","3:00","4:00","5:00","6:00","7:00","8:00","12:30"]],
                "", "", "", "",
            ])
    print(f"  追記: {len(all_judo)}行")

    print(f"\n=== 完了 ===")
    print(f"合計児童数: {len(march_attendance)}名")
    print(f"合計実績行: {len(all_combined)}行")
    print(f"合計重度支援行: {len(all_judo)}行")


if __name__ == "__main__":
    main()
