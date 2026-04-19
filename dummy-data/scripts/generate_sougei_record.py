"""
送迎記録表生成スクリプト
- 実績報告書CSV（児童別・日別）から送迎記録表CSVを生成する
- 複数月の一括生成に対応
- 仕様: docs/送迎記録表_仕様.md を参照
"""

import csv
import glob
import random
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# === 設定 ===
TARGET_MONTHS = [
    "2025-03", "2025-04", "2025-05", "2025-06",
    "2025-07", "2025-08", "2025-09", "2025-10",
    "2025-11", "2025-12", "2026-01", "2026-02",
]
INPUT_DIR = Path("dummy-data/児童別/実績報告書")

# 固定値
DRIVER = "只津"          # 送迎職員
ASSISTANT = "溝口"       # 介助者
TRANSPORT_METHOD = "車"  # 送迎方法（徒歩・自動車）
CAR_NUMBER = "707"       # 車番
ROLL_CALL_CHECK = "なし"  # 点呼確認による異常（疾病・疲労・飲酒等）

# 迎え順の最後尾に配置する児童（他施設利用のため）
LAST_PICKUP = {"寺田菫", "園田柚真", "園田葵"}

# 櫻井兄弟: 月曜・水曜は夕方送迎（迎え）不要
# デイサービスが「よりまど」まで送ってくれるため
SAKURAI_NO_EVENING_PICKUP = {"櫻井麻智", "櫻井斗麻", "櫻井帆斗"}
SAKURAI_SKIP_WEEKDAYS = {"月", "水"}  # 月曜・水曜は夕迎え対象から除外

# 曜日変換
WEEKDAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]

# 乱数シード（再現性のため）
random.seed(42)


def load_attendance(target_month: str) -> dict:
    """実績CSVから日別の出席情報を集約する"""
    files = glob.glob(str(INPUT_DIR / f"*{target_month}*.csv"))
    # date -> list of {name, has_lunch, has_breakfast, has_dinner}
    daily = defaultdict(list)

    for f in sorted(files):
        with open(f, encoding="utf-8-sig") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                date = row["日付"]
                name = row["児童名"]
                lunch = row.get("昼食", "−")
                breakfast = row.get("朝食", "−")
                dinner = row.get("夕食", "−")
                daily[date].append({
                    "name": name,
                    "has_lunch": lunch not in ("−", "", None),
                    "has_breakfast": breakfast not in ("−", "", None),
                    "has_dinner": dinner not in ("−", "", None),
                })

    return daily


def get_weekday(date_str: str) -> str:
    """日付文字列から曜日を返す"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return WEEKDAY_NAMES[dt.weekday()]


def random_time_between(start_h, start_m, end_h, end_m):
    """指定範囲内でランダムな時刻を生成（5分刻み）"""
    start_total = start_h * 60 + start_m
    end_total = end_h * 60 + end_m
    candidates = list(range(start_total, end_total + 1, 5))
    chosen = random.choice(candidates)
    h, m = divmod(chosen, 60)
    return f"{h}:{m:02d}"


def order_children(names: list[str]) -> list[str]:
    """児童をランダム順に並べ、指定の3名を最後尾に配置する"""
    last = [n for n in names if n in LAST_PICKUP]
    others = [n for n in names if n not in LAST_PICKUP]
    random.shuffle(others)
    random.shuffle(last)
    return others + last


def has_late_pickup_children(names: list[str]) -> bool:
    """最後尾配置の児童（寺田・園田兄弟）が迎え対象にいるか"""
    return any(n in LAST_PICKUP for n in names)


def format_children_list(names: list[str]) -> str:
    """児童名リストを「①○○ ②○○ ...」形式の文字列にする"""
    circled = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
    parts = []
    for i, name in enumerate(names):
        if i < len(circled):
            parts.append(f"{circled[i]}{name}")
        else:
            parts.append(f"({i+1}){name}")
    return " ".join(parts)


def make_record(date, weekday, category, children_str, start, end, remarks=""):
    """1レコード分の辞書を生成する"""
    return {
        "日付": date,
        "曜日": weekday,
        "送迎区分": category,
        "送迎職員": DRIVER,
        "介助者": ASSISTANT,
        "送迎方法": TRANSPORT_METHOD,
        "車番": CAR_NUMBER,
        "点呼確認による異常": ROLL_CALL_CHECK,
        "利用者": children_str,
        "開始時刻": start,
        "終了時刻": end,
        "備考": remarks,
    }


def generate_records(daily: dict) -> list[dict]:
    """送迎記録のレコードを生成する（日付×送迎区分ごとに1行）"""
    records = []

    for date in sorted(daily.keys()):
        entries = daily[date]
        all_names = [e["name"] for e in entries]
        weekday = get_weekday(date)

        # --- 夕迎え ---
        # 櫻井兄弟は月曜・水曜の夕迎えから除外
        if weekday in SAKURAI_SKIP_WEEKDAYS:
            evening_candidates = [n for n in all_names if n not in SAKURAI_NO_EVENING_PICKUP]
        else:
            evening_candidates = list(all_names)

        evening_names = order_children(evening_candidates)
        has_late = has_late_pickup_children(evening_names)
        pickup_start = random_time_between(17, 30, 17, 40)
        if has_late:
            pickup_end = "18:40"
        else:
            pickup_end = random.choice(["18:20", "18:30", "18:40"])

        records.append(make_record(
            date, weekday, "夕迎え",
            format_children_list(evening_names),
            pickup_start, pickup_end,
        ))

        # --- 朝送り（朝食ありの児童） ---
        morning_names = [e["name"] for e in entries if e["has_breakfast"]]
        if morning_names:
            morning_ordered = order_children(morning_names)
            morning_start = random_time_between(7, 40, 7, 50)
            morning_end = random_time_between(8, 40, 8, 45)
            records.append(make_record(
                date, weekday, "朝送り",
                format_children_list(morning_ordered),
                morning_start, morning_end,
            ))

        # --- 昼送り（昼食ありの児童） ---
        lunch_names = [e["name"] for e in entries if e["has_lunch"]]
        if lunch_names:
            lunch_ordered = order_children(lunch_names)
            records.append(make_record(
                date, weekday, "昼送り",
                format_children_list(lunch_ordered),
                "12:40", "13:20",
            ))

    return records


FIELDNAMES = [
    "日付", "曜日", "送迎区分",
    "送迎職員", "介助者", "送迎方法", "車番",
    "点呼確認による異常", "利用者",
    "開始時刻", "終了時刻", "備考",
]


def main():
    from collections import Counter

    all_records = []

    for month in TARGET_MONTHS:
        random.seed(42)  # 月ごとにシードをリセットして再現性確保

        daily = load_attendance(month)
        records = generate_records(daily)
        all_records.extend(records)

        print(f"{month}: {len(records)}件")
        by_type = Counter(r["送迎区分"] for r in records)
        for t, c in sorted(by_type.items()):
            print(f"  {t}: {c}件")

    # 全月まとめて1ファイルに出力
    output_file = Path("dummy-data/送迎記録表_全期間.csv")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(all_records)

    print(f"\n生成完了: {output_file}")
    print(f"合計レコード数: {len(all_records)}")


if __name__ == "__main__":
    main()
