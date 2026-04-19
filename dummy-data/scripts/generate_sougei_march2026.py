"""
2026年3月分の送迎記録表を生成
既存の generate_sougei_record.py のロジックを使い、3月分のみ生成
"""

import csv
import sys
from pathlib import Path

# 既存スクリプトのロジックをインポートするため、パスを追加
sys.path.insert(0, str(Path(__file__).parent))

import generate_sougei_record as gen
import random

MONTH = "2026-03"
FIELDNAMES = gen.FIELDNAMES


def main():
    from collections import Counter

    random.seed(42)

    daily = gen.load_attendance(MONTH)
    if not daily:
        print(f"エラー: {MONTH}の出席データが見つかりません")
        return

    records = gen.generate_records(daily)

    print(f"{MONTH}: {len(records)}件")
    by_type = Counter(r["送迎区分"] for r in records)
    for t, c in sorted(by_type.items()):
        print(f"  {t}: {c}件")

    # 月別ファイル出力
    monthly_file = Path(f"dummy-data/送迎記録表_{MONTH}.csv")
    with open(monthly_file, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)
    print(f"\n月別ファイル生成: {monthly_file}")

    # 全期間ファイルに追記
    all_file = Path("dummy-data/送迎記録表_全期間.csv")
    with open(all_file, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writerows(records)
    print(f"全期間ファイルに追記: {len(records)}件")


if __name__ == "__main__":
    main()
