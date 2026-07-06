# -*- coding: utf-8 -*-
"""
遊戯王PSA10 buylist.xlsm から Mycaアップロード用CSVを出力するスクリプト

保存先:
  C:\\Users\\user\\ClimaxGit\\climax6\\yugioh_psa10_myca_upload.csv

出力形式:
  マスタID,MycaID,カード名,販売店舗価格
  ,,【PSA10】ブラック・マジシャン・ガール,498000

特徴:
  - Excelのヘッダー名に依存しない
  - シート内を全行スキャンして、カード名っぽい文字列と価格を拾う
  - マスタID/MycaIDは空欄で出力
  - CP932で保存
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
INPUT_XLSM = BASE_DIR / "buylist.xlsm"
OUTPUT_CSV = BASE_DIR / "yugioh_psa10_myca_upload.csv"

CSV_ENCODING = "cp932"

CSV_HEADERS = [
    "マスタID",
    "MycaID",
    "カード名",
    "販売店舗価格",
]

MIN_PRICE = 1


NG_NAME_WORDS = [
    "カード名",
    "商品名",
    "買取",
    "買取表",
    "販売店舗価格",
    "価格",
    "金額",
    "マスタ",
    "myca",
    "更新日",
    "合計",
    "平均",
    "件数",
    "なし",
    "未入力",
    "nan",
]


def normalize_text(value) -> str:
    if value is None:
        return ""

    s = str(value)
    s = s.replace("\u3000", " ")
    s = s.replace("\r", " ")
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def parse_price(value) -> Optional[int]:
    if value is None:
        return None

    if isinstance(value, int):
        return value if value > 0 else None

    if isinstance(value, float):
        if value <= 0:
            return None
        return int(value)

    s = normalize_text(value)

    if not s:
        return None

    # 498,000円 / ¥498000 / 498000 など対応
    s = s.replace(",", "")
    s = s.replace("円", "")
    s = s.replace("￥", "")
    s = s.replace("¥", "")
    s = s.strip()

    # パーセントや型番っぽいものは除外
    if "%" in s:
        return None

    m = re.fullmatch(r"\d+", s)
    if not m:
        return None

    price = int(s)

    if price < MIN_PRICE:
        return None

    # 型番や小さい数字を価格として拾いすぎないための最低限フィルタ
    if price <= 0:
        return None

    return price


def looks_like_name(value) -> bool:
    s = normalize_text(value)

    if not s:
        return False

    if len(s) < 2:
        return False

    lower = s.lower()

    for ng in NG_NAME_WORDS:
        if ng.lower() in lower:
            return False

    # 数字だけ、価格だけっぽいものは除外
    price = parse_price(s)
    if price is not None:
        return False

    # URLやファイルパスは除外
    if "http://" in lower or "https://" in lower:
        return False

    if "\\" in s and ":" in s:
        return False

    # 拡張子っぽいものは除外
    if re.search(r"\.(jpg|jpeg|png|webp|gif|html|csv|xlsx|xlsm)$", lower):
        return False

    # 遊戯王カード名に多い文字が含まれていればOK
    if re.search(r"[ぁ-んァ-ン一-龥ー・ヴＡ-Ｚａ-ｚA-Za-z]", s):
        return True

    return False


def clean_name(name: str) -> str:
    s = normalize_text(name)

    # 既に【PSA10】が付いている場合はそのまま
    if s.startswith("【PSA10】"):
        return s

    # 既に PSA10 が含まれている場合は整形だけ
    compact = s.upper().replace(" ", "").replace("　", "")
    if "PSA10" in compact or "ＰＳＡ１０" in s:
        s = re.sub(r"(?i)\s*PSA\s*10\s*", "", s).strip()
        s = s.replace("ＰＳＡ１０", "").strip()

    return f"【PSA10】{s}"


def find_price_cells(row_values: List[object]) -> List[Tuple[int, int]]:
    prices = []

    for idx, value in enumerate(row_values):
        price = parse_price(value)
        if price is not None:
            prices.append((idx, price))

    return prices


def find_name_for_price(row_values: List[object], price_idx: int) -> Optional[str]:
    """
    価格セルの左側を優先してカード名候補を探す。
    同じ行にカード名と価格が並んでいる想定。
    """

    # まず価格の左側を近い順に見る
    for idx in range(price_idx - 1, -1, -1):
        value = row_values[idx]
        if looks_like_name(value):
            return clean_name(str(value))

    # 次に価格の右側も一応見る
    for idx in range(price_idx + 1, len(row_values)):
        value = row_values[idx]
        if looks_like_name(value):
            return clean_name(str(value))

    return None


def extract_from_workbook() -> List[Dict[str, str]]:
    wb = load_workbook(INPUT_XLSM, data_only=True, keep_vba=True)

    rows_out: List[Dict[str, str]] = []
    seen = set()

    for ws in wb.worksheets:
        print(f"[INFO] シート確認中: {ws.title}")

        for row in ws.iter_rows():
            row_values = [cell.value for cell in row]

            price_cells = find_price_cells(row_values)
            if not price_cells:
                continue

            for price_idx, price in price_cells:
                name = find_name_for_price(row_values, price_idx)

                if not name:
                    continue

                key = (name, price)
                if key in seen:
                    continue
                seen.add(key)

                rows_out.append({
                    "マスタID": "",
                    "MycaID": "",
                    "カード名": name,
                    "販売店舗価格": str(price),
                })

    return rows_out


def write_csv(rows: List[Dict[str, str]]) -> None:
    with OUTPUT_CSV.open("w", newline="", encoding=CSV_ENCODING, errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    print("==================================================")
    print("遊戯王PSA10 Myca CSV 出力開始")
    print(f"入力: {INPUT_XLSM}")
    print(f"出力: {OUTPUT_CSV}")
    print("==================================================")

    if not INPUT_XLSM.exists():
        print(f"[ERROR] buylist.xlsm が見つかりません: {INPUT_XLSM}")
        return 1

    try:
        rows = extract_from_workbook()
    except Exception as e:
        print(f"[ERROR] Excel読み込み・抽出に失敗しました: {e}")
        return 1

    if not rows:
        print("[ERROR] カードデータを1件も抽出できませんでした。")
        print("[HINT] buylist.xlsm のカード名と価格が同じ行にあるか確認してください。")
        return 1

    try:
        write_csv(rows)
    except Exception as e:
        print(f"[ERROR] CSV保存に失敗しました: {e}")
        return 1

    print("==================================================")
    print("[OK] Mycaアップロード用CSVを出力しました")
    print(f"件数: {len(rows)}")
    print(f"CSV: {OUTPUT_CSV}")
    print("==================================================")

    return 0


if __name__ == "__main__":
    sys.exit(main())