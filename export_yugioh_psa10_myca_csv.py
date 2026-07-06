# -*- coding: utf-8 -*-
"""
遊戯王PSA10 buylist.xlsm から Mycaアップロード用CSVを出力するスクリプト

出力先:
  docs/myca/yugioh_psa10_myca_upload.csv

前提:
  - このファイルを C:\\Users\\user\\ClimaxGit\\climax6 に置く
  - 同じフォルダに buylist.xlsm がある
  - openpyxl が必要
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook


# ==================================================
# 設定
# ==================================================

BASE_DIR = Path(__file__).resolve().parent
INPUT_XLSM = BASE_DIR / "buylist.xlsm"

OUTPUT_DIR = BASE_DIR / "docs" / "myca"
OUTPUT_CSV = OUTPUT_DIR / "yugioh_psa10_myca_upload.csv"

# Myca側のCSVヘッダー
# Mycaのテンプレート列名が違う場合はここだけ変更してください。
CSV_HEADERS = [
    "マスタID",
    "MycaID",
    "カード名",
    "販売店舗価格",
]

# 価格がこの金額未満なら出力しない
MIN_PRICE = 1

# カード名にPSA10表記を付けたい場合 True
# すでにカード名に「PSA10」が含まれている場合は重複追加しません。
ADD_PSA10_TO_NAME = True

# カード名の末尾に付ける表記
PSA10_SUFFIX = " PSA10"

# CSV文字コード
# Mycaアップロード用は cp932 が無難です。
CSV_ENCODING = "cp932"


# ==================================================
# Excel列名ゆれ対応
# ==================================================

HEADER_CANDIDATES = {
    "master_id": [
        "マスタID",
        "マスターID",
        "master_id",
        "masterid",
        "商品ID",
        "管理ID",
    ],
    "myca_id": [
        "MycaID",
        "MYCAID",
        "myca_id",
        "mycaid",
        "Myca ID",
        "MYCA ID",
    ],
    "name": [
        "カード名",
        "商品名",
        "名前",
        "name",
        "product_name",
    ],
    "price": [
        "販売店舗価格",
        "買取価格",
        "買取金額",
        "価格",
        "金額",
        "amount",
        "price",
        "buy_price",
        "kaitori_price",
    ],
}


# ==================================================
# 共通処理
# ==================================================

def normalize_text(value) -> str:
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\u3000", " ")
    s = s.strip()
    return s


def normalize_header(value) -> str:
    s = normalize_text(value)
    s = s.lower()
    s = re.sub(r"[\s_\-　]+", "", s)
    return s


def parse_price(value) -> Optional[int]:
    if value is None:
        return None

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        return int(value)

    s = str(value)
    s = s.replace(",", "")
    s = s.replace("円", "")
    s = s.replace("￥", "")
    s = s.replace("¥", "")
    s = s.strip()

    m = re.search(r"-?\d+", s)
    if not m:
        return None

    try:
        return int(m.group(0))
    except ValueError:
        return None


def clean_card_name(name: str) -> str:
    name = normalize_text(name)

    # 余計な連続スペースだけ整理
    name = re.sub(r"\s+", " ", name).strip()

    if ADD_PSA10_TO_NAME and "PSA10" not in name.upper().replace(" ", ""):
        name = f"{name}{PSA10_SUFFIX}"

    return name


def find_header_row_and_columns(ws) -> Tuple[int, Dict[str, int]]:
    """
    Excelの上から30行を見て、必要列が一番多く見つかる行をヘッダー行として採用。
    戻り値:
      header_row: 1始まりの行番号
      columns: {"name": 3, "price": 7 ...} のような1始まり列番号
    """

    normalized_candidates = {
        key: [normalize_header(x) for x in values]
        for key, values in HEADER_CANDIDATES.items()
    }

    best_row = None
    best_columns: Dict[str, int] = {}
    best_score = 0

    max_scan_row = min(ws.max_row, 30)

    for row_idx in range(1, max_scan_row + 1):
        row_values = [ws.cell(row_idx, col_idx).value for col_idx in range(1, ws.max_column + 1)]

        found: Dict[str, int] = {}

        for col_idx, value in enumerate(row_values, start=1):
            h = normalize_header(value)
            if not h:
                continue

            for key, candidates in normalized_candidates.items():
                if key in found:
                    continue
                if h in candidates:
                    found[key] = col_idx

        score = len(found)

        # カード名と価格がある行を優先
        if "name" in found and "price" in found:
            score += 10

        if score > best_score:
            best_score = score
            best_row = row_idx
            best_columns = found

    if best_row is None or "name" not in best_columns or "price" not in best_columns:
        raise RuntimeError(
            "ヘッダー行を特定できませんでした。"
            " Excel内に「カード名」と「買取価格」または「販売店舗価格」系の列があるか確認してください。"
        )

    return best_row, best_columns


def get_cell(ws, row_idx: int, columns: Dict[str, int], key: str):
    col_idx = columns.get(key)
    if not col_idx:
        return None
    return ws.cell(row_idx, col_idx).value


def make_output_row(ws, row_idx: int, columns: Dict[str, int]) -> Optional[Dict[str, str]]:
    raw_name = get_cell(ws, row_idx, columns, "name")
    raw_price = get_cell(ws, row_idx, columns, "price")

    name = clean_card_name(normalize_text(raw_name))
    price = parse_price(raw_price)

    if not name:
        return None

    if price is None:
        return None

    if price < MIN_PRICE:
        return None

    master_id = normalize_text(get_cell(ws, row_idx, columns, "master_id"))
    myca_id = normalize_text(get_cell(ws, row_idx, columns, "myca_id"))

    row = {
        "マスタID": master_id,
        "MycaID": myca_id,
        "カード名": name,
        "販売店舗価格": str(price),
    }

    return row


def choose_sheet(wb):
    """
    基本はアクティブシート。
    ただし、アクティブシートでヘッダー検出できない場合は全シートを順に試す。
    """

    sheets = [wb.active] + [ws for ws in wb.worksheets if ws.title != wb.active.title]

    last_error = None

    for ws in sheets:
        try:
            header_row, columns = find_header_row_and_columns(ws)
            return ws, header_row, columns
        except Exception as e:
            last_error = e

    raise RuntimeError(f"有効なシートが見つかりませんでした: {last_error}")


def write_csv(rows: List[Dict[str, str]]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with OUTPUT_CSV.open("w", newline="", encoding=CSV_ENCODING, errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    print("==================================================")
    print("遊戯王PSA10 Mycaアップロード用CSV 出力開始")
    print(f"入力: {INPUT_XLSM}")
    print(f"出力: {OUTPUT_CSV}")
    print("==================================================")

    if not INPUT_XLSM.exists():
        print(f"[ERROR] buylist.xlsm が見つかりません: {INPUT_XLSM}")
        return 1

    try:
        wb = load_workbook(INPUT_XLSM, data_only=True, keep_vba=True)
    except Exception as e:
        print(f"[ERROR] Excelファイルを開けませんでした: {e}")
        return 1

    try:
        ws, header_row, columns = choose_sheet(wb)
    except Exception as e:
        print(f"[ERROR] {e}")
        return 1

    print(f"[INFO] 使用シート: {ws.title}")
    print(f"[INFO] ヘッダー行: {header_row}")
    print(f"[INFO] 検出列: {columns}")

    rows: List[Dict[str, str]] = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        row = make_output_row(ws, row_idx, columns)
        if row:
            rows.append(row)

    if not rows:
        print("[WARN] 出力対象データが0件でした。")
        print("[WARN] カード名列・価格列・価格の値を確認してください。")
        return 1

    try:
        write_csv(rows)
    except Exception as e:
        print(f"[ERROR] CSV出力に失敗しました: {e}")
        return 1

    print("==================================================")
    print("[OK] Mycaアップロード用CSVを出力しました")
    print(f"件数: {len(rows)}")
    print(f"CSV: {OUTPUT_CSV}")
    print("==================================================")

    return 0


if __name__ == "__main__":
    sys.exit(main())