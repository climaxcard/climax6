# -*- coding: utf-8 -*-
"""
遊戯王PSA10 buylist.xlsm から Mycaアップロード用CSVを出力するスクリプト

保存先:
  C:\\Users\\user\\ClimaxGit\\climax6\\yugioh_psa10_myca_upload.csv

出力形式:
  マスタID,MycaID,カード名,販売店舗価格
  1145665,600565960,【PSA10】カオス・ソルジャー －開闢の使者－,14000

修正ポイント:
  - マスタID / MycaID を販売価格として拾わない
  - カード名が「PSA」だけになる行を出力しない
  - 元データ上の列の意味を保って Myca CSV に出力する
  - 出力CSVの列順は固定: マスタID, MycaID, カード名, 販売店舗価格
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook


# ==================================================
# パス設定
# ==================================================

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


# ==================================================
# 設定
# ==================================================

MIN_PRICE = 1

# 価格としてあり得る最大値
# ID誤検出防止用。遊戯王PSA10の買取価格として通常ここまでで十分。
MAX_REASONABLE_PRICE = 10_000_000

# MycaIDっぽい大きな数字を価格扱いしないためのしきい値
# 例: 600565960 は価格ではなくMycaID
ID_LIKE_MIN = 100_000_000

NG_NAME_EXACT = {
    "",
    "psa",
    "psa10",
    "ＰＳＡ",
    "ＰＳＡ１０",
    "カード名",
    "商品名",
    "買取",
    "買取価格",
    "販売店舗価格",
    "価格",
    "金額",
    "マスタid",
    "マスターid",
    "mycaid",
    "myca id",
}

NG_NAME_CONTAINS = [
    "買取表",
    "更新日",
    "販売店舗価格",
    "マスタ",
    "myca",
    "合計",
    "平均",
    "件数",
    "未入力",
]


# ==================================================
# 基本関数
# ==================================================

def normalize_text(value) -> str:
    if value is None:
        return ""

    s = str(value)
    s = s.replace("\u3000", " ")
    s = s.replace("\r", " ")
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def normalize_key(value) -> str:
    s = normalize_text(value)
    s = s.lower()
    s = s.replace("　", " ")
    s = re.sub(r"\s+", "", s)
    return s


def parse_int(value) -> Optional[int]:
    if value is None:
        return None

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return int(value)

    s = normalize_text(value)
    if not s:
        return None

    s = s.replace(",", "")
    s = s.replace("円", "")
    s = s.replace("￥", "")
    s = s.replace("¥", "")
    s = s.strip()

    if "%" in s:
        return None

    if not re.fullmatch(r"\d+", s):
        return None

    return int(s)


def is_psa_marker(value) -> bool:
    s = normalize_key(value)
    return s in {
        "psa",
        "psa10",
        "psa鑑定",
        "鑑定",
        "グレード",
        "grade",
    }


def is_header_like(value) -> bool:
    s = normalize_key(value)
    return s in {
        "マスタid",
        "マスターid",
        "mycaid",
        "カード名",
        "商品名",
        "販売店舗価格",
        "買取価格",
        "価格",
        "金額",
        "psa",
    }


def looks_like_name(value) -> bool:
    s = normalize_text(value)
    if not s:
        return False

    key = normalize_key(s)

    if key in NG_NAME_EXACT:
        return False

    for ng in NG_NAME_CONTAINS:
        if normalize_key(ng) in key:
            return False

    # 数字だけはカード名ではない
    if parse_int(s) is not None:
        return False

    # URLやパスは除外
    lower = s.lower()
    if "http://" in lower or "https://" in lower:
        return False

    if "\\" in s and ":" in s:
        return False

    if re.search(r"\.(jpg|jpeg|png|webp|gif|html|csv|xlsx|xlsm)$", lower):
        return False

    # PSAだけのセルはカード名ではない
    if key in {"psa", "psa10"}:
        return False

    # 遊戯王カード名っぽい文字があるか
    if re.search(r"[ぁ-んァ-ン一-龥ー・ヴA-Za-zＡ-Ｚａ-ｚ]", s):
        return True

    return False


def clean_name(name: str) -> str:
    s = normalize_text(name)

    # 既存のPSA表記を除去してから、先頭に【PSA10】を統一付与
    s = re.sub(r"^【\s*PSA\s*10\s*】", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(r"^【\s*ＰＳＡ\s*１０\s*】", "", s).strip()
    s = re.sub(r"(?i)\s*PSA\s*10\s*$", "", s).strip()
    s = s.replace("ＰＳＡ１０", "").strip()

    # 念のため「PSA」だけになったものは空にする
    if normalize_key(s) in {"psa", "psa10", "ｐｓａ", "ｐｓａ１０"}:
        return ""

    return f"【PSA10】{s}"


def is_reasonable_price(n: int) -> bool:
    if n < MIN_PRICE:
        return False

    if n > MAX_REASONABLE_PRICE:
        return False

    # 9桁以上はMycaIDなどのIDとみなす
    if n >= ID_LIKE_MIN:
        return False

    return True


# ==================================================
# 行解析
# ==================================================

def values_from_row(row) -> List[object]:
    return [cell.value for cell in row]


def find_card_name_index(values: List[object]) -> Optional[int]:
    """
    行内のカード名候補を探す。
    基本的に一番左側のカード名っぽい文字列を採用。
    """

    candidates: List[Tuple[int, str]] = []

    for idx, value in enumerate(values):
        s = normalize_text(value)
        if looks_like_name(s):
            candidates.append((idx, s))

    if not candidates:
        return None

    # 「PSA」より前にある候補を優先
    psa_indexes = [i for i, v in enumerate(values) if is_psa_marker(v)]
    if psa_indexes:
        first_psa = min(psa_indexes)
        before_psa = [(i, s) for i, s in candidates if i < first_psa]
        if before_psa:
            return before_psa[0][0]

    return candidates[0][0]


def find_psa_index(values: List[object]) -> Optional[int]:
    for idx, value in enumerate(values):
        if is_psa_marker(value):
            return idx
    return None


def find_price(values: List[object], name_idx: Optional[int], psa_idx: Optional[int]) -> Optional[int]:
    """
    販売店舗価格を探す。

    優先順位:
      1. PSA列の右側にある妥当な価格
      2. カード名より右側にある妥当な価格のうち、IDっぽくない最後の価格
    """

    # 1. PSA列の右側を最優先
    if psa_idx is not None:
        for idx in range(psa_idx + 1, len(values)):
            n = parse_int(values[idx])
            if n is None:
                continue
            if is_reasonable_price(n):
                return n

    # 2. カード名より右側から探す
    start = 0
    if name_idx is not None:
        start = name_idx + 1

    nums: List[int] = []

    for idx in range(start, len(values)):
        n = parse_int(values[idx])
        if n is None:
            continue

        if not is_reasonable_price(n):
            continue

        nums.append(n)

    if not nums:
        return None

    # ID列が先にあって最後に価格があるケースが多いので最後を採用
    return nums[-1]


def find_ids(values: List[object], name_idx: Optional[int], psa_idx: Optional[int], price: Optional[int]) -> Tuple[str, str]:
    """
    マスタID / MycaID を推定する。

    基本想定:
      カード名, マスタID, MycaID, PSA, 価格

    価格は除外して、PSA列より左側の数字をIDとして拾う。
    """

    master_id = ""
    myca_id = ""

    if name_idx is None:
        return master_id, myca_id

    end = len(values)
    if psa_idx is not None:
        end = psa_idx

    nums: List[int] = []

    for idx in range(name_idx + 1, end):
        n = parse_int(values[idx])
        if n is None:
            continue

        if price is not None and n == price:
            continue

        nums.append(n)

    # 例:
    # カード名, 1145665, 600565960, PSA, 14000
    if len(nums) >= 1:
        master_id = str(nums[0])

    if len(nums) >= 2:
        myca_id = str(nums[1])

    return master_id, myca_id


def parse_data_row(values: List[object]) -> Optional[Dict[str, str]]:
    if not values:
        return None

    # ヘッダー行っぽいものは除外
    joined_key = "".join(normalize_key(v) for v in values if v is not None)
    if "カード名" in joined_key and ("販売店舗価格" in joined_key or "買取価格" in joined_key):
        return None

    name_idx = find_card_name_index(values)
    if name_idx is None:
        return None

    raw_name = normalize_text(values[name_idx])
    name = clean_name(raw_name)

    # カード名がPSAだけ等になった場合は除外
    if not name or normalize_key(name) in {"【psa10】psa", "【psa10】psa10"}:
        return None

    psa_idx = find_psa_index(values)
    price = find_price(values, name_idx, psa_idx)

    if price is None:
        return None

    master_id, myca_id = find_ids(values, name_idx, psa_idx, price)

    return {
        "マスタID": master_id,
        "MycaID": myca_id,
        "カード名": name,
        "販売店舗価格": str(price),
    }


# ==================================================
# Excel抽出 / CSV出力
# ==================================================

def extract_rows() -> List[Dict[str, str]]:
    wb = load_workbook(INPUT_XLSM, data_only=True, keep_vba=True)

    rows_out: List[Dict[str, str]] = []
    seen = set()

    for ws in wb.worksheets:
        print(f"[INFO] シート確認中: {ws.title}")

        for row in ws.iter_rows():
            values = values_from_row(row)
            parsed = parse_data_row(values)

            if not parsed:
                continue

            key = (
                parsed["マスタID"],
                parsed["MycaID"],
                parsed["カード名"],
                parsed["販売店舗価格"],
            )

            if key in seen:
                continue

            seen.add(key)
            rows_out.append(parsed)

    return rows_out


def write_csv(rows: List[Dict[str, str]]) -> None:
    with OUTPUT_CSV.open("w", newline="", encoding=CSV_ENCODING, errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS, extrasaction="ignore")
        writer.writeheader()

        for row in rows:
            # 列順固定
            writer.writerow({
                "マスタID": row.get("マスタID", ""),
                "MycaID": row.get("MycaID", ""),
                "カード名": row.get("カード名", ""),
                "販売店舗価格": row.get("販売店舗価格", ""),
            })


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
        rows = extract_rows()
    except Exception as e:
        print(f"[ERROR] Excel読み込み・抽出に失敗しました: {e}")
        return 1

    if not rows:
        print("[ERROR] カードデータを1件も抽出できませんでした。")
        print("[HINT] buylist.xlsm の並びが想定と違う可能性があります。")
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

    # 先頭5件プレビュー
    print("")
    print("[PREVIEW]")
    for row in rows[:5]:
        print(row)

    return 0


if __name__ == "__main__":
    sys.exit(main())