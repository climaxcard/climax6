# -*- coding: utf-8 -*-
"""
buylist.xlsm の中身をそのまま Mycaアップロード用CSVとして出力するスクリプト

入力:
  C:\\Users\\user\\ClimaxGit\\climax6\\buylist.xlsm

出力:
  C:\\Users\\user\\ClimaxGit\\climax6\\yugioh_psa10_myca_upload.csv

仕様:
  - Sheet1を使用
  - ExcelのA列〜最終列までをそのままCSV化
  - 1行目〜最終行までそのまま出力
  - 列位置をズラさない
  - 販売店舗価格列、買取価格列も元Excelと同じ位置のまま
  - 空欄セルも空欄として保持
  - 文字コードはMyca向けに cp932
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent

INPUT_XLSM = BASE_DIR / "buylist.xlsm"
OUTPUT_CSV = BASE_DIR / "yugioh_psa10_myca_upload.csv"

SHEET_NAME = "Sheet1"

CSV_ENCODING = "cp932"


def cell_to_csv_value(value: Any) -> str:
    """
    Excelセル値をCSV用の文字列に変換する。
    Noneは空欄。
    数値は余計な .0 が出ないようにする。
    """

    if value is None:
        return ""

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)

    return str(value)


def find_last_used_row(ws) -> int:
    """
    実データがある最終行を探す。
    openpyxlのmax_rowだけだと装飾だけの行を拾うことがあるため、
    値が入っている最終行を下から探す。
    """

    for row_idx in range(ws.max_row, 0, -1):
        for col_idx in range(1, ws.max_column + 1):
            if ws.cell(row_idx, col_idx).value not in (None, ""):
                return row_idx
    return 1


def find_last_used_col(ws) -> int:
    """
    実データがある最終列を探す。
    """

    for col_idx in range(ws.max_column, 0, -1):
        for row_idx in range(1, ws.max_row + 1):
            if ws.cell(row_idx, col_idx).value not in (None, ""):
                return col_idx
    return 1


def export_sheet_to_csv() -> int:
    print("==================================================")
    print("buylist.xlsm そのままCSV出力 開始")
    print(f"入力: {INPUT_XLSM}")
    print(f"出力: {OUTPUT_CSV}")
    print("==================================================")

    if not INPUT_XLSM.exists():
        print(f"[ERROR] buylist.xlsm が見つかりません: {INPUT_XLSM}")
        return 1

    try:
        wb = load_workbook(INPUT_XLSM, data_only=True, keep_vba=True)
    except Exception as e:
        print(f"[ERROR] buylist.xlsm を開けませんでした: {e}")
        return 1

    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb.active
        print(f"[WARN] {SHEET_NAME} が見つからないため、アクティブシートを使用します: {ws.title}")

    last_row = find_last_used_row(ws)
    last_col = find_last_used_col(ws)

    print(f"[INFO] 使用シート: {ws.title}")
    print(f"[INFO] 出力範囲: 1行目〜{last_row}行目 / 1列目〜{last_col}列目")

    try:
        with OUTPUT_CSV.open("w", newline="", encoding=CSV_ENCODING, errors="replace") as f:
            writer = csv.writer(f)

            for row_idx in range(1, last_row + 1):
                row_values = []

                for col_idx in range(1, last_col + 1):
                    value = ws.cell(row_idx, col_idx).value
                    row_values.append(cell_to_csv_value(value))

                writer.writerow(row_values)

    except Exception as e:
        print(f"[ERROR] CSV保存に失敗しました: {e}")
        return 1

    print("==================================================")
    print("[OK] buylist.xlsm の中身をそのままCSV出力しました")
    print(f"CSV: {OUTPUT_CSV}")
    print("==================================================")

    return 0


if __name__ == "__main__":
    sys.exit(export_sheet_to_csv())