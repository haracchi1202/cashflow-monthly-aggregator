"""月別集計・期間フィルター・ファイル別集計。

excel_reader.parse_file() が返す FileResult のリストから、Streamlit 表示・
比較レポート出力で使う DataFrame をまとめて構築する。
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from excel_reader import FileResult


def build_dataframes(
    results: list[FileResult],
    restored_keys: set[str],
    month_from: str | None,
    month_to: str | None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """(月別集計, 入金明細, 出金明細, 除外行, ファイル別集計) を返す。

    Parameters
    ----------
    results
        excel_reader.parse_file() の結果リスト。
    restored_keys
        ユーザーが手動で「復活」した除外行のキー ("ファイル名::Excel行番号") 集合。
    month_from / month_to
        集計対象の期間フィルター。None の場合はフィルターしない。
    """
    detail_records: list[dict[str, Any]] = []
    excluded_records: list[dict[str, Any]] = []

    for r in results:
        for row in r.detail_rows:
            detail_records.append(row)
        for row in r.excluded_rows:
            key = f"{row['file_name']}::{row['excel_row']}"
            if key in restored_keys and row.get("month") and row.get("amount") is not None:
                # 復活分は明細に追加（除外理由は付けたまま記録に残す）
                restored_row = dict(row)
                restored_row["restored"] = True
                detail_records.append(restored_row)
            else:
                excluded_records.append(row)

    detail_df = pd.DataFrame(detail_records)
    excluded_df = pd.DataFrame(excluded_records)

    if not detail_df.empty:
        if month_from:
            detail_df = detail_df[detail_df["month"] >= month_from]
        if month_to:
            detail_df = detail_df[detail_df["month"] <= month_to]

    income_df = (
        detail_df[detail_df["kind"] == "入金"].copy() if not detail_df.empty else pd.DataFrame()
    )
    expense_df = (
        detail_df[detail_df["kind"] == "出金"].copy() if not detail_df.empty else pd.DataFrame()
    )

    if detail_df.empty:
        monthly_df = pd.DataFrame(columns=["month", "入金合計", "出金合計", "差額"])
    else:
        inc = (
            income_df.groupby("month", as_index=False)["amount"].sum().rename(
                columns={"amount": "入金合計"}
            )
            if not income_df.empty
            else pd.DataFrame(columns=["month", "入金合計"])
        )
        exp = (
            expense_df.groupby("month", as_index=False)["amount"].sum().rename(
                columns={"amount": "出金合計"}
            )
            if not expense_df.empty
            else pd.DataFrame(columns=["month", "出金合計"])
        )
        monthly_df = pd.merge(inc, exp, on="month", how="outer").fillna(0.0)
        monthly_df["差額"] = monthly_df["入金合計"] - monthly_df["出金合計"]
        monthly_df = monthly_df.sort_values("month").reset_index(drop=True)

    # ファイル別集計
    file_records = []
    for r in results:
        included_amount = sum((row.get("amount") or 0) for row in r.detail_rows)
        included_count = len(r.detail_rows)
        excluded_count = len(r.excluded_rows)

        restored_in_file = [
            row
            for row in r.excluded_rows
            if f"{row['file_name']}::{row['excel_row']}" in restored_keys
            and row.get("amount") is not None
        ]
        included_amount += sum((row.get("amount") or 0) for row in restored_in_file)
        included_count += len(restored_in_file)
        excluded_count -= len(restored_in_file)

        file_records.append(
            {
                "ファイル名": r.file_name,
                "区分": r.kind,
                "合計金額": included_amount,
                "集計件数": included_count,
                "除外件数": excluded_count,
            }
        )
    file_summary_df = pd.DataFrame(file_records)

    return monthly_df, income_df, expense_df, excluded_df, file_summary_df


def monthly_summary_to_dict(monthly_df: pd.DataFrame) -> dict[str, dict[str, float]]:
    """月別集計を {YYYY-MM: {"入金合計": .., "出金合計": ..}} の辞書に変換。"""
    out: dict[str, dict[str, float]] = {}
    if monthly_df.empty:
        return out
    for _, row in monthly_df.iterrows():
        ym = str(row["month"])
        out[ym] = {
            "入金合計": float(row.get("入金合計", 0) or 0),
            "出金合計": float(row.get("出金合計", 0) or 0),
        }
    return out
