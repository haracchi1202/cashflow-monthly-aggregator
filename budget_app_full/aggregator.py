"""月別集計とファイル別集計。

入力は excel_reader.results_to_records() で得た「集計対象レコード」リスト。
出力は次の DataFrame 群:

    - monthly_df : 月別の確定/予測/合算 入金・支払・差額
    - by_source_dfs: 4種 (確定入金/確定支払/予測入金/予測支払) ごとの明細 DataFrame
    - file_summary_df: ファイルごとの件数・合計金額・除外件数
    - excluded_df: 除外行
    - monthly_summary_for_update: 更新モードごとに {YYYY-MM: {入金合計, 出金合計}} の辞書
"""
from __future__ import annotations

from typing import Any, Iterable

import pandas as pd

from excel_reader import FileResult


SOURCE_GROUPS = ["確定入金", "確定支払", "予測入金", "予測支払"]


def build_dataframes(
    results: list[FileResult],
    restored_keys: set[str],
    month_from: str | None,
    month_to: str | None,
) -> tuple[
    pd.DataFrame,           # monthly_df (確定/予測/合算 列をすべて持つ)
    dict[str, pd.DataFrame],  # by_source_dfs: {"確定入金": df, ...}
    pd.DataFrame,           # excluded_df
    pd.DataFrame,           # file_summary_df
]:
    detail_records: list[dict[str, Any]] = []
    excluded_records: list[dict[str, Any]] = []
    for r in results:
        for row in r.detail_rows:
            detail_records.append(row)
        for row in r.excluded_rows:
            key = f"{row['source_file']}::{row['raw_row_index']}"
            if (
                key in restored_keys
                and row.get("target_month")
                and row.get("amount") is not None
            ):
                restored = dict(row)
                restored["restored"] = True
                detail_records.append(restored)
            else:
                excluded_records.append(row)

    detail_df = pd.DataFrame(detail_records)
    excluded_df = pd.DataFrame(excluded_records)

    if not detail_df.empty:
        if month_from:
            detail_df = detail_df[detail_df["target_month"] >= month_from]
        if month_to:
            detail_df = detail_df[detail_df["target_month"] <= month_to]

    # ソースグループ別 DataFrame
    by_source_dfs: dict[str, pd.DataFrame] = {}
    for grp in SOURCE_GROUPS:
        if detail_df.empty:
            by_source_dfs[grp] = pd.DataFrame()
        else:
            sub = detail_df[detail_df["source_group"] == grp].copy()
            if not sub.empty:
                sub = sub.sort_values(["target_month", "raw_row_index"]).reset_index(drop=True)
            by_source_dfs[grp] = sub

    # 月別集計
    if detail_df.empty:
        cols = [
            "target_month",
            "確定入金", "確定支払", "確定差額",
            "予測入金", "予測支払", "予測差額",
            "合算入金", "合算支払", "合算差額",
        ]
        monthly_df = pd.DataFrame(columns=cols)
    else:
        def _sum(grp_name: str) -> pd.DataFrame:
            sub = by_source_dfs[grp_name]
            if sub.empty:
                return pd.DataFrame(columns=["target_month", grp_name])
            agg = sub.groupby("target_month", as_index=False)["amount"].sum()
            agg = agg.rename(columns={"amount": grp_name})
            return agg

        confirmed_in = _sum("確定入金")
        confirmed_out = _sum("確定支払")
        forecast_in = _sum("予測入金")
        forecast_out = _sum("予測支払")

        monthly_df = confirmed_in
        for other in (confirmed_out, forecast_in, forecast_out):
            monthly_df = pd.merge(monthly_df, other, on="target_month", how="outer")
        monthly_df = monthly_df.fillna(0.0)

        monthly_df["確定差額"] = monthly_df["確定入金"] - monthly_df["確定支払"]
        monthly_df["予測差額"] = monthly_df["予測入金"] - monthly_df["予測支払"]
        monthly_df["合算入金"] = monthly_df["確定入金"] + monthly_df["予測入金"]
        monthly_df["合算支払"] = monthly_df["確定支払"] + monthly_df["予測支払"]
        monthly_df["合算差額"] = monthly_df["合算入金"] - monthly_df["合算支払"]
        monthly_df = monthly_df.sort_values("target_month").reset_index(drop=True)
        monthly_df = monthly_df[
            [
                "target_month",
                "確定入金", "確定支払", "確定差額",
                "予測入金", "予測支払", "予測差額",
                "合算入金", "合算支払", "合算差額",
            ]
        ]

    # ファイル別集計
    file_records = []
    for r in results:
        included_amount = sum((row.get("amount") or 0) for row in r.detail_rows)
        included_count = len(r.detail_rows)
        excluded_count = len(r.excluded_rows)
        restored_in_file = [
            row
            for row in r.excluded_rows
            if f"{row['source_file']}::{row['raw_row_index']}" in restored_keys
            and row.get("amount") is not None
        ]
        included_amount += sum((row.get("amount") or 0) for row in restored_in_file)
        included_count += len(restored_in_file)
        excluded_count -= len(restored_in_file)

        file_records.append(
            {
                "ファイル名": r.file_name,
                "区分": r.source_group,
                "確定/予測": "確定" if r.transaction_status == "confirmed" else "予測",
                "合計金額": included_amount,
                "集計件数": included_count,
                "除外件数": excluded_count,
            }
        )
    file_summary_df = pd.DataFrame(file_records)

    return monthly_df, by_source_dfs, excluded_df, file_summary_df


def confirmed_summary_dict(monthly_df: pd.DataFrame) -> dict[str, dict[str, float]]:
    """{YYYY-MM: {"入金合計": 確定入金, "出金合計": 確定支払}} を返す。"""
    out: dict[str, dict[str, float]] = {}
    if monthly_df.empty:
        return out
    for _, row in monthly_df.iterrows():
        ym = str(row["target_month"])
        out[ym] = {
            "入金合計": float(row.get("確定入金", 0) or 0),
            "出金合計": float(row.get("確定支払", 0) or 0),
        }
    return out


def forecast_summary_dict(monthly_df: pd.DataFrame) -> dict[str, dict[str, float]]:
    """{YYYY-MM: {"入金合計": 予測入金, "出金合計": 予測支払}} を返す。"""
    out: dict[str, dict[str, float]] = {}
    if monthly_df.empty:
        return out
    for _, row in monthly_df.iterrows():
        ym = str(row["target_month"])
        out[ym] = {
            "入金合計": float(row.get("予測入金", 0) or 0),
            "出金合計": float(row.get("予測支払", 0) or 0),
        }
    return out


def update_scope_label(scope: str) -> str:
    return {
        "both": "確定行＋予測行 両方に反映",
        "confirmed_only": "確定行のみに反映",
        "forecast_only": "予測行のみに反映",
    }.get(scope, scope)
