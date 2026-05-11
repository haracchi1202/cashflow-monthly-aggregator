"""比較レポート Excel の生成。

11 シート構成:
    1. 月別集計
    2. 入金明細
    3. 出金明細
    4. 除外行一覧
    5. ファイル別集計
    6. 上書き前データ
    7. 上書き後データ
    8. 上書き前後比較
    9. 更新対象一覧
    10. 未照合月一覧
    11. 処理ログ
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Any

import pandas as pd

from excel_reader import format_month_jp
from budget_updater import BudgetCellMap, PlanRow


YEN_FMT = "#,##0\"円\""
NEG_YEN_FMT = "#,##0\"円\";[Red]-#,##0\"円\""


def _yen_format(workbook):
    return workbook.add_format({"num_format": YEN_FMT})


def _neg_yen_format(workbook):
    return workbook.add_format({"num_format": NEG_YEN_FMT})


def _header_format(workbook):
    return workbook.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})


def _flag_format(workbook, color: str):
    return workbook.add_format({"bg_color": color, "border": 1})


def _write_headers(ws, columns: list[str], fmt) -> None:
    for col_num, val in enumerate(columns):
        ws.write(0, col_num, val, fmt)


def _detail_sheet(writer, sheet_name: str, df: pd.DataFrame, yen_fmt, header_fmt) -> None:
    if df.empty:
        out = pd.DataFrame(columns=["ファイル名", "行番号", "月", "金額", "商談名"])
    else:
        out = df.copy()
        keep = ["file_name", "excel_row", "month", "amount", "deal_name"]
        out = out[[c for c in keep if c in out.columns]]
        out.columns = ["ファイル名", "行番号", "月", "金額", "商談名"]
        if "月" in out.columns:
            out["月"] = out["月"].apply(lambda v: format_month_jp(v) if v else "")
    out.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.set_column("A:A", 36)
    ws.set_column("B:B", 8)
    ws.set_column("C:C", 12)
    ws.set_column("D:D", 18, yen_fmt)
    ws.set_column("E:E", 40)
    _write_headers(ws, list(out.columns), header_fmt)


def build_report(
    monthly_df: pd.DataFrame,
    income_df: pd.DataFrame,
    expense_df: pd.DataFrame,
    excluded_df: pd.DataFrame,
    file_summary_df: pd.DataFrame,
    cell_maps: list[BudgetCellMap],
    plan_rows: list[PlanRow],
    summary_only_months: list[str],
    budget_only_months: list[str],
    written_log: list[dict[str, Any]],
    log_lines: list[str],
    config: dict[str, Any],
    month_from: str | None,
    month_to: str | None,
) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        yen_fmt = _yen_format(wb)
        neg_yen_fmt = _neg_yen_format(wb)
        header_fmt = _header_format(wb)
        update_fmt = _flag_format(wb, "#E2EFDA")  # 緑系（更新対象）
        skip_fmt = _flag_format(wb, "#FFF2CC")    # 黄系（スキップ）

        # ---- 1. 月別集計 ----
        m = monthly_df.copy()
        if not m.empty:
            m["月"] = m["month"].apply(format_month_jp)
            m = m[["月", "入金合計", "出金合計", "差額"]]
        else:
            m = pd.DataFrame(columns=["月", "入金合計", "出金合計", "差額"])
        m.to_excel(writer, sheet_name="月別集計", index=False)
        ws = writer.sheets["月別集計"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 18, yen_fmt)
        ws.set_column("D:D", 18, neg_yen_fmt)
        _write_headers(ws, list(m.columns), header_fmt)

        # ---- 2. 入金明細 / 3. 出金明細 ----
        _detail_sheet(writer, "入金明細", income_df, yen_fmt, header_fmt)
        _detail_sheet(writer, "出金明細", expense_df, yen_fmt, header_fmt)

        # ---- 4. 除外行一覧 ----
        if not excluded_df.empty:
            ex = excluded_df.copy()
            keep = ["file_name", "excel_row", "month", "amount", "exclude_reason", "deal_name"]
            ex = ex[[c for c in keep if c in ex.columns]]
            ex.columns = ["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"]
            if "月" in ex.columns:
                ex["月"] = ex["月"].apply(lambda v: format_month_jp(v) if v else "")
        else:
            ex = pd.DataFrame(columns=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名"])
        ex.to_excel(writer, sheet_name="除外行一覧", index=False)
        ws = writer.sheets["除外行一覧"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 8)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 18, yen_fmt)
        ws.set_column("E:E", 22)
        ws.set_column("F:F", 36)
        _write_headers(ws, list(ex.columns), header_fmt)

        # ---- 5. ファイル別集計 ----
        fs = file_summary_df.copy() if not file_summary_df.empty else pd.DataFrame(
            columns=["ファイル名", "区分", "合計金額", "集計件数", "除外件数"]
        )
        fs.to_excel(writer, sheet_name="ファイル別集計", index=False)
        ws = writer.sheets["ファイル別集計"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 8)
        ws.set_column("C:C", 18, yen_fmt)
        ws.set_column("D:E", 12)
        _write_headers(ws, list(fs.columns), header_fmt)

        # ---- 6. 上書き前データ ----
        before_rows: list[dict[str, Any]] = []
        for cm in cell_maps:
            before_rows.append(
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "区分": "確定" + cm.kind,
                    "セル": cm.cell_address,
                    "上書き前値": cm.before_value,
                    "元データ": str(cm.before_raw) if cm.before_raw is not None else "",
                }
            )
        before_df = pd.DataFrame(before_rows, columns=["シート", "月", "区分", "セル", "上書き前値", "元データ"])
        before_df.to_excel(writer, sheet_name="上書き前データ", index=False)
        ws = writer.sheets["上書き前データ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 10)
        ws.set_column("E:E", 18, neg_yen_fmt)
        ws.set_column("F:F", 30)
        _write_headers(ws, list(before_df.columns), header_fmt)

        # ---- 7. 上書き後データ ----
        plan_by_key = {(p.sheet, p.month): p for p in plan_rows}
        after_rows: list[dict[str, Any]] = []
        for cm in cell_maps:
            plan = plan_by_key.get((cm.sheet, cm.month))
            if plan and plan.update_flag:
                after_val = plan.income_after if cm.kind == "入金" else plan.expense_after
            else:
                after_val = cm.before_value
            after_rows.append(
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "区分": "確定" + cm.kind,
                    "セル": cm.cell_address,
                    "上書き後値": after_val,
                    "更新対象": "○" if (plan and plan.update_flag) else "—",
                }
            )
        after_df = pd.DataFrame(after_rows, columns=["シート", "月", "区分", "セル", "上書き後値", "更新対象"])
        after_df.to_excel(writer, sheet_name="上書き後データ", index=False)
        ws = writer.sheets["上書き後データ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 10)
        ws.set_column("E:E", 18, neg_yen_fmt)
        ws.set_column("F:F", 12)
        _write_headers(ws, list(after_df.columns), header_fmt)

        # ---- 8. 上書き前後比較 ----
        cmp_records: list[dict[str, Any]] = []
        for plan in plan_rows:
            cmp_records.append(
                {
                    "シート": plan.sheet,
                    "月": format_month_jp(plan.month),
                    "入金セル": plan.income_cell,
                    "上書き前 確定入金 予算": plan.income_before,
                    "上書き後 確定入金 予算": plan.income_after,
                    "入金差額": plan.income_diff,
                    "出金セル": plan.expense_cell,
                    "上書き前 確定支払 予算": plan.expense_before,
                    "上書き後 確定支払 予算": plan.expense_after,
                    "支払差額": plan.expense_diff,
                    "更新対象": "○" if plan.update_flag else "×",
                    "更新理由": plan.reason,
                }
            )
        cmp_df = pd.DataFrame(cmp_records)
        cmp_df.to_excel(writer, sheet_name="上書き前後比較", index=False)
        ws = writer.sheets["上書き前後比較"]
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:F", 18, neg_yen_fmt)
        ws.set_column("G:G", 10)
        ws.set_column("H:J", 18, neg_yen_fmt)
        ws.set_column("K:K", 10)
        ws.set_column("L:L", 16)
        _write_headers(ws, list(cmp_df.columns), header_fmt)
        # 更新対象行に色を付ける
        for i, plan in enumerate(plan_rows):
            row_idx = i + 1  # 1始まり（0はヘッダ）
            fmt = update_fmt if plan.update_flag else skip_fmt
            ws.write(row_idx, 10, "○" if plan.update_flag else "×", fmt)

        # ---- 9. 更新対象一覧 ----
        target_records = [
            {
                "シート": p.sheet,
                "月": format_month_jp(p.month),
                "入金セル": p.income_cell,
                "上書き後 確定入金 予算": p.income_after,
                "出金セル": p.expense_cell,
                "上書き後 確定支払 予算": p.expense_after,
            }
            for p in plan_rows
            if p.update_flag
        ]
        if not target_records:
            target_records = [{"シート": "(更新対象なし)", "月": "", "入金セル": "", "上書き後 確定入金 予算": None, "出金セル": "", "上書き後 確定支払 予算": None}]
        target_df = pd.DataFrame(target_records)
        target_df.to_excel(writer, sheet_name="更新対象一覧", index=False)
        ws = writer.sheets["更新対象一覧"]
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 18, neg_yen_fmt)
        ws.set_column("E:E", 10)
        ws.set_column("F:F", 18, neg_yen_fmt)
        _write_headers(ws, list(target_df.columns), header_fmt)

        # ---- 10. 未照合月一覧 ----
        unmatched_rows: list[dict[str, str]] = []
        for m_ in summary_only_months:
            unmatched_rows.append({"区分": "集計のみ（予算ファイルに該当月なし）", "月": format_month_jp(m_), "YYYY-MM": m_})
        for m_ in budget_only_months:
            unmatched_rows.append({"区分": "予算のみ（集計データなし）", "月": format_month_jp(m_), "YYYY-MM": m_})
        if not unmatched_rows:
            unmatched_rows = [{"区分": "(未照合なし)", "月": "", "YYYY-MM": ""}]
        unmatched_df = pd.DataFrame(unmatched_rows)
        unmatched_df.to_excel(writer, sheet_name="未照合月一覧", index=False)
        ws = writer.sheets["未照合月一覧"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 14)
        ws.set_column("C:C", 12)
        _write_headers(ws, list(unmatched_df.columns), header_fmt)

        # ---- 11. 処理ログ ----
        cfg_lines = [
            f"出力日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"更新期間: {month_from or '(未指定)'} ～ {month_to or '(未指定)'}",
            f"月列キーワード: {', '.join(config.get('month_column_keywords', []))}",
            f"金額列キーワード: {', '.join(config.get('amount_column_keywords', []))}",
            f"商談列キーワード: {', '.join(config.get('deal_column_keywords', []))}",
            f"除外キーワード: {', '.join(config.get('exclude_keywords', []))}",
            f"デフォルトヘッダー行: {config.get('default_header_row')}",
            f"対象シート: {', '.join(config.get('budget_target_sheets', []))}",
        ]
        write_lines = [
            f"{w['sheet']}!{w['cell']} [{w['kind']}] "
            f"{format_month_jp(w['month'])}: {w['before']} → {w['after']}"
            for w in written_log
        ]
        all_lines = cfg_lines + [""] + ["== 処理ログ =="] + log_lines + [""] + ["== 書き込み詳細 =="] + write_lines
        log_df = pd.DataFrame({"ログ": all_lines})
        log_df.to_excel(writer, sheet_name="処理ログ", index=False)
        ws = writer.sheets["処理ログ"]
        ws.set_column("A:A", 100)
        _write_headers(ws, ["ログ"], header_fmt)

    bio.seek(0)
    return bio.read()
