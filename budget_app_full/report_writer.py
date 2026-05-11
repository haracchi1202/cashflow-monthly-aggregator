"""比較レポート Excel の生成。

14 シート構成:
    1.  月別集計（確定/予測/合算）
    2.  確定入金明細
    3.  確定支払明細
    4.  予測入金明細
    5.  予測支払明細
    6.  確定＋予測月別集計（合算のみ）
    7.  除外行一覧
    8.  ファイル別集計
    9.  上書き前データ
    10. 上書き後データ
    11. 上書き前後比較
    12. 更新対象一覧
    13. 未照合月一覧
    14. 処理ログ
"""
from __future__ import annotations

import io
from datetime import datetime
from typing import Any

import pandas as pd

from excel_reader import format_month_jp
from aggregator import update_scope_label
from budget_updater import BudgetCellMap, CellPlan


YEN_FMT = '#,##0"円"'
NEG_YEN_FMT = '#,##0"円";[Red]-#,##0"円"'


def _yen_format(wb):
    return wb.add_format({"num_format": YEN_FMT})


def _neg_yen_format(wb):
    return wb.add_format({"num_format": NEG_YEN_FMT})


def _header_format(wb):
    return wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})


def _flag_format(wb, color: str):
    return wb.add_format({"bg_color": color, "border": 1})


def _write_headers(ws, columns: list[str], fmt) -> None:
    for col_num, val in enumerate(columns):
        ws.write(0, col_num, val, fmt)


def _detail_sheet(writer, sheet_name: str, df: pd.DataFrame, yen_fmt, header_fmt) -> None:
    columns_template = [
        "source_file", "raw_row_index", "target_month",
        "client_name", "deal_name", "amount", "transaction_date",
    ]
    out_columns_jp = ["ファイル名", "行番号", "月", "顧客名", "商談名", "金額", "元月セル"]
    if df.empty:
        out = pd.DataFrame(columns=out_columns_jp)
    else:
        out = df.copy()
        out = out[[c for c in columns_template if c in out.columns]]
        out.columns = out_columns_jp[: len(out.columns)]
        if "月" in out.columns:
            out["月"] = out["月"].apply(lambda v: format_month_jp(v) if v else "")
    out.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.set_column("A:A", 36)
    ws.set_column("B:B", 8)
    ws.set_column("C:C", 12)
    ws.set_column("D:D", 24)
    ws.set_column("E:E", 36)
    ws.set_column("F:F", 18, yen_fmt)
    ws.set_column("G:G", 20)
    _write_headers(ws, list(out.columns), header_fmt)


def build_report(
    monthly_df: pd.DataFrame,
    by_source_dfs: dict[str, pd.DataFrame],
    excluded_df: pd.DataFrame,
    file_summary_df: pd.DataFrame,
    cell_maps: list[BudgetCellMap],
    plans: list[CellPlan],
    summary_only_months: list[str],
    budget_only_months: list[str],
    written_log: list[dict[str, Any]],
    log_lines: list[str],
    config: dict[str, Any],
    month_from: str | None,
    month_to: str | None,
    update_scope: str,
) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        wb = writer.book
        yen_fmt = _yen_format(wb)
        neg_yen_fmt = _neg_yen_format(wb)
        header_fmt = _header_format(wb)
        update_fmt = _flag_format(wb, "#E2EFDA")
        skip_fmt = _flag_format(wb, "#FFF2CC")

        # ---- 1. 月別集計（全列まとめ） ----
        m = monthly_df.copy()
        if not m.empty:
            m.insert(0, "月", m["target_month"].apply(format_month_jp))
            m = m.drop(columns=["target_month"])
        else:
            m = pd.DataFrame(
                columns=[
                    "月", "確定入金", "確定支払", "確定差額",
                    "予測入金", "予測支払", "予測差額",
                    "合算入金", "合算支払", "合算差額",
                ]
            )
        m.to_excel(writer, sheet_name="月別集計", index=False)
        ws = writer.sheets["月別集計"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 16, yen_fmt)
        ws.set_column("D:D", 16, neg_yen_fmt)
        ws.set_column("E:F", 16, yen_fmt)
        ws.set_column("G:G", 16, neg_yen_fmt)
        ws.set_column("H:I", 16, yen_fmt)
        ws.set_column("J:J", 16, neg_yen_fmt)
        _write_headers(ws, list(m.columns), header_fmt)

        # ---- 1b. 月別集計_確定のみ ----
        if not monthly_df.empty:
            conf = monthly_df[["target_month", "確定入金", "確定支払", "確定差額"]].copy()
            conf.insert(0, "月", conf["target_month"].apply(format_month_jp))
            conf = conf.drop(columns=["target_month"])
        else:
            conf = pd.DataFrame(columns=["月", "確定入金", "確定支払", "確定差額"])
        conf.to_excel(writer, sheet_name="月別集計_確定のみ", index=False)
        ws = writer.sheets["月別集計_確定のみ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 18, yen_fmt)
        ws.set_column("D:D", 18, neg_yen_fmt)
        _write_headers(ws, list(conf.columns), header_fmt)

        # ---- 1c. 月別集計_予測のみ ----
        if not monthly_df.empty:
            fc = monthly_df[["target_month", "予測入金", "予測支払", "予測差額"]].copy()
            fc.insert(0, "月", fc["target_month"].apply(format_month_jp))
            fc = fc.drop(columns=["target_month"])
        else:
            fc = pd.DataFrame(columns=["月", "予測入金", "予測支払", "予測差額"])
        fc.to_excel(writer, sheet_name="月別集計_予測のみ", index=False)
        ws = writer.sheets["月別集計_予測のみ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 18, yen_fmt)
        ws.set_column("D:D", 18, neg_yen_fmt)
        _write_headers(ws, list(fc.columns), header_fmt)

        # ---- 2-5. ソースグループ別明細 ----
        for grp in ["確定入金", "確定支払", "予測入金", "予測支払"]:
            sub_df = by_source_dfs.get(grp, pd.DataFrame())
            sheet_name = grp + "明細"
            _detail_sheet(writer, sheet_name, sub_df, yen_fmt, header_fmt)

        # ---- 6. 確定＋予測 月別集計（合算のみ抜粋） ----
        if not monthly_df.empty:
            combined = monthly_df[["target_month", "合算入金", "合算支払", "合算差額"]].copy()
            combined.insert(0, "月", combined["target_month"].apply(format_month_jp))
            combined = combined.drop(columns=["target_month"]).rename(
                columns={"合算入金": "確定＋予測 入金合計", "合算支払": "確定＋予測 支払合計", "合算差額": "確定＋予測 差額"}
            )
        else:
            combined = pd.DataFrame(
                columns=["月", "確定＋予測 入金合計", "確定＋予測 支払合計", "確定＋予測 差額"]
            )
        combined.to_excel(writer, sheet_name="確定＋予測月別集計", index=False)
        ws = writer.sheets["確定＋予測月別集計"]
        ws.set_column("A:A", 14)
        ws.set_column("B:C", 20, yen_fmt)
        ws.set_column("D:D", 20, neg_yen_fmt)
        _write_headers(ws, list(combined.columns), header_fmt)

        # ---- 7. 除外行一覧 ----
        if not excluded_df.empty:
            ex = excluded_df.copy()
            keep = [
                "source_file", "raw_row_index", "target_month",
                "amount", "exclude_reason", "deal_name", "source_group",
            ]
            ex = ex[[c for c in keep if c in ex.columns]]
            ex.columns = ["ファイル名", "行番号", "月", "金額", "除外理由", "商談名", "区分"]
            if "月" in ex.columns:
                ex["月"] = ex["月"].apply(lambda v: format_month_jp(v) if v else "")
        else:
            ex = pd.DataFrame(
                columns=["ファイル名", "行番号", "月", "金額", "除外理由", "商談名", "区分"]
            )
        ex.to_excel(writer, sheet_name="除外行一覧", index=False)
        ws = writer.sheets["除外行一覧"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 8)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 16, yen_fmt)
        ws.set_column("E:E", 22)
        ws.set_column("F:F", 36)
        ws.set_column("G:G", 12)
        _write_headers(ws, list(ex.columns), header_fmt)

        # ---- 8. ファイル別集計 ----
        fs = (
            file_summary_df.copy()
            if not file_summary_df.empty
            else pd.DataFrame(
                columns=["ファイル名", "区分", "確定/予測", "合計金額", "集計件数", "除外件数"]
            )
        )
        fs.to_excel(writer, sheet_name="ファイル別集計", index=False)
        ws = writer.sheets["ファイル別集計"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 10)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 18, yen_fmt)
        ws.set_column("E:F", 12)
        _write_headers(ws, list(fs.columns), header_fmt)

        # ---- 9. 上書き前データ ----
        before_rows: list[dict[str, Any]] = []
        for cm in cell_maps:
            status_jp = "確定" if cm.transaction_status == "confirmed" else "予測"
            before_rows.append(
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "ステータス": status_jp,
                    "区分": cm.kind,
                    "セル": cm.cell_address,
                    "上書き前値": cm.before_value,
                    "元データ": str(cm.before_raw) if cm.before_raw is not None else "",
                }
            )
        before_df = pd.DataFrame(
            before_rows,
            columns=["シート", "月", "ステータス", "区分", "セル", "上書き前値", "元データ"],
        )
        before_df.to_excel(writer, sheet_name="上書き前データ", index=False)
        ws = writer.sheets["上書き前データ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 8)
        ws.set_column("E:E", 10)
        ws.set_column("F:F", 18, neg_yen_fmt)
        ws.set_column("G:G", 30)
        _write_headers(ws, list(before_df.columns), header_fmt)

        # ---- 10. 上書き後データ ----
        plan_by_key = {(p.sheet, p.month, p.transaction_status, p.kind): p for p in plans}
        after_rows: list[dict[str, Any]] = []
        for cm in cell_maps:
            plan = plan_by_key.get((cm.sheet, cm.month, cm.transaction_status, cm.kind))
            if plan and plan.update_flag:
                after_val = plan.after
            else:
                after_val = cm.before_value
            status_jp = "確定" if cm.transaction_status == "confirmed" else "予測"
            after_rows.append(
                {
                    "シート": cm.sheet,
                    "月": format_month_jp(cm.month),
                    "ステータス": status_jp,
                    "区分": cm.kind,
                    "セル": cm.cell_address,
                    "上書き後値": after_val,
                    "更新対象": "○" if (plan and plan.update_flag) else "—",
                }
            )
        after_df = pd.DataFrame(
            after_rows,
            columns=["シート", "月", "ステータス", "区分", "セル", "上書き後値", "更新対象"],
        )
        after_df.to_excel(writer, sheet_name="上書き後データ", index=False)
        ws = writer.sheets["上書き後データ"]
        ws.set_column("A:A", 14)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 8)
        ws.set_column("E:E", 10)
        ws.set_column("F:F", 18, neg_yen_fmt)
        ws.set_column("G:G", 12)
        _write_headers(ws, list(after_df.columns), header_fmt)

        # ---- 11. 上書き前後比較 ----
        cmp_records: list[dict[str, Any]] = []
        for p in plans:
            status_jp = "確定" if p.transaction_status == "confirmed" else "予測"
            cmp_records.append(
                {
                    "シート": p.sheet,
                    "月": format_month_jp(p.month),
                    "ステータス": status_jp,
                    "区分": p.kind,
                    "セル": p.cell_address,
                    "上書き前": p.before,
                    "上書き後": p.after,
                    "差額": p.diff,
                    "更新対象": "○" if p.update_flag else "×",
                    "更新理由": p.reason,
                }
            )
        cmp_df = pd.DataFrame(cmp_records)
        cmp_df.to_excel(writer, sheet_name="上書き前後比較", index=False)
        ws = writer.sheets["上書き前後比較"]
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 8)
        ws.set_column("E:E", 10)
        ws.set_column("F:H", 18, neg_yen_fmt)
        ws.set_column("I:I", 10)
        ws.set_column("J:J", 16)
        _write_headers(ws, list(cmp_df.columns), header_fmt)
        for i, p in enumerate(plans):
            row_idx = i + 1
            fmt = update_fmt if p.update_flag else skip_fmt
            ws.write(row_idx, 8, "○" if p.update_flag else "×", fmt)

        # ---- 12. 更新対象一覧 ----
        target_records = [
            {
                "シート": p.sheet,
                "月": format_month_jp(p.month),
                "ステータス": "確定" if p.transaction_status == "confirmed" else "予測",
                "区分": p.kind,
                "セル": p.cell_address,
                "上書き後": p.after,
            }
            for p in plans
            if p.update_flag
        ]
        if not target_records:
            target_records = [
                {
                    "シート": "(更新対象なし)",
                    "月": "",
                    "ステータス": "",
                    "区分": "",
                    "セル": "",
                    "上書き後": None,
                }
            ]
        target_df = pd.DataFrame(target_records)
        target_df.to_excel(writer, sheet_name="更新対象一覧", index=False)
        ws = writer.sheets["更新対象一覧"]
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 8)
        ws.set_column("E:E", 10)
        ws.set_column("F:F", 18, neg_yen_fmt)
        _write_headers(ws, list(target_df.columns), header_fmt)

        # ---- 13. 未照合月一覧 ----
        unmatched_rows: list[dict[str, str]] = []
        for m_ in summary_only_months:
            unmatched_rows.append(
                {
                    "区分": "集計のみ（予算ファイルに該当月なし）",
                    "月": format_month_jp(m_),
                    "YYYY-MM": m_,
                }
            )
        for m_ in budget_only_months:
            unmatched_rows.append(
                {
                    "区分": "予算のみ（集計データなし）",
                    "月": format_month_jp(m_),
                    "YYYY-MM": m_,
                }
            )
        if not unmatched_rows:
            unmatched_rows = [{"区分": "(未照合なし)", "月": "", "YYYY-MM": ""}]
        unmatched_df = pd.DataFrame(unmatched_rows)
        unmatched_df.to_excel(writer, sheet_name="未照合月一覧", index=False)
        ws = writer.sheets["未照合月一覧"]
        ws.set_column("A:A", 36)
        ws.set_column("B:B", 14)
        ws.set_column("C:C", 12)
        _write_headers(ws, list(unmatched_df.columns), header_fmt)

        # ---- 14. 処理ログ ----
        cfg_lines = [
            f"出力日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"更新スコープ: {update_scope_label(update_scope)} ({update_scope})",
            f"更新期間: {month_from or '(未指定)'} ～ {month_to or '(未指定)'}",
            f"月列キーワード: {', '.join(config.get('month_column_keywords', []))}",
            f"金額列キーワード: {', '.join(config.get('amount_column_keywords', []))}",
            f"商談列キーワード: {', '.join(config.get('deal_column_keywords', []))}",
            f"除外キーワード: {', '.join(config.get('exclude_keywords', []))}",
            f"デフォルトヘッダー行: {config.get('default_header_row')}",
            f"対象シート: {', '.join(config.get('budget_target_sheets', []))}",
        ]
        write_lines = [
            f"{w['sheet']}!{w['cell']} "
            f"[{'確定' if w['transaction_status'] == 'confirmed' else '予測'}{w['kind']}] "
            f"{format_month_jp(w['month'])}: {w['before']} → {w['after']}"
            for w in written_log
        ]
        all_lines = (
            cfg_lines
            + [""]
            + ["== 処理ログ =="]
            + log_lines
            + [""]
            + ["== 書き込み詳細 =="]
            + write_lines
        )
        log_df = pd.DataFrame({"ログ": all_lines})
        log_df.to_excel(writer, sheet_name="処理ログ", index=False)
        ws = writer.sheets["処理ログ"]
        ws.set_column("A:A", 100)
        _write_headers(ws, ["ログ"], header_fmt)

    bio.seek(0)
    return bio.read()
