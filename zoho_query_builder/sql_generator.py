"""UNION ALL SQL ジェネレータ。

各 transaction ペアに対して 1 SELECT を生成し、UNION ALL で連結する。
除外条件: 日付 NULL / 金額 NULL / 金額 0 / 商談名 NULL。

Zoho Analytics の SQL は ANSI 標準ベース。識別子はダブルクォートで囲む。
"""
from __future__ import annotations

from detector import DetectionResult, TransactionPair


def _q(name: str) -> str:
    """Zoho Analytics 用の識別子クォート。"""
    return '"' + name.replace('"', '""') + '"'


def build_transaction_sql(
    source_table: str,
    detection: DetectionResult,
    transaction_status: str,
    stage_column: str | None = None,
    stage_values: list[str] | None = None,
    qualify_column: str | None = None,
    qualify_date_from: str | None = None,
    qualify_date_to: str | None = None,
    extra_where: str | None = None,
    include_deal_id: bool = False,
    deal_id_column: str = "Id",
    resolve_client_name: bool = False,
    client_lookup_table: str = "取引先",
    client_lookup_id_col: str = "Id",
    client_lookup_name_col: str = "取引先名",
    include_case_number: bool = False,
    case_lookup_table: str = "deal_case_numbers",
    case_lookup_id_col: str = "deal_id",
    case_lookup_name_col: str = "案件番号",
    resolve_payee: bool = False,
    payee_lookup_table: str = "商談",
    payee_lookup_id_col: str = "Id",
    deal_join_id_col: str = "Id",
) -> str:
    """確定/予測 共通の UNION ALL SQL を生成。

    Parameters
    ----------
    source_table
        元テーブル名（Zoho Analytics 上のテーブル名）
    detection
        detector.detect_columns() の戻り値
    transaction_status
        "confirmed" または "forecast"
    stage_column
        フィルタに使う列名（例: "ステージ"）。None なら全行対象。
    stage_values
        含めるステージ値のリスト。stage_column 必須。
    extra_where
        追加 WHERE 条件文字列（自由形式、必要なら）
    """
    if detection.deal_column is None:
        raise ValueError("商談名 列を特定できませんでした")
    if not detection.pairs:
        raise ValueError("transaction ペア（日付+金額）を1つも検出できませんでした")

    src_alias = "d"     # 商談
    lkp_alias = "a"     # 取引先 (account)
    src = f"{_q(source_table)} {src_alias}"

    def _src_col(col: str) -> str:
        return f"{src_alias}.{_q(col)}"

    deal = _src_col(detection.deal_column)

    # クライアント名: JOIN で取引先テーブルから取得
    if resolve_client_name:
        client = f"{lkp_alias}.{_q(client_lookup_name_col)}"
        join_clause = (
            f"\nLEFT JOIN {_q(client_lookup_table)} {lkp_alias} "
            f"ON {src_alias}.{_q(detection.client_column or 'Id')} "
            f"= {lkp_alias}.{_q(client_lookup_id_col)}"
        )
    else:
        client = _src_col(detection.client_column) if detection.client_column else "NULL"
        join_clause = ""

    # 案件番号
    # - include_case_number=True なら deal_case_numbers (CRM の案件番号) を JOIN
    # - そうでなければ include_deal_id=True なら 商談.Id を 案件番号 として出す
    case_alias = "n"   # case-number lookup
    if include_case_number:
        deal_id_select = f"    {case_alias}.{_q(case_lookup_name_col)} AS \"案件番号\",\n"
        case_join = (
            f"\nLEFT JOIN {_q(case_lookup_table)} {case_alias} "
            f"ON {src_alias}.{_q(deal_id_column)} "
            f"= {case_alias}.{_q(case_lookup_id_col)}"
        )
    elif include_deal_id:
        deal_id_select = f"    {_src_col(deal_id_column)} AS \"案件番号\",\n"
        case_join = ""
    else:
        deal_id_select = ""
        case_join = ""

    # 支払先 を 商談 raw テーブルから取りに行く JOIN
    # （確定/予測 テーブルに 支払先 列を sync させていない場合のフォールバック）
    payee_alias = "m"
    payee_join = ""
    if resolve_payee:
        payee_join = (
            f"\nLEFT JOIN {_q(payee_lookup_table)} {payee_alias} "
            f"ON {src_alias}.{_q(deal_join_id_col)} "
            f"= {payee_alias}.{_q(payee_lookup_id_col)}"
        )

    # ステージフィルタ
    stage_filter = ""
    if stage_column and stage_values:
        quoted_vals = ",".join("'" + v.replace("'", "''") + "'" for v in stage_values)
        stage_filter = f"\n  AND {_src_col(stage_column)} IN ({quoted_vals})"

    # 商談単位の期間フィルタ（例: 国内仕入１　支払日 BETWEEN '2025-06-01' AND '2026-12-31'）
    qualify_filter = ""
    if qualify_column and (qualify_date_from or qualify_date_to):
        col_q = _src_col(qualify_column)
        if qualify_date_from and qualify_date_to:
            qualify_filter = (
                f"\n  AND {col_q} BETWEEN DATE '{qualify_date_from}'"
                f" AND DATE '{qualify_date_to}'"
            )
        elif qualify_date_from:
            qualify_filter = f"\n  AND {col_q} >= DATE '{qualify_date_from}'"
        else:
            qualify_filter = f"\n  AND {col_q} <= DATE '{qualify_date_to}'"

    extra = f"\n  AND {extra_where}" if extra_where else ""

    blocks: list[str] = []
    for pair in detection.pairs:
        date_q = _src_col(pair.date_column)
        amount_q = _src_col(pair.amount_column)
        # 支払先名: payment 行のみ採用
        # 優先順位: (1) resolve_payee=True なら 商談 raw を JOIN し、
        #             date_column の "支払日" を "支払先" に差し替えた列を引く
        #          (2) detector が pair.payee_column を検出していればそれ
        #          (3) どちらもなければ NULL
        if pair.transaction_type == "payment":
            if resolve_payee:
                payee_col_name = pair.date_column.replace("支払日", "支払先")
                payee_q = f"{payee_alias}.{_q(payee_col_name)}"
            elif pair.payee_column:
                payee_q = _src_col(pair.payee_column)
            else:
                payee_q = "NULL"
        else:
            payee_q = "NULL"
        # payment_round は元の項目名（日付列名）を入れる
        # SQL 文字列リテラル内のシングルクォートをエスケープ
        round_label = pair.date_column.replace("'", "''")
        block = f"""SELECT
    {deal}         AS "商談名",
{deal_id_select}    {client}       AS "クライアント名",
    {payee_q}      AS "支払先名",
    {date_q}       AS "取引日",
    {amount_q}     AS "金額",
    '{transaction_status}'   AS "transaction_status",
    '{pair.transaction_type}' AS "transaction_type",
    '{round_label}' AS "payment_round"
FROM {src}{join_clause}{case_join}{payee_join}
WHERE {deal}     IS NOT NULL
  AND {date_q}   IS NOT NULL
  AND {amount_q} IS NOT NULL
  AND {amount_q} <> 0{stage_filter}{qualify_filter}{extra}"""
        blocks.append(block)

    return "\n\nUNION ALL\n\n".join(blocks)
