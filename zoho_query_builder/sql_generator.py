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

    deal = _q(detection.deal_column)
    client = _q(detection.client_column) if detection.client_column else "CAST(NULL AS VARCHAR)"
    src = _q(source_table)

    # ステージフィルタ
    stage_filter = ""
    if stage_column and stage_values:
        quoted_vals = ",".join("'" + v.replace("'", "''") + "'" for v in stage_values)
        stage_filter = f"\n  AND {_q(stage_column)} IN ({quoted_vals})"

    # 商談単位の期間フィルタ（例: 国内仕入１　支払日 BETWEEN '2025-06-01' AND '2026-12-31'）
    qualify_filter = ""
    if qualify_column and (qualify_date_from or qualify_date_to):
        if qualify_date_from and qualify_date_to:
            qualify_filter = (
                f"\n  AND {_q(qualify_column)} BETWEEN DATE '{qualify_date_from}'"
                f" AND DATE '{qualify_date_to}'"
            )
        elif qualify_date_from:
            qualify_filter = f"\n  AND {_q(qualify_column)} >= DATE '{qualify_date_from}'"
        else:
            qualify_filter = f"\n  AND {_q(qualify_column)} <= DATE '{qualify_date_to}'"

    extra = f"\n  AND {extra_where}" if extra_where else ""

    blocks: list[str] = []
    for pair in detection.pairs:
        date_q = _q(pair.date_column)
        amount_q = _q(pair.amount_column)
        # payment_round は元の項目名（日付列名）を入れる
        # SQL 文字列リテラル内のシングルクォートをエスケープ
        round_label = pair.date_column.replace("'", "''")
        block = f"""SELECT
    {deal}         AS "商談名",
    {client}       AS "クライアント名",
    {date_q}       AS "取引日",
    {amount_q}     AS "金額",
    '{transaction_status}'   AS "transaction_status",
    '{pair.transaction_type}' AS "transaction_type",
    '{round_label}' AS "payment_round"
FROM {src}
WHERE {deal}     IS NOT NULL
  AND {date_q}   IS NOT NULL
  AND {amount_q} IS NOT NULL
  AND {amount_q} <> 0{stage_filter}{qualify_filter}{extra}"""
        blocks.append(block)

    return "\n\nUNION ALL\n\n".join(blocks)
