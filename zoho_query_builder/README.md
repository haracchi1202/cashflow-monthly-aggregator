# Zoho Analytics: 横持ち → 縦持ち Query Table ビルダー

Zoho Analytics の **「確定」「予測」レポート**（1 商談 1 行の横持ち構造）から、
transaction 形式（1 transaction 1 行）の Query Table を自動生成するツール。

```
[元テーブル / 横持ち]                     [生成される Query Table / 縦持ち]
商談名 / 入金日１ / 入金額１ / 入金日２ /    商談名  / クライアント / 取引日 / 金額 / status / type / round
       入金額２ / 仕入１支払日 / 仕入１原価
       / ...                                A案件 / AAA / 2025-09-01 / 1,000,000 / confirmed / income / income_1
                                            A案件 / AAA / 2025-10-01 /   500,000 / confirmed / payment / payment_1
                                            ...
```

---

## ファイル構成

| ファイル | 役割 |
|---|---|
| `zoho_client.py` | OAuth refresh + Zoho Analytics API v2 ラッパー |
| `detector.py` | 列名パターンマッチ（全角/半角・税込表記の揺れ対応） |
| `sql_generator.py` | UNION ALL SQL ビルダー（NULL/0 除外フィルタつき） |
| `main.py` | CLI エントリ（list / inspect / build / workspaces） |
| `.env.example` | 認証情報テンプレート |
| `requirements.txt` | 依存ライブラリ（requests のみ） |

---

## セットアップ

### 1. 依存をインストール

```powershell
cd CFCREATOR\zoho_query_builder
pip install -r requirements.txt
```

### 2. Zoho 認証情報を取得して `.env` に書く

`.env.example` をコピーして `.env` を作成し、以下の 6 項目を入れる:

| 変数 | 取得元 |
|---|---|
| `ZOHO_REGION` | `zoho.jp` / `zoho.com` / `zoho.eu` |
| `ZOHO_CLIENT_ID` | https://api-console.zoho.jp で Self Client 作成 → Client Details |
| `ZOHO_CLIENT_SECRET` | 同上 |
| `ZOHO_REFRESH_TOKEN` | Self Client → Generate Code → curl で交換（下記参照） |
| `ZOHO_ORG_ID` | Zoho Analytics → Settings → Account Info、または `python main.py workspaces` |
| `ZOHO_WORKSPACE_ID` | Zoho Analytics のワークスペース URL `/workspace/<ID>/` |

#### Refresh Token の取得手順

1. https://api-console.zoho.jp にログイン
2. **ADD CLIENT → Self Client → Create**
3. **Generate Code** タブ:
   - Scope: `ZohoAnalytics.data.all,ZohoAnalytics.metadata.read,ZohoAnalytics.share.read,ZohoAnalytics.modeling.all`
   - Time Duration: 10 minutes
   - Scope Description: 任意
   - **CREATE** → 表示された Grant Code をコピー（10 分以内）
4. PowerShell で交換:

```powershell
curl.exe -X POST "https://accounts.zoho.jp/oauth/v2/token?code=<GRANT_CODE>&client_id=<CLIENT_ID>&client_secret=<CLIENT_SECRET>&grant_type=authorization_code"
```

5. レスポンスの `refresh_token` を `.env` に書く。

---

## 使い方

### A. ワークスペース内の view 一覧を見る

```powershell
python main.py list
python main.py list --filter 確定
```

出力例:
```
=== Workspace 内 view 一覧（42 件）===
  [TABLE       ]      123456789  確定
  [TABLE       ]      234567890  予測
  [REPORT      ]      345678901  確定レポート
  ...
```

### B. 列構造と自動判定を確認

```powershell
python main.py inspect --source-table 確定
```

出力例:
```
=== 確定 の列一覧（27 列）===
  - 商談名
  - クライアント名
  - クライアント入金日１
  - クライアント入金額１(税込)
  ...

=== 自動判定結果 ===
  商談名 列      : 商談名
  クライアント名 : クライアント名
  入金 ペア      : 2 件
  支払 ペア      : 5 件

  - income_1:  日付=クライアント入金日１ / 金額=クライアント入金額１(税込)
  - income_2:  日付=クライアント入金日２ / 金額=クライアント入金額２(税込)
  - payment_1: 日付=国内仕入１ 支払日   / 金額=国内仕入１ 原価総額(税込)
  - payment_2: ...
```

### C. SQL のみ確認（dry-run）

```powershell
python main.py build --source-table 確定 --output-name confirmed_transactions --status confirmed --dry-run
```

### D. Query Table を作成

```powershell
python main.py build --source-table 確定 --output-name confirmed_transactions --status confirmed
python main.py build --source-table 予測 --output-name forecast_transactions --status forecast
```

---

## 生成される SQL の形式

```sql
SELECT
    "商談名"                    AS "商談名",
    "クライアント名"            AS "クライアント名",
    "クライアント入金日１"      AS "取引日",
    "クライアント入金額１(税込)" AS "金額",
    'confirmed'                 AS "transaction_status",
    'income'                    AS "transaction_type",
    'income_1'                  AS "payment_round"
FROM "確定"
WHERE "商談名"                IS NOT NULL
  AND "クライアント入金日１"   IS NOT NULL
  AND "クライアント入金額１(税込)" IS NOT NULL
  AND "クライアント入金額１(税込)" <> 0

UNION ALL

SELECT ... -- income_2

UNION ALL

SELECT ... -- payment_1

... (合計 入金2 + 支払5 = 7 UNION)
```

---

## 生成 Query Table の構造

| 列 | 型 | 値 |
|---|---|---|
| 商談名 | varchar | 元テーブルから引用 |
| クライアント名 | varchar | 元テーブルから引用 |
| 支払先名 | varchar | payment 行のみ。元テーブルの `国内仕入N　支払先` / `予測{初回\|残金\|予備}支払先` を引用。income 行は NULL |
| 取引日 | date | 各日付列を縦持ち化 |
| 金額 | number | 税込金額 |
| transaction_status | varchar | `confirmed` / `forecast` |
| transaction_type | varchar | `income` / `payment` |
| payment_round | varchar | `income_1`, `income_2`, `payment_1` 〜 `payment_5` |

---

## 除外条件

以下の行は SELECT の WHERE で除外されます（transaction 化しない）:

- `商談名` が NULL
- 日付が NULL
- 金額が NULL
- 金額が 0

---

## 支払先名（`--resolve-payee` オプション）

確定/予測テーブル本体に「支払先」列が同期されていない場合、`--resolve-payee` を付けると `商談` raw テーブルを LEFT JOIN し、各 payment 行に対応する `国内仕入N　支払先` 列を `支払先名` として出力します。

```powershell
python main.py build --source-table 商談 --output-name confirmed_transactions ^
                    --status confirmed --resolve-payee
```

### 前提条件（重要）

`--resolve-payee` を有効にするには、**JOIN 先テーブル（既定: `商談`）の支払先 列名が、payment の date_column と同じ表記**である必要があります。具体的には：

- date_column が `国内仕入１　支払日`（全角１＋全角空白）なら、JOIN 先にも `国内仕入１　支払先`（同じ全角１＋同じ全角空白）が存在する必要がある
- 内部実装は `date_column.replace("支払日", "支払先")` で列名を導出する単純なロジック

この前提が崩れると **「列が存在しない」エラーで Query Table 作成自体が失敗** します。事前に以下で揃いを確認してください：

```powershell
python main.py inspect --source-table 商談 | findstr 支払
```

JOIN 先が `商談` 以外（独自 lookup テーブル等）の場合は `--payee-lookup-table <name>` で指定できますが、列名規約（`〜支払日` ↔ `〜支払先`）は同様に守る必要があります。

### `--resolve-payee` を使わない場合

確定/予測テーブル本体に直接 `〜支払先` 列が同期されていれば、`detector.py` が自動で `pair.payee_column` として拾い、JOIN なしで `支払先名` を出します。

---

## CSV エクスポート

作成された Query Table は Zoho Analytics の画面で:
1. 左メニューから `confirmed_transactions` / `forecast_transactions` を開く
2. **Export** → **CSV** で書き出し可能

API 経由でエクスポートしたい場合は別途エクスポート用エンドポイントを使います（このツールでは未実装）。

---

## トラブルシューティング

| エラー | 原因 | 対処 |
|---|---|---|
| `OAuth refresh failed` | Client ID/Secret 違い、refresh_token 期限切れ | Self Client を再作成 |
| `view が見つかりません` | テーブル名のタイポ、または REPORT のみ存在し TABLE がない | `python main.py list` で正確な名前を確認 |
| `transaction ペアが検出できません` | 列名の命名規則が想定外 | `detector.py` の正規表現を編集、または `inspect` で列名を確認して報告 |
| `Zoho API ... 400 ...` | SQL が不正、または既存の同名 Query Table がある | dry-run の SQL を Zoho Analytics の SQL エディタで手動実行してエラー詳細を確認 |

---

## ファイル別 SQL を Zoho Analytics に手動貼り付けたい場合

`--dry-run` で出力された SQL をコピー → Zoho Analytics の **Create > New Query Table** に貼り付けて保存も可能。
API 作成が落ちる時のフォールバック手段として有効。
