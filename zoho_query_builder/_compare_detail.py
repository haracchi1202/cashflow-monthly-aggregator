"""2026-04 以降に絞って、月別 × 商談単位で旧/新の金額差を出す。"""
import pandas as pd
import re

def normalize_month(v):
    if pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp,)):
        return f'{v.year:04d}-{v.month:02d}'
    s = str(v).strip()
    m = re.match(r'^\s*(20\d{2})[\s/\-.年]\s*(\d{1,2})', s)
    if m:
        return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}'
    m = re.search(r'(\d{1,2})\s*月.*?(20\d{2})', s)
    if m:
        return f'{int(m.group(2)):04d}-{int(m.group(1)):02d}'
    try:
        ts = pd.to_datetime(s, errors='raise')
        return f'{ts.year:04d}-{ts.month:02d}'
    except Exception:
        return None


def read_old(fp, kind_prefix):
    df = pd.read_excel(fp, header=6, engine='openpyxl', dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    date_col = amt_col = None
    for c in df.columns:
        if '入金日' in c or '支払日' in c:
            date_col = c; break
    for c in df.columns:
        if ('入金額' in c or '原価総額' in c) and '合計' in c:
            amt_col = c; break
    rows = []
    last_month = None
    for _, row in df.iterrows():
        raw_month = row.get(date_col); amt = row.get(amt_col); deal = row.get('商談名')
        ym = normalize_month(raw_month)
        if ym: last_month = ym
        else: ym = last_month
        if amt is None or (isinstance(amt, float) and pd.isna(amt)): continue
        if not deal or (isinstance(deal, float) and pd.isna(deal)): continue
        if ym is None: continue
        try: a = float(amt)
        except: continue
        if a == 0: continue
        rows.append({'month': ym, 'deal': str(deal).strip(), 'amount': a, 'kind': kind_prefix})
    return rows


files_income = [
    r'C:/Users/hara/Downloads/入金日１_2024年9月～2025年8月_複製済み (1).xlsx',
    r'C:/Users/hara/Downloads/入金日２_2023年9月～2024年8月_複製済み (1).xlsx',
]
files_payment = [
    r'C:/Users/hara/Downloads/国内支払１_2024年9月～2025年8月_複製済み (1).xlsx',
    r'C:/Users/hara/Downloads/国内支払２_2023年9月～2024年8月_複製済み (1).xlsx',
    r'C:/Users/hara/Downloads/国内支払３_2023年9月～2024年8月_複製済み (1).xlsx',
    r'C:/Users/hara/Downloads/国内支払４_2023年9月～2024年8月_複製済み (1).xlsx',
    r'C:/Users/hara/Downloads/国内支払５_2023年9月～2024年8月_複製済み (1).xlsx',
]

old_income = []
for fp in files_income: old_income.extend(read_old(fp, '入金'))
old_payment = []
for fp in files_payment: old_payment.extend(read_old(fp, '支払'))

ct = pd.read_excel(r'C:/Users/hara/Downloads/confirmed_transactions.xlsx')
ct['月'] = ct['取引日'].apply(lambda d: f'{d.year:04d}-{d.month:02d}')

# 2026-04 以降に絞り、月別×商談で集計
TARGET_FROM = '2026-04'

def collect_to_dict(records, kind_filter):
    out = {}  # (月, 商談) → amount
    for r in records:
        if r['month'] < TARGET_FROM: continue
        key = (r['month'], r['deal'])
        out[key] = out.get(key, 0) + r['amount']
    return out

old_in_dict = collect_to_dict(old_income, '入金')
old_pay_dict = collect_to_dict(old_payment, '支払')

new_in = ct[(ct['transaction_type']=='income') & (ct['月'] >= TARGET_FROM)]
new_pay = ct[(ct['transaction_type']=='payment') & (ct['月'] >= TARGET_FROM)]
new_in_dict = new_in.groupby(['月', '商談名'])['金額'].sum().to_dict()
new_pay_dict = new_pay.groupby(['月', '商談名'])['金額'].sum().to_dict()


def print_diff(old_d, new_d, label):
    print(f'\\n=== {label} 商談単位 差分（{TARGET_FROM} 以降）===')
    all_keys = set(old_d) | set(new_d)
    diffs = []
    for k in all_keys:
        o = old_d.get(k, 0); n = new_d.get(k, 0)
        d = n - o
        if abs(d) > 1:
            diffs.append((k[0], k[1], o, n, d))
    diffs.sort(key=lambda x: (x[0], -abs(x[4])))
    print(f'{"月":<10}{"商談名":<48}{"旧":>14}{"新":>14}{"差":>14}')
    for m, deal, o, n, d in diffs:
        print(f'{m:<10}{deal[:46]:<48}{int(o):>14,}{int(n):>14,}{int(d):>+14,}')
    print(f'  ({len(diffs)} 件の差異商談)')

print_diff(old_in_dict, new_in_dict, '入金')
print_diff(old_pay_dict, new_pay_dict, '支払')
