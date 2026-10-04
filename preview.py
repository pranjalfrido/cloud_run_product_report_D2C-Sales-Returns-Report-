import sys
sys.stdout.reconfigure(encoding='utf-8')

import os, json, base64
from datetime import date, timedelta
from premailer import transform
from google.cloud import bigquery
from google.oauth2 import service_account

# Load credentials from local sa_key.json
sa_path = os.path.join(os.path.dirname(__file__), '..', 'sa_key.json')
with open(sa_path) as f:
    sa_info = json.load(f)
creds = service_account.Credentials.from_service_account_info(
    sa_info, scopes=['https://www.googleapis.com/auth/cloud-platform']
)
client = bigquery.Client(credentials=creds, project=sa_info['project_id'])

today      = date.today()
cutoff     = today - timedelta(days=15)
start_date = cutoff - timedelta(days=30)
date_label = f"{start_date.strftime('%d %b %Y')} – {(cutoff - timedelta(days=1)).strftime('%d %b %Y')} (30 days)"

def q(sql):
    return list(client.query(sql).result())

def pct(val):
    if val is None: return '—'
    return f"{val}%"

def num(val):
    if val is None: return '—'
    return f"{val:,}"

def lakh(val):
    if val is None: return '—'
    if val >= 10000000: return f'₹{val/10000000:.2f} Cr'
    if val >= 100000: return f'₹{val/100000:.1f}L'
    if val >= 1000: return f'₹{val/1000:.0f}k'
    return f'₹{int(val):,}'

D2C_FILTER = f"""
    Channel = 'Shopify'
    AND SubChannel IN ('MyFrido', 'Mobility')
    AND OrderId NOT LIKE '%_EX%'
    AND DATE(OrderDate) >= '{start_date}'
    AND DATE(OrderDate) < '{cutoff}'
"""

ITEM_MASTER_JOIN = ""  # SubCategory and Category come directly from the fact table

print("Fetching KPIs...")
kpi_rows = q(f"""
SELECT
  SUM(SellingPrice_Inc_GST) AS gross_inc_gst,
  SUM(IF(Order_Status = 'Cancelled', SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(Order_Status IN ('RTO','Return'), SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(Order_Status = 'CIR', SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF(Order_Status = 'Exchange', SellingPrice_Inc_GST, 0)) AS exch_rev,
  -- Net exc GST: excl cancel/rto/cir only (exchange+NULL included) — matches dashboard "Net Revenue Ex Gst" ≈₹6.02Cr
  SUM(IF((Order_Status IS NULL OR Order_Status NOT IN ('Cancelled','CIR')) AND (Order_Status IS NULL OR Order_Status NOT IN ('RTO','Return')), SellingPrice_Exc_GST, 0)) AS net_exc_gst,
  -- Net inc GST: excl cancel/rto/cir (exchange included, NULL status included)
  SUM(IF((Order_Status IS NULL OR Order_Status NOT IN ('Cancelled','CIR')) AND (Order_Status IS NULL OR Order_Status NOT IN ('RTO','Return')), SellingPrice_Inc_GST, 0)) AS net_inc_gst_raw,
  -- Net units: excl cancelled/rto/cir (exchange included, NULL included)
  SUM(IF((Order_Status IS NULL OR Order_Status NOT IN ('Cancelled','CIR')) AND (Order_Status IS NULL OR Order_Status NOT IN ('RTO','Return')), ItemQty, 0)) AS net_units,
  COUNT(DISTINCT OrderId) AS orders,
  SUM(ItemQty) AS gross_units
FROM `frido-429506.production.fact_all_platform_sales_report` f
WHERE {D2C_FILTER}
""")

kpi_raw = kpi_rows[0] if kpi_rows else None
def _safe(val, default=0):
    return val if val is not None else default

if kpi_raw and _safe(kpi_raw.gross_inc_gst) > 0:
    gross_inc_gst   = _safe(kpi_raw.gross_inc_gst)
    cancel_rev      = _safe(kpi_raw.cancel_rev)
    rto_rev         = _safe(kpi_raw.rto_rev)
    cir_rev         = _safe(kpi_raw.cir_rev)
    exch_rev        = _safe(kpi_raw.exch_rev)
    net_inc_gst     = _safe(kpi_raw.net_inc_gst_raw)       # excl cancel/rto/cir/exchange — matches dashboard
    net_exc_gst     = _safe(kpi_raw.net_exc_gst)          # matches dashboard "Net Revenue Ex Gst" ≈ ₹6.02Cr
    net_units       = _safe(kpi_raw.net_units)
    gross_units     = _safe(kpi_raw.gross_units)
    orders          = _safe(kpi_raw.orders)
    asp             = round(gross_inc_gst / gross_units, 0) if gross_units else None   # gross rev / gross units (matches dashboard)
    aov             = round(gross_inc_gst / orders, 0) if orders else None              # gross rev / all orders (matches dashboard)
    leakage_rev     = cancel_rev + rto_rev + cir_rev
    overall_ret_pct = round(leakage_rev / gross_inc_gst * 100, 1) if gross_inc_gst else None  # cancel+rto+cir only (matches dashboard "Returns %")
    cancel_pct      = round(cancel_rev / gross_inc_gst * 100, 1) if gross_inc_gst else None
    rto_pct_kpi     = round(rto_rev    / gross_inc_gst * 100, 1) if gross_inc_gst else None
    cir_pct_kpi     = round(cir_rev    / gross_inc_gst * 100, 1) if gross_inc_gst else None
    exch_pct_kpi    = round(exch_rev   / gross_inc_gst * 100, 1) if gross_inc_gst else None
    predispatch_pct = cancel_pct   # cancel only = pre-dispatch loss
    returns_pct     = round((rto_rev + cir_rev) / gross_inc_gst * 100, 1) if gross_inc_gst else None  # RTO+CIR only = post-dispatch returns
    gst_amt         = net_inc_gst - net_exc_gst
    gst_pct         = round(gst_amt / net_inc_gst * 100, 1) if net_inc_gst else None
    daily_avg_net   = round(net_exc_gst / 30, 0) if net_exc_gst else None
    daily_avg_units = round(net_units / 30, 0) if net_units else None
    retained_pct    = round((gross_inc_gst - leakage_rev) / gross_inc_gst * 100, 1) if gross_inc_gst else None
    ret_threshold   = round((returns_pct or 0) * 1.5, 1)  # red if > 1.5x overall returns %
    no_data         = False
else:
    gross_inc_gst = net_inc_gst = net_exc_gst = gross_units = net_units = orders = None
    asp = aov = overall_ret_pct = cancel_pct = rto_pct_kpi = cir_pct_kpi = exch_pct_kpi = None
    gst_amt = gst_pct = daily_avg_net = None
    leakage_rev = 0; predispatch_pct = None; returns_pct = None
    daily_avg_units = None; retained_pct = None; cancel_rev = rto_rev = cir_rev = 0
    ret_threshold = 15
    no_data = True

print("Fetching leakage by category...")
leakage_rows = q(f"""
SELECT
  COALESCE(f.Category, 'Others') AS category,
  SUM(f.SellingPrice_Inc_GST) AS gross_rev,
  SUM(f.ItemQty) AS gross_units,
  SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) AS exch_rev
FROM `frido-429506.production.fact_all_platform_sales_report` f
{ITEM_MASTER_JOIN}
WHERE {D2C_FILTER}
GROUP BY 1
ORDER BY gross_rev DESC
""")

print("Fetching COD vs Prepaid summary...")
pay_split = q(f"""
SELECT
  CASE WHEN payment_type = 'COD' THEN 'COD' ELSE 'Prepaid' END AS pay_mode,
  SUM(ItemQty) AS qty,
  COUNT(DISTINCT OrderId) AS orders,
  SUM(SellingPrice_Inc_GST) AS gross_rev,
  SUM(IF(Order_Status = 'Cancelled', SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(Order_Status IN ('RTO','Return'), SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(Order_Status = 'CIR', SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF((Order_Status IS NULL OR Order_Status NOT IN ('Cancelled','CIR')) AND (Order_Status IS NULL OR Order_Status NOT IN ('RTO','Return')), SellingPrice_Inc_GST, 0)) AS net_rev
FROM `frido-429506.production.fact_all_platform_sales_report`
WHERE {D2C_FILTER}
GROUP BY 1
""")

print("Fetching top 20 products...")
top_products = q(f"""
SELECT
  COALESCE(f.SubCategory, f.masterskucode) AS product,
  COALESCE(f.Category, 'Others') AS category,
  SUM(f.SellingPrice_Inc_GST) AS gross_rev,
  SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) AS exch_rev,
  SUM(IF((f.Order_Status IS NULL OR f.Order_Status NOT IN ('Cancelled','CIR','Exchange','Return')) AND (f.Order_Status IS NULL OR f.Order_Status NOT IN ('RTO')), f.SellingPrice_Exc_GST, 0)) AS net_exc_gst,
  SUM(f.ItemQty) AS gross_units
FROM `frido-429506.production.fact_all_platform_sales_report` f
{ITEM_MASTER_JOIN}
WHERE {D2C_FILTER}
  AND f.masterskucode IS NOT NULL AND TRIM(f.masterskucode) != ''
GROUP BY 1, 2
HAVING gross_units >= 3
ORDER BY gross_rev DESC
LIMIT 20
""")

print("Fetching concerning products...")
concern_products = q(f"""
SELECT
  COALESCE(f.SubCategory, f.masterskucode) AS product,
  COALESCE(f.Category, 'Others') AS category,
  SUM(f.SellingPrice_Inc_GST) AS gross_rev,
  SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) AS exch_rev,
  SUM(IF((f.Order_Status IS NULL OR f.Order_Status NOT IN ('Cancelled','CIR','Exchange','Return')) AND (f.Order_Status IS NULL OR f.Order_Status NOT IN ('RTO')), f.SellingPrice_Exc_GST, 0)) AS net_exc_gst,
  SUM(f.ItemQty) AS gross_units
FROM `frido-429506.production.fact_all_platform_sales_report` f
{ITEM_MASTER_JOIN}
WHERE {D2C_FILTER}
  AND f.masterskucode IS NOT NULL AND TRIM(f.masterskucode) != ''
GROUP BY 1, 2
HAVING gross_units >= 20
ORDER BY
  (SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0))
 + SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0))
 + SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0))) DESC
LIMIT 15
""")

print("Fetching all products by category...")
all_products_by_cat = q(f"""
SELECT
  CASE
    WHEN COALESCE(f.Category, 'Others') IN ('Covers', 'Frido', 'Spare Parts') THEN 'Others'
    ELSE COALESCE(f.Category, 'Others')
  END AS category,
  COALESCE(f.SubCategory, f.masterskucode) AS product,
  SUM(f.SellingPrice_Inc_GST) AS gross_rev,
  SUM(f.ItemQty) AS gross_units,
  SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0)) AS cir_rev,
  SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) AS exch_rev
FROM `frido-429506.production.fact_all_platform_sales_report` f
WHERE {D2C_FILTER}
  AND f.masterskucode IS NOT NULL AND TRIM(f.masterskucode) != ''
GROUP BY 1, 2
HAVING gross_units >= 1
ORDER BY 1, gross_rev DESC
""")

print("Fetching COD risk SKUs...")
cod_risk = q(f"""
WITH base AS (
  SELECT
    COALESCE(f.SubCategory, f.masterskucode) AS product,
    COALESCE(f.Category, 'Others') AS category,
    f.SellingPrice_Inc_GST, f.Order_Status, f.payment_type, f.ItemQty
  FROM `frido-429506.production.fact_all_platform_sales_report` f
  {ITEM_MASTER_JOIN}
  WHERE {D2C_FILTER}
),
agg AS (
  SELECT product, category,
    SUM(SellingPrice_Inc_GST) AS gross_rev,
    SUM(IF(payment_type = 'COD', SellingPrice_Inc_GST, 0)) AS cod_rev,
    SUM(IF(Order_Status IN ('RTO','Return'), SellingPrice_Inc_GST, 0)) AS rto_rev_all,
    SUM(ItemQty) AS gross_units
  FROM base GROUP BY 1, 2
)
SELECT product, category, gross_rev, cod_rev, gross_units,
  ROUND(cod_rev / NULLIF(gross_rev, 0) * 100, 1) AS cod_pct,
  ROUND(rto_rev_all / NULLIF(gross_rev, 0) * 100, 1) AS rto_pct,
  ROUND(cod_rev / NULLIF(gross_rev, 0) * rto_rev_all / NULLIF(gross_rev, 0) * gross_rev, 0) AS rev_at_risk
FROM agg
WHERE ROUND(cod_rev / NULLIF(gross_rev, 0) * 100, 1) > 30
  AND ROUND(rto_rev_all / NULLIF(gross_rev, 0) * 100, 1) > 8
  AND gross_units >= 20
ORDER BY rev_at_risk DESC
""")

print("Fetching exchange watch SKUs...")
exch_watch = q(f"""
SELECT
  COALESCE(f.SubCategory, f.masterskucode) AS product,
  COALESCE(f.Category, 'Others') AS category,
  SUM(f.ItemQty) AS gross_units,
  SUM(f.SellingPrice_Inc_GST) AS gross_rev,
  SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) AS exch_rev,
  ROUND(SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) / NULLIF(SUM(f.SellingPrice_Inc_GST), 0) * 100, 1) AS exch_pct,
  SUM(IF(f.Order_Status = 'Cancelled', f.SellingPrice_Inc_GST, 0)) AS cancel_rev,
  SUM(IF(f.Order_Status IN ('RTO','Return'), f.SellingPrice_Inc_GST, 0)) AS rto_rev,
  SUM(IF(f.Order_Status = 'CIR', f.SellingPrice_Inc_GST, 0)) AS cir_rev
FROM `frido-429506.production.fact_all_platform_sales_report` f
{ITEM_MASTER_JOIN}
WHERE {D2C_FILTER}
  AND f.masterskucode IS NOT NULL AND TRIM(f.masterskucode) != ''
GROUP BY 1, 2
HAVING SUM(IF(f.Order_Status = 'Exchange', f.SellingPrice_Inc_GST, 0)) > 0
ORDER BY exch_rev DESC
LIMIT 20
""")

print("All data fetched. Building HTML...")

def td(val, red_if=None, suffix=''):
    if val is None: return '<td>—</td>'
    s = f"{val}{suffix}"
    if red_if and red_if(val):
        return f'<td style="color:#000000;font-weight:bold;">{s}</td>'
    return f'<td>{s}</td>'

def td_name(val):
    return f'<td class="name">{val}</td>'

TS  = 'bgcolor="#1F3864" color="#ffffff" style="background:#1F3864 !important;color:#fff !important;font-weight:bold;padding:4px 7px;border:1px solid #1F3864;text-align:center;white-space:nowrap;"'
TSL = 'bgcolor="#1F3864" color="#ffffff" style="background:#1F3864 !important;color:#fff !important;font-weight:bold;padding:4px 7px;border:1px solid #1F3864;text-align:left;white-space:nowrap;"'

def fmt_rev(val):
    if val is None: return '—'
    if val >= 10000000: return f'₹{val/10000000:.2f} Cr'
    if val >= 100000: return f'₹{val/100000:.1f}L'
    if val >= 1000:   return f'₹{val/1000:.0f}k'
    return f'₹{int(val):,}'

def fmt_units(val):
    if val is None: return '—'
    return f'{int(val):,}'

def fmt_asp(val):
    if val is None: return '—'
    return f'₹{int(val):,}'

ret_red   = ' red' if (overall_ret_pct or 0) > 15 else ''
ret_style = ' style="color:#c00000;font-weight:bold;"' if (overall_ret_pct or 0) > 15 else ''

def build_leakage_rows():
    rows = ''
    tot_gross = tot_can = tot_rto = tot_cir = tot_exch = tot_u = 0
    data = []
    for r in leakage_rows:
        g  = _safe(r.gross_rev);  u  = _safe(r.gross_units)
        c  = _safe(r.cancel_rev); rt = _safe(r.rto_rev)
        ci = _safe(r.cir_rev);    ex = _safe(r.exch_rev)
        net_ret  = g - c - rt - ci
        ret_pct  = round(net_ret / g * 100, 1) if g else None
        can_pct  = round(c  / g * 100, 1) if g else None
        rto_pct  = round(rt / g * 100, 1) if g else None
        cir_pct  = round(ci / g * 100, 1) if g else None
        exch_pct = round(ex / g * 100, 1) if g else None
        asp_v    = round(g / u, 0) if u else None
        tot_gross += g; tot_can += c; tot_rto += rt
        tot_cir += ci; tot_exch += ex; tot_u += u
        data.append((r.category, g, u, asp_v, can_pct, rto_pct, cir_pct, exch_pct, net_ret, ret_pct))
    tot_net_ret  = tot_gross - tot_can - tot_rto - tot_cir
    avg_ret_pct  = round(tot_net_ret / tot_gross * 100, 1) if tot_gross else 100
    tot_ret_pct  = round(tot_net_ret / tot_gross * 100, 1) if tot_gross else 0
    tot_can_pct  = round(tot_can  / tot_gross * 100, 1) if tot_gross else 0
    tot_rto_pct  = round(tot_rto  / tot_gross * 100, 1) if tot_gross else 0
    tot_cir_pct  = round(tot_cir  / tot_gross * 100, 1) if tot_gross else 0
    tot_exch_pct = round(tot_exch / tot_gross * 100, 1) if tot_gross else 0
    def bld(val, avg):
        if val is None: return '<td>—</td>'
        is_bold = avg and val > avg * 1.2
        s = f'{val}%'
        return f'<td style="font-weight:bold;">{s}</td>' if is_bold else f'<td>{s}</td>'
    for (cat, g, u, asp_v, can_pct, rto_pct, cir_pct, exch_pct, net_ret, ret_pct) in data:
        flag = ret_pct is not None and ret_pct < avg_ret_pct
        ret_td = f'<td style="color:#000000;font-weight:bold;">{ret_pct}%</td>' if flag else f'<td>{ret_pct if ret_pct is not None else "—"}%</td>'
        asp_str = '—' if asp_v is None else f'₹{int(asp_v):,}'
        rows += f"""<tr>
          {td_name(cat)}
          <td>{lakh(g)}</td>
          <td>{fmt_units(int(u))}</td>
          <td>{asp_str}</td>
          {bld(can_pct,  tot_can_pct)}
          {bld(rto_pct,  tot_rto_pct)}
          {bld(cir_pct,  tot_cir_pct)}
          {bld(exch_pct, tot_exch_pct)}
          <td>{lakh(net_ret)}</td>
          {ret_td}
        </tr>"""
    tot_asp      = round(tot_gross / tot_u, 0) if tot_u else None
    tot_asp_str  = '—' if tot_asp is None else f'₹{int(tot_asp):,}'
    rows += f"""<tr>
      <td {TSL}>Total</td>
      <td {TS}>{lakh(tot_gross)}</td>
      <td {TS}>{fmt_units(int(tot_u))}</td>
      <td {TS}>{tot_asp_str}</td>
      <td {TS}>{tot_can_pct}%</td>
      <td {TS}>{tot_rto_pct}%</td>
      <td {TS}>{tot_cir_pct}%</td>
      <td {TS}>{tot_exch_pct}%</td>
      <td {TS}>{lakh(tot_net_ret)}</td>
      <td {TS}>{tot_ret_pct}%</td>
    </tr>"""
    return rows

def build_product_rows(rows_data, total_gross=None):
    rows = ''
    tot_g = tot_c = tot_rt = tot_ci = tot_ex = tot_u = 0
    for r in rows_data:
        g   = _safe(r.gross_rev);  c  = _safe(r.cancel_rev)
        rt  = _safe(r.rto_rev);    ci = _safe(r.cir_rev)
        ex  = _safe(r.exch_rev)
        gross_u  = _safe(r.gross_units)
        net_ret  = g - c - rt - ci
        asp_v    = round(g / gross_u, 0) if gross_u else None
        asp_str  = '—' if asp_v is None else f'₹{int(asp_v):,}'
        can_p    = round(c  / g * 100, 1) if g else None
        rto_p    = round(rt / g * 100, 1) if g else None
        cir_p    = round(ci / g * 100, 1) if g else None
        exch_p   = round(ex / g * 100, 1) if g else None
        ret_p    = round((c + rt + ci) / g * 100, 1) if g else None
        ret_pct_v = round(net_ret / g * 100, 1) if g else None
        rows += f"""<tr>
          {td_name(r.product)}
          <td>{fmt_units(int(gross_u))}</td>
          <td>{lakh(g)}</td>
          <td>{asp_str}</td>
          {td(can_p,  None, '%')}
          {td(rto_p,  None, '%')}
          {td(cir_p,  None, '%')}
          {td(exch_p, None, '%')}
          <td>{lakh(net_ret)}</td>
          {td(ret_p, lambda x: x > ret_threshold, '%')}
        </tr>"""
        tot_g += g; tot_c += c; tot_rt += rt; tot_ci += ci; tot_ex += ex; tot_u += gross_u
    tot_net_ret = tot_g - tot_c - tot_rt - tot_ci
    t_asp       = round(tot_g / tot_u, 0) if tot_u else None
    t_asp_str   = '—' if t_asp is None else f'₹{int(t_asp):,}'
    t_can_p     = round(tot_c  / tot_g * 100, 1) if tot_g else None
    t_rto_p     = round(tot_rt / tot_g * 100, 1) if tot_g else None
    t_cir_p     = round(tot_ci / tot_g * 100, 1) if tot_g else None
    t_exch_p    = round(tot_ex / tot_g * 100, 1) if tot_g else None
    t_ret_p     = round((tot_c + tot_rt + tot_ci) / tot_g * 100, 1) if tot_g else None
    def bld(val, avg):
        if val is None: return '<td>—</td>'
        s = f'{val}%'
        return f'<td style="font-weight:bold;">{s}</td>' if (avg and val > avg * 1.2) else f'<td>{s}</td>'
    # Re-render rows with bold flagging now that we have totals
    rows = ''
    for r in rows_data:
        g   = _safe(r.gross_rev);  c  = _safe(r.cancel_rev)
        rt  = _safe(r.rto_rev);    ci = _safe(r.cir_rev)
        ex  = _safe(r.exch_rev)
        gross_u  = _safe(r.gross_units)
        net_ret  = g - c - rt - ci
        asp_v    = round(g / gross_u, 0) if gross_u else None
        asp_str  = '—' if asp_v is None else f'₹{int(asp_v):,}'
        can_p    = round(c  / g * 100, 1) if g else None
        rto_p    = round(rt / g * 100, 1) if g else None
        cir_p    = round(ci / g * 100, 1) if g else None
        exch_p   = round(ex / g * 100, 1) if g else None
        ret_p    = round((c + rt + ci) / g * 100, 1) if g else None
        rows += f"""<tr>
          {td_name(r.product)}
          <td>{fmt_units(int(gross_u))}</td>
          <td>{lakh(g)}</td>
          <td>{asp_str}</td>
          {bld(can_p,  t_can_p)}
          {bld(rto_p,  t_rto_p)}
          {bld(cir_p,  t_cir_p)}
          {bld(exch_p, t_exch_p)}
          <td>{lakh(net_ret)}</td>
          {td(ret_p, lambda x: x > ret_threshold, '%')}
        </tr>"""
    rows += f"""<tr>
      <td {TSL}>Total</td>
      <td {TS}>{fmt_units(int(tot_u))}</td>
      <td {TS}>{lakh(tot_g)}</td>
      <td {TS}>{t_asp_str}</td>
      <td {TS}>{pct(t_can_p)}</td>
      <td {TS}>{pct(t_rto_p)}</td>
      <td {TS}>{pct(t_cir_p)}</td>
      <td {TS}>{pct(t_exch_p)}</td>
      <td {TS}>{lakh(tot_net_ret)}</td>
      <td {TS}>{pct(t_ret_p)}</td>
    </tr>"""
    return rows, tot_g

def build_pay_split_html():
    cod = next((r for r in pay_split if r.pay_mode == 'COD'), None)
    pre = next((r for r in pay_split if r.pay_mode == 'Prepaid'), None)
    def ps(r):
        if not r: return {}
        qty     = int(_safe(r.qty))
        orders  = int(_safe(r.orders))
        gross_r = _safe(r.gross_rev)
        net_r   = _safe(r.net_rev)
        can_r   = _safe(r.cancel_rev)
        rto_r   = _safe(r.rto_rev)
        cir_r   = _safe(r.cir_rev)
        exch_r  = gross_r - net_r - can_r - rto_r - cir_r
        aov_v   = round(gross_r / orders, 0) if orders else None
        can_p   = round(can_r  / gross_r * 100, 1) if gross_r else 0
        rto_p   = round(rto_r  / gross_r * 100, 1) if gross_r else 0
        cir_p   = round(cir_r  / gross_r * 100, 1) if gross_r else 0
        exch_p  = round(exch_r / gross_r * 100, 1) if gross_r and exch_r > 0 else 0
        net_p   = round(net_r  / gross_r * 100, 1) if gross_r else 0
        return dict(qty=qty, orders=orders, gross_r=gross_r, aov_v=aov_v,
                    can_p=can_p, rto_p=rto_p, cir_p=cir_p, exch_p=exch_p, net_p=net_p)
    c = ps(cod); p = ps(pre)
    tot_rev = (c.get('gross_r', 0)) + (p.get('gross_r', 0))
    cod_rev_pct = round(c.get('gross_r', 0) / tot_rev * 100, 1) if tot_rev else 0
    pre_rev_pct = round(p.get('gross_r', 0) / tot_rev * 100, 1) if tot_rev else 0
    def cell(d, label, color, bg, rev_pct):
        if not d: return f'<td style="width:50%;border:1px solid #dde3ee;border-radius:4px;padding:10px 12px;background:{bg};vertical-align:top;">No data</td>'
        aov_str = f'₹{int(d["aov_v"]):,}' if d.get('aov_v') else '—'
        return f"""<td style="width:50%;border:1px solid #dde3ee;border-radius:4px;padding:10px 12px;background:{bg};vertical-align:top;">
          <div style="font-size:9px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:0.3px;margin-bottom:6px;">{label}</div>
          <div style="font-size:18px;font-weight:bold;color:{color};">{lakh(d['gross_r'])} &nbsp;<span style="font-size:11px;color:#666;font-weight:normal;">{fmt_units(d['qty'])} units ({rev_pct}%)</span></div>
          <div style="font-size:9px;color:#555;margin-top:6px;line-height:1.8;">
            Orders: {fmt_units(d['orders'])} &nbsp;|&nbsp; AOV: {aov_str}<br>
            Cancel: {d['can_p']}% &nbsp;|&nbsp; RTO: {d['rto_p']}% &nbsp;|&nbsp; CIR: {d['cir_p']}% &nbsp;|&nbsp; Exch: {d['exch_p']}%<br>
            Net Retained: {d['net_p']}% of gross rev
          </div>
        </td>"""
    return f"""
    <table style="width:100%;border-collapse:separate;border-spacing:6px;margin-top:10px;">
      <tr>
        {cell(c, 'COD', '#c00000', '#fff8e1', cod_rev_pct)}
        {cell(p, 'Prepaid', '#1a7a1a', '#f0f7f0', pre_rev_pct)}
      </tr>
    </table>"""

def build_exch_watch_html():
    if not exch_watch:
        return ''
    rows = ''
    tot_g = tot_ex = tot_can = tot_rt = tot_ci = tot_u = 0
    for r in exch_watch:
        g    = _safe(r.gross_rev)
        ex   = _safe(r.exch_rev)
        can  = _safe(r.cancel_rev)
        rt   = _safe(r.rto_rev)
        ci   = _safe(r.cir_rev)
        u    = _safe(r.gross_units)
        ep   = float(r.exch_pct) if r.exch_pct is not None else 0
        canp = round(can / g * 100, 1) if g else 0
        rtop = round(rt  / g * 100, 1) if g else 0
        cirp = round(ci  / g * 100, 1) if g else 0
        rows += f"""<tr>
          {td_name(r.product)}
          <td>{fmt_units(int(u))}</td>
          <td>{lakh(g)}</td>
          <td style="font-weight:bold;">{lakh(ex)}</td>
          <td style="font-weight:bold;">{ep}%</td>
          <td>{canp}%</td>
          <td>{rtop}%</td>
          <td>{cirp}%</td>
        </tr>"""
        tot_g += g; tot_ex += ex; tot_can += can; tot_rt += rt; tot_ci += ci; tot_u += u
    t_ep   = round(tot_ex  / tot_g * 100, 1) if tot_g else 0
    t_canp = round(tot_can / tot_g * 100, 1) if tot_g else 0
    t_rtop = round(tot_rt  / tot_g * 100, 1) if tot_g else 0
    t_cirp = round(tot_ci  / tot_g * 100, 1) if tot_g else 0
    rows += f"""<tr>
      <td {TSL}>Total</td>
      <td {TS}>{fmt_units(int(tot_u))}</td>
      <td {TS}>{lakh(tot_g)}</td>
      <td {TS}>{lakh(tot_ex)}</td>
      <td {TS}>{t_ep}%</td>
      <td {TS}>{t_canp}%</td>
      <td {TS}>{t_rtop}%</td>
      <td {TS}>{t_cirp}%</td>
    </tr>"""
    return f"""
  <div class="section">
    <div class="section-title">Exchange Watch — Top 20 by Exchange Value</div>
    <p class="callout">Products ranked by exchange revenue — high exchange % often signals size-chart issues.</p>
    <div class="tbl-wrap">
    <table class="dtbl">
      <tr>
        <th class="left">Product</th><th>Sold Units</th>
        <th>Gross Rev</th><th>Exch Rev</th><th>Exch %</th>
        <th>Cancel %</th><th>RTO %</th><th>CIR %</th>
      </tr>
      {rows}
    </table>
    </div>
  </div>"""

def build_cod_rows():
    if not cod_risk:
        return '<tr><td colspan="6" style="color:#888;font-size:11px;padding:8px;text-align:left;">No SKUs flagged today.</td></tr>'
    rows = ''
    tot_gross = tot_risk = 0
    for r in cod_risk:
        rows += f"""<tr>
          {td_name(r.product)}
          <td>{r.category}</td>
          <td>{lakh(r.gross_rev)}</td>
          {td(r.cod_pct, lambda x: x > 30, '%')}
          {td(r.rto_pct, lambda x: x > 8, '%')}
          <td>{'—' if r.rev_at_risk is None else lakh(float(r.rev_at_risk))}</td>
        </tr>"""
        tot_gross += _safe(r.gross_rev)
        tot_risk  += float(r.rev_at_risk) if r.rev_at_risk is not None else 0
    rows += f"""<tr class="total-row">
      <td class="name">Total</td>
      <td>—</td>
      <td>{lakh(tot_gross)}</td>
      <td>—</td>
      <td>—</td>
      <td>{lakh(tot_risk)}</td>
    </tr>"""
    return rows

PRODUCT_TH = """
  <th class="left">Product</th>
  <th>Qty</th>
  <th>Gross Rev</th>
  <th>ASP</th>
  <th>Cancel %</th>
  <th>RTO %</th>
  <th>CIR %</th>
  <th>Exch %</th>
  <th>Net Retained</th>
  <th>Tot.Ret.%</th>
"""

def build_category_product_rows(rows_data):
    rows = ''
    tot_g = tot_c = tot_rt = tot_ci = tot_ex = tot_u = 0
    computed_rows = []
    for r in rows_data:
        g  = _safe(r.gross_rev);  c  = _safe(r.cancel_rev)
        rt = _safe(r.rto_rev);   ci = _safe(r.cir_rev)
        ex = _safe(r.exch_rev);  u  = _safe(r.gross_units)
        net_ret = g - c - rt - ci
        ret_pct = round(net_ret / g * 100, 1) if g else None
        asp_v   = round(g / u, 0) if u else None
        can_p   = round(c  / g * 100, 1) if g else None
        rto_p   = round(rt / g * 100, 1) if g else None
        cir_p   = round(ci / g * 100, 1) if g else None
        exch_p  = round(ex / g * 100, 1) if g else None
        computed_rows.append((r.product, int(u), g, asp_v, can_p, rto_p, cir_p, exch_p, net_ret, ret_pct))
        tot_g += g; tot_c += c; tot_rt += rt; tot_ci += ci; tot_ex += ex; tot_u += u
    tot_net_ret = tot_g - tot_c - tot_rt - tot_ci
    t_asp     = round(tot_g / tot_u, 0) if tot_u else None
    t_can_p   = round(tot_c  / tot_g * 100, 1) if tot_g else None
    t_rto_p   = round(tot_rt / tot_g * 100, 1) if tot_g else None
    t_cir_p   = round(tot_ci / tot_g * 100, 1) if tot_g else None
    t_exch_p  = round(tot_ex / tot_g * 100, 1) if tot_g else None
    t_ret_p   = round(tot_net_ret / tot_g * 100, 1) if tot_g else None
    def bld(val, avg):
        if val is None: return '<td>—</td>'
        s = f'{val}%'
        return f'<td style="font-weight:bold;">{s}</td>' if (avg and val > avg * 1.2) else f'<td>{s}</td>'
    for (product, u, g, asp_v, can_p, rto_p, cir_p, exch_p, net_ret, ret_pct) in computed_rows:
        asp_str = '—' if asp_v is None else f'₹{int(asp_v):,}'
        rows += f"""<tr>
          {td_name(product)}
          <td>{lakh(g)}</td>
          <td>{fmt_units(u)}</td>
          <td>{asp_str}</td>
          {bld(can_p,  t_can_p)}
          {bld(rto_p,  t_rto_p)}
          {bld(cir_p,  t_cir_p)}
          {bld(exch_p, t_exch_p)}
          <td>{lakh(net_ret)}</td>
          <td>{'—' if ret_pct is None else f'{ret_pct}%'}</td>
        </tr>"""
    t_asp_str = '—' if t_asp is None else f'₹{int(t_asp):,}'
    rows += f"""<tr>
      <td {TSL}>Total</td>
      <td {TS}>{lakh(tot_g)}</td>
      <td {TS}>{fmt_units(int(tot_u))}</td>
      <td {TS}>{t_asp_str}</td>
      <td {TS}>{pct(t_can_p)}</td>
      <td {TS}>{pct(t_rto_p)}</td>
      <td {TS}>{pct(t_cir_p)}</td>
      <td {TS}>{pct(t_exch_p)}</td>
      <td {TS}>{lakh(tot_net_ret)}</td>
      <td {TS}>{pct(t_ret_p)}</td>
    </tr>"""
    return rows


def build_category_sections():
    from collections import defaultdict
    by_cat = defaultdict(list)
    for r in all_products_by_cat:
        by_cat[r.category].append(r)
    sections = ''
    for cat in sorted(by_cat.keys()):
        rows = build_category_product_rows(by_cat[cat])
        sections += f"""
  <div class="section">
    <div class="section-title">{cat} — All Products</div>
    <div class="tbl-wrap">
    <table class="dtbl">
      <tr>
        <th class="left">Product</th>
        <th>Gross Rev</th><th>Qty</th><th>ASP</th>
        <th>Cancel %</th><th>RTO %</th><th>CIR %</th><th>Exch %</th>
        <th>Net Retained</th><th>Retained %</th>
      </tr>
      {rows}
    </table>
    </div>
    <p class="note">All products in {cat}. Sorted by Gross Revenue. % columns bold if &gt;1.2× category avg. Net Retained = Gross Rev − Cancelled − RTO − CIR.</p>
  </div>"""
    return sections


_top20_rows, _top20_gross = build_product_rows(top_products)
_top20_share = round(_top20_gross / gross_inc_gst * 100, 1) if gross_inc_gst and _top20_gross else None
_top20_share_note = f'These top 20 products represent {_top20_share}% of total D2C gross revenue ({fmt_rev(gross_inc_gst)}).' if _top20_share else ''
_top20_section = f"""<div class="section">
    <div class="section-title">Top 20 Products — Revenue View</div>
    <div class="tbl-wrap">
    <table class="dtbl">
      <tr>{PRODUCT_TH}</tr>
      {_top20_rows}
    </table>
    </div>
    <p class="note">Sorted by Gross Revenue. Min 3 units. Tot.Ret.% bold if &gt; 1.5× overall ({ret_threshold}%). {_top20_share_note}</p>
  </div>"""

_leakage_rows, _leakage_gross = build_product_rows(concern_products)
_leakage_section = f"""<div class="section">
    <div class="section-title">Top Leakage Products — by &#8377; Lost</div>
    <div class="tbl-wrap">
    <table class="dtbl">
      <tr>{PRODUCT_TH}</tr>
      {_leakage_rows}
    </table>
    </div>
    <p class="note">Sorted by &#8377; lost (Cancel + RTO + CIR). Min 20 units.</p>
  </div>"""

html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: Calibri, Arial, sans-serif; font-size: 13px; color: #222; background: #f4f6f9; }}
  .wrapper {{ width: 100%; max-width: 700px; margin: 0 auto; background: #fff; }}
  .header {{ background: #1F3864; color: #fff; padding: 10px 20px; }}
  .header h1 {{ font-size: 17px; font-weight: bold; }}
  .header .sub {{ font-size: 11px; color: #aac4e0; margin-top: 5px; }}
  .section {{ padding: 14px 12px; border-bottom: 1px solid #e4e8ef; }}
  .section-title {{ font-size: 11px; font-weight: bold; color: #1F3864; border-left: 3px solid #2E75B6; padding-left: 8px; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }}
  .kpi-outer {{ width: 100%; border-collapse: separate; border-spacing: 5px; margin-bottom: 6px; table-layout: fixed; }}
  .kpi-cell {{ border: 1px solid #dde3ee; border-radius: 4px; padding: 8px 8px; background: #fafbfd; text-align: left; vertical-align: top; width: 25%; overflow: hidden; }}
  .kpi-cell .kl {{ font-size: 9px; color: #555; font-weight: 700; text-transform: uppercase; letter-spacing: 0.3px; line-height: 1.3; display: block; margin-bottom: 3px; }}
  .kpi-cell .kv {{ font-size: 18px; font-weight: bold; color: #1F3864; display: block; line-height: 1.2; }}
  .kpi-cell .kv.red {{ color: #c00000; font-weight: bold; }}
  .kpi-cell .ksub {{ font-size: 9px; color: #999; display: block; margin-top: 2px; }}
  .tbl-wrap {{ width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }}
  .dtbl {{ width: 100%; border-collapse: collapse; font-size: 10px; white-space: nowrap; }}
  .dtbl th {{ background: #1F3864; color: #fff; padding: 5px 7px; text-align: center; font-size: 9px; font-weight: bold; white-space: nowrap; }}
  .dtbl th.left {{ text-align: left; }}
  .dtbl td {{ padding: 4px 5px; border: 1px solid #e0e5ef; text-align: center; color: #222; background: #fff; white-space: nowrap; }}
  .dtbl td.name {{ text-align: left; font-weight: bold; color: #1F3864; min-width: 125px; max-width: 155px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .dtbl tr:nth-child(even) td {{ background: #f7f9fc; }}
  .dtbl tr.total-row td {{ background: #1F3864 !important; color: #fff !important; font-weight: bold; border-color: #1F3864; }}
  .dtbl tr.total-row td.name {{ color: #fff !important; }}
  .callout {{ font-size: 10px; color: #444; margin-bottom: 8px; padding: 6px 10px; background: #fff8e1; border-left: 3px solid #f9a825; border-radius: 2px; }}
  .note {{ font-size: 10px; color: #888; margin-top: 6px; font-style: italic; }}
  .footer {{ background: #1F3864; color: #aac4e0; font-size: 10px; padding: 10px 16px; text-align: center; }}
  @media (max-width: 768px) {{
    .wrapper {{ max-width: 100%; }}
    .section {{ padding: 10px 4px; }}
    .kpi-cell .kv {{ font-size: 15px; }}
    .dtbl {{ font-size: 8px; }}
    .dtbl th {{ font-size: 7px; padding: 3px 2px; }}
    .dtbl td {{ font-size: 8px; padding: 3px 2px; }}
    .dtbl td.name {{ white-space: normal; }}
  }}
</style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <h1>Frido D2C Sales & Returns Report</h1>
    <div class="sub">{date_label} &nbsp;|&nbsp; D2C: MyFrido + Mobility</div>
  </div>

  <!-- SECTION 1: KPI CARDS -->
  <div class="section">
    <div class="section-title">Overall KPIs</div>
    <!-- Row 1: Revenue KPIs -->
    <table class="kpi-outer" role="presentation">
      <tr>
        <td class="kpi-cell">
          <span class="kl">Gross Rev (incl GST)</span>
          <span class="kv">{fmt_rev(gross_inc_gst)}</span>
          <span class="ksub">{fmt_units(gross_units)} units sold</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">Net Rev (incl GST)</span>
          <span class="kv">{fmt_rev(net_inc_gst)}</span>
          <span class="ksub">{fmt_units(net_units)} net units</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">Net Rev (excl GST)</span>
          <span class="kv">{fmt_rev(net_exc_gst)}</span>
          <span class="ksub">{fmt_units(net_units)} net units</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">Daily Avg Net Rev</span>
          <span class="kv">{fmt_rev(daily_avg_net)}</span>
          <span class="ksub">{fmt_units(daily_avg_units)} units/day · exc GST</span>
        </td>
      </tr>
    </table>
    <!-- Row 2: Loss % KPIs -->
    <table class="kpi-outer" role="presentation" style="margin-top:0;">
      <tr>
        <td class="kpi-cell">
          <span class="kl">Cancellation %</span>
          <span class="kv">{pct(cancel_pct)}</span>
          <span class="ksub">{fmt_rev(cancel_rev)} lost</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">RTO %</span>
          <span class="kv">{pct(rto_pct_kpi)}</span>
          <span class="ksub">{fmt_rev(rto_rev)} lost</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">CIR %</span>
          <span class="kv">{pct(cir_pct_kpi)}</span>
          <span class="ksub">{fmt_rev(cir_rev)} lost</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">Exchange %</span>
          <span class="kv">{pct(exch_pct_kpi)}</span>
          <span class="ksub">{fmt_rev(exch_rev)} exchanged</span>
        </td>
      </tr>
    </table>
    <!-- Row 3: Summary KPIs -->
    <table class="kpi-outer" role="presentation" style="margin-top:0;">
      <tr>
        <td class="kpi-cell">
          <span class="kl">Total Return %</span>
          <span class="kv{ret_red}"{ret_style}>{pct(overall_ret_pct)}</span>
          <span class="ksub">cir + rto + canc.</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">Total Retained %</span>
          <span class="kv">{pct(retained_pct)}</span>
          <span class="ksub">{fmt_rev(gross_inc_gst - leakage_rev if gross_inc_gst else None)} retained</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">ASP</span>
          <span class="kv">{fmt_asp(asp)}</span>
          <span class="ksub">gross rev ÷ gross units</span>
        </td>
        <td class="kpi-cell">
          <span class="kl">AOV</span>
          <span class="kv">{fmt_asp(aov)}</span>
          <span class="ksub">gross rev ÷ all orders</span>
        </td>
      </tr>
    </table>
    {build_pay_split_html()}
  </div>

  <!-- SECTION 2: LEAKAGE FUNNEL -->
  <div class="section">
    <div class="section-title">Revenue Leakage by Category</div>
    <div class="tbl-wrap">
    <table class="dtbl">
      <tr>
        <th class="left">Category</th>
        <th>Gross Rev</th><th>Qty</th><th>ASP</th>
        <th>Cancel %</th><th>RTO %</th><th>CIR %</th><th>Exch %</th>
        <th>Net Retained</th><th>Retained %</th>
      </tr>
      {build_leakage_rows()}
    </table>
    </div>
    <p class="note">Retained % highlighted red if below overall average. All values incl. GST. Net Retained = Gross Rev − Cancelled − RTO − CIR.</p>
  </div>

  <!-- SECTION 3: TOP 20 PRODUCTS -->
  {_top20_section}

  <!-- SECTION 4: TOP LEAKAGE -->
  {_leakage_section}

  <!-- SECTION 5: CATEGORY-WISE ALL PRODUCTS -->
  {build_category_sections()}

  {build_exch_watch_html()}

  <div class="section" style="background:#f9fafc;">
    <div class="section-title">Definitions &amp; Data Notes</div>
    <p style="font-size:10px;color:#444;line-height:1.8;">
      <b>Gross Rev</b>: Total revenue incl. GST across all order statuses.<br>
      <b>Net Retained</b>: Gross Rev − Cancelled − RTO − CIR (exchange not deducted — revenue recovered via replacement).<br>
      <b>Tot.Ret.%</b>: (Cancel + RTO + CIR) ÷ Gross Rev — total revenue lost.<br>
      <b>Total Retained %</b>: (Gross Rev − Cancel − RTO − CIR) ÷ Gross Rev.<br>
      <b>ASP</b>: Gross Rev ÷ Gross Units sold (all statuses).<br>
      <b>AOV</b>: Gross Rev ÷ Total Orders (all statuses, incl. cancelled).<br>
      <b>Daily Avg Net Rev</b>: Net Rev (excl GST) ÷ 30 days.<br>
      <b>Window</b>: {date_label}. D2C = Shopify (MyFrido + Mobility). Exchange replacement orders (_EX) excluded.
    </p>
  </div>

  <div class="footer">
    Frido D2C Sales & Returns Report &nbsp;|&nbsp; {date_label} &nbsp;|&nbsp; D2C = Shopify (MyFrido + Mobility)
  </div>

</div>
</body>
</html>"""

print("Inlining CSS...")
inlined = transform(html, remove_classes=False, strip_important=False)

out_path = os.path.join(os.path.dirname(__file__), 'preview.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(inlined)

print(f"Done! HTML saved to: {out_path}")
