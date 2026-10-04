# cloud_run_product_report_D2C-Sales-Returns-Report-

A scheduled Google Cloud Run service that generates and emails a daily **D2C Sales & Returns Report** for Frido.

## What it does

- Runs daily via **Cloud Scheduler**
- Queries **BigQuery** (`frido-429506.production.fact_all_platform_sales_report`) for D2C Shopify orders (MyFrido + Mobility channels)
- Computes **30+ KPIs** — gross revenue, net revenue, cancellation %, RTO %, CIR %, exchange %, ASP, AOV, COD vs Prepaid split
- Generates a **responsive HTML email** with 8 data sections and red/bold threshold highlighting
- Attaches a full **Excel file** (`Frido_D2C_Orders_{start}_{end}.xlsx`) with 21 columns of raw order data
- Sends to stakeholders via **Microsoft Graph API** (Microsoft 365)

## Project structure

```
main.py             # Core logic: queries, KPI computation, HTML + Excel generation, email dispatch
send_test.py        # Test sender — HTML email only (no Excel), sends to hardcoded test addresses
preview.py          # Local preview — writes preview.html, no email sent
preview.html        # Output of preview.py — open in browser to check layout
requirements.txt    # Python dependencies
Dockerfile          # Container definition (python:3.11-slim)
env_flags.yaml      # Cloud Run deploy config (env vars + secrets). Never commit real credentials.
DOCUMENTATION.docx  # Full developer documentation
```

## Quick start

```bash
# Generate HTML preview locally (no email)
python preview.py
# Open preview.html in browser

# Send test email (HTML only, no Excel) to test address
python send_test.py

# Full run (HTML + Excel attachment)
python main.py
```

Requires GCP service account key at `../sa_key.json` for local runs, and env vars from `env_flags.yaml`.

## Deploy

```bash
bash deploy.sh
```

Requires `gcloud` CLI authenticated with Cloud Run Admin + Secret Manager Admin on `frido-429506`.

## Data source

Single BigQuery table: `frido-429506.production.fact_all_platform_sales_report`

**D2C filter applied to all queries:**
```sql
Channel = 'Shopify'
AND SubChannel IN ('MyFrido', 'Mobility')
AND OrderId NOT LIKE '%_EX%'
AND DATE(OrderDate) >= '{start_date}'
AND DATE(OrderDate) < '{cutoff}'
```

**Date window:** Rolling 30 days, with 15-day lag for order maturity.

## Report sections

1. KPI Cards (Gross Rev, Net Rev, Cancel %, RTO %, CIR %, Exch %, ASP, AOV, COD vs Prepaid)
2. Monthly Trend (6-month rolling)
3. Revenue Leakage by Category
4. Top 20 Products by Revenue
5. Top Leakage Products (highest ₹ lost to returns)
6. Exchange Watch (size-chart issue signals)
7. Category-wise All Products
8. Definitions

## SQL queries (10 total)

| Query | Purpose |
|---|---|
| Q1 | Overall KPI Cards |
| Q2 | Leakage by Category |
| Q3 | Monthly Trend (6 months) |
| Q4 | COD vs Prepaid Split |
| Q5 | Top 20 Products by Revenue |
| Q6 | Top Leakage Products |
| Q7 | Exchange Watch SKUs |
| Q8 | All Products by Category |
| Q9 | COD Risk SKUs |
| Q10 | Raw order data (Excel export) |

## Documentation

See `DOCUMENTATION.docx` for full developer documentation: all SQL queries, KPI definitions, threshold rules, Excel schema, deployment steps, and architecture diagram.
