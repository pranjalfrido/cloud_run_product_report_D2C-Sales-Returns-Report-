"""
Generate DOCUMENTATION.docx for cloud_run_product_report.
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import os

OUTPUT = r"c:\Users\PranjalTripati\OneDrive - Arcatron Mobility Pvt Ltd\Desktop\frido\MIS\cloud_run_product_report\DOCUMENTATION.docx"

DARK_BLUE = RGBColor(31, 56, 100)
LIGHT_GREY = RGBColor(242, 242, 242)
ALT_ROW = RGBColor(235, 241, 250)
WHITE = RGBColor(255, 255, 255)
RED = RGBColor(192, 0, 0)

# ── helpers ────────────────────────────────────────────────────────────────────

def set_cell_bg(cell, rgb: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    hex_val = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_val)
    tcPr.append(shd)


def set_cell_borders(cell, border_color="C0C0C0"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def make_table(doc, headers, rows, col_widths=None):
    """Create a styled table. headers = list of str, rows = list of lists."""
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # header row
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_bg(cell, DARK_BLUE)
        set_cell_borders(cell, "FFFFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(str(h))
        run.bold = True
        run.font.color.rgb = WHITE
        run.font.size = Pt(9)
        run.font.name = 'Calibri'
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # data rows
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        bg = ALT_ROW if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            set_cell_bg(cell, bg)
            set_cell_borders(cell, "CCCCCC")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val) if val is not None else '')
            run.font.size = Pt(9)
            run.font.name = 'Calibri'
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # column widths
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)

    return table


def add_code_block(doc, code: str):
    """Add a paragraph styled as a code block."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.2)

    # grey background via paragraph shading
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)

    run = p.add_run(code)
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(40, 40, 40)
    return p


def add_bullet(doc, text: str, level: int = 0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.name = 'Calibri'
    return p


def add_numbered(doc, text: str, level: int = 0):
    p = doc.add_paragraph(style='List Number')
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.name = 'Calibri'
    return p


def add_h1(doc, text):
    p = doc.add_heading(text, level=1)
    for run in p.runs:
        run.font.name = 'Calibri'
    return p


def add_h2(doc, text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        run.font.name = 'Calibri'
    return p


def add_h3(doc, text):
    p = doc.add_heading(text, level=3)
    for run in p.runs:
        run.font.name = 'Calibri'
    return p


def add_body(doc, text: str):
    p = doc.add_paragraph(text)
    p.style = 'Normal'
    for run in p.runs:
        run.font.size = Pt(10)
        run.font.name = 'Calibri'
    return p


def page_break(doc):
    doc.add_page_break()


# ── document setup ──────────────────────────────────────────────────────────────

def setup_document():
    doc = Document()

    # 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # default font
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(10)

    return doc


# ── title page ──────────────────────────────────────────────────────────────────

def add_title_page(doc):
    # spacer
    for _ in range(6):
        doc.add_paragraph()

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('cloud_run_product_report')
    run.bold = True
    run.font.size = Pt(28)
    run.font.color.rgb = DARK_BLUE
    run.font.name = 'Calibri'

    # Subtitle
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run('D2C Sales & Returns Report — Developer Documentation')
    run2.font.size = Pt(16)
    run2.font.color.rgb = RGBColor(68, 68, 68)
    run2.font.name = 'Calibri'

    # Date
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = p3.add_run('October 2026')
    run3.font.size = Pt(13)
    run3.font.color.rgb = RGBColor(100, 100, 100)
    run3.font.name = 'Calibri'
    run3.italic = True

    # spacer + horizontal rule via bottom border on a paragraph
    for _ in range(4):
        doc.add_paragraph()

    # Divider line
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p_line._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    hex_blue = f"{DARK_BLUE[0]:02X}{DARK_BLUE[1]:02X}{DARK_BLUE[2]:02X}"
    bottom.set(qn('w:color'), hex_blue)
    pBdr.append(bottom)
    pPr.append(pBdr)

    doc.add_page_break()


# ── section 1 ───────────────────────────────────────────────────────────────────

def add_section1(doc):
    add_h1(doc, 'Section 1 — Project Overview')

    add_body(doc,
        'This is a scheduled Google Cloud Run microservice that generates and emails a daily '
        'D2C (Direct-to-Consumer) Sales & Returns Report for Frido. It focuses on revenue '
        'leakage analysis — quantifying how much revenue is lost to cancellations, RTOs, '
        'CIR (Customer Initiated Returns), and exchanges across Frido\'s two D2C Shopify '
        'channels: MyFrido and Mobility.'
    )

    add_h2(doc, 'Why It Exists')
    add_body(doc,
        'The D2C team needs a daily consolidated view of sales performance, return rates by '
        'product/category, and risk signals (COD-heavy SKUs, high-exchange products). Without '
        'this, identifying leakage patterns would require manual BigQuery queries.'
    )

    add_h2(doc, 'What It Outputs')
    add_bullet(doc, 'A responsive HTML email with 8 data sections covering 30+ KPIs')
    add_bullet(doc, 'An Excel attachment (Frido_D2C_Orders_{start}_{end}.xlsx) with full raw order-level data (21 columns)')

    add_h2(doc, 'Trigger')
    add_body(doc,
        'Google Cloud Scheduler fires daily (configured in deploy.sh). The Cloud Run HTTP '
        'endpoint receives the request, runs the full pipeline, and returns 200 OK.'
    )

    add_h2(doc, 'Recipients')
    add_body(doc, 'Configured via EMAIL_TO / EMAIL_CC environment variables.')


# ── section 2 ───────────────────────────────────────────────────────────────────

ASCII_FLOW = """\
┌─────────────────────────┐
│   Cloud Scheduler       │
│   Daily trigger         │
└────────────┬────────────┘
             │  HTTP GET (OIDC)
             ▼
┌─────────────────────────┐
│   Cloud Run (main.py)   │
│   Python 3.11-slim      │
└────────────┬────────────┘
             │  10 SQL queries
             ▼
┌──────────────────────────────────────────┐
│   BigQuery (frido-429506)                │
│   production.fact_all_platform_sales_    │
│   report                                 │
│   Filter: Channel='Shopify'              │
│           SubChannel IN ('MyFrido',      │
│           'Mobility')                    │
│           OrderId NOT LIKE '%_EX%'       │
└────────────┬─────────────────────────────┘
             │  30+ KPIs + raw order data
             ▼
┌─────────────────────────┐
│   Build HTML Report     │
│   + Excel Attachment    │
│   premailer inlines CSS │
│   openpyxl builds xlsx  │
└────────────┬────────────┘
             │  OAuth token
             ├──────────────► Azure AD (OTP_FRIDO tenant)
             │  access_token
             ◄──────────────
             │  POST /sendMail
             ▼
┌─────────────────────────┐
│   Microsoft Graph API   │
│   data.team@myfrido.com │
└────────────┬────────────┘
             │
             ▼
┌───────────────────────────────────────┐
│   Stakeholders — HTML Email           │
│   + Excel attachment                  │
└───────────────────────────────────────┘"""


def add_section2(doc):
    add_h1(doc, 'Section 2 — Architecture & Data Flow')

    add_body(doc,
        'Cloud Scheduler fires a GET to Cloud Run. main.py loads GCP credentials, runs 10 '
        'BigQuery queries against the D2C sales fact table, builds an HTML email with 8 '
        'sections and a full raw-data Excel file, obtains an OAuth token from Azure AD, and '
        'sends both via Microsoft Graph API.'
    )

    add_h2(doc, 'System Flow Diagram')
    add_code_block(doc, ASCII_FLOW)


# ── section 3 ───────────────────────────────────────────────────────────────────

def add_section3(doc):
    add_h1(doc, 'Section 3 — File Structure')

    add_body(doc, 'The project contains the following files:')

    headers = ['File', 'Role', 'Notes']
    rows = [
        ['main.py', 'Core application',
         'All data fetching, KPI computation, HTML generation, Excel creation, and email dispatch. Production entry point.'],
        ['send_test.py', 'Test sender',
         'Same queries as main.py but sends HTML-only email (no Excel) to a hardcoded test address. Faster for iterating on HTML layout.'],
        ['preview.py', 'Local preview',
         'Runs all queries and writes preview.html locally — no email sent. Requires sa_key.json in parent directory.'],
        ['preview.html', 'Generated preview',
         'Output of preview.py. Open in browser to check report rendering before deploying.'],
        ['requirements.txt', 'Python dependencies',
         'google-cloud-bigquery, db-dtypes, requests, premailer, google-auth, openpyxl'],
        ['Dockerfile', 'Container definition',
         'python:3.11-slim, runs main.py as entrypoint'],
        ['env_flags.yaml', 'Cloud Run deploy config',
         'Contains --set-env-vars block for gcloud deploy command. Never commit real credentials.'],
    ]
    make_table(doc, headers, rows, col_widths=[1.4, 1.4, 3.6])

    add_h2(doc, 'Entry Point')
    add_body(doc,
        'main.py runs as a standalone Python script (CMD ["python", "-u", "main.py"] in Dockerfile). '
        'It runs the full pipeline on startup and exits. Cloud Run restarts it on each scheduler trigger.'
    )


# ── section 4 ───────────────────────────────────────────────────────────────────

def add_section4(doc):
    add_h1(doc, 'Section 4 — Configuration & Environment Variables')

    headers = ['Variable', 'Source', 'Purpose']
    rows = [
        ['GCP_CREDENTIALS_B64', 'Secret Manager',
         'Base64-encoded GCP service account JSON. Decoded at runtime, written to /tmp/gcp_creds.json for BigQuery auth.'],
        ['OTP_FRIDO_TENANT_ID', 'env_flags.yaml / Secret',
         'Azure AD tenant ID (7e5ccb3a-...) for OAuth token requests.'],
        ['OTP_FRIDO_CLIENT_ID', 'env_flags.yaml / Secret',
         'Azure AD application (client) ID.'],
        ['OTP_FRIDO_CLIENT_SECRET', 'env_flags.yaml / Secret',
         'Azure AD client secret. Rotates ~12 months — update when it expires.'],
        ['FROM_EMAIL', 'env_flags.yaml',
         'Sender address (data.team@myfrido.com). Must have Mail.Send permission in Azure AD.'],
        ['EMAIL_TO', 'env_flags.yaml',
         'Comma-separated TO recipients.'],
        ['EMAIL_CC', 'env_flags.yaml',
         'Comma-separated CC recipients (can be empty).'],
    ]
    make_table(doc, headers, rows, col_widths=[1.8, 1.6, 3.0])

    add_h2(doc, 'How Credentials Load')
    add_body(doc,
        'main.py checks for GCP_CREDENTIALS_B64. If present, base64-decodes it, writes to '
        '/tmp/gcp_creds.json, sets GOOGLE_APPLICATION_CREDENTIALS env var, and initialises '
        'the BigQuery client. Falls back to Application Default Credentials if absent '
        '(useful on GCP-managed infra).'
    )


# ── section 5 ───────────────────────────────────────────────────────────────────

D2C_FILTER_CODE = """\
Channel = 'Shopify'
AND SubChannel IN ('MyFrido', 'Mobility')
AND OrderId NOT LIKE '%_EX%'
AND DATE(OrderDate) >= '{start_date}'
AND DATE(OrderDate) < '{cutoff}'"""


def add_section5(doc):
    add_h1(doc, 'Section 5 — BigQuery Table & Schema')

    add_body(doc,
        'Single primary table: frido-429506.production.fact_all_platform_sales_report\n'
        'All 10 queries read from this one table. No joins to other tables.'
    )

    add_h2(doc, 'Column Reference')
    headers = ['Column', 'Type', 'Meaning']
    rows = [
        ['OrderId', 'STRING',
         'Unique order identifier. Exchange replacement orders have suffix _EX and are excluded by the D2C filter.'],
        ['OrderDate', 'TIMESTAMP',
         'Order creation datetime. All date filtering is on DATE(OrderDate).'],
        ['Dispatch_Date', 'DATE', 'Date order was dispatched from warehouse.'],
        ['Delivered_Date', 'DATE', 'Date order was delivered to customer.'],
        ['Cancelled_At', 'DATE', 'Date order was cancelled (if applicable).'],
        ['Channel', 'STRING',
         "Always 'Shopify' for D2C orders. Other values: Amazon, Retail, offline_sales, etc."],
        ['SubChannel', 'STRING',
         "'MyFrido' or 'Mobility' for D2C. Filter is SubChannel IN ('MyFrido', 'Mobility')."],
        ['payment_type', 'STRING', "'COD' or prepaid values. Code normalises to COD vs Prepaid."],
        ['State', 'STRING', "Customer's state."],
        ['City', 'STRING', "Customer's city."],
        ['Pincode', 'STRING', "Customer's pincode."],
        ['masterskucode', 'STRING', 'Product SKU code. NULL rows excluded from product-level queries.'],
        ['SubCategory', 'STRING',
         'Product name / subcategory. Used as the display name (falls back to masterskucode if NULL).'],
        ['Category', 'STRING',
         "Product category. 'Covers', 'Frido', 'Spare Parts' are grouped into 'Others' in reports."],
        ['ItemQty', 'INTEGER', 'Quantity of units in this line item.'],
        ['SellingPrice_Inc_GST', 'FLOAT', 'Revenue including 18% GST.'],
        ['SellingPrice_Exc_GST', 'FLOAT', 'Revenue excluding GST (= net revenue when not returned/cancelled).'],
        ['Order_Status', 'STRING',
         'Fulfillment/return status. Key values: NULL (delivered), Cancelled, RTO, Return, CIR, Exchange.'],
        ['OrderTrackingStatus', 'STRING', 'Carrier tracking status.'],
        ['Customer_Return_Reason', 'STRING', 'High-level return reason given by customer.'],
        ['Customer_Sub_Reason', 'STRING', 'Granular return sub-reason.'],
    ]
    make_table(doc, headers, rows, col_widths=[1.8, 0.8, 3.8])

    add_h2(doc, 'Order_Status Values Explained')
    headers2 = ['Value', 'Meaning', 'Impact on Revenue']
    rows2 = [
        ['NULL (or non-return status)', 'Delivered successfully', 'Counts as Net Revenue'],
        ['Cancelled', 'Order cancelled before dispatch', 'Revenue lost — Cancel Rev'],
        ['RTO', 'Return to Origin — courier returned parcel', 'Revenue lost — RTO Rev'],
        ['Return', 'Customer-initiated return via courier', 'Revenue lost — RTO Rev (grouped with RTO)'],
        ['CIR', 'Customer Initiated Return (direct)', 'Revenue lost — CIR Rev'],
        ['Exchange', 'Customer exchanged for different product', 'Revenue at risk — Exch Rev'],
    ]
    make_table(doc, headers2, rows2, col_widths=[1.8, 2.2, 2.4])

    add_h2(doc, 'D2C Filter (used in all queries)')
    add_code_block(doc, D2C_FILTER_CODE)
    add_body(doc,
        'The OrderId NOT LIKE \'%_EX%\' clause excludes exchange replacement shipments to '
        'avoid double-counting revenue.'
    )

    add_h2(doc, 'Date Window Logic')
    add_bullet(doc, 'today = date.today()')
    add_bullet(doc, 'cutoff = today - 15 days  (15-day lag ensures orders have mature status)')
    add_bullet(doc, 'start_date = cutoff - 30 days')
    add_bullet(doc, 'reporting period = start_date to cutoff - 1 day  (30 calendar days)')
    add_bullet(doc, 'Monthly trend: last 6 months from cutoff')
    add_body(doc, 'Example: job fires 2026-10-04 → window = 2026-08-20 to 2026-09-18.')


# ── section 6 ───────────────────────────────────────────────────────────────────

def add_section6(doc):
    add_h1(doc, 'Section 6 — SQL Queries (All 10)')

    add_body(doc,
        'All queries run against frido-429506.production.fact_all_platform_sales_report '
        'using the D2C_FILTER. Run sequentially via client.query(...).result().'
    )

    # Q1
    add_h2(doc, 'Q1 — Overall KPI Cards')
    add_body(doc, 'Purpose: Compute all headline metrics shown in the KPI cards at the top of the report.')
    add_body(doc, 'Filters: D2C_FILTER (Shopify, MyFrido+Mobility, 30-day window)')
    headers = ['Metric', 'Formula']
    rows = [
        ['gross_inc_gst', 'SUM(SellingPrice_Inc_GST) — total gross revenue incl GST'],
        ['cancel_rev', "SUM(SellingPrice_Inc_GST) WHERE Order_Status = 'Cancelled'"],
        ['rto_rev', "SUM(SellingPrice_Inc_GST) WHERE Order_Status IN ('RTO','Return')"],
        ['cir_rev', "SUM(SellingPrice_Inc_GST) WHERE Order_Status = 'CIR'"],
        ['exch_rev', "SUM(SellingPrice_Inc_GST) WHERE Order_Status = 'Exchange'"],
        ['net_exc_gst', 'SUM(SellingPrice_Exc_GST) excluding Cancelled/CIR/RTO/Return rows'],
        ['net_inc_gst', 'SUM(SellingPrice_Inc_GST) excluding Cancelled/CIR/RTO/Return rows'],
        ['net_units', 'SUM(ItemQty) excluding Cancelled/CIR/RTO/Return rows'],
        ['orders', 'COUNT(DISTINCT OrderId)'],
        ['gross_units', 'SUM(ItemQty)'],
    ]
    make_table(doc, headers, rows, col_widths=[1.6, 4.8])
    add_body(doc,
        'Derived in Python: ASP = gross_inc_gst / gross_units, '
        'AOV = gross_inc_gst / orders, '
        'returns_pct = (cancel+rto+cir) / gross_inc_gst × 100'
    )

    # Q2
    add_h2(doc, 'Q2 — Leakage Funnel by Category')
    add_body(doc, "Purpose: Show which product categories have the highest return rates and revenue leakage.")
    add_body(doc, "Grouping: Category (with 'Covers', 'Frido', 'Spare Parts' → 'Others')")
    add_body(doc, 'Order: ORDER BY gross_rev DESC')
    add_body(doc, 'Metrics: category, gross_rev, gross_units, cancel_rev, rto_rev, cir_rev, exch_rev')
    add_body(doc, 'Derived in Python: cancel_pct, rto_pct, cir_pct, exch_pct, total_ret_pct, retained_rev, retained_pct per category')

    # Q3
    add_h2(doc, 'Q3 — Monthly Trend')
    add_body(doc, 'Purpose: 6-month rolling trend to spot improvement or regression in returns.')
    add_body(doc, 'Filters: Channel=\'Shopify\', SubChannel IN (\'MyFrido\',\'Mobility\'), OrderId NOT LIKE \'%_EX%\', last 6 months from cutoff')
    add_body(doc, 'Grouping: DATE_TRUNC(DATE(OrderDate), MONTH) — one row per calendar month')
    add_body(doc, 'Metrics per month: month_label, gross_rev, gross_units, cancel_rev, rto_rev, cir_rev, exch_rev')
    add_body(doc, 'Derived: ASP, cancel_pct, rto_pct, cir_pct, exch_pct, total_ret_pct, retained_rev, retained_pct')

    # Q4
    add_h2(doc, 'Q4 — COD vs Prepaid Split')
    add_body(doc, 'Purpose: Compare cancellation, RTO, and CIR patterns between COD and Prepaid orders.')
    add_body(doc, "Grouping: CASE WHEN payment_type = 'COD' THEN 'COD' ELSE 'Prepaid' END")
    add_body(doc, 'Metrics: pay_mode, qty, orders, gross_rev, cancel_rev, rto_rev, cir_rev, net_rev')
    add_body(doc, 'Derived: AOV, cancel_pct, rto_pct, cir_pct, retained_pct per payment mode')
    add_body(doc, 'Note: COD orders typically show higher RTO rates; this section quantifies the gap.')

    # Q5
    add_h2(doc, 'Q5 — Top 20 Products by Revenue')
    add_body(doc, 'Purpose: Revenue-ranked best-sellers with their return breakdown.')
    add_body(doc, 'Filters: D2C_FILTER + masterskucode IS NOT NULL + HAVING gross_units >= 3')
    add_body(doc, 'Order: ORDER BY gross_rev DESC LIMIT 20')
    add_body(doc, 'Display name: COALESCE(SubCategory, masterskucode)')
    add_body(doc, 'Metrics: product, category, gross_rev, cancel_rev, rto_rev, cir_rev, exch_rev, net_exc_gst, gross_units')
    add_body(doc, 'Highlighting: Return % bolded red if > overall returns_pct × 1.5')

    # Q6
    add_h2(doc, 'Q6 — Top Leakage Products')
    add_body(doc, 'Purpose: Products losing the most absolute ₹ to returns — quality/fulfilment concern signals.')
    add_body(doc, 'Filters: D2C_FILTER + masterskucode IS NOT NULL + HAVING gross_units >= 20')
    add_body(doc, 'Order: ORDER BY (cancel_rev + rto_rev + cir_rev) DESC LIMIT 15')
    add_body(doc, 'Same metrics as Q5 but ranked by total leakage ₹ rather than gross revenue.')
    add_body(doc, 'Note: Minimum 20 units ensures statistical significance.')

    # Q7
    add_h2(doc, 'Q7 — Exchange Watch SKUs')
    add_body(doc, 'Purpose: Products with high exchange rates — often indicates size chart issues or product description mismatch.')
    add_body(doc, 'Filters: D2C_FILTER + masterskucode IS NOT NULL + HAVING exch_rev > 0')
    add_body(doc, 'Order: ORDER BY exch_rev DESC LIMIT 20')
    add_body(doc, 'Extra metric: exch_pct = exch_rev / gross_rev × 100')
    add_body(doc, 'Metrics: product, category, gross_units, gross_rev, exch_rev, exch_pct, cancel_rev, rto_rev, cir_rev')

    # Q8
    add_h2(doc, 'Q8 — All Products by Category')
    add_body(doc, 'Purpose: Complete product-level drill-down grouped by category for detailed analysis.')
    add_body(doc, 'Filters: D2C_FILTER + masterskucode IS NOT NULL + HAVING gross_units >= 1')
    add_body(doc, 'Grouping: category first, then product within category')
    add_body(doc, 'Order: ORDER BY category, gross_rev DESC')
    add_body(doc, 'This powers Section 7 (Category-wise All Products) in the HTML report.')
    add_body(doc, 'Highlighting: % columns bolded if > 1.2× the category average for that metric.')

    # Q9
    add_h2(doc, 'Q9 — COD Risk SKUs')
    add_body(doc, 'Purpose: Identify products with both high COD share AND high RTO rate — highest expected revenue loss.')
    add_h3(doc, 'Thresholds Applied in SQL')
    add_bullet(doc, 'cod_pct > 30%  (majority of orders are COD)')
    add_bullet(doc, 'rto_pct > 8%  (elevated return rate)')
    add_bullet(doc, 'gross_units >= 20  (statistical significance)')
    add_body(doc, 'Risk formula: rev_at_risk = (cod_pct/100) × (rto_pct/100) × gross_rev')
    add_body(doc, '  This estimates the expected revenue at risk if the COD+RTO trend continues.')
    add_body(doc, 'Order: ORDER BY rev_at_risk DESC')
    add_body(doc, 'Metrics: product, category, gross_rev, cod_rev, gross_units, cod_pct, rto_pct, rev_at_risk')

    # Q10
    add_h2(doc, 'Q10 — Raw D2C Order Data (for Excel)')
    add_body(doc, 'Purpose: Transactional-level export for offline analysis. This powers the Excel attachment.')
    add_body(doc, 'Filters: D2C_FILTER')
    add_body(doc, 'Order: ORDER BY OrderDate DESC, OrderId')
    add_body(doc,
        'All 21 columns fetched: OrderId, OrderDate, Dispatch_Date, Delivered_Date, Cancelled_At, '
        'Channel, SubChannel, payment_type, State, City, Pincode, masterskucode, SubCategory, '
        'Category, ItemQty, SellingPrice_Inc_GST, SellingPrice_Exc_GST, Order_Status, '
        'OrderTrackingStatus, Customer_Return_Reason, Customer_Sub_Reason'
    )


# ── section 7 ───────────────────────────────────────────────────────────────────

HTML_GEN_STEPS = [
    'All 10 BigQuery queries run and results stored in Python dicts/lists.',
    'Python string templates build each section as HTML tables and div blocks.',
    "Threshold checks apply inline style='...' with red/bold formatting on individual <td> cells.",
    'All sections concatenated into a full <!DOCTYPE html> document with embedded <style> block.',
    'premailer.transform(html) inlines every CSS rule into style="" attributes.',
    'openpyxl builds the Excel attachment in write-only mode (memory efficient).',
    'Excel bytes are base64-encoded and embedded in the Graph API payload as a fileAttachment.',
    'Email sent via Microsoft Graph API POST /sendMail.',
]


def add_section7(doc):
    add_h1(doc, 'Section 7 — Report HTML Structure')

    add_body(doc,
        'Single self-contained HTML email. CSS inlined by premailer for Gmail/Outlook compatibility. '
        'Visual theme: Dark blue header (#1F3864), white body, max-width 680px, mobile-responsive.'
    )

    add_h2(doc, 'Email Sections')
    headers = ['Section', 'Data Source', 'What It Shows']
    rows = [
        ['Header', 'Computed dates',
         '"Frido D2C Sales & Returns Report", period label, "D2C: MyFrido + Mobility"'],
        ['KPI Cards', 'Q1',
         '4 groups of KPI cards: Revenue (Gross/Net incl & excl GST), Loss % (Cancel/RTO/CIR/Exch), Summary (Total Ret%, Retained%, ASP, AOV), COD vs Prepaid side-by-side'],
        ['Monthly Trend', 'Q3',
         '6-month table: Month, Gross Rev, Qty, ASP, Cancel%, RTO%, CIR%, Exch%, Tot Ret ₹, Tot Ret%, Retained ₹, Retained%'],
        ['Revenue Leakage by Category', 'Q2',
         'Category-level leakage: same columns as monthly trend. Retained% highlighted red if below category average.'],
        ['Top 20 Products', 'Q5',
         'Revenue-ranked top sellers with return breakdown. Tot Ret% bolded red if > 1.5× overall threshold.'],
        ['Top Leakage Products', 'Q6',
         'Products losing most ₹ to returns (min 20 units). Sorted by absolute leakage ₹.'],
        ['Exchange Watch', 'Q7',
         'Top 20 by exchange revenue. Callout note about size-chart signals.'],
        ['Category-wise All Products', 'Q8',
         'One subsection per category; all products ≥1 unit. % columns bolded if > 1.2× category avg.'],
        ['Definitions', 'Static',
         'Explains all metric abbreviations (Gross Rev, Net Retained, Tot Ret%, ASP, AOV, etc.)'],
        ['Footer', 'Static',
         'Dark blue footer with report title, date range, and D2C definition'],
    ]
    make_table(doc, headers, rows, col_widths=[1.8, 1.2, 3.4])

    add_h2(doc, 'HTML Generation Steps')
    for step in HTML_GEN_STEPS:
        add_numbered(doc, step)


# ── section 8 ───────────────────────────────────────────────────────────────────

THRESHOLD_CODE = """\
ret_threshold = (returns_pct or 0) * 1.5
if ret_p > ret_threshold:
    style = 'font-weight:bold;color:#c00'"""


def add_section8(doc):
    add_h1(doc, 'Section 8 — Threshold Rules & Highlighting')

    headers = ['Rule', 'Condition', 'Visual Effect', 'Where Applied']
    rows = [
        ['Overall Return %', '> 15%', 'Red text on KPI card', 'Overall KPI section'],
        ['Product Return %', '> overall_ret_pct × 1.5', 'Bold red text', 'Top 20 Products, Top Leakage Products tables'],
        ['Category/Product % columns', '> category_avg × 1.2', 'Bold text', 'Category-wise All Products table'],
        ['Retained %', '< overall avg retained %', 'Bold red', 'Revenue Leakage by Category table'],
        ['COD Risk SKU', 'cod_pct > 30% AND rto_pct > 8% AND units ≥ 20', 'Included in COD Risk section', 'COD Risk SKUs (Q9)'],
    ]
    make_table(doc, headers, rows, col_widths=[1.5, 1.8, 1.5, 1.7])

    add_h2(doc, 'How to Change a Threshold')
    add_body(doc, 'Search main.py for the current value. Example:')
    add_code_block(doc, THRESHOLD_CODE)


# ── section 9 ───────────────────────────────────────────────────────────────────

def add_section9(doc):
    add_h1(doc, 'Section 9 — Excel Attachment')

    add_body(doc, 'The email includes a .xlsx attachment with full raw order data for the 30-day period.')

    add_h2(doc, 'File Details')
    add_bullet(doc, 'File name: Frido_D2C_Orders_{start_date}_{cutoff}.xlsx')
    add_bullet(doc, 'Sheet name: D2C Orders')
    add_bullet(doc, 'Library: openpyxl in write_only=True mode (memory efficient for large datasets)')

    add_h2(doc, '21-Column Schema')
    headers = ['Column', 'Source Field', 'Notes']
    rows = [
        ['Order ID', 'OrderId', 'Unique identifier'],
        ['Order Date', 'OrderDate', 'Order creation datetime'],
        ['Dispatch Date', 'Dispatch_Date', 'When dispatched'],
        ['Delivered Date', 'Delivered_Date', 'When delivered'],
        ['Cancelled At', 'Cancelled_At', 'Cancellation date if applicable'],
        ['Channel', 'Channel', "Always 'Shopify' for D2C"],
        ['Sub Channel', 'SubChannel', 'MyFrido or Mobility'],
        ['Payment Type', 'payment_type', 'COD or Prepaid'],
        ['State', 'State', 'Customer state'],
        ['City', 'City', 'Customer city'],
        ['Pincode', 'Pincode', 'Customer pincode'],
        ['SKU Code', 'masterskucode', 'Product SKU'],
        ['Product', 'SubCategory', 'Product name'],
        ['Category', 'Category', 'Product category'],
        ['Qty', 'ItemQty', 'Units ordered'],
        ['Selling Price (incl GST)', 'SellingPrice_Inc_GST', 'Revenue incl 18% GST'],
        ['Selling Price (excl GST)', 'SellingPrice_Exc_GST', 'Revenue excl GST'],
        ['Order Status', 'Order_Status', 'Delivered/Cancelled/RTO/CIR/Exchange'],
        ['Tracking Status', 'OrderTrackingStatus', 'Carrier tracking status'],
        ['Return Reason', 'Customer_Return_Reason', 'High-level return reason'],
        ['Return Sub Reason', 'Customer_Sub_Reason', 'Granular return sub-reason'],
    ]
    make_table(doc, headers, rows, col_widths=[1.8, 1.7, 2.9])

    add_h2(doc, 'Styling')
    add_body(doc,
        'Header row dark blue (#1F3864) with white bold text, size 10. '
        'Numeric columns right-aligned, text columns left-aligned.'
    )


# ── section 10 ──────────────────────────────────────────────────────────────────

OAUTH_STEPS = [
    'POST to https://login.microsoftonline.com/{OTP_FRIDO_TENANT_ID}/oauth2/v2.0/token',
    'Body: grant_type=client_credentials, client_id, client_secret, scope=https://graph.microsoft.com/.default',
    'Response: access_token (valid ~60 minutes)',
    'POST to https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/sendMail with Authorization: Bearer {token}',
    'Payload includes HTML body + base64-encoded Excel attachment',
    'saveToSentItems: false — email does not appear in sender\'s Sent folder',
]


def add_section10(doc):
    add_h1(doc, 'Section 10 — Email Delivery (Microsoft Graph API)')

    add_body(doc, 'Uses Microsoft Graph API (not SMTP) to avoid GCP firewall restrictions.')

    add_h2(doc, 'OAuth Flow')
    for step in OAUTH_STEPS:
        add_numbered(doc, step)

    add_h2(doc, 'Email Subject Format')
    add_body(doc, 'Subject: "Frido D2C Sales & Returns Report | {date_label}"')
    add_body(doc, 'Example: "Frido D2C Sales & Returns Report | 20 Aug – 18 Sep 2026"')

    add_h2(doc, 'Common Failure Modes')
    headers = ['Symptom', 'Cause', 'Fix']
    rows = [
        ['401 Unauthorized', 'OTP_FRIDO_CLIENT_SECRET expired',
         'Rotate in Azure AD, update Secret Manager'],
        ['403 Forbidden', 'App missing Mail.Send permission',
         'Add Mail.Send application permission in Azure AD'],
        ['400 Bad Request', 'Invalid recipient address',
         'Check EMAIL_TO/EMAIL_CC for stray spaces'],
        ['Large attachment rejected', 'Excel file too large (>4MB limit)',
         'Reduce date window or filter columns'],
        ['Email to spam', 'Sender domain not trusted',
         'Mark data.team@myfrido.com as safe sender'],
    ]
    make_table(doc, headers, rows, col_widths=[1.8, 2.2, 2.4])


# ── section 11 ──────────────────────────────────────────────────────────────────

DEPLOY_CMD = 'bash deploy.sh'

DEPLOY_STEPS = [
    'docker build -t gcr.io/frido-429506/product-report .',
    'docker push gcr.io/frido-429506/product-report',
    'gcloud run deploy product-report with: --region asia-south1, --memory 1Gi, --timeout 540, --no-allow-unauthenticated, --set-secrets, --set-env-vars from env_flags.yaml',
    'Creates/updates Cloud Scheduler job for daily trigger',
]

PREREQS = [
    'gcloud CLI authenticated with Cloud Run Admin, Storage Admin, Secret Manager Admin on frido-429506',
    'Docker daemon running locally',
    'Secrets created in Secret Manager: OTP_FRIDO_TENANT_ID, OTP_FRIDO_CLIENT_ID, OTP_FRIDO_CLIENT_SECRET, GCP_CREDENTIALS_B64',
]


def add_section11(doc):
    add_h1(doc, 'Section 11 — Deployment')

    add_h2(doc, 'Prerequisites')
    for p in PREREQS:
        add_bullet(doc, p)

    add_h2(doc, 'Deploy Command')
    add_code_block(doc, DEPLOY_CMD)

    add_h2(doc, 'deploy.sh Steps')
    for step in DEPLOY_STEPS:
        add_numbered(doc, step)

    add_h2(doc, 'Cloud Run Configuration')
    headers = ['Setting', 'Value']
    rows = [
        ['Image', 'gcr.io/frido-429506/product-report'],
        ['Region', 'asia-south1 (Mumbai)'],
        ['Memory', '1 GB'],
        ['Timeout', '540 seconds'],
        ['Min instances', '0 (scales to zero)'],
        ['Authentication', 'OIDC (no public access)'],
    ]
    make_table(doc, headers, rows, col_widths=[2.0, 4.4])


# ── section 12 ──────────────────────────────────────────────────────────────────

PREVIEW_CODE = 'python preview.py'
SEND_TEST_CODE = '# Set env vars first (from env_flags.yaml values)\npython send_test.py'
MAIN_CODE = 'python main.py'
DATE_OVERRIDE_CODE = """\
today = date(2026, 9, 30)   # replace date.today()
cutoff = today - timedelta(days=15)
start_date = cutoff - timedelta(days=30)"""


def add_section12(doc):
    add_h1(doc, 'Section 12 — Local Development')

    add_body(doc, 'Three ways to test locally:')

    add_h2(doc, '1. preview.py — HTML Only, No Email')
    add_body(doc, 'Requires: sa_key.json in parent directory (c:\\...\\frido\\MIS\\sa_key.json)')
    add_code_block(doc, PREVIEW_CODE)
    add_body(doc, 'Output: preview.html written to same directory. Open in browser to check layout.')

    add_h2(doc, '2. send_test.py — Email Only, No Excel')
    add_body(doc, 'Faster than main.py — skips Excel generation.')
    add_body(doc, 'Hardcoded TO: pranjal.t@myfrido.com, suraj.s@myfrido.com')
    add_body(doc, 'Requires env vars set (OTP_FRIDO_* and GCP_CREDENTIALS_B64):')
    add_code_block(doc, SEND_TEST_CODE)

    add_h2(doc, '3. Full Local Run of main.py')
    add_body(doc, 'Set all env vars from env_flags.yaml, then:')
    add_code_block(doc, MAIN_CODE)

    add_h2(doc, 'Changing the Test Date')
    add_body(doc, 'The date window is computed dynamically. To test a specific period, temporarily hardcode in the script:')
    add_code_block(doc, DATE_OVERRIDE_CODE)


# ── section 13 ──────────────────────────────────────────────────────────────────

def add_section13(doc):
    add_h1(doc, 'Section 13 — Error Handling & Monitoring')

    add_h2(doc, 'What Surfaces as Errors')
    headers = ['Failure', 'How It Surfaces']
    rows = [
        ['BigQuery auth failure', 'Exception on client init, Cloud Run logs, job fails in Scheduler'],
        ['BigQuery query error', 'Exception with BQ error message, job fails'],
        ['Graph API 4xx/5xx', 'resp.raise_for_status() raises exception, job fails'],
        ['Azure AD token failure', 'HTTP error on token POST, job fails'],
        ['openpyxl write error', 'Exception during Excel build, job fails'],
        ['premailer CSS error', 'Exception during HTML transform, job fails'],
    ]
    make_table(doc, headers, rows, col_widths=[2.2, 4.2])

    add_h2(doc, 'What Fails Silently')
    silents = [
        'Invalid CC email address: Graph API may silently skip. Check EMAIL_CC carefully.',
        'NULL SubCategory: Falls back to masterskucode for display. No crash.',
        "NULL Category: Shows as 'Others'. Grouped correctly.",
        'Orders with NULL Order_Status: Treated as delivered/retained (correct business logic — NULL = successfully delivered).',
        'masterskucode IS NULL rows: Excluded from product-level queries but included in KPI totals.',
    ]
    for s in silents:
        add_bullet(doc, s)

    add_h2(doc, 'Monitoring')
    monitoring = [
        'Cloud Scheduler: Job history in GCP Console shows success/failure per run.',
        'Cloud Run logs: Script prints "Running D2C report for {date}...", "All queries done", "HTML built", "Email sent (HTTP 202)".',
        'No alerting configured. Add Cloud Monitoring alert on Cloud Scheduler failure or Cloud Run error log count > 0 for production monitoring.',
    ]
    for m in monitoring:
        add_bullet(doc, m)


# ── section 14 ──────────────────────────────────────────────────────────────────

def add_section14(doc):
    add_h1(doc, 'Section 14 — KPI Glossary')

    headers = ['Abbreviation', 'Full Name', 'Definition']
    rows = [
        ['Gross Rev', 'Gross Revenue',
         'Total revenue (incl GST) across all orders regardless of status'],
        ['Net Rev', 'Net Revenue',
         'Revenue from successfully delivered orders only (excl Cancelled, CIR, RTO, Return)'],
        ['Cancel Rev', 'Cancellation Revenue',
         'Revenue lost to orders cancelled before dispatch'],
        ['RTO Rev', 'Return to Origin Revenue',
         'Revenue lost to courier-returned parcels (RTO + Return status)'],
        ['CIR Rev', 'Customer Initiated Return Revenue',
         'Revenue lost to direct customer returns'],
        ['Exch Rev', 'Exchange Revenue',
         'Revenue from orders placed as exchanges (at risk — may become CIR)'],
        ['Tot Ret %', 'Total Return %',
         '(Cancel + RTO + CIR) / Gross Rev × 100. Red if > 15%.'],
        ['Retained %', 'Retained Revenue %',
         'Net Rev / Gross Rev × 100. Inverse of Tot Ret%.'],
        ['ASP', 'Average Selling Price',
         'Gross Rev / Gross Units'],
        ['AOV', 'Average Order Value',
         'Gross Rev / Order Count'],
        ['COD', 'Cash on Delivery',
         'Customer pays at door. Higher RTO risk than Prepaid.'],
        ['Prepaid', 'Prepaid Order',
         'Customer paid online. Lower RTO and cancellation risk.'],
        ['D2C', 'Direct to Consumer',
         "Sales via Frido's own Shopify channels (MyFrido + Mobility)"],
        ['MyFrido', 'MyFrido SubChannel',
         "Frido's primary D2C website (myfrido.com)"],
        ['Mobility', 'Mobility SubChannel',
         "Frido's mobility/healthcare product channel"],
        ['RTO', 'Return to Origin',
         'Courier returns parcel to sender (customer absent, refused, undeliverable)'],
        ['CIR', 'Customer Initiated Return',
         'Customer requests return directly (not via courier failure)'],
        ['Exchange', 'Exchange Order',
         'Customer swaps product for a different variant/size'],
        ['_EX suffix', 'Exchange Replacement',
         'OrderIds ending in _EX are replacement shipments for exchanges. Excluded from revenue calculations to avoid double-counting.'],
        ['rev_at_risk', 'Revenue at Risk',
         'COD Risk metric: (cod_pct/100) × (rto_pct/100) × gross_rev. Expected revenue loss if COD+RTO trend continues.'],
        ['ret_threshold', 'Return Threshold',
         'overall_ret_pct × 1.5. Products above this are flagged red in tables.'],
    ]
    make_table(doc, headers, rows, col_widths=[1.3, 1.7, 3.4])


# ── main ────────────────────────────────────────────────────────────────────────

def main():
    doc = setup_document()
    add_title_page(doc)

    add_section1(doc)
    add_section2(doc)
    add_section3(doc)
    add_section4(doc)
    add_section5(doc)
    add_section6(doc)
    add_section7(doc)
    add_section8(doc)
    add_section9(doc)
    add_section10(doc)
    add_section11(doc)
    add_section12(doc)
    add_section13(doc)
    add_section14(doc)

    doc.save(OUTPUT)
    size = os.path.getsize(OUTPUT)
    print(f"Saved: {OUTPUT}")
    print(f"File size: {size:,} bytes ({size/1024:.1f} KB)")


if __name__ == '__main__':
    main()
