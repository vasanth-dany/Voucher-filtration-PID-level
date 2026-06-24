# ============================================================
# VOUCHER FILTRATION — SHOPEE CAMPAIGN CONFIG
# Used by voucher_filtration.py (CLI script).
# The Streamlit dashboard (app.py) defines mechanics in its own UI
# and does NOT read this file — this is for CLI workflow only.
# ============================================================

# ------------------------------------------------------------
# 1. FILE PATHS — update every campaign run
# ------------------------------------------------------------
PRICE_FILE   = 'inputs/Shopee_SellerPriceTemplate.xlsx'
CONTENT_FILE = 'inputs/Content_file.xlsx'
ZECOM_FILE   = 'inputs/zeCOM_Tracking_File.xlsx'
AM_EXCL_FILE = None    # e.g. 'inputs/AM_Exclusion.xlsx'

OUTPUT_FILE  = 'outputs/Shopee_Voucher_Filtration.xlsx'

# ------------------------------------------------------------
# 2. SHOPEE FILE SETTINGS
# ------------------------------------------------------------
PRODUCT_ID_COL_NAME = 'Product ID'   # col 1 — groups all variants of a product
SKU_COL_NAME        = 'SKU'          # col 6 — contains the EAN barcode
PRICE_COL_NAME      = 'Price'        # col 7
DATA_START_ROW      = 2              # headers at row 1, data from row 2

# ------------------------------------------------------------
# 3. CAMPAIGN TYPE — determines which zeCOM price columns to use
#    Options: 'BAU' | 'Mid Month / Payday' | 'Double Digit' | 'Mega'
# ------------------------------------------------------------
CAMPAIGN_TYPE = 'Mid Month / Payday'

CAMPAIGN_PRICE_COLS = {
    'BAU':               {'price': 54, 'special_price': 55, 'disc_pct': 56},
    'Mid Month / Payday':{'price': 57, 'special_price': 58, 'disc_pct': 59},
    'Double Digit':      {'price': 60, 'special_price': 61, 'disc_pct': 62},
    'Mega':              {'price': 63, 'special_price': 64, 'disc_pct': 65},
}

# ------------------------------------------------------------
# 4. ZECOM COLUMN POSITIONS (0-based index)
#    Update if the zeCOM tracking file layout changes
# ------------------------------------------------------------
_price_cols = CAMPAIGN_PRICE_COLS[CAMPAIGN_TYPE]

ZECOM_COLS = {
    'style':         2,   # C  — Style# = ALU_NO
    'launch_date':  21,   # V  — Shopee & Lazada Launch Dates -> output "Live"
    'status':       24,   # Y  — Shopee (YES/NO) -> output "Status"
    'price':        _price_cols['price'],           # MY RRP
    'special_price':_price_cols['special_price'],   # MY EC SRP
    'disc_pct':     _price_cols['disc_pct'],        # DISC %
    'exclusion':    66,   # Shopee Exclusion column -> output "Exclusions"
}

# ------------------------------------------------------------
# 5. ELIGIBILITY RULES
# ------------------------------------------------------------
# A mechanic YES requires ALL of:
#   - Status (zeCOM Shopee) == 'YES'
#   - Live (Launch Date) <= today
#   - ALU_NO not in AM Exclusion file
#   - PID-level rule: ALL SKUs under the same Product ID must pass above
#   - Exclusion text matches the mechanic's rule
#
# Price threshold is OFF by default — enable and set minimum below if needed.
MIN_PRICE_THRESHOLD       = 39
APPLY_MIN_PRICE_THRESHOLD = False

# ------------------------------------------------------------
# 6. CAMPAIGN MECHANICS — update every campaign run
#    match_type: 'contains' or 'exact'
#    match_values: list of zeCOM Exclusion column values to match
#    excludes: substrings that block a match even if match_values matched
# ------------------------------------------------------------
CAMPAIGNS = [

    {
        'name'        : '20% NMS (All 20% VC Remark)',
        'match_type'  : 'contains',
        'match_values': ['20% VC'],
        'excludes'    : [],
    },

    {
        'name'        : '35% NMS (open for all + OPEN for 10days up to 50%)',
        'match_type'  : 'exact',
        'match_values': ['Open for all', 'OPEN for 10days up to 50%'],
        'excludes'    : [],
    },

    # ── TEMPLATE: copy-paste to add a new mechanic ──────────
    # {
    #     'name'        : 'NEW MECHANIC NAME',
    #     'match_type'  : 'contains',   # or 'exact'
    #     'match_values': ['EXCLUSION VALUE'],
    #     'excludes'    : [],
    # },

]
