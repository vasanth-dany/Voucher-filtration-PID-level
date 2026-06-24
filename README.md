# Shopee Voucher Filtration Automation

Automated voucher eligibility filtering for Shopee campaigns.

**PID-level rule:** if ANY SKU under a Product ID is ineligible, ALL SKUs under that Product ID are blanked — even if other variants would individually qualify.

**Output format:** your Shopee SellerPriceTemplate with `ALU_NO / Live / Status / RRP / SRP / RRP check / % / Exclusions` plus one YES/blank column per mechanic appended directly onto the same sheet.

---

## Run the Dashboard

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Every Campaign Run

### Step 1 — Select campaign type in sidebar
BAU / Mid Month Payday / Double Digit / Mega — this auto-sets the correct zeCOM price columns.

### Step 2 — Define mechanics
Enter the mechanic names and the matching zeCOM Exclusion values (check column AZ of your zeCOM MY sheet).

### Step 3 — Upload files
1. Shopee SellerPriceTemplate
2. Content file
3. zeCOM Tracking File
4. AM Exclusion file *(optional)*

### Step 4 — Run & download

---

## zeCOM Column Reference (Shopee)

| Field | Column | Index | Notes |
|---|---|---|---|
| Style# (ALU_NO) | C | 2 | |
| Launch Date | V | 21 | Shared with Lazada |
| Status Shopee | Y | 24 | YES / NO |
| Exclusion | AZ | 66 | Shopee-specific |

**Price columns by campaign type:**

| Campaign | RRP idx | SRP idx | Disc idx |
|---|---|---|---|
| BAU | 54 | 55 | 56 |
| Mid Month / Payday | 57 | 58 | 59 |
| Double Digit | 60 | 61 | 62 |
| Mega | 63 | 64 | 65 |

---

## Eligibility Rules

| Rule | Notes |
|---|---|
| Status Shopee = YES | gates every mechanic |
| Live (Launch Date) ≤ today | future-dated launch = ineligible |
| ALU_NO not in AM Exclusion | optional file |
| **PID-level rule** | if any SKU in a Product ID fails → whole PID blanked |
| Exclusion text matches mechanic | contains / exact + excludes |
| *(optional)* RRP & SRP ≥ minimum | toggle in sidebar, default OFF |
