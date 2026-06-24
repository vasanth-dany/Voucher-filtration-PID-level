"""
Voucher Filtration Automation — Shopee
=======================================
Run: streamlit run app.py

PID-level rule: if ANY SKU under a Product ID is ineligible
(Status ≠ YES, future launch date, AM excluded, or price below threshold),
ALL SKUs under that Product ID are blanked — even if other SKUs would
individually qualify.

Output: the original Shopee SellerPriceTemplate with ALU_NO / Live / Status /
RRP / SRP / RRP check / % / Exclusions + one YES/blank column per mechanic
appended directly onto the same sheet.
"""

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import warnings
from datetime import date, datetime
from collections import defaultdict

warnings.filterwarnings("ignore")

NA = "#N/A"
TODAY = date.today()

st.set_page_config(page_title="Shopee Voucher Filtration", page_icon="🟠", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    .stButton > button {
        background-color: #ee4d2d; color: white; font-weight: 700;
        border: none; padding: 0.6rem 2rem; border-radius: 6px;
        font-size: 16px; width: 100%;
    }
    .stButton > button:hover { background-color: #ff6347; }
</style>
""", unsafe_allow_html=True)

st.title("🟠 Shopee Voucher Filtration")
st.caption(f"PID-level campaign filtering  ·  Run date: {TODAY}")
st.divider()

tab_run, tab_clear = st.tabs(["▶  Run Filtration", "🗑️  Clear Output Columns"])

# ── SIDEBAR ────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    st.subheader("Eligibility Rules")
    st.caption(
        "**PID-level rule:** if ANY SKU in a Product ID group fails eligibility, "
        "ALL SKUs in that group are blanked."
    )
    apply_threshold = st.checkbox("Require RRP & SRP ≥ minimum price", value=False)
    min_price = st.number_input("Minimum price", value=39, min_value=0, step=1, disabled=not apply_threshold)

    st.divider()
    st.subheader("Campaign Type")
    st.caption("Selects the correct price columns from zeCOM automatically")

    CAMPAIGN_PRICE_COLS = {
        "BAU":               {"price": 54, "special_price": 55, "disc_pct": 56},
        "Mid Month / Payday":{"price": 57, "special_price": 58, "disc_pct": 59},
        "Double Digit":      {"price": 60, "special_price": 61, "disc_pct": 62},
        "Mega":              {"price": 63, "special_price": 64, "disc_pct": 65},
    }
    campaign_type = st.selectbox("Shopee campaign type", list(CAMPAIGN_PRICE_COLS.keys()), index=1)
    price_cols = CAMPAIGN_PRICE_COLS[campaign_type]

    st.divider()
    st.subheader("zeCOM Column Positions")
    st.caption("0-based index — update if the tracking file layout changes")
    col_style  = st.number_input("Style# (ALU_NO)",    value=2,  min_value=0)
    col_launch = st.number_input("Launch Date",         value=21, min_value=0)
    col_status = st.number_input("Status Shopee",       value=24, min_value=0)
    col_excl   = st.number_input("Exclusion",           value=66, min_value=0)

    st.caption(f"Price columns auto-set for **{campaign_type}**:")
    st.caption(f"RRP idx={price_cols['price']}  SRP idx={price_cols['special_price']}  Disc idx={price_cols['disc_pct']}")

    st.divider()
    st.subheader("Shopee File Column Names")
    st.caption("Match these to your Shopee SellerPriceTemplate row 1 headers exactly")
    sku_col_name   = st.text_input("SKU / EAN column name", value="SKU",
        help="Column containing the EAN/barcode used to look up ALU_NO")
    price_col_name = st.text_input("Price column name", value="Price")
    pid_col_name   = st.text_input("Product ID column name", value="Product ID",
        help="Column that groups variants under one product listing")
    data_start_row = st.number_input("Data starts at row", value=2, min_value=2)

ZECOM_COLS = {
    "style": int(col_style), "launch_date": int(col_launch),
    "status": int(col_status), "exclusion": int(col_excl),
    "price": price_cols["price"], "special_price": price_cols["special_price"],
    "disc_pct": price_cols["disc_pct"],
}


# ── HELPERS ────────────────────────────────────────────────────
def normalise_ean(v):
    return str(v).strip().split(".")[0].strip()


def matches_mechanic(excl_val, mech):
    el = str(excl_val).strip().lower()
    mv = [v.lower() for v in mech["match_values"]]
    ex = [v.lower() for v in mech.get("excludes", [])]
    matched = any(m in el for m in mv) if mech["match_type"] == "contains" else el in mv
    return matched and not any(x in el for x in ex if x)


def build_mechanics(raw):
    out = []
    for m in raw:
        name = m["name"].strip()
        mvs = [v.strip() for v in m["match_values"].splitlines() if v.strip()]
        if name and mvs:
            out.append({"name": name, "match_type": m["match_type"], "match_values": mvs,
                        "excludes": [v.strip() for v in m["excludes"].split(",") if v.strip()]})
    return out


def is_eligible_sku(status, live, alu_no, rrp, srp, am_set):
    if status != "YES":
        return False
    if isinstance(live, (date, datetime)):
        ld = live.date() if isinstance(live, datetime) else live
        if ld > TODAY:
            return False
    if alu_no and alu_no in am_set:
        return False
    if apply_threshold:
        for pv in (rrp, srp):
            if pv not in (NA, None):
                try:
                    if float(pv) < min_price:
                        return False
                except (TypeError, ValueError):
                    pass
    return True


def mechanic_ui(state_key):
    if state_key not in st.session_state:
        st.session_state[state_key] = [
            {"name": "", "match_type": "contains", "match_values": "", "excludes": ""}
        ]
    _remove = None
    for i, mech in enumerate(st.session_state[state_key]):
        with st.container(border=True):
            c1, c2, c3 = st.columns([4, 2, 1])
            mech["name"] = c1.text_input("Column header name", value=mech["name"],
                key=f"{state_key}_name_{i}", placeholder="e.g. 20% NMS (All 20% VC Remark)")
            mech["match_type"] = c2.selectbox("Match type", ["contains", "exact"],
                index=["contains","exact"].index(mech["match_type"]),
                key=f"{state_key}_type_{i}")
            c3.markdown("&nbsp;")
            if c3.button("🗑️", key=f"{state_key}_rm_{i}"):
                _remove = i
            mech["match_values"] = st.text_area("zeCOM Exclusion value(s) — one per line",
                value=mech["match_values"], key=f"{state_key}_mv_{i}", height=68,
                placeholder="Open for all")
            mech["excludes"] = st.text_input(
                "Exclude if text also contains (comma-separated, optional)",
                value=mech["excludes"], key=f"{state_key}_ex_{i}",
                placeholder="e.g. No platform VC")
    if _remove is not None:
        st.session_state[state_key].pop(_remove)
        st.rerun()
    if st.button("➕ Add mechanic", key=f"{state_key}_add"):
        st.session_state[state_key].append(
            {"name": "", "match_type": "contains", "match_values": "", "excludes": ""})
        st.rerun()


# ── TAB 1: RUN FILTRATION ──────────────────────────────────────
with tab_run:
    st.subheader("🎯 Voucher Mechanics")
    st.caption("Define one entry per campaign mechanic — names and match values change every campaign.")
    mechanic_ui("sh_mechanics")

    st.divider()
    st.subheader("📂 Upload Input Files")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**1. Price / Stock File** *(required)*")
        price_file = st.file_uploader("Shopee SellerPriceTemplate", type=["xlsx","xls"], key="price")
    with c2:
        st.markdown("**2. Content File** *(required)*")
        content_file = st.file_uploader("Content_file", type=["xlsx","xls"], key="content")
    with c3:
        st.markdown("**3. zeCOM Tracking File** *(required)*")
        zecom_file = st.file_uploader("zeCOM_Tracking_File", type=["xlsx","xls"], key="zecom")
    c4, c5 = st.columns(2)
    with c4:
        st.markdown("**4. AM Exclusion File** *(optional)*")
        am_file = st.file_uploader("AM_Exclusion.xlsx", type=["xlsx","xls"], key="am")
    with c5:
        st.markdown("**Output filename** *(optional)*")
        out_name = st.text_input("out", value="", label_visibility="collapsed",
            placeholder="e.g. Shopee_June_Payday_filtration.xlsx", key="out_name")

    st.divider()

    # PID rule explainer
    with st.expander("ℹ️ How the PID-level rule works"):
        st.markdown("""
**Pass 1 — per-SKU eligibility check:**
Each SKU is checked individually: Status = YES, Live date ≤ today, not AM excluded, (optionally) RRP & SRP ≥ minimum.

**Pass 2 — PID-level gate:**
All SKUs are grouped by Product ID. If **any** SKU in a group failed Pass 1,
the **entire group** is marked ineligible and all mechanic columns are left blank.

This ensures Shopee's product-level voucher logic: vouchers apply to a whole
product listing, so a product with even one ineligible variant cannot participate.
        """)

    mechanics = build_mechanics(st.session_state.get("sh_mechanics", []))
    ready = price_file and content_file and zecom_file and mechanics

    if not ready:
        st.info("👆 Upload 3 required files and define at least one mechanic to get started")

    if st.button("▶  Run Shopee Filtration", disabled=not ready):
        with st.spinner("Processing (PID-level)..."):
            try:
                logs = []
                def log(m): logs.append(m)

                # Load Content map
                content_file.seek(0)
                cwb = openpyxl.load_workbook(content_file, read_only=True, data_only=True)
                if "content" not in cwb.sheetnames:
                    raise ValueError("Content file must have a sheet named 'content'.")
                cws = cwb["content"]
                ch = [c.value for c in next(cws.iter_rows(min_row=1, max_row=1))]
                ean_i   = [str(h).strip().lower() for h in ch].index("ean")
                color_i = [str(h).strip().lower() for h in ch].index("color_no")
                content_map = {}
                for row in cws.iter_rows(min_row=2, values_only=True):
                    if row[ean_i] is not None:
                        content_map[normalise_ean(row[ean_i])] = str(row[color_i]).strip()
                cwb.close()
                log(f"✓ {len(content_map):,} EAN → ALU_NO mappings")

                # Load zeCOM map
                zecom_file.seek(0)
                zwb = openpyxl.load_workbook(zecom_file, read_only=True, data_only=True)
                if "MY" not in zwb.sheetnames:
                    raise ValueError("zeCOM file must have a sheet named 'MY'.")
                zws = zwb["MY"]
                zecom_map = {}
                for row in zws.iter_rows(min_row=5, values_only=True):
                    style = row[ZECOM_COLS["style"]] if len(row) > ZECOM_COLS["style"] else None
                    if not style: continue
                    def g(k, row=row):
                        i = ZECOM_COLS.get(k)
                        return row[i] if i is not None and len(row) > i else None
                    zecom_map[str(style).strip()] = {
                        "Launch_Date": g("launch_date"), "Status": g("status"),
                        "Price": g("price"), "Special_Price": g("special_price"),
                        "Disc_Pct": g("disc_pct"), "Exclusion": g("exclusion"),
                    }
                zwb.close()
                log(f"✓ {len(zecom_map):,} styles from zeCOM ({campaign_type})")

                # Load AM exclusion
                am_set = set()
                if am_file:
                    am_file.seek(0)
                    awb = openpyxl.load_workbook(am_file, read_only=True, data_only=True)
                    for row in awb.active.iter_rows(min_row=2, values_only=True):
                        if row and row[0]: am_set.add(str(row[0]).strip())
                    awb.close()
                log(f"✓ {len(am_set):,} AM-excluded ALU_NOs")

                # Load Price/Stock workbook
                price_file.seek(0)
                wb = openpyxl.load_workbook(price_file)
                ws = wb[wb.sheetnames[0]]
                headers = [c.value for c in ws[1]]

                def fc(name):
                    for idx, h in enumerate(headers, 1):
                        if h and str(h).strip().lower() == name.lower(): return idx
                    return None

                sku_col   = fc(sku_col_name)
                price_col = fc(price_col_name)
                pid_col   = fc(pid_col_name)
                if not sku_col:  raise ValueError(f"Column '{sku_col_name}' not found in row 1 — check the 'SKU / EAN column name' in the sidebar")
                if not price_col: raise ValueError(f"Column '{price_col_name}' not found in row 1 — check the 'Price column name' in the sidebar")
                if not pid_col:
                    raise ValueError(
                        f"Column '{pid_col_name}' not found in row 1. "
                        f"Check the 'Product ID column name' setting in the sidebar."
                    )
                last_col = max((i for i,h in enumerate(headers,1) if h), default=len(headers))
                log(f"✓ Sheet='{ws.title}' | PID col='{pid_col_name}' (col {pid_col})")
                log(f"  Row 1 headers: {[h for h in headers if h]}")

                # ── PASS 1: per-SKU eligibility ──────────────────
                pass1 = []  # (r, pid, alu_no, row_vals, excl_v, elig)
                total = matched_alu = matched_zecom = 0

                for r in range(int(data_start_row), ws.max_row + 1):
                    sku = ws.cell(row=r, column=sku_col).value
                    if sku is None: continue
                    total += 1
                    orig_price = ws.cell(row=r, column=price_col).value
                    pid = ws.cell(row=r, column=pid_col).value
                    alu_no = content_map.get(normalise_ean(sku))

                    if alu_no is None:
                        row_vals = [NA] * 8
                    else:
                        matched_alu += 1
                        rec = zecom_map.get(alu_no)
                        if rec is None:
                            row_vals = [alu_no] + [NA] * 7
                        else:
                            matched_zecom += 1
                            live, status = rec["Launch_Date"], rec["Status"]
                            rrp, srp, pct, excl = rec["Price"], rec["Special_Price"], rec["Disc_Pct"], rec["Exclusion"]
                            try:
                                orig_p = float(str(orig_price).replace("'","").strip()) if orig_price is not None else None
                                rrp_check = round(float(rrp),2) == round(orig_p,2) if orig_p is not None else NA
                            except: rrp_check = NA
                            row_vals = [alu_no, live, status, rrp, srp, rrp_check, pct, excl]

                    status_v, live_v, rrp_v, srp_v, excl_v = row_vals[2], row_vals[1], row_vals[3], row_vals[4], row_vals[7]
                    # ineligible if lookup failed OR individual eligibility fails
                    elig = (excl_v != NA) and is_eligible_sku(status_v, live_v, alu_no, rrp_v, srp_v, am_set)
                    pass1.append((r, pid, alu_no, row_vals, excl_v, elig))

                # ── PASS 2: PID-level gate ────────────────────────
                pid_elig = defaultdict(lambda: True)
                for _, pid, _, _, _, elig in pass1:
                    if not elig:
                        pid_elig[pid] = False

                pid_total    = len(pid_elig)
                pid_ok       = sum(1 for v in pid_elig.values() if v)
                pid_blocked  = pid_total - pid_ok
                log(f"✓ {total:,} SKUs | {matched_alu:,} ALU_NO | {matched_zecom:,} zeCOM")
                log(f"✓ {pid_total:,} Product IDs | {pid_ok:,} eligible | {pid_blocked:,} blocked by PID rule")

                # ── Write output ──────────────────────────────────
                N_LOOKUP = 8
                new_headers = ["ALU_NO","Live","Status","RRP","SRP","RRP check","%","Exclusions"] + [m["name"] for m in mechanics]
                base_font   = Font(name="Aptos Narrow", size=11, bold=False)
                yellow_fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")

                for i, h in enumerate(new_headers):
                    col  = last_col + 1 + i
                    cell = ws.cell(row=1, column=col, value=h)
                    cell.font = base_font
                    if i >= N_LOOKUP: cell.fill = yellow_fill
                    ws.column_dimensions[get_column_letter(col)].width = max(
                        ws.column_dimensions[get_column_letter(col)].width or 0, 14)

                counts = {m["name"]: 0 for m in mechanics}

                for r, pid, alu_no, row_vals, excl_v, sku_elig in pass1:
                    final_elig = sku_elig and pid_elig.get(pid, True)
                    mech_vals = []
                    for m in mechanics:
                        if not final_elig:
                            mech_vals.append("")
                        else:
                            yes = matches_mechanic(excl_v, m)
                            if yes: counts[m["name"]] += 1
                            mech_vals.append("YES" if yes else "")

                    for i, val in enumerate(row_vals + mech_vals):
                        col  = last_col + 1 + i
                        cell = ws.cell(row=r, column=col, value=val)
                        cell.font = base_font
                        if i == 1 and isinstance(val, (date, datetime)):
                            cell.number_format = "mm-dd-yy"

                for n, c in counts.items(): log(f"  {n}: {c:,} YES")

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                st.success("✅ Shopee filtration complete!")
                mc = st.columns(4 + len(mechanics))
                with mc[0]: st.metric("Total SKUs", f"{total:,}")
                with mc[1]: st.metric("ALU_NO matched", f"{matched_alu:,}")
                with mc[2]: st.metric("PIDs eligible", f"{pid_ok:,}")
                with mc[3]: st.metric("PIDs blocked", f"{pid_blocked:,}")
                for i,(n,c) in enumerate(counts.items()):
                    with mc[4+i]: st.metric(n[:22], f"{c:,}")
                with st.expander("📝 Processing Log"):
                    for l in logs: st.text(l)

                fname = out_name.strip() or price_file.name.rsplit(".",1)[0] + "_shopee_filtration.xlsx"
                st.download_button("⬇️  Download Shopee Output", buf, fname,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            except Exception as e:
                st.error(f"❌ {e}"); st.exception(e)


# ── TAB 2: CLEAR OUTPUT COLUMNS ────────────────────────────────
with tab_clear:
    st.subheader("🗑️ Clear Output Columns from Processed File")
    st.caption("Strips ALU_NO and all appended columns, returning the file to the original Shopee SellerPriceTemplate.")
    st.divider()

    clear_file = st.file_uploader("Upload processed Excel to clear", type=["xlsx","xls"], key="clear_upload")

    if clear_file:
        clear_file.seek(0)
        try:
            wb_p = openpyxl.load_workbook(clear_file, read_only=True)
            ws_p = wb_p[wb_p.sheetnames[0]]
            row1 = [ws_p.cell(row=1, column=c).value for c in range(1, ws_p.max_column + 1)]
            wb_p.close()

            strip_from = next(
                (i+1 for i,h in enumerate(row1) if h and str(h).strip().upper() == "ALU_NO"), None)

            if strip_from:
                n_strip = len(row1) - strip_from + 1
                stripped = [str(row1[c-1]) for c in range(strip_from, len(row1)+1)]
                st.success(f"✅ Found ALU_NO at column **{get_column_letter(strip_from)}** — will remove **{n_strip}** column(s).")
                with st.expander("Columns to be removed"):
                    for name in stripped: st.markdown(f"- `{name}`")
                st.divider()
                if st.button("🗑️  Strip & prepare download", use_container_width=True):
                    with st.spinner("Removing columns..."):
                        clear_file.seek(0)
                        wb_out = openpyxl.load_workbook(clear_file)
                        wb_out[wb_out.sheetnames[0]].delete_cols(strip_from, n_strip)
                        buf = io.BytesIO(); wb_out.save(buf); buf.seek(0)
                    clean_name = clear_file.name.replace("_shopee_filtration","").rsplit(".",1)[0] + "_clean.xlsx"
                    st.download_button("⬇️  Download Clean File", buf, clean_name,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                    st.success(f"✅ Done — {n_strip} column(s) removed.")
            else:
                st.warning("⚠️ No ALU_NO column found — file may already be clean.")
        except Exception as e:
            st.error(f"Could not open file: {e}")
