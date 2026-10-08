"""
Medicine Stock Review – Streamlit app
Upload the e-BMSIS "View Stock Balance" report (PDF printout or Excel/CSV export)
and get duplicates, expiry alerts and stock levels, with an Excel report to download.

Run locally:   streamlit run app.py
"""

import io
from datetime import date

import pandas as pd
import streamlit as st

import stock_core as sc

st.set_page_config(page_title="Medicine Stock Review", page_icon="💊", layout="wide")


# ----------------------------------------------------------------------------------
# Cached readers (so changing a setting doesn't re-read the PDFs)
# ----------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def read_pdf(name: str, data: bytes) -> pd.DataFrame:
    return sc.parse_pdf(data, source=name)


@st.cache_data(show_spinner=False)
def read_table(name: str, data: bytes) -> pd.DataFrame:
    if name.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(data))
    return sc._read_excel_find_header(data)


FIELD_LABELS = {
    "item_name": "Item Name *",
    "expiry": "Expiry Date *",
    "batch": "Batch No.",
    "supply_type": "Supply Type",
    "balance_qty": "Balance Qty",
    "monthly_use": "Average monthly use (AMC)",
    "row_no": "Row # (optional)",
}

# ----------------------------------------------------------------------------------
# Sidebar
# ----------------------------------------------------------------------------------
with st.sidebar:
    st.header("Settings")
    facility = st.text_input("Facility name", value="")
    report_date = st.date_input("Report date (expiry is checked against this)", value=date.today(),
                                format="DD/MM/YYYY")
    st.markdown("**Stock level limits** (used only when a monthly-use column is uploaded)")
    low_months = st.number_input("Low stock if less than … months of stock", 0.0, 24.0, 1.0, 0.5)
    over_months = st.number_input("Overstock if more than … months of stock", 1.0, 60.0, 12.0, 1.0)
    st.divider()
    st.caption("Duplicate rule: judged on **Item Name only**. Same name + same dosage = duplicate. "
               "Same name with a different dosage = a separate medicine.")

# ----------------------------------------------------------------------------------
# Upload
# ----------------------------------------------------------------------------------
st.title("💊 Medicine Stock Review")
st.write("Upload the e-BMSIS **View Stock Balance** report. You can upload all PDF pages at once "
         "(e.g. 1–100, 101–200 …) or a single Excel/CSV export.")

files = st.file_uploader("Stock balance file(s)", type=["pdf", "xlsx", "xls", "csv"],
                         accept_multiple_files=True)
if not files:
    st.info("Tip: an Excel export (or a PDF printed in landscape, 'fit to page') that includes "
            "**Balance Qty** lets the app also check zero, low and overstock.")
    st.stop()

frames, problems = [], []
with st.spinner("Reading files…"):
    for f in files:
        data = f.getvalue()
        try:
            if f.name.lower().endswith(".pdf"):
                frames.append(read_pdf(f.name, data))
                continue
            raw = read_table(f.name, data)
            auto = sc.map_columns(raw)
            options = ["— not in file —"] + [str(c) for c in raw.columns]
            with st.expander(f"Columns in {f.name}", expanded=not {"item_name", "expiry"} <= set(auto)):
                st.caption("Check the app picked the right columns. * = required.")
                mapping = {}
                cols = st.columns(4)
                for i, (field, label) in enumerate(FIELD_LABELS.items()):
                    current = str(auto[field]) if field in auto else options[0]
                    pick = cols[i % 4].selectbox(label, options, index=options.index(current),
                                                 key=f"{f.name}-{field}")
                    if pick != options[0]:
                        mapping[field] = raw.columns[options.index(pick) - 1]
            frames.append(sc.standardise_table(raw, source=f.name, mapping=mapping))
        except Exception as e:  # show the problem but keep going with the other files
            problems.append(f"**{f.name}**: {e}")

for p in problems:
    st.error(p)
if not frames:
    st.stop()

df = sc.combine(frames)
res = sc.analyse(df, report_date, low_months, over_months)
s = res["summary"]

# ----------------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------------
st.subheader("Summary" + (f" – {facility}" if facility else ""))
c = st.columns(5)
c[0].metric("Stock lines", s["lines"])
c[1].metric("Distinct medicines", s["medicines"])
c[2].metric("Medicines listed more than once", s["medicines_repeated"])
c[3].metric("Expired lines", s["expiry_counts"][sc.EXPIRED])
c[4].metric("Expiring ≤3 months", s["expiry_counts"][sc.EXP_3M])

d = s["dup_counts"]
problem_lines = d[sc.EXACT_DUP] + d[sc.SAME_BATCH_TYPES] + d[sc.EXPIRY_CONFLICT]
c = st.columns(4)
c[0].metric("Exact duplicate entries", d[sc.EXACT_DUP], help=sc.FINDING_ACTION[sc.EXACT_DUP])
c[1].metric("Same batch, 2 supply types", d[sc.SAME_BATCH_TYPES], help=sc.FINDING_ACTION[sc.SAME_BATCH_TYPES])
c[2].metric("Same batch, different expiry", d[sc.EXPIRY_CONFLICT], help=sc.FINDING_ACTION[sc.EXPIRY_CONFLICT])
c[3].metric("Different batches (normal)", d[sc.MULTI_BATCH], help=sc.FINDING_ACTION[sc.MULTI_BATCH])

if not s["has_qty"]:
    st.warning("No **Balance Qty** in the upload, so stock levels (zero / low / overstock) were not checked. "
               "PDF printouts usually cut this column off – use the Excel export or print in landscape.")

st.download_button(
    "⬇️ Download Excel report",
    data=sc.excel_report(res, report_date, facility),
    file_name=f"Medicine_Stock_Review_{report_date:%Y-%m-%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# ----------------------------------------------------------------------------------
# Detail tabs
# ----------------------------------------------------------------------------------
DATE = st.column_config.DateColumn("Expiry Date", format="DD/MM/YYYY")
COMMON = {
    "row_no": st.column_config.NumberColumn("Row #", format="%d"),
    "supply_type": "Supply Type", "item_name": "Item Name", "batch": "Batch No.", "expiry": DATE,
    "days_to_expiry": st.column_config.NumberColumn("Days to Expiry", format="%d"),
    "expiry_status": "Expiry Status", "balance_qty": "Balance Qty",
    "duplicate_finding": "Duplicate Finding", "group_id": "Group", "action": "What to do",
}


def show(frame, cols, height=480):
    cols = [c for c in cols if c in frame.columns and not (c == "balance_qty" and not s["has_qty"])]
    st.dataframe(frame[cols], column_config=COMMON, hide_index=True, use_container_width=True, height=height)


tab_names = ["🔁 Duplicates", "⏰ Expiry alerts"] + (["📦 Stock level"] if s["has_qty"] else []) + ["📋 All items"]
tabs = st.tabs(tab_names)

with tabs[0]:
    dups = res["duplicates"]
    st.caption("Medicines whose **Item Name** (name + dosage) appears on more than one line.")
    choice = st.multiselect("Show findings", sc.FINDING_ORDER,
                            default=[f for f in sc.FINDING_ORDER if f != sc.MULTI_BATCH] if problem_lines
                            else sc.FINDING_ORDER)
    # keep whole groups together: show a group if any of its lines has a chosen finding
    groups = dups.loc[dups["duplicate_finding"].isin(choice), "group_id"].unique()
    view = dups[dups["group_id"].isin(groups)]
    st.write(f"{view['group_id'].nunique()} medicines, {len(view)} lines")
    show(view, ["group_id", "item_name", "row_no", "supply_type", "batch", "expiry", "expiry_status",
                "balance_qty", "duplicate_finding", "action"])

with tabs[1]:
    exp = res["expiry"]
    pick = st.multiselect("Show", [sc.EXPIRED, sc.EXP_3M, sc.EXP_6M], default=[sc.EXPIRED, sc.EXP_3M, sc.EXP_6M])
    show(exp[exp["expiry_status"].isin(pick)],
         ["row_no", "item_name", "batch", "supply_type", "expiry", "days_to_expiry", "expiry_status",
          "balance_qty", "action"])

i = 2
if s["has_qty"]:
    with tabs[i]:
        stock = res["stock"].copy()
        stock["stock_status"] = stock["stock_status"].str[2:]
        st.caption("Quantities of the same medicine + dosage are added across batches "
                   "(expired stock and exact duplicate lines are not counted as usable).")
        if "monthly_use" not in stock:
            st.info("Add an average monthly use (AMC) column to the upload to flag low and overstock. "
                    "Without it, only zero stock is flagged.")
        pick = st.multiselect("Show", sorted(stock["stock_status"].unique()),
                              default=[x for x in sorted(stock["stock_status"].unique()) if x != "OK"])
        st.dataframe(stock[stock["stock_status"].isin(pick)], hide_index=True, use_container_width=True,
                     column_config={"item_name": "Medicine (name + dosage)", "balance_qty": "Usable Balance Qty",
                                    "expired_qty": "Expired Qty", "monthly_use": "Monthly use",
                                    "months_of_stock": "Months of stock", "stock_status": "Stock status",
                                    "rows": "Rows"})
    i += 1

with tabs[i]:
    q = st.text_input("Search item name or batch")
    view = res["all"]
    if q:
        view = view[view["item_name"].str.contains(q, case=False, regex=False)
                    | view["batch"].str.contains(q, case=False, regex=False)]
    show(view, ["row_no", "supply_type", "item_name", "batch", "expiry", "days_to_expiry",
                "expiry_status", "balance_qty", "duplicate_finding"], height=560)
