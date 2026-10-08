"""
Medicine Stock Review – Streamlit app
Upload the e-BMSIS "View Stock Balance" report (PDF printout or Excel/CSV export)
and get duplicates (same medicine + same dosage), expiry alerts and stock levels,
with PDF, Excel and interactive (HTML) reports to download.

Run locally:   streamlit run app.py
"""

import io
from datetime import date

import pandas as pd
import streamlit as st

import reports as rp
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
    st.caption("**Duplicate rule:** same medicine name **and** same dosage on more than one line. "
               "The same medicine with a different dosage (e.g. Amoxicillin 250mg and 500mg capsule) "
               "is counted as a separate item, not a duplicate.")

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
d = s["dup_counts"]
e = s["expiry_counts"]

# ----------------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------------
st.subheader("Summary" + (f" – {facility}" if facility else ""))

st.markdown("**Items**")
c = st.columns(4)
c[0].metric("Stock lines in report", s["lines"])
c[1].metric("Medicine names", s["medicine_names"], help="Each medicine name counted once, whatever its dosage.")
c[2].metric("Total items (medicine + dosage)", s["medicines"],
            help="Each dosage of a medicine is counted as its own item.")
c[3].metric("Medicines with more than one dosage", s["names_multi_dosage"],
            help="Counted separately – these are NOT duplicates. See the 'Same name, different dosage' tab.")

st.markdown("**Duplicates – same medicine + same dosage**")
c = st.columns(4)
c[0].metric("Items listed more than once", s["medicines_repeated"])
c[1].metric("Lines in these items", s["dup_lines"])
c[2].metric("Extra lines", s["extra_lines"], help="Lines beyond one line per item.")
c[3].metric("Exact duplicate entries (lines)", d[sc.EXACT_DUP], help=sc.FINDING_ACTION[sc.EXACT_DUP])
c = st.columns(4)
c[0].metric("Same batch, 2 supply types (lines)", d[sc.SAME_BATCH_TYPES], help=sc.FINDING_ACTION[sc.SAME_BATCH_TYPES])
c[1].metric("Same batch, different expiry (lines)", d[sc.EXPIRY_CONFLICT], help=sc.FINDING_ACTION[sc.EXPIRY_CONFLICT])
c[2].metric("Different batches – normal (lines)", d[sc.MULTI_BATCH], help=sc.FINDING_ACTION[sc.MULTI_BATCH])

st.markdown("**Expiry**")
c = st.columns(4)
c[0].metric("Expired lines", e[sc.EXPIRED])
c[1].metric("Expiring ≤3 months", e[sc.EXP_3M])
c[2].metric("Expiring ≤6 months", e[sc.EXP_6M])
c[3].metric("OK", e[sc.EXP_OK])

if not s["has_qty"]:
    st.warning("No **Balance Qty** in the upload, so stock levels (zero / low / overstock) were not checked. "
               "PDF printouts usually cut this column off – use the Excel export or print in landscape.")

# ----------------------------------------------------------------------------------
# Downloads
# ----------------------------------------------------------------------------------
st.markdown("**Download report**")
stem = f"Medicine_Stock_Review_{report_date:%Y-%m-%d}"
c = st.columns(3)
with st.spinner("Preparing reports…"):
    c[0].download_button("⬇️ PDF report", data=rp.pdf_report(res, report_date, facility),
                         file_name=f"{stem}.pdf", mime="application/pdf",
                         use_container_width=True, type="primary")
    c[1].download_button("⬇️ Excel report", data=rp.excel_report(res, report_date, facility),
                         file_name=f"{stem}.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         use_container_width=True)
    c[2].download_button("⬇️ Interactive report (HTML)", data=rp.html_report(res, report_date, facility),
                         file_name=f"{stem}.html", mime="text/html", use_container_width=True,
                         help="Opens in any web browser, also offline. Search, filter and sort every table.")

# ----------------------------------------------------------------------------------
# Detail tabs
# ----------------------------------------------------------------------------------
DATE = st.column_config.DateColumn("Expiry Date", format="DD/MM/YYYY")
COMMON = {
    "row_no": st.column_config.NumberColumn("Row #", format="%d"),
    "supply_type": "Supply Type", "item_name": "Item Name", "medicine": "Medicine", "dosage": "Dosage",
    "batch": "Batch No.", "expiry": DATE,
    "days_to_expiry": st.column_config.NumberColumn("Days Left", format="%d"),
    "expiry_status": "Expiry Status", "balance_qty": "Balance Qty",
    "duplicate_finding": "Finding", "group_id": "Group", "action": "What to do",
    "lines": "Lines", "extra_lines": "Extra Lines", "batches": "Batches", "supply_types": "Supply Types",
    "finding": "Main Finding", "expired_lines": "Expired Lines", "rows": "Rows", "duplicated": "Duplicated?",
}


def show(frame, cols, height=480):
    cols = [c for c in cols if c in frame.columns and not (c == "balance_qty" and not s["has_qty"])]
    st.dataframe(frame[cols], column_config=COMMON, hide_index=True, use_container_width=True, height=height)


tab_names = ["🔁 Duplicate summary", "🔎 Duplicate detail", "💊 Same name, different dosage",
             "⏰ Expiry alerts"] + (["📦 Stock level"] if s["has_qty"] else []) + ["📋 All items"]
tabs = st.tabs(tab_names)

with tabs[0]:
    st.caption("One row per **medicine + dosage** that is listed on more than one line.")
    ds = res["dup_summary"]
    pick = st.multiselect("Main finding", sc.FINDING_ORDER, default=sc.FINDING_ORDER, key="f_sum")
    show(ds[ds["finding"].isin(pick)], ["medicine", "dosage", "lines", "extra_lines", "batches",
                                        "supply_types", "finding", "expired_lines", "rows"])

with tabs[1]:
    dups = res["duplicates"]
    st.caption("Every line of the duplicated items, with what to do.")
    choice = st.multiselect("Show findings", sc.FINDING_ORDER, default=sc.FINDING_ORDER, key="f_det")
    # keep whole groups together: show a group if any of its lines has a chosen finding
    groups = dups.loc[dups["duplicate_finding"].isin(choice), "group_id"].unique()
    view = dups[dups["group_id"].isin(groups)]
    st.write(f"{view['group_id'].nunique()} items, {len(view)} lines")
    show(view, ["group_id", "medicine", "dosage", "row_no", "supply_type", "batch", "expiry", "expiry_status",
                "balance_qty", "duplicate_finding", "action"])

with tabs[2]:
    nd = res["name_dosage"]
    st.caption(f"{s['names_multi_dosage']} medicines are stocked in more than one dosage. Each dosage is "
               "counted as a separate item – these are **not** duplicates. "
               "'Duplicated?' shows whether that exact dosage is itself repeated.")
    q = st.text_input("Search medicine", key="q_nd")
    if q:
        nd = nd[nd["medicine"].str.contains(q, case=False, regex=False)]
    show(nd, ["medicine", "dosage", "lines", "duplicated", "rows"])

with tabs[3]:
    exp = res["expiry"]
    pick = st.multiselect("Show", [sc.EXPIRED, sc.EXP_3M, sc.EXP_6M], default=[sc.EXPIRED, sc.EXP_3M, sc.EXP_6M])
    show(exp[exp["expiry_status"].isin(pick)],
         ["row_no", "medicine", "dosage", "batch", "supply_type", "expiry", "days_to_expiry", "expiry_status",
          "balance_qty", "action"])

i = 4
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
                     column_config={"item_name": "Item (medicine + dosage)", "balance_qty": "Usable Balance Qty",
                                    "expired_qty": "Expired Qty", "monthly_use": "Monthly use",
                                    "months_of_stock": "Months of stock", "stock_status": "Stock status",
                                    "rows": "Rows"})
    i += 1

with tabs[i]:
    q = st.text_input("Search medicine, dosage or batch", key="q_all")
    view = res["all"]
    if q:
        view = view[view["item_name"].str.contains(q, case=False, regex=False)
                    | view["batch"].str.contains(q, case=False, regex=False)]
    show(view, ["row_no", "supply_type", "medicine", "dosage", "batch", "expiry", "days_to_expiry",
                "expiry_status", "balance_qty", "duplicate_finding"], height=560)
