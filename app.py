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
import ui

st.set_page_config(page_title="Medicine Stock Review", page_icon="💊", layout="wide")
st.markdown(ui.CSS, unsafe_allow_html=True)


def html(s: str):
    st.markdown(s, unsafe_allow_html=True)


# ----------------------------------------------------------------------------------
# Cached readers (so changing a setting doesn't re-read the files)
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
    st.markdown("### Settings")
    facility = st.text_input("Facility name", value="", placeholder="e.g. Tsirang Hospital")
    report_date = st.date_input("Report date", value=date.today(), format="DD/MM/YYYY",
                                help="Expiry is checked against this date.")
    st.markdown("**Stock level limits**")
    st.caption("Used only when the upload has an average monthly use (AMC) column.")
    low_months = st.number_input("Low stock: less than … months", 0.0, 24.0, 1.0, 0.5)
    over_months = st.number_input("Overstock: more than … months", 1.0, 60.0, 12.0, 1.0)
    st.divider()
    st.caption("**Duplicate rule:** the same medicine name **and** the same dosage on more than one line. "
               "The same medicine in a different dosage (e.g. Amoxicillin 250mg and 1000mg capsule) "
               "is a separate item, not a duplicate.")

head = st.empty()
TITLE = "Medicine Stock Review"
SUB = "e-BMSIS stock balance: duplicates, expiry and stock levels"

# ----------------------------------------------------------------------------------
# Upload
# ----------------------------------------------------------------------------------
files = st.file_uploader("Upload the View Stock Balance report (PDF pages, Excel or CSV)",
                         type=["pdf", "xlsx", "xls", "csv"], accept_multiple_files=True)
if not files:
    head.markdown(ui.header(TITLE, SUB, [facility, f"Report date {report_date:%d/%m/%Y}"]),
                  unsafe_allow_html=True)
    html(ui.steps())
    st.caption("Tip: an Excel export, or a PDF printed in landscape with 'fit to page', includes Balance Qty, "
               "so the app can also flag zero, low and overstock.")
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
    head.markdown(ui.header(TITLE, SUB, [facility]), unsafe_allow_html=True)
    st.stop()

df = sc.combine(frames)
res = sc.analyse(df, report_date, low_months, over_months)
s = res["summary"]
d, e = s["dup_counts"], s["expiry_counts"]
content = rp.build_content(res, report_date, facility)
tables = {name: tbl for name, _, tbl in content["tables"]}
descs = {name: desc for name, desc, _ in content["tables"]}

head.markdown(ui.header(
    TITLE + (f" – {facility}" if facility else ""), SUB,
    [f"Report date {report_date:%d/%m/%Y}", f"{len(files)} file{'s' if len(files) > 1 else ''}",
     f"{s['lines']} stock lines"]), unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# Overview
# ----------------------------------------------------------------------------------
html(ui.section("Overview", "Each dosage of a medicine is counted as its own item"))
html(ui.cards([
    (s["lines"], "Stock lines", ""),
    (s["medicine_names"], "Medicine names", ""),
    (s["medicines"], "Total items (medicine + dosage)", ""),
    (s["names_multi_dosage"], "Medicines with several dosages", "calm"),
    (s["medicines_repeated"], "Duplicated items", "alert"),
    (len(tables["Exact Duplicates"]), "Exact duplicate medicines", "alert"),
    (e[sc.EXPIRED], "Expired lines", "alert"),
    (e[sc.EXP_3M], "Expiring within 3 months", "alert"),
]))

exact, check = tables["Exact Duplicates"], tables["Other Entries to Check"]
extra = int(exact["Copies"].sum() - len(exact)) if len(exact) else 0
html(ui.section("Entries to correct in e-BMSIS",
                "Same medicine and dosage entered more than once for the same batch"))
html('<div class="msr-grid2">'
     + ui.problem_list("Exact duplicate medicines",
                       f"Same dosage, batch, expiry and supply type entered more than once. "
                       f"Keep one line each: {extra} extra line{'s' if extra != 1 else ''} to remove.", exact)
     + ui.problem_list("Other entries to check",
                       "Same batch under two supply types, or with two expiry dates.", check, show_finding=True)
     + "</div>")

html(ui.section("Analysis"))
panels = [
    ui.bars("Duplicate lines by finding", [
        (sc.EXPIRY_CONFLICT, d[sc.EXPIRY_CONFLICT], "alert"),
        (sc.EXACT_DUP, d[sc.EXACT_DUP], "alert"),
        (sc.SAME_BATCH_TYPES, d[sc.SAME_BATCH_TYPES], "warn"),
        (sc.MULTI_BATCH, d[sc.MULTI_BATCH], ""),
    ]),
    ui.bars("Lines by expiry status", [
        (sc.EXPIRED, e[sc.EXPIRED], "alert"), (sc.EXP_3M, e[sc.EXP_3M], "warn"),
        (sc.EXP_6M, e[sc.EXP_6M], "soft"), (sc.EXP_OK, e[sc.EXP_OK], ""),
    ]),
]
if s["has_qty"]:
    tone = {"Zero stock": "alert", "Zero usable stock (only expired)": "alert", "Low stock": "warn",
            "Overstock": "soft", "OK": ""}
    panels.append(ui.bars("Items by stock level", [
        (k[2:], v, tone.get(k[2:], "")) for k, v in sorted(s["stock_counts"].items())]))
html('<div class="msr-grid2">' + "".join(panels) + "</div>")

note_html = [ui.esc(n) for n in content["notes"]]
note_html.append("<b>Exact duplicate</b> and <b>same batch, different expiry</b> lines need correcting in e-BMSIS. "
                 "<b>Different batches</b> are normal: issue the earliest expiry first (FEFO).")
html('<div class="msr-grid2">' + ui.summary_table(content["sections"]) + ui.notes(note_html) + "</div>")

# ----------------------------------------------------------------------------------
# Downloads
# ----------------------------------------------------------------------------------
html(ui.section("Download report", "Same content in all three formats"))
stem = f"Medicine_Stock_Review_{report_date:%Y-%m-%d}"
c = st.columns(3)
with st.spinner("Preparing reports…"):
    c[0].download_button("PDF report", data=rp.pdf_report(res, report_date, facility),
                         file_name=f"{stem}.pdf", mime="application/pdf", type="primary",
                         help="Printable report for filing or sharing.")
    c[1].download_button("Excel workbook", data=rp.excel_report(res, report_date, facility),
                         file_name=f"{stem}.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         help="One sheet per table, with filters.")
    c[2].download_button("Interactive report", data=rp.html_report(res, report_date, facility),
                         file_name=f"{stem}.html", mime="text/html",
                         help="One file that opens in any web browser, also offline. Search, filter and sort.")


# ----------------------------------------------------------------------------------
# Details
# ----------------------------------------------------------------------------------
def search(frame: pd.DataFrame, q: str) -> pd.DataFrame:
    if not q:
        return frame
    text = frame.astype(str).agg(" ".join, axis=1)
    return frame[text.str.contains(q, case=False, regex=False)]


html(ui.section("Details", "Search and filter each table"))
names = ["Exact Duplicates", "Other Entries to Check", "Duplicate Summary", "Duplicate Detail",
         "Same Name, Different Dosage", "Expiry Alerts"]
if "Stock Level" in tables:
    names.append("Stock Level")
names.append("All Items")
tabs = st.tabs([f"{n}  ({len(tables[n])})" for n in names])

for tab, name in zip(tabs, names):
    t = tables[name]
    with tab:
        st.caption(descs[name])
        a, b = st.columns([2, 1])
        q = a.text_input("Search", key=f"q-{name}", placeholder="Search medicine, dosage, batch …",
                         label_visibility="collapsed")
        view = t
        if name == "Duplicate Summary":
            pick = b.multiselect("Main finding", sc.FINDING_ORDER, default=sc.FINDING_ORDER, key="f-sum",
                                 label_visibility="collapsed", placeholder="Main finding")
            view = view[view["Main Finding"].isin(pick)]
        elif name == "Duplicate Detail":
            pick = b.multiselect("Finding", sc.FINDING_ORDER, default=sc.FINDING_ORDER, key="f-det",
                                 label_visibility="collapsed", placeholder="Finding")
            groups = view.loc[view["Finding"].isin(pick), "Group"].unique()   # keep whole items together
            view = view[view["Group"].isin(groups)]
        elif name == "Same Name, Different Dosage":
            pick = b.selectbox("Show", ["All dosages", "Only repeated dosages", "Only single-line dosages"],
                               key="f-nd", label_visibility="collapsed")
            if pick == "Only repeated dosages":
                view = view[view["Duplicated?"].str.startswith("Yes")]
            elif pick == "Only single-line dosages":
                view = view[view["Duplicated?"] == "No"]
        elif name == "Expiry Alerts":
            opts = [sc.EXPIRED, sc.EXP_3M, sc.EXP_6M]
            pick = b.multiselect("Status", opts, default=opts, key="f-exp", label_visibility="collapsed",
                                 placeholder="Expiry status")
            view = view[view["Expiry Status"].isin(pick)]
        elif name == "Stock Level":
            opts = sorted(view["Stock Status"].unique())
            pick = b.multiselect("Status", opts, default=[o for o in opts if o != "OK"] or opts, key="f-stk",
                                 label_visibility="collapsed", placeholder="Stock status")
            view = view[view["Stock Status"].isin(pick)]
        elif name == "All Items":
            pick = b.selectbox("Show", ["All lines", "Only duplicated items", "Only expiry alerts"], key="f-all",
                               label_visibility="collapsed")
            if pick == "Only duplicated items":
                view = view[view["Duplicate Finding"] != ""]
            elif pick == "Only expiry alerts":
                view = view[view["Expiry Status"].isin([sc.EXPIRED, sc.EXP_3M, sc.EXP_6M])]
        view = search(view, q)
        html(ui.count(len(view), len(t)))
        html(ui.table(view, group_col="Group" if name == "Duplicate Detail" else None))
