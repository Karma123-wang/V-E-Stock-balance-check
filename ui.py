"""
Look and feel for the Medicine Stock Review app.

Matches the interactive HTML report: navy header band, white KPI cards, bar panels,
blue-headed tables with coloured expiry status. Everything is plain HTML/CSS rendered
with st.markdown, so it needs no extra packages.
"""

from __future__ import annotations

import html

import pandas as pd

import stock_core as sc

NAVY = "#1f4e78"
BLUE = "#2e75b6"
INK = "#1d2433"
MUTED = "#5b6475"
LINE = "#e3e7ee"
BG = "#f6f8fb"
ALERT = "#c00000"

CSS = f"""
<style>
:root {{ --navy:{NAVY}; --blue:{BLUE}; --ink:{INK}; --muted:{MUTED}; --line:{LINE}; --bg:{BG}; --alert:{ALERT}; }}

/* ---------- page frame ---------- */
#MainMenu, footer, [data-testid="stDecoration"] {{ visibility:hidden; height:0; }}
[data-testid="stHeader"] {{ background:transparent; }}
.stApp {{ background:var(--bg); color:var(--ink); }}
.block-container {{ padding-top:1.2rem; padding-bottom:3rem; max-width:1320px; }}
h1, h2, h3 {{ color:var(--ink); letter-spacing:-.01em; }}

/* ---------- header band ---------- */
.msr-head {{ background:var(--navy); color:#fff; border-radius:14px; padding:22px 26px 20px; margin-bottom:18px;
  display:flex; justify-content:space-between; align-items:flex-end; gap:18px; flex-wrap:wrap; }}
.msr-head h1 {{ color:#fff; margin:0; font-size:1.65rem; font-weight:700; line-height:1.15; padding:0; }}
.msr-head p {{ margin:6px 0 0; color:rgba(255,255,255,.82); font-size:.92rem; }}
.msr-meta {{ display:flex; gap:8px; flex-wrap:wrap; }}
.msr-meta span {{ background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.22); color:#fff;
  border-radius:999px; padding:4px 11px; font-size:.8rem; white-space:nowrap; }}

/* ---------- section titles ---------- */
.msr-sec {{ display:flex; align-items:baseline; gap:10px; margin:22px 0 10px; }}
.msr-sec h3 {{ margin:0; font-size:1.08rem; font-weight:700; padding:0; }}
.msr-sec small {{ color:var(--muted); font-size:.85rem; }}

/* ---------- KPI cards ---------- */
.msr-cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(138px,1fr)); gap:10px; }}
.msr-card {{ background:#fff; border:1px solid var(--line); border-left:4px solid var(--blue); border-radius:10px;
  padding:12px 14px 11px; }}
.msr-card .v {{ font-size:1.75rem; font-weight:750; color:var(--navy); line-height:1.1; font-variant-numeric:tabular-nums; }}
.msr-card .l {{ font-size:.82rem; color:var(--muted); margin-top:3px; line-height:1.3; }}
.msr-card.alert {{ border-left-color:var(--alert); }}
.msr-card.alert .v {{ color:var(--alert); }}
.msr-card.calm {{ border-left-color:#9fb6cc; }}

/* ---------- panels ---------- */
.msr-grid2 {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(340px,1fr)); gap:12px; margin-top:12px; align-items:start; }}
.msr-panel {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:15px 17px; }}
.msr-panel h4 {{ margin:0 0 10px; font-size:.95rem; font-weight:700; color:var(--ink); padding:0; }}
.msr-bar {{ display:grid; grid-template-columns:minmax(150px,46%) 1fr 38px; gap:9px; align-items:center;
  margin:6px 0; font-size:.84rem; color:var(--ink); }}
.msr-bar .track {{ background:#eef3f9; border-radius:4px; height:14px; overflow:hidden; }}
.msr-bar .fill {{ background:var(--blue); height:100%; border-radius:4px; }}
.msr-bar .fill.alert {{ background:#d9534f; }}
.msr-bar .fill.warn {{ background:#f0a35e; }}
.msr-bar .fill.soft {{ background:#e9c46a; }}
.msr-bar .n {{ text-align:right; font-weight:700; font-variant-numeric:tabular-nums; }}

table.msr-sum {{ width:100%; border-collapse:collapse; font-size:.86rem; }}
table.msr-sum td {{ padding:5px 7px; border-bottom:1px solid var(--line); vertical-align:top; }}
table.msr-sum td.v {{ text-align:center; font-weight:700; width:62px; font-variant-numeric:tabular-nums; }}
table.msr-sum tr.sec td {{ background:#ddebf7; font-weight:700; color:var(--navy); }}
table.msr-sum td.note {{ color:var(--muted); font-size:.8rem; }}
table.msr-sum td.ind {{ padding-left:20px; }}
.msr-note p {{ margin:0 0 8px; color:var(--muted); font-size:.88rem; line-height:1.5; }}
.msr-note b {{ color:var(--ink); }}

/* ---------- problem lists ---------- */
.msr-panel h4 .cnt {{ display:inline-block; min-width:24px; text-align:center; background:var(--alert); color:#fff;
  border-radius:999px; font-size:.78rem; padding:1px 8px; margin-left:6px; vertical-align:1px; }}
.msr-panel p.hint {{ margin:-4px 0 10px; color:var(--muted); font-size:.84rem; }}
ul.msr-list {{ list-style:none; margin:0; padding:0; max-height:500px; overflow:auto; }}
ul.msr-list li {{ display:grid; grid-template-columns:1fr auto; column-gap:12px; padding:8px 0;
  border-top:1px solid var(--line); }}
ul.msr-list li .nm {{ font-size:.9rem; color:var(--ink); }}
ul.msr-list li .nm b {{ font-weight:650; }}
ul.msr-list li .mt {{ grid-column:1; font-size:.8rem; color:var(--muted); margin-top:2px; }}
ul.msr-list li .cp {{ grid-column:2; grid-row:1 / span 2; align-self:center; font-weight:700; color:var(--alert);
  font-variant-numeric:tabular-nums; }}
ul.msr-list .tag {{ display:inline-block; margin-left:6px; font-size:.72rem; color:#8a4b00; background:#fce4d6;
  border-radius:4px; padding:1px 6px; vertical-align:1px; }}
.msr-ok {{ color:var(--muted); font-size:.88rem; }}

/* ---------- data tables ---------- */
.msr-wrap {{ overflow:auto; max-height:560px; border:1px solid var(--line); border-radius:12px; background:#fff; }}
table.msr-data {{ border-collapse:collapse; width:100%; font-size:.84rem; }}
table.msr-data th {{ position:sticky; top:0; z-index:1; background:var(--navy); color:#fff; text-align:left;
  padding:8px 9px; font-weight:600; white-space:nowrap; }}
table.msr-data td {{ padding:7px 9px; border-bottom:1px solid var(--line); vertical-align:top; color:var(--ink); }}
table.msr-data tr:hover td {{ background:#f3f7fc; }}
table.msr-data tr.alt td {{ background:#f8fafc; }}
table.msr-data th.w-xl, table.msr-data td.w-xl {{ min-width:300px; }}
table.msr-data th.w-l, table.msr-data td.w-l {{ min-width:190px; }}
table.msr-data th.w-m, table.msr-data td.w-m {{ min-width:140px; }}
table.msr-data td.w-s {{ max-width:150px; word-break:break-word; }}
table.msr-data td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
table.msr-data td.st-x {{ background:#f8cbad !important; font-weight:600; }}
table.msr-data td.st-3 {{ background:#fce4d6 !important; }}
table.msr-data td.st-6 {{ background:#fff2cc !important; }}
table.msr-data td.flag {{ color:var(--alert); font-weight:600; }}
.msr-count {{ color:var(--muted); font-size:.84rem; margin:2px 0 8px; }}
.msr-empty {{ background:#fff; border:1px dashed var(--line); border-radius:12px; padding:22px; color:var(--muted); text-align:center; }}

/* ---------- landing ---------- */
.msr-steps {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:12px; margin:6px 0 16px; }}
.msr-step {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:15px 17px; }}
.msr-step .k {{ display:inline-flex; width:26px; height:26px; border-radius:50%; background:var(--navy); color:#fff;
  align-items:center; justify-content:center; font-weight:700; font-size:.85rem; margin-bottom:8px; }}
.msr-step b {{ display:block; margin-bottom:3px; }}
.msr-step span {{ color:var(--muted); font-size:.87rem; line-height:1.45; }}

/* ---------- Streamlit widgets ---------- */
.stTabs [data-baseweb="tab-list"] {{ gap:6px; flex-wrap:wrap; border-bottom:none; }}
.stTabs [data-baseweb="tab"] {{ background:#fff; border:1px solid var(--line); border-radius:999px; padding:6px 14px;
  height:auto; font-size:.86rem; }}
.stTabs [aria-selected="true"] {{ background:var(--navy); color:#fff; border-color:var(--navy); }}
.stTabs [aria-selected="true"] p {{ color:#fff; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}
.stDownloadButton button {{ width:100%; border-radius:10px; border:1px solid var(--navy); background:#fff;
  color:var(--navy); font-weight:600; padding:.6rem 1rem; }}
.stDownloadButton button:hover {{ background:#eef3f9; color:var(--navy); border-color:var(--navy); }}
.stDownloadButton button[kind="primary"] {{ background:var(--navy); color:#fff; }}
.stDownloadButton button[kind="primary"]:hover {{ background:#173d5f; color:#fff; }}
[data-testid="stFileUploaderDropzone"] {{ background:#fff; border:1.5px dashed #9fb6cc; border-radius:12px; }}
[data-testid="stSidebar"] {{ background:#fff; border-right:1px solid var(--line); }}
[data-testid="stExpander"] {{ background:#fff; border-radius:12px; }}
</style>
"""


def esc(v) -> str:
    # "$" would start maths formatting in st.markdown
    return html.escape("" if v is None else str(v)).replace("$", "&#36;")


# ----------------------------------------------------------------------------------
# blocks (each returns an HTML string for st.markdown(..., unsafe_allow_html=True))
# ----------------------------------------------------------------------------------
def header(title: str, subtitle: str, chips: list[str]) -> str:
    chip_html = "".join(f"<span>{esc(c)}</span>" for c in chips if c)
    return (f'<div class="msr-head"><div><h1>{esc(title)}</h1><p>{esc(subtitle)}</p></div>'
            f'<div class="msr-meta">{chip_html}</div></div>')


def section(title: str, hint: str = "") -> str:
    return f'<div class="msr-sec"><h3>{esc(title)}</h3><small>{esc(hint)}</small></div>'


def cards(items: list[tuple]) -> str:
    """items: (value, label, tone) with tone '', 'alert' or 'calm'."""
    out = []
    for value, label, tone in items:
        tone = tone if (tone != "alert" or value) else ""
        out.append(f'<div class="msr-card {tone}"><div class="v">{esc(value)}</div><div class="l">{esc(label)}</div></div>')
    return f'<div class="msr-cards">{"".join(out)}</div>'


def bars(title: str, rows: list[tuple]) -> str:
    """rows: (label, value, tone)."""
    mx = max([1] + [v for _, v, _ in rows])
    body = "".join(
        f'<div class="msr-bar"><span>{esc(l)}</span><div class="track"><div class="fill {t}" '
        f'style="width:{v / mx * 100:.1f}%"></div></div><span class="n">{v}</span></div>'
        for l, v, t in rows)
    return f'<div class="msr-panel"><h4>{esc(title)}</h4>{body}</div>'


def summary_table(sections: list) -> str:
    rows = []
    for name, items in sections:
        rows.append(f'<tr class="sec"><td colspan="3">{esc(name)}</td></tr>')
        for label, value, note in items:
            ind = " ind" if label.startswith("  ") else ""
            rows.append(f'<tr><td class="{ind.strip()}">{esc(label.strip())}</td><td class="v">{value}</td>'
                        f'<td class="note">{esc(note)}</td></tr>')
    return f'<div class="msr-panel"><h4>Summary</h4><table class="msr-sum">{"".join(rows)}</table></div>'


def notes(paragraphs: list[str]) -> str:
    return ('<div class="msr-panel msr-note"><h4>How to read this</h4>'
            + "".join(f"<p>{p}</p>" for p in paragraphs) + "</div>")


def steps() -> str:
    items = [
        ("1", "Export from e-BMSIS", "Open Ledger, then Stock, then View Stock Balance. Print every page to PDF, "
                                     "or export to Excel to include quantities."),
        ("2", "Upload the files", "Drop all pages at once (1–100, 101–200 …) or a single Excel/CSV export."),
        ("3", "Review and download", "See duplicates, expiry and stock levels, then download the PDF, Excel "
                                     "or interactive report."),
    ]
    return '<div class="msr-steps">' + "".join(
        f'<div class="msr-step"><div class="k">{k}</div><b>{esc(t)}</b><span>{esc(d)}</span></div>'
        for k, t, d in items) + "</div>"


def table(df: pd.DataFrame, group_col: str | None = None) -> str:
    """Styled read-only table; expects display-formatted columns (from reports.build_content)."""
    if df.empty:
        return '<div class="msr-empty">Nothing to show for this selection.</div>'
    cols = list(df.columns)
    status_i = cols.index("Expiry Status") if "Expiry Status" in cols else -1
    flag_i = [cols.index(c) for c in ("Main Finding", "Finding", "Duplicate Finding") if c in cols]
    num_cols = {i for i, c in enumerate(cols)
                if c in ("Row #", "Group", "Lines", "Extra Lines", "Batches", "Expired Lines", "Days Left",
                         "Balance Qty", "Usable Balance Qty", "Copies", "Expired Qty", "Monthly Use", "Months of Stock")}
    g_i = cols.index(group_col) if group_col and group_col in cols else -1
    width = {c: "w-xl" for c in ("What to do", "Item (medicine + dosage)")}
    width.update({c: "w-l" for c in ("Dosage", "Main Finding", "Finding", "Duplicate Finding")})
    width.update({c: "w-m" for c in ("Medicine", "Expiry Status", "Duplicated?", "Stock Status")})
    width.update({c: "w-s" for c in ("Batch No.", "Rows")})
    head = "".join(f'<th class="{width.get(c, "")}">{esc(c)}</th>' for c in cols)
    body = []
    for row in df.itertuples(index=False):
        vals = [_fmt(v) for v in row]
        alt = ' class="alt"' if g_i >= 0 and vals[g_i].isdigit() and int(vals[g_i]) % 2 == 0 else ""
        tds = []
        for i, v in enumerate(vals):
            cls = [width.get(cols[i], "")]
            if i in num_cols:
                cls.append("num")
            if i == status_i:
                cls.append({sc.EXPIRED: "st-x", sc.EXP_3M: "st-3", sc.EXP_6M: "st-6"}.get(v, ""))
            if i in flag_i and v and not v.startswith("Different"):
                cls.append("flag")
            c = f' class="{" ".join(x for x in cls if x)}"' if any(cls) else ""
            tds.append(f"<td{c}>{esc(v)}</td>")
        body.append(f"<tr{alt}>{''.join(tds)}</tr>")
    return f'<div class="msr-wrap"><table class="msr-data"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def problem_list(title: str, hint: str, df: pd.DataFrame, show_finding: bool = False) -> str:
    """Compact list of entries to correct: medicine + dosage, batch, rows."""
    if df.empty:
        body = '<div class="msr-ok">None found.</div>'
    else:
        items = []
        for g in df.to_dict("records"):
            tag = f'<span class="tag">{esc(g["Finding"])}</span>' if show_finding and "Finding" in g else ""
            items.append(
                f'<li><div class="nm"><b>{esc(g["Medicine"])}</b> {esc(g["Dosage"])}{tag}</div>'
                f'<div class="mt">Batch {esc(g["Batch No."])}, rows {esc(g["Rows"])}</div>'
                f'<div class="cp">{esc(g["Copies"])}×</div></li>')
        body = f'<ul class="msr-list">{"".join(items)}</ul>'
    return (f'<div class="msr-panel"><h4>{esc(title)} <span class="cnt">{len(df)}</span></h4>'
            f'<p class="hint">{esc(hint)}</p>{body}</div>')


def count(shown: int, total: int, noun: str = "rows") -> str:
    return f'<div class="msr-count">Showing {shown} of {total} {noun}</div>'


def _fmt(v) -> str:
    if v is None or v is pd.NA or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)
