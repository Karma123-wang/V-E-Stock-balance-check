"""
Medicine Stock Review - core logic (no Streamlit here, so it can be tested on its own).

Reads e-BMSIS "View Stock Balance" exports:
  * PDF printouts (browser "Print to PDF" of the stock balance page)
  * Excel / CSV exports

Then checks:
  * Duplicates  - same medicine name + same dosage listed more than once.
                  Same name with a DIFFERENT dosage is a separate medicine, not a duplicate.
  * Expiry      - expired, expiring within 3 months, within 6 months.
  * Stock level - zero stock, low / over stock (only when a Balance Qty column exists).
"""

from __future__ import annotations

import io
import re
from datetime import date

import pandas as pd

SUPPLY_TYPES = ("Annual", "Additional", "Mobilized", "Mobilised")

# ----------------------------------------------------------------------------------
# Finding labels (used in the app and in the Excel report)
# ----------------------------------------------------------------------------------
EXACT_DUP = "Exact duplicate entry"
SAME_BATCH_TYPES = "Same batch under 2 supply types"
EXPIRY_CONFLICT = "Same batch, different expiry (data error)"
MULTI_BATCH = "Different batches (normal, use FEFO)"

FINDING_ACTION = {
    EXACT_DUP: "Same medicine, dosage, batch, expiry and supply type listed more than once. "
               "Likely double entry in e-BMSIS: check ledger and physical count, then remove the extra line.",
    SAME_BATCH_TYPES: "Same batch shown under two supply types (e.g. Annual + Additional). "
                      "Confirm there were two separate receipts; if not, correct it to avoid double-counting.",
    EXPIRY_CONFLICT: "Same batch number has two different expiry dates. Check the actual pack and correct e-BMSIS.",
    MULTI_BATCH: "Different batches of the same medicine and dosage (normal). Issue earliest expiry first (FEFO).",
}
FINDING_ORDER = [EXPIRY_CONFLICT, EXACT_DUP, SAME_BATCH_TYPES, MULTI_BATCH]

EXPIRED = "EXPIRED"
EXP_3M = "Expires ≤3 months"
EXP_6M = "Expires ≤6 months"
EXP_OK = "OK"


# ==================================================================================
# 1. READING FILES
# ==================================================================================
def read_any(file_name: str, data: bytes) -> pd.DataFrame:
    """Read one uploaded file into the standard table."""
    name = file_name.lower()
    if name.endswith(".pdf"):
        return parse_pdf(data, source=file_name)
    if name.endswith((".xlsx", ".xls")):
        raw = _read_excel_find_header(data)
    elif name.endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(data))
    else:
        raise ValueError(f"{file_name}: please upload a PDF, Excel (.xlsx/.xls) or CSV file.")
    return standardise_table(raw, source=file_name)


def combine(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Join several uploads (e.g. the 4 PDF pages of one report)."""
    df = pd.concat(frames, ignore_index=True)
    # When pages overlap, the same printed row number can appear twice.
    if df["row_no"].notna().all():
        df = (df.sort_values(["row_no", "name_len"], ascending=[True, False])
                .drop_duplicates("row_no", keep="first")
                .sort_values("row_no"))
    return df.drop(columns="name_len").reset_index(drop=True)


# ---------------------------------- PDF ------------------------------------------
def parse_pdf(data: bytes, source: str = "") -> pd.DataFrame:
    import pdfplumber
    import logging
    logging.getLogger("pdfminer").setLevel(logging.ERROR)

    rows: list[dict] = []
    cols = None          # column boundaries, taken from the header row
    cur = None           # row currently being filled

    # ---- 1. collect printed lines page by page ----
    pages: list[list[list[dict]]] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(extra_attrs=["fontname", "size"], x_tolerance=1.5)
            header = _find_pdf_header(words)
            if header:
                cols = header
            if cols is None:
                pages.append([])
                continue
            body = [w for w in words
                    if "Bold" not in w["fontname"] and 8.5 <= w["size"] <= 11
                    and w["x0"] >= cols["row_no"] - 5
                    and (not header or w["top"] > header["header_top"] + 5)]
            body.sort(key=lambda w: (round(w["top"]), w["x0"]))
            lines: list[list[dict]] = []
            for w in body:
                if lines and abs(lines[-1][0]["top"] - w["top"]) < 3:
                    lines[-1].append(w)
                else:
                    lines.append([w])
            lines = [ln for ln in lines if not _is_pager_line(ln)]
            pages.append(lines)
    if cols is None:
        raise ValueError(f"{source}: table header (Item Name / Batch No. / Expiry Date) not found. "
                         "Is this an e-BMSIS 'View Stock Balance' PDF?")

    # ---- 2. a line cut by a page break is printed on both pages: keep the better copy ----
    for a, b in zip(pages, pages[1:]):
        if not a or not b:
            continue
        bottom, top = a[-1], b[0]
        if bottom[0]["top"] > 700 and top[0]["top"] < 75 and _same_print_line(bottom, top):
            if _line_quality(bottom) >= _line_quality(top):
                b.pop(0)
            else:
                a.pop()
    for lines in pages:            # leftover clipped scraps such as "p g" or "l ti b ttl"
        lines[:] = [ln for ln in lines if not _is_clipped_scrap(ln)]

    # ---- 3. build rows ----
    for lines in pages:
            for line in lines:
                first = line[0]
                is_new = (first["x0"] < cols["supply"] - 2 and first["text"].isdigit()
                          and any(w["text"] in SUPPLY_TYPES for w in line))
                if is_new:
                    cur = {"row_no": int(first["text"]), "parts": {k: [] for k in cols["order"]}}
                    rows.append(cur)
                    line = line[1:]
                if cur is None:
                    continue
                for w in line:
                    col = _column_for(w["x0"], cols)
                    if col:
                        cur["parts"][col].append(w["text"])

    out = []
    for r in rows:
        p = r["parts"]
        get = lambda k: " ".join(p.get(k, []))
        supply = " ".join(t for t in p.get("supply", []) if t != "Qty")
        out.append({
            "row_no": r["row_no"],
            "supply_type": supply.replace("Mobilised", "Mobilized"),
            "item_name": _clean_name(get("item")),
            "batch": get("batch").strip(),
            "expiry": _first_date(get("expiry")),
            "received_batch": get("received_batch").strip(),
            "received_expiry": get("received_expiry").strip(),
            "balance_qty": _to_number(get("balance")) if "balance" in p else None,
            "source": source,
        })
    df = pd.DataFrame(out)
    if df.empty:
        raise ValueError(f"{source}: no stock rows found. Is this an e-BMSIS 'View Stock Balance' PDF?")
    df["name_len"] = df["item_name"].str.len()
    return df


def _find_pdf_header(words):
    """Find column x-positions from the bold table header ('Item Name', 'Batch No.', ...)."""
    bold = [w for w in words if "Bold" in w["fontname"] and w["text"] not in ("↑↓",)]
    item = next((w for w in bold if w["text"] == "Item"), None)
    if not item:
        return None
    # same font size as 'Item' keeps out the left menu ("Reports" etc.) that overlaps the page
    band = [w for w in bold if abs(w["top"] - item["top"]) < 16 and abs(w["size"] - item["size"]) < 0.5]
    band.sort(key=lambda w: w["x0"])
    # cluster header words into columns
    clusters: list[list[dict]] = []
    for w in band:
        if clusters and w["x0"] - max(c["x1"] for c in clusters[-1]) < 8:
            clusters[-1].append(w)
        elif clusters and any(abs(w["x0"] - c["x0"]) < 3 for c in clusters[-1]):
            clusters[-1].append(w)
        else:
            clusters.append([w])
    cols = {"order": [], "header_top": item["top"]}
    seen_expiry = seen_batch = False
    for c in clusters:
        text = " ".join(w["text"] for w in sorted(c, key=lambda w: (w["top"], w["x0"]))).lower()
        x = min(w["x0"] for w in c)
        key = None
        if text.startswith("#"):
            key = "row_no"
        elif "supply" in text:
            key = "supply"
        elif "item" in text:
            key = "item"
        elif "received" in text and "batch" in text:
            key = "received_batch"
        elif "received" in text and "expiry" in text:
            key = "received_expiry"
        elif "batch" in text and not seen_batch:
            key, seen_batch = "batch", True
        elif "expiry" in text and not seen_expiry:
            key, seen_expiry = "expiry", True
        elif "balance" in text or text in ("qty", "quantity", "stock"):
            key = "balance"
        elif text.startswith("re"):
            key = "received_batch"      # header cut off at page edge ("Re...")
        if key and key not in cols:
            cols[key] = x
            cols["order"].append(key)
    if not {"item", "batch", "expiry"} <= set(cols):
        return None
    cols.setdefault("row_no", 50)
    cols.setdefault("supply", cols["item"] - 60)
    if "supply" not in cols["order"]:
        cols["order"].insert(0, "supply")
    cols["order"] = sorted([k for k in cols["order"] if k != "row_no"], key=lambda k: cols[k])
    return cols


PAGER_WORDS = {"showing", "entries", "previous", "next", "to", "of"}
SHORT_REAL_WORDS = {"qty", "jar", "for", "and", "inj", "gel", "iv", "im", "tab", "cap", "bag", "kit", "box", "set"}


def _is_pager_line(line):
    """'Showing 1 to 100 of 314 entries   Previous 1 2 3 4 Next' at the bottom of the table."""
    texts = [w["text"].lower() for w in line]
    if any(t in ("showing", "entries", "previous", "next") for t in texts):
        return True
    return (all(t.isdigit() or t in PAGER_WORDS for t in texts)
            and any(t in ("to", "of") for t in texts) and texts.count("of") >= 1)


def _line_quality(line):
    return sum(len(w["text"]) for w in line if len(w["text"]) >= 4) + \
        sum(10 for w in line if re.fullmatch(r"\d{2}/\d{2}/\d{4}", w["text"]))


def _same_print_line(a, b):
    small, big = (a, b) if len(a) <= len(b) else (b, a)
    def alike(p, q):          # a clipped word is the start of the full word ("Additi" / "Additional")
        p, q = p.lower(), q.lower()
        return p.startswith(q) or q.startswith(p)
    hits = sum(1 for w in small
               if any(abs(w["x0"] - v["x0"]) < 4 and alike(w["text"], v["text"]) for v in big))
    return hits / max(len(small), 1) > 0.5


def _is_clipped_scrap(line):
    texts = [w["text"] for w in line]
    return all(t.isalpha() and len(t) <= 3 and t.lower() not in SHORT_REAL_WORDS for t in texts)


def _column_for(x, cols):
    best = None
    for k in cols["order"]:
        if x >= cols[k] - 3:
            best = k
    return best


# -------------------------------- Excel / CSV ------------------------------------
COLUMN_HINTS = {
    "item_name": ["item name", "item", "medicine", "drug", "description", "product"],
    "batch": ["batch no", "batch"],
    "expiry": ["expiry date", "expiry", "exp date", "exp"],
    "supply_type": ["supply type", "supply"],
    "balance_qty": ["balance qty", "balance quantity", "stock balance", "balance", "closing stock",
                    "available qty", "current stock", "qty", "quantity"],
    "monthly_use": ["average monthly consumption", "monthly consumption", "amc", "avg monthly",
                    "monthly use", "consumption"],
    "row_no": ["#", "sl no", "s/n", "sl. no", "sno"],
}


def _read_excel_find_header(data: bytes) -> pd.DataFrame:
    """e-BMSIS exports sometimes have title rows above the table; find the real header."""
    preview = pd.read_excel(io.BytesIO(data), header=None, nrows=30)
    header_row = 0
    for i, row in preview.iterrows():
        text = " ".join(str(v).lower() for v in row.values)
        if "item" in text and ("batch" in text or "expiry" in text):
            header_row = i
            break
    return pd.read_excel(io.BytesIO(data), header=header_row)


def map_columns(raw: pd.DataFrame) -> dict:
    lower = {c: re.sub(r"\s+", " ", str(c)).strip().lower() for c in raw.columns}
    mapping = {}
    for field, hints in COLUMN_HINTS.items():
        for hint in hints:
            match = [c for c, l in lower.items()
                     if (l == hint or hint in l) and c not in mapping.values()
                     and not (field in ("batch", "expiry") and "received" in l)
                     and not (field == "balance_qty" and ("received" in l or "issue" in l))]
            if match:
                mapping[field] = match[0]
                break
    return mapping


def standardise_table(raw: pd.DataFrame, source: str = "", mapping: dict | None = None) -> pd.DataFrame:
    raw = raw.dropna(how="all")
    mapping = mapping or map_columns(raw)
    missing = [f for f in ("item_name", "expiry") if f not in mapping]
    if missing:
        raise ValueError(f"{source}: could not find column(s) for {', '.join(missing)}. "
                         f"Columns found: {', '.join(map(str, raw.columns))}")
    df = pd.DataFrame({
        "row_no": pd.to_numeric(raw[mapping["row_no"]], errors="coerce") if "row_no" in mapping else None,
        "supply_type": raw[mapping["supply_type"]].astype(str).str.replace("Qty", "").str.strip()
                       if "supply_type" in mapping else "",
        "item_name": raw[mapping["item_name"]].astype(str).map(_clean_name),
        "batch": raw[mapping["batch"]].astype(str).str.strip() if "batch" in mapping else "",
        "expiry": raw[mapping["expiry"]].map(_first_date),
        "received_batch": "",
        "received_expiry": "",
        "balance_qty": raw[mapping["balance_qty"]].map(_to_number) if "balance_qty" in mapping else None,
        "source": source,
    })
    if "monthly_use" in mapping:
        df["monthly_use"] = raw[mapping["monthly_use"]].map(_to_number)
    df = df[df["item_name"].str.len() > 1]
    df = df[~df["item_name"].str.lower().isin(["nan", "none", "total"])]
    if df["row_no"].isna().any():
        df["row_no"] = range(1, len(df) + 1)
    df["name_len"] = df["item_name"].str.len()
    return df.reset_index(drop=True)


# ---------------------------------- helpers --------------------------------------
def _clean_name(s: str) -> str:
    s = re.sub(r"\s+", " ", str(s)).strip()
    s = re.sub(r"\s*([µμ])\s*g\b", r"\1g", s)            # "125 μ g" -> "125μg"
    return s.replace("( ", "(").replace(" )", ")")


def _first_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (pd.Timestamp, date)):
        return pd.Timestamp(v)
    m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", str(v))
    if not m:
        return pd.to_datetime(str(v), errors="coerce", dayfirst=True)
    d, mth, y = map(int, m.groups())
    try:
        return pd.Timestamp(year=y, month=mth, day=d)
    except ValueError:          # e.g. 31/09/2027 - use last day of that month
        return pd.Timestamp(year=y, month=mth, day=1) + pd.offsets.MonthEnd(0)


def _to_number(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


# ==================================================================================
# 2. ANALYSIS
# ==================================================================================
# ---------- splitting "Amoxicillin 250mg capsule" into name / strength / form ----------
_UNIT = r"(?:mg|mcg|µg|μg|g|kg|ml|l|iu|%|meq|mmol|million\s*iu|actuations?)"
STRENGTH_RE = re.compile(
    r"\d[\d,]*(?:\.\d+)?\s*" + _UNIT + r"(?:\s*/\s*(?:\d+(?:\.\d+)?\s*)?(?:ml|l|g|actuation|dose))?(?![a-z])"
    r"|\d+\s*:\s*[\d,]+"                       # 1:80,000   30:70
    r"|\b1\s+in\s+[\d,]+",                     # 1 in 200,000
    re.I)
FORM_RE = re.compile(
    r"\b(tablets?|capsules?|injection|inj|infusion|powder|cream|ointment|syrup|solution|suspension|drops?|"
    r"eye|nasal|gel|inhaler|inhalation|suppository|spray|pump|sachet|paste|crystal|lotion|jar|packet|bottle|"
    r"vial|ampoule|pre-filled|respiratory|cartridge|w/w|sustained|extended|enteric|chewable|retard|"
    r"dried|vaginal|viginal|sublingual|intradermal|plastic|preservative)\b", re.I)


def _in_brackets(text, m):
    return text[:m.start()].rstrip().endswith("(") and text[m.end():].lstrip().startswith(")")


def split_item(name: str) -> tuple[str, str, str]:
    """Return (medicine name, strength, form) from an e-BMSIS item description.

    'Amoxicillin 250mg capsule'                      -> ('Amoxicillin', '250mg', 'capsule')
    'Lignocaine 2% + Adrenaline 1:80,000 injection'  -> ('Lignocaine + Adrenaline', '2% + 1:80,000', 'injection')
    'Morphine 10mg/ml (1mL) injection ampoule'       -> ('Morphine', '10mg/ml', '(1mL) injection ampoule')
    """
    s = re.sub(r"\s+", " ", name).strip()
    form_at = len(s)
    for m in FORM_RE.finditer(s):
        if m.start() == 0 or s[:m.start()].rstrip().lower().endswith(" for"):   # "Water for injection"
            continue
        form_at = m.start()
        break
    head, form = s[:form_at], s[form_at:]

    ms = list(STRENGTH_RE.finditer(head))
    name_parts, strengths, tail_from = [head], [], len(head)
    if ms:
        name_parts = [head[:ms[0].start()]]
        tail_from = ms[0].start()
        for i, m in enumerate(ms):
            if _in_brackets(head, m):                     # "(1mL)", "(200 actuations)" = pack size
                tail_from = head.rfind("(", 0, m.start())
                if i == 0:
                    name_parts = [head[:tail_from]]
                break
            strengths.append(m.group().strip())
            tail_from = m.end()
            if i + 1 < len(ms):
                between = head[m.end():ms[i + 1].start()]
                if "(" in between or ")" in between:       # "(equivalent to 500mg ...)" -> description
                    break
                if between.strip(" ,"):
                    name_parts.append(between)
                tail_from = ms[i + 1].start()
    tail = head[tail_from:] if ms else ""
    nm = " ".join(name_parts)
    nm = re.sub(r"\s*\+\s*(\+\s*)*", " + ", nm)
    nm = re.sub(r"\s+", " ", nm).strip(" +,-")
    if nm.count("(") > nm.count(")"):
        nm = nm.rstrip() + ")"
    nm = nm.replace("( ", "(").replace(" )", ")")
    form = re.sub(r"\s+", " ", (tail + " " + form)).strip(" ,")
    if form.startswith(")") and nm.endswith(")"):
        form = form[1:].strip()
    if not strengths:              # e.g. "Glycerin suppository 4g": strength written after the form
        strengths = [m.group().strip() for m in STRENGTH_RE.finditer(form) if not _in_brackets(form, m)][:1]
    return nm or s, " + ".join(strengths), form


def dosage_text(strength: str, form: str) -> str:
    if strength and strength.replace(" ", "").lower() in form.replace(" ", "").lower():
        return form                              # "suppository 4g", not "4g suppository 4g"
    return (strength + " " + form).strip()


def name_key(medicine_name: str) -> str:
    return re.sub(r"[^a-z0-9+]", "", medicine_name.lower())


def medicine_key(name: str) -> str:
    """Key for 'same medicine + same dosage'.

    Uses the full e-BMSIS item description (name + strength + form + pack),
    so 'Amoxicillin 250mg capsule' and 'Amoxicillin 500mg capsule' stay separate,
    while two lines of 'Amoxicillin 250mg capsule' are grouped as duplicates.
    """
    s = name.lower()
    s = re.sub(r"\binj\b\.?", "injection", s)
    s = s.replace("µg", "mcg").replace("μg", "mcg")
    s = re.sub(r"(\d)\s+(mg|mcg|g|ml|iu|%)", r"\1\2", s)
    s = re.sub(r"[^a-z0-9%.+/:]", "", s)
    return s


FORM_WORDS = r"(tablets?|capsules?|injection|inj|syrup|suspension|cream|ointment|powder|solution|drops?|gel|" \
             r"inhaler|inhalation|suppository|sachet|spray|crystal|paste|lotion|respiratory|vial|ampoule|" \
             r"infusion|pump|eye|nasal|viginal|vaginal|enteric|soluble|sustained|extended|retard|chewable)"


def split_name_dosage(item: str) -> tuple[str, str]:
    """'Amoxicillin 250mg capsule' -> ('Amoxicillin', '250mg capsule').

    The medicine name is the text before the first strength (a number) or dosage form word.
    Used to show 'same name, different dosage' – it does NOT decide duplicates;
    duplicates are always the full Item Name (name + dosage) repeated.
    """
    words = item.split()
    cut = len(words)
    depth = 0                       # inside brackets? e.g. "Thiamine (vitamin B1)" – keep it in the name
    for i, w in enumerate(words):
        before = depth
        depth = max(0, depth + w.count("(") - w.count(")"))
        if i == 0 or before > 0:
            continue
        lw = w.lower().strip(",:;()")
        if re.search(r"\d", w) or re.fullmatch(FORM_WORDS, lw):
            if lw == "injection" and words[i - 1].lower() == "for":   # "Water for injection"
                continue
            cut = i
            break
    name = " ".join(words[:cut])
    name = re.sub(r"\s*\([^)]*$", "", name)          # drop an unclosed "(" left by the cut
    name = name.rstrip(" ,;:+-/(").strip() or item
    dosage = " ".join(words[cut:]).strip()
    return name, dosage


def _norm_batch(b) -> str:
    return re.sub(r"[^a-z0-9]", "", str(b).lower())


def analyse(df: pd.DataFrame, report_date: date, low_months: float = 1.0,
            over_months: float = 12.0) -> dict:
    df = df.copy()
    today = pd.Timestamp(report_date)
    df["days_to_expiry"] = (df["expiry"] - today).dt.days
    df["expiry_status"] = df["days_to_expiry"].map(_expiry_status)
    df["med_key"] = df["item_name"].map(medicine_key)          # name + dosage  -> decides duplicates
    df["batch_key"] = df["batch"].map(_norm_batch)
    # split_item keeps combinations together ("Amoxicillin + Clavulanic acid"), so they are not
    # shown as another dosage of the single medicine
    split = df["item_name"].map(split_item)
    df["medicine"] = split.str[0]
    df["strength"] = split.str[1]
    df["form"] = split.str[2]
    df["dosage"] = [dosage_text(s, f) for s, f in zip(df["strength"], df["form"])]
    df["name_key"] = df["medicine"].map(name_key)

    # ---------------- duplicates ----------------
    df["duplicate_finding"] = ""
    df["group_id"] = pd.NA
    counts = df["med_key"].value_counts()
    dup_keys = [k for k in df["med_key"].unique() if counts[k] > 1]
    for gid, key in enumerate(dup_keys, 1):
        idx = df.index[df["med_key"] == key]
        df.loc[idx, "group_id"] = gid
        for i in idx:
            r = df.loc[i]
            same = df.loc[[j for j in idx if j != i and df.at[j, "batch_key"] == r["batch_key"]
                           and r["batch_key"] != ""]]
            if same.empty:
                f = MULTI_BATCH
            elif (same["expiry"] != r["expiry"]).any():
                f = EXPIRY_CONFLICT
            elif (same["supply_type"] != r["supply_type"]).any():
                f = SAME_BATCH_TYPES
            else:
                f = EXACT_DUP
            df.at[i, "duplicate_finding"] = f

    dups = df[df["duplicate_finding"] != ""].copy()
    dups["action"] = dups.apply(_dup_action, axis=1)
    dups = dups.sort_values(["group_id", "row_no"])

    # one row per duplicated medicine + dosage
    def _main_finding(fs):
        for f in FINDING_ORDER:
            if f in set(fs):
                return f
    dup_summary = (dups.groupby("med_key", sort=False)
                   .agg(medicine=("medicine", "first"), dosage=("dosage", "first"),
                        item_name=("item_name", "first"), lines=("row_no", "size"),
                        batches=("batch_key", "nunique"),
                        supply_types=("supply_type", lambda s: ", ".join(sorted(set(s)))),
                        rows=("row_no", lambda s: ", ".join(map(str, s))),
                        finding=("duplicate_finding", _main_finding),
                        expired_lines=("expiry_status", lambda s: int((s == EXPIRED).sum())))
                   .reset_index(drop=True))
    dup_summary["extra_lines"] = dup_summary["lines"] - 1
    dup_summary = dup_summary[["medicine", "dosage", "lines", "extra_lines", "batches", "supply_types",
                               "finding", "expired_lines", "rows", "item_name"]] \
        .sort_values(["finding", "medicine"], key=lambda c: c.map(FINDING_ORDER.index) if c.name == "finding" else c)

    # same medicine name with different dosages (NOT duplicates – shown for information)
    per_name = df.groupby("name_key").agg(medicine=("medicine", "first"),
                                          dosages=("med_key", "nunique"), lines=("row_no", "size"))
    multi = per_name[per_name["dosages"] > 1]
    name_dosage = []
    for key, r in multi.iterrows():
        sub = df[df["name_key"] == key]
        for mk, g in sub.groupby("med_key", sort=False):
            name_dosage.append({"medicine": r["medicine"], "dosage": g["dosage"].iloc[0],
                                "item_name": g["item_name"].iloc[0], "lines": len(g),
                                "duplicated": "Yes – same dosage repeated" if len(g) > 1 else "No",
                                "rows": ", ".join(map(str, g["row_no"]))})
    name_dosage = pd.DataFrame(name_dosage, columns=["medicine", "dosage", "item_name", "lines",
                                                     "duplicated", "rows"])

    # ---------------- expiry ----------------
    expiry = df[df["expiry_status"].isin([EXPIRED, EXP_3M, EXP_6M])].copy()
    expiry["action"] = expiry["expiry_status"].map({
        EXPIRED: "Remove from usable stock, quarantine and write off.",
        EXP_3M: "Use first. If balance is high, mobilize to other wards/facilities before expiry.",
        EXP_6M: "Prioritise use (FEFO); review quantity against monthly use.",
    })
    expiry = expiry.sort_values("expiry")

    # ---------------- stock level (per medicine + dosage) ----------------
    has_qty = df["balance_qty"].notna().any()
    stock = None
    if has_qty:
        usable = df[df["expiry_status"] != EXPIRED]
        # exact duplicates would double count: keep one line per batch/expiry/type
        usable = usable.drop_duplicates(["med_key", "batch_key", "expiry", "supply_type", "balance_qty"])
        agg = {"item_name": "first", "balance_qty": "sum", "row_no": lambda s: ", ".join(map(str, s))}
        if "monthly_use" in df:
            agg["monthly_use"] = "max"
        stock = usable.groupby("med_key").agg(agg).rename(columns={"row_no": "rows"})
        expired_qty = df[df["expiry_status"] == EXPIRED].groupby("med_key")["balance_qty"].sum()
        stock["expired_qty"] = expired_qty.reindex(stock.index).fillna(0)
        all_keys = df.groupby("med_key")["item_name"].first()
        missing = all_keys[~all_keys.index.isin(stock.index)]       # only expired stock left
        if len(missing):
            extra = pd.DataFrame({"item_name": missing, "balance_qty": 0.0, "rows": "",
                                  "expired_qty": expired_qty.reindex(missing.index).fillna(0)})
            stock = pd.concat([stock, extra])
        if "monthly_use" in stock:
            stock["months_of_stock"] = stock.apply(
                lambda r: round(r["balance_qty"] / r["monthly_use"], 1) if r.get("monthly_use") else None, axis=1)
        stock["stock_status"] = stock.apply(lambda r: _stock_status(r, low_months, over_months), axis=1)
        stock = stock.reset_index(drop=True).sort_values(["stock_status", "item_name"])

    summary = {
        "lines": len(df),
        "medicine_names": df["name_key"].nunique(),
        "medicines": df["med_key"].nunique(),
        "names_multi_dosage": len(multi),
        "medicines_repeated": len(dup_keys),
        "dup_lines": len(dups),
        "extra_lines": len(dups) - len(dup_keys),
        "dup_counts": {f: int((df["duplicate_finding"] == f).sum()) for f in FINDING_ORDER},
        "expiry_counts": {s: int((df["expiry_status"] == s).sum()) for s in [EXPIRED, EXP_3M, EXP_6M, EXP_OK]},
        "has_qty": has_qty,
        "stock_counts": stock["stock_status"].value_counts().to_dict() if stock is not None else {},
        "bad_received": int(_received_mismatch(df).sum()),
    }
    df["received_batch_differs"] = _received_mismatch(df)
    return {"all": df, "duplicates": dups, "dup_summary": dup_summary, "name_dosage": name_dosage,
            "expiry": expiry, "stock": stock, "summary": summary}


def _expiry_status(days):
    if pd.isna(days):
        return "No expiry date"
    if days < 0:
        return EXPIRED
    if days <= 92:
        return EXP_3M
    if days <= 183:
        return EXP_6M
    return EXP_OK


def _dup_action(r):
    if r["duplicate_finding"] == MULTI_BATCH and r["expiry_status"] == EXPIRED:
        return "This batch has EXPIRED. Separate it from usable stock and write off; issue the other batch."
    return FINDING_ACTION[r["duplicate_finding"]]


def _stock_status(r, low_m, over_m):
    q = r["balance_qty"] or 0
    if q <= 0:
        return "1 Zero stock" if not r.get("expired_qty") else "1 Zero usable stock (only expired)"
    mos = r.get("months_of_stock")
    if mos is not None and not pd.isna(mos):
        if mos < low_m:
            return "2 Low stock"
        if mos > over_m:
            return "3 Overstock"
    return "4 OK"


def _received_mismatch(df):
    """Received batch printed but clearly different from current batch.
    The PDF often cuts the received batch short, so only compare the visible start."""
    out = []
    for b, rb in zip(df["batch"].map(_norm_batch), df["received_batch"].map(_norm_batch)):
        if len(rb) < 4 or not b:
            out.append(False)
        else:
            n = min(len(b), len(rb))
            out.append(b[:n] != rb[:n])
    return pd.Series(out, index=df.index)


# ==================================================================================
# 3. REPORTS  (Excel, PDF and interactive HTML live in reports.py)
# ==================================================================================
def excel_report(res: dict, report_date: date, facility: str = "") -> bytes:
    import reports
    return reports.excel_report(res, report_date, facility)
