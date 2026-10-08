# Medicine Stock Review

A Streamlit app that reviews the e-BMSIS **View Stock Balance** report.

## What it checks

| Check | How |
|---|---|
| **Items** | Each Item Name is split into **Medicine** and **Dosage** (strength + form). Counts: stock lines, medicine names, total items (medicine + dosage), medicines with more than one dosage. |
| **Duplicates** | Same medicine **and** same dosage on more than one line = duplicate. The same medicine with a different dosage (e.g. Amoxicillin 250mg capsule vs 1000mg capsule) is a separate item, not a duplicate. |
| Duplicate type | *Exact duplicate entry* (same batch, expiry, supply type) · *Same batch under 2 supply types* · *Same batch, different expiry* (data error) · *Different batches* (normal, use FEFO) |
| **Expiry** | Expired · expires within 3 months · within 6 months (against the report date you choose) |
| **Stock level** | Only when the upload has a Balance Qty column: zero stock, and low / overstock when an average monthly use (AMC) column is also present. Quantities are added across batches of the same medicine; expired stock and exact duplicate lines are not counted as usable. |

Results show on screen and download in three formats with the same content:

- **PDF** – printable report (summary, duplicate summary and detail, same name/different dosage, expiry alerts, stock level, all items)
- **Excel** – one sheet per table, with filters
- **Interactive (HTML)** – one file that opens in any browser, also offline; search, filter and sort every table

## Files you can upload

- **PDF** – the browser printout of *Ledger → Stock → View Stock Balance*. Upload all pages at once (1–100, 101–200 …). Portrait printouts usually cut off the Balance Qty column, so only duplicates and expiry are checked.
- **Excel / CSV** – an export of the same page. Column names are detected automatically, and you can correct them on screen.

## Run on your computer

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Create a new GitHub repository and upload `app.py`, `stock_core.py`, `reports.py` and `requirements.txt`.
2. On share.streamlit.io choose **New app**, pick the repository, and set the main file to `app.py`.

## Files

- `app.py` – the screens (upload, summary, tabs, download).
- `stock_core.py` – reading files, splitting medicine name / dosage, and all the analysis rules. Change rules here.
- `reports.py` – the PDF, Excel and interactive HTML downloads.
