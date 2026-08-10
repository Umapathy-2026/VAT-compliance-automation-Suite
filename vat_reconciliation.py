"""
VAT Transaction vs Sub-ledger Reconciliation
=============================================
Reads:
    1) VAT_Transaction.xls  (or .xlsx)
    2) Sub-ledger.xlsx

Produces:
    VAT_Reconciliation_Output.xlsx

Run:
    python vat_reconciliation.py
(Place this script in the same folder as the two input files, or edit the
 file paths in the CONFIG section below.)
"""

import os
import sys
import copy
import time
import re
import random
import pandas as pd
import requests
import urllib3
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ----------------------------------------------------------------------
# CONFIG - edit these if your files live somewhere else / have different names
# ----------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

VAT_FILE = os.path.join(SCRIPT_DIR, "VAT_Transaction.xls")
SUBLEDGER_FILE = os.path.join(SCRIPT_DIR, "Sub-ledger.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "VAT_Reconciliation_Output.xlsx")
VAT_SUMMARY_TEMPLATE = os.path.join(SCRIPT_DIR, "VAT Summary.xlsx")
SUMMARY_SHEET_NAME = "VAT Summary"
LOCAL_PURCHASE_FILE = os.path.join(SCRIPT_DIR, "Local Purchase File.xlsx")
LOCAL_PURCHASE_SHEET_NAME = "Local Purchase"

# VIES (VAT Information Exchange System) - official EU API used to
# validate VAT Registration No. values online
VIES_API_URL = "https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country}/vat/{vat}"
VIES_DELAY = 1.0
VIES_MAX_RETRIES = 3
VIES_TIMEOUT_SECONDS = 8
VIES_RETRY_CODES = {
    "MS_MAX_CONCURRENT_REQ", "MS_UNAVAILABLE", "SERVICE_UNAVAILABLE",
    "GLOBAL_MAX_CONCURRENT_REQ", "TIMEOUT",
}
EU_MEMBER_STATES = {
    'AT', 'BE', 'BG', 'CY', 'CZ', 'DE', 'DK', 'EE', 'ES', 'FI', 'FR', 'GR',
    'HR', 'HU', 'IE', 'IT', 'LT', 'LU', 'LV', 'MT', 'NL', 'PL', 'PT', 'RO',
    'SE', 'SI', 'SK', 'XI',
}

TAX_REG_NO = "ESW0394952F"
TAX_CODES = [
    "ES_REG/DSG_21",
    "ES_REG/ES_DSG_0_RC",
    "ES_REG/ES_ICAG_21_RCP",
    "ES_REG/ES_ICAG_21_RCR",
]

ACCOUNT_SALES = "283-00-0000-32310000-000-46695-000"
ACCOUNT_ICAG_RCR = "283-00-0000-16030000-000-36895-000"

ROUND_DP = 2  # decimal places used when comparing amounts


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def read_vat_file(path):
    """Read the VAT Transaction file, handling legacy .xls as well as .xlsx."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        try:
            return pd.read_excel(path, dtype={"Invoice No.": str})
        except ImportError:
            # xlrd not installed -> try converting via LibreOffice, which is
            # usually available on most machines / servers.
            print("xlrd not found - attempting to convert .xls to .xlsx via LibreOffice ...")
            converted = os.path.splitext(path)[0] + ".xlsx"
            os.system(
                f'soffice --headless --convert-to xlsx --outdir "{os.path.dirname(path)}" "{path}"'
            )
            if os.path.exists(converted):
                return pd.read_excel(converted, dtype={"Invoice No.": str})
            raise RuntimeError(
                "Could not read the .xls file. Please install 'xlrd' "
                "(pip install xlrd) or re-save the file as .xlsx and update VAT_FILE."
            )
    return pd.read_excel(path, dtype={"Invoice No.": str})


def sum_ledger(sub_df, invoice_no, account_number, categories):
    """Sum 'Func Total' in the sub-ledger for a given invoice/account/category filter."""
    mask = (
        (sub_df["Invoice NO."] == invoice_no)
        & (sub_df["Account Number"] == account_number)
        & (sub_df["Category"].isin(categories))
    )
    return round(sub_df.loc[mask, "Func Total"].sum(), ROUND_DP)


def vat_description_lookup(sub_df, invoice_no):
    """
    For Case 1 mismatches: find sub-ledger rows for this invoice whose
    'Account Description' contains 'VAT', and build a
    '(Account Description_Account Number)_Value Not Matching' style comment.
    """
    mask = (sub_df["Invoice NO."] == invoice_no) & (
        sub_df["Account Description"].astype(str).str.contains("VAT", case=False, na=False)
    )
    matches = sub_df.loc[mask, ["Account Description", "Account Number"]].drop_duplicates()
    if matches.empty:
        return "No VAT account found in Sub-ledger_Value Not Matching"
    parts = [
        f"({desc}_{acc})_Value Not Matching"
        for desc, acc in zip(matches["Account Description"], matches["Account Number"])
    ]
    return "; ".join(parts)


def neutralize_formula_like_strings(ws, start_row=2):
    """
    openpyxl (and pandas' openpyxl writer) auto-detects any string cell
    value starting with '=', '+', '-', or '@' as a formula. If that string
    isn't actually valid formula syntax - e.g. a reference number, VAT
    number, or name that just happens to start with one of those characters
    - Excel will silently strip it on open with a
    "Removed Records: Formula from ..." repair warning. This scans a sheet's
    plain data (not our own deliberately-written formulas) and forces any
    such cell back to a literal string so it displays correctly instead of
    being stripped.
    """
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith(("=", "+", "-", "@")):
                cell.data_type = "s"


def write_dataframe_sheet(wb, sheet_name, df):
    """Create a new sheet (replacing one of the same name if present) and
    write a dataframe into it as plain values, with a bold header row."""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in df.itertuples(index=False):
        clean_row = [None if (isinstance(v, float) and pd.isna(v)) else v for v in row]
        ws.append(clean_row)
    neutralize_formula_like_strings(ws)
    ws.freeze_panes = "A2"
    return ws


def parse_vat(raw):
    """Split a raw VAT Registration No. into (country_code, number), or
    (None, None) if it doesn't look like a valid EU-prefixed VAT number."""
    raw = str(raw).strip().replace(" ", "").replace("-", "").upper()
    if len(raw) < 3:
        return None, None
    country, number = raw[:2], raw[2:]
    if not re.match(r'^[A-Z]{2}$', country) or country not in EU_MEMBER_STATES:
        return None, None
    return country, number


def check_vat_via_vies(vat_reg_no, cache):
    """
    Validate a single VAT Registration No. against the official EU VIES
    service (https://ec.europa.eu/taxation_customs/vies/#/vat-validation),
    with retries on VIES's known transient/overload error codes.

    The Member State is taken from the VAT number's own 2-letter country
    prefix (e.g. "ESB03933835" -> ES, "PL6443245867" -> PL), NOT hardcoded
    to Spain - VIES rejects a number if the selected member state doesn't
    match its actual prefix, so this is required for non-Spanish supplier
    VAT numbers (which do appear in this data, e.g. Polish suppliers) to
    validate correctly at all.

    Returns a short status string, e.g. "VALID", "INVALID", "INVALID - Not
    found in VIES", "Skipped - blank", or "ERROR: <details>".
    """
    key = str(vat_reg_no).strip().upper().replace(" ", "").replace("-", "")
    if not key or key == "NAN":
        return "Skipped - blank"
    if key in cache:
        return cache[key]

    country, number = parse_vat(vat_reg_no)
    if not country:
        status = "Skipped - invalid format or non-EU prefix"
        cache[key] = status
        return status

    url = VIES_API_URL.format(country=country, vat=number)
    status = "ERROR: Max retries exceeded"
    for attempt in range(1, VIES_MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url, timeout=VIES_TIMEOUT_SECONDS, verify=False,
                headers={"Accept": "application/json"},
            )
            if resp.status_code == 404:
                status = "INVALID - Not found in VIES"
                break
            if resp.status_code in (429, 503):
                if attempt < VIES_MAX_RETRIES:
                    time.sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5))
                    continue
                status = "ERROR: VIES service overloaded after retries"
                break

            resp.raise_for_status()
            data = resp.json()
            is_valid = data.get("isValid", False)
            err = (data.get("userError", "") or "").strip().upper()

            if err in VIES_RETRY_CODES:
                if attempt < VIES_MAX_RETRIES:
                    time.sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5))
                    continue
                status = f"ERROR: VIES busy ({err}) - please re-run"
                break

            if is_valid or err == "VALID":
                status = "VALID"
            elif err == "INVALID" and not is_valid:
                status = "INVALID"
            else:
                if attempt < VIES_MAX_RETRIES:
                    time.sleep(VIES_DELAY + random.uniform(0.3, 1.0))
                    continue
                status = f"INVALID - {err}" if err else "INVALID"
            break

        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt < VIES_MAX_RETRIES:
                time.sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5))
                continue
            status = "ERROR: Connection/timeout after retries"
            break
        except Exception as e:
            status = f"ERROR: {e}"
            break

    cache[key] = status
    time.sleep(VIES_DELAY)
    return status


def build_vat_validation_sheet(wb, vat_out):
    """Create the 'VAT Validation' sheet: one row per unique VAT
    Registration No. found in the reconciliation data, with its VIES
    validation status."""
    unique_vats = sorted(
        v for v in vat_out["VAT Registration No."].astype(str).str.strip().unique()
        if v and v.lower() != "nan"
    )

    print(f"Checking {len(unique_vats)} unique VAT Registration No. value(s) against VIES ...")
    cache = {}
    rows = []
    for i, vat_no in enumerate(unique_vats, start=1):
        status = check_vat_via_vies(vat_no, cache)
        print(f"  [{i}/{len(unique_vats)}] {vat_no} -> {status}")
        rows.append({"VAT Registration No.": vat_no, "VAT Status": status})

    df = pd.DataFrame(rows, columns=["VAT Registration No.", "VAT Status"])
    return write_dataframe_sheet(wb, "VAT Validation", df)


def add_vat_status_column(ws):
    """Append a 'VAT Status' column that looks up each row's VAT
    Registration No. against the VAT Validation sheet."""
    header = [c.value for c in ws[1]]
    if "VAT Registration No." not in header or ws.max_row < 2:
        return
    vat_col_letter = get_column_letter(header.index("VAT Registration No.") + 1)
    new_col_idx = ws.max_column + 1
    ws.cell(row=1, column=new_col_idx, value="VAT Status").font = Font(bold=True)
    for r in range(2, ws.max_row + 1):
        formula = (
            f"=IFERROR(VLOOKUP(TRIM({vat_col_letter}{r}),'VAT Validation'!A:B,2,FALSE),\"\")"
        )
        ws.cell(row=r, column=new_col_idx, value=formula)


def add_local_purchase_sheet(wb):
    """Copy the 'Local Purchase' sheet from Local Purchase File.xlsx into wb
    as-is (values/formulas, formatting, merged cells, column widths).
    Returns the new worksheet, or None if the source file/sheet isn't found."""
    if not os.path.exists(LOCAL_PURCHASE_FILE):
        print(f"NOTE: Local Purchase File not found at '{LOCAL_PURCHASE_FILE}' - skipping that sheet.")
        return None

    source_wb = load_workbook(LOCAL_PURCHASE_FILE)
    if LOCAL_PURCHASE_SHEET_NAME not in source_wb.sheetnames:
        print(f"NOTE: Sheet '{LOCAL_PURCHASE_SHEET_NAME}' not found in '{LOCAL_PURCHASE_FILE}' - skipping.")
        return None

    source_ws = source_wb[LOCAL_PURCHASE_SHEET_NAME]
    if LOCAL_PURCHASE_SHEET_NAME in wb.sheetnames:
        del wb[LOCAL_PURCHASE_SHEET_NAME]
    return copy_sheet(source_ws, wb, LOCAL_PURCHASE_SHEET_NAME)


def copy_sheet(source_ws, target_wb, title):
    """Copy a worksheet (values/formulas, styles, merged cells, column widths,
    row heights) from another workbook into target_wb as a new sheet."""
    target_ws = target_wb.create_sheet(title=title)

    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.border = copy.copy(cell.border)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.alignment = copy.copy(cell.alignment)

    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    for col_letter, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_idx].height = dim.height

    return target_ws


def add_vat_summary_sheet(wb, data_sheet_name, vat_out):
    """Copy the VAT Summary template into wb and populate the required
    formula cells, referencing the reconciliation data on `data_sheet_name`."""
    if not os.path.exists(VAT_SUMMARY_TEMPLATE):
        print(f"NOTE: VAT Summary template not found at '{VAT_SUMMARY_TEMPLATE}' - skipping summary sheet.")
        return

    if SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[SUMMARY_SHEET_NAME]

    template_wb = load_workbook(VAT_SUMMARY_TEMPLATE)
    template_ws = template_wb[SUMMARY_SHEET_NAME] if SUMMARY_SHEET_NAME in template_wb.sheetnames else template_wb.active
    ws = copy_sheet(template_ws, wb, SUMMARY_SHEET_NAME)

    columns = vat_out.columns
    # Resolve the data sheet's column letters dynamically, so this keeps
    # working even if OUTPUT_COLUMNS is reordered later.
    col = {name: get_column_letter(i + 1) for i, name in enumerate(columns)}
    taxable = col["Taxable Amt (Tax Curr)"]
    recoverable = col["Recoverable Amt (Tax Curr)"]
    doc_type = col["Doc type"]
    tax_code = col["Tax Code"]
    supplier = col["Customer/Supplier"]

    def sumifs(sum_col, *conditions, negate=False):
        parts = [f"'{data_sheet_name}'!${sum_col}:${sum_col}"]
        for cond_col, cond_val in conditions:
            parts.append(f"'{data_sheet_name}'!${cond_col}:${cond_col}")
            parts.append(f'"{cond_val}"')
        formula = "SUMIFS(" + ",".join(parts) + ")"
        return f"=({formula})*-1" if negate else f"={formula}"

    # Local sales, credit notes - box07/box09
    ws["C9"] = sumifs(taxable, (doc_type, "Credit note"), (tax_code, "ES_REG/DSG_21"), negate=True)
    ws["I9"] = sumifs(recoverable, (doc_type, "Credit note"), (tax_code, "ES_REG/DSG_21"), negate=True)

    # Intra-community acquisition (Johnson Electric Poland, RCR) - appears in
    # both the output-VAT section (row 10) and the deductible section (row 23)
    ws["C10"] = sumifs(taxable, (supplier, "JOHNSON ELECTRIC POLAND SP ZOO"), (tax_code, "ES_REG/ES_ICAG_21_RCR"))
    ws["I10"] = sumifs(recoverable, (supplier, "JOHNSON ELECTRIC POLAND SP ZOO"), (tax_code, "ES_REG/ES_ICAG_21_RCR"))
    ws["C23"] = "=C10"
    ws["I23"] = "=I10"

    # Domestic sales, invoices - box07/box09
    ws["C12"] = sumifs(taxable, (doc_type, "Invoice"), (tax_code, "ES_REG/DSG_21"), negate=True)
    ws["I12"] = sumifs(recoverable, (doc_type, "Invoice"), (tax_code, "ES_REG/DSG_21"), negate=True)

    # Totals and balance
    # NOTE: interpreted "Cell I7" / "Cell I29" as the *Total* rows (I17 / I29),
    # since I7 is itself one of the cells being summed (I7:K16) - using I7 as
    # both the target and part of the range would be a circular reference.
    # If you actually meant a different cell, let me know and I'll adjust.
    ws["I17"] = "=SUM(I7:K16)"
    ws["I29"] = "=SUM(I19:K28)"
    ws["I30"] = "=I17-I29"
    ws["I36"] = "=I17-I29"
    ws["I37"] = "=I17-I29"

    # Local sales subject to reverse charge - box122 holds the taxable base directly (no filter/sign flip)
    ws["I34"] = sumifs(taxable, (tax_code, "ES_REG/ES_DSG_0_RC"))

    # Total VAT amount / amount to be paid or reclaimed
    ws["I42"] = "=I37+I38-I40"
    ws["I44"] = "=I37+I38-I40"

    # J2 - "Period: Q<n>_<year>", derived from the latest
    # 'Registered Date / Invoice Creation Date:' in the data sheet. Computed
    # in Python (not as a live formula) so it's not affected by how Excel
    # happens to interpret the date column's type.
    date_col = "Registered Date / Invoice Creation Date:"
    period_label = "Period: "
    if date_col in vat_out.columns:
        reg_dates = pd.to_datetime(vat_out[date_col], errors="coerce")
        if reg_dates.notna().any():
            period_date = reg_dates.max()
            quarter = (period_date.month - 1) // 3 + 1
            period_label = f"Period: Q{quarter}_{period_date.year}"
    ws["J2"] = period_label


def main():
    print("Reading VAT Transaction file ...")
    vat = read_vat_file(VAT_FILE)

    print("Reading Sub-ledger Report file ...")
    sub = pd.read_excel(SUBLEDGER_FILE, dtype={"Invoice NO.": str})

    # --------------------------------------------------------------
    # Data-quality fix: the source export shifts Curr..Journal Description
    # one column to the right whenever 'Journal Description' is blank,
    # spilling the true value into a stray 'Unnamed: 24' column. This
    # means 'Func Total' is wrong for affected rows - the real value is
    # sitting in 'Description' instead. Detect and correct it here so
    # every downstream calculation uses the right number.
    # --------------------------------------------------------------
    if "Unnamed: 24" in sub.columns:
        shifted = sub["Curr"].isna() & sub["Unnamed: 24"].notna()
        n_shifted = int(shifted.sum())
        if n_shifted:
            print(f"NOTE: {n_shifted} sub-ledger rows have a column-shift export "
                  f"defect - correcting 'Func Total' from the 'Description' column.")
            sub.loc[shifted, "Func Total"] = pd.to_numeric(
                sub.loc[shifted, "Description"], errors="coerce"
            )

    # Make sure key join columns are clean strings
    vat["Invoice No."] = vat["Invoice No."].astype(str).str.strip()
    sub["Invoice NO."] = sub["Invoice NO."].astype(str).str.strip()
    sub["Account Number"] = sub["Account Number"].astype(str).str.strip()

    # --------------------------------------------------------------
    # Step 1: Filter File 1 (VAT Transaction)
    # --------------------------------------------------------------
    vat_f = vat[
        (vat["Our Tax Registration No."] == TAX_REG_NO) & (vat["Tax Code"].isin(TAX_CODES))
    ].copy()

    if vat_f.empty:
        print("WARNING: No rows matched the VAT filter conditions. Exiting.")
        sys.exit(1)

    # Sum per Invoice No. + Tax Code + Inv Curr (also used for the case-by-case matching below)
    invoice_taxcode_totals = (
        vat_f.groupby(["Invoice No.", "Tax Code", "Inv Curr"])["Recoverable Amt (Tax Curr)"]
        .sum()
        .round(ROUND_DP)
    )
    vat_f["Recoverable Amt (Tax Curr)"] = vat_f.apply(
        lambda r: invoice_taxcode_totals.get((r["Invoice No."], r["Tax Code"], r["Inv Curr"]), 0), axis=1
    )

    # Same grouping logic applied to Taxable Amt (Tax Curr)
    invoice_taxcode_taxable_totals = (
        vat_f.groupby(["Invoice No.", "Tax Code", "Inv Curr"])["Taxable Amt (Tax Curr)"]
        .sum()
        .round(ROUND_DP)
    )
    vat_f["Taxable Amt (Tax Curr)"] = vat_f.apply(
        lambda r: invoice_taxcode_taxable_totals.get((r["Invoice No."], r["Tax Code"], r["Inv Curr"]), 0), axis=1
    )

    # VAT rate % = Recoverable Amt (Tax Curr) / Taxable Amt (Tax Curr) * 100
    def calc_vat_rate(row):
        taxable = row["Taxable Amt (Tax Curr)"]
        recoverable = row["Recoverable Amt (Tax Curr)"]
        if taxable == 0:
            return 0
        return round((recoverable / taxable) * 100, ROUND_DP)

    vat_f["VAT rate %"] = vat_f.apply(calc_vat_rate, axis=1)

    # Doc type: Credit note if Taxable Amt (Tax Curr) < 0, else Invoice
    vat_f["Doc type"] = vat_f["Taxable Amt (Tax Curr)"].apply(lambda x: "Credit note" if x < 0 else "Invoice")

    # --------------------------------------------------------------
    # Step 2: Filter File 2 (Sub-ledger Report)
    # --------------------------------------------------------------
    sub_f = sub[sub["Category"] != "Addition"].copy()

    # --------------------------------------------------------------
    # Step 3: Case-by-case comparison / tagging
    # --------------------------------------------------------------
    comments = []
    mappings = []
    func_totals = []  # NEW: sum of 'Func Total' from File 2 for the applicable condition

    for _, row in vat_f.iterrows():
        inv = row["Invoice No."]
        tax_code = row["Tax Code"]
        inv_curr = row["Inv Curr"]
        vat_amt = invoice_taxcode_totals.get((inv, tax_code, inv_curr), 0)

        vat_rate = row["VAT rate %"]

        if tax_code == "ES_REG/DSG_21":
            # Case 1 - value must match AND VAT rate % must be 21
            ledger_amt = sum_ledger(sub_f, inv, ACCOUNT_SALES, ["Sales Invoices", "Credit Memos"])
            if vat_amt == ledger_amt and vat_rate == 21:
                comments.append("Okay-Value Matching")
            else:
                # not matching -> same lookup logic as before
                comments.append(vat_description_lookup(sub_f, inv))
            mappings.append("Domestic sale of Goods")
            func_totals.append(ledger_amt)

        elif tax_code == "ES_REG/ES_DSG_0_RC":
            # Case 2 - no ledger lookup performed; VAT rate % must be 0
            if vat_rate == 0:
                comments.append("Okay")
            else:
                comments.append("Not Okay - VAT rate % must be 0")
            mappings.append("Domestic sale of Goods-Reverse charge")
            func_totals.append(None)

        elif tax_code == "ES_REG/ES_ICAG_21_RCP":
            # Case 3 - value must match AND VAT rate % must be 0
            ledger_amt = sum_ledger(sub_f, inv, ACCOUNT_SALES, ["Purchase Invoices"])
            if vat_amt == ledger_amt and vat_rate == 0:
                comments.append("Okay-Value Matching")
            else:
                comments.append("Value Not Matching")
            mappings.append("Intra community acquisition of goods- Reverse charge payable")
            func_totals.append(ledger_amt)

        elif tax_code == "ES_REG/ES_ICAG_21_RCR":
            # Case 4 - value must match AND VAT rate % must be 21
            ledger_amt = sum_ledger(sub_f, inv, ACCOUNT_ICAG_RCR, ["Purchase Invoices"])
            if vat_amt == ledger_amt and vat_rate == 21:
                comments.append("Okay-Value Matching")
            else:
                comments.append("Value Not Matching")
            mappings.append("Intra community acquisition of goods - Reverse charge Recoverable")
            func_totals.append(ledger_amt)

        else:
            comments.append("")
            mappings.append("")
            func_totals.append(None)

    vat_f["Comments"] = comments
    vat_f["Mapping"] = mappings
    vat_f["Func Total"] = func_totals

    # --------------------------------------------------------------
    # Step 4: Keep only the requested columns, in order
    # --------------------------------------------------------------
    OUTPUT_COLUMNS = [
        "Doc type",
        "Invoice No.",
        "Invoice Date",
        "Registered Date / Invoice Creation Date:",
        "Delivery Date",
        "Incoterm",
        "Customer/Supplier No",
        "Customer/Supplier",
        "Our Tax Registration No.",
        "VAT Registration No.",
        "Tax Code",
        "Taxable Amt (Tax Curr)",
        "Recoverable Amt (Tax Curr)",
        "Func Total",
        "Inv Curr",
        "VAT rate %",
        "Mapping",
        "Comments",
    ]
    vat_out = vat_f[OUTPUT_COLUMNS].copy()

    # Keep only unique records per Invoice No. + Tax Code + Inv Curr
    # (the source VAT file can contain multiple transaction lines for the
    # same invoice/tax code/currency combination; the aggregate amount,
    # Comments and Mapping are identical for all of them, so one row is kept)
    vat_out = vat_out.drop_duplicates(subset=["Invoice No.", "Tax Code", "Inv Curr"], keep="first").reset_index(
        drop=True
    )

    # --------------------------------------------------------------
    # Step 5: Save output
    # --------------------------------------------------------------
    DATA_SHEET_NAME = "Overall VAT Comparison"
    vat_out.to_excel(OUTPUT_FILE, index=False, sheet_name=DATA_SHEET_NAME)

    wb = load_workbook(OUTPUT_FILE)
    ws = wb[DATA_SHEET_NAME]
    neutralize_formula_like_strings(ws)

    def apply_highlighting(target_ws):
        """
        Yellow:
          - blank 'VAT Registration No.'
          - 'Invoice Date' falls in a different quarter than
            'Registered Date / Invoice Creation Date:' on the same row
          - for Tax Code ES_REG/DSG_21 or ES_REG/ES_DSG_0_RC, 'VAT
            Registration No.' does not start with "ES"
        Red: 'Value Not Matching' in Comments.
        """
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        header = [c.value for c in target_ws[1]]
        if "VAT Registration No." not in header or "Comments" not in header:
            return

        vat_col_idx = header.index("VAT Registration No.") + 1
        comments_col_idx = header.index("Comments") + 1
        invoice_date_col_idx = header.index("Invoice Date") + 1 if "Invoice Date" in header else None
        reg_date_col_idx = (
            header.index("Registered Date / Invoice Creation Date:") + 1
            if "Registered Date / Invoice Creation Date:" in header else None
        )
        tax_code_col_idx = header.index("Tax Code") + 1 if "Tax Code" in header else None

        def quarter_of(value):
            dt = pd.to_datetime(value, errors="coerce")
            if pd.isna(dt):
                return None
            return (dt.year, (dt.month - 1) // 3 + 1)

        ES_PREFIX_TAX_CODES = {"ES_REG/DSG_21", "ES_REG/ES_DSG_0_RC"}

        for r in range(2, target_ws.max_row + 1):
            vat_cell = target_ws.cell(row=r, column=vat_col_idx)
            v = vat_cell.value
            is_blank = v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""
            if is_blank:
                vat_cell.fill = yellow
            elif tax_code_col_idx:
                # Check 2: for domestic tax codes, VAT Registration No. must start with "ES"
                tax_code = str(target_ws.cell(row=r, column=tax_code_col_idx).value or "").strip()
                if tax_code in ES_PREFIX_TAX_CODES and not str(v).strip().upper().startswith("ES"):
                    vat_cell.fill = yellow

            comment_cell = target_ws.cell(row=r, column=comments_col_idx)
            if comment_cell.value and "Value Not Matching" in str(comment_cell.value):
                comment_cell.fill = red

            # Check 1: Invoice Date must fall in the same quarter as
            # Registered Date / Invoice Creation Date:
            if invoice_date_col_idx and reg_date_col_idx:
                invoice_date_cell = target_ws.cell(row=r, column=invoice_date_col_idx)
                reg_date_cell = target_ws.cell(row=r, column=reg_date_col_idx)
                inv_q = quarter_of(invoice_date_cell.value)
                reg_q = quarter_of(reg_date_cell.value)
                if inv_q is not None and reg_q is not None and inv_q != reg_q:
                    invoice_date_cell.fill = yellow

    apply_highlighting(ws)

    # --------------------------------------------------------------
    # Step 6: Tax-code specific breakout sheets (same columns/rows as the
    # main comparison sheet, just split out by tax code for easier review) -
    # with the same yellow/red highlighting applied
    # --------------------------------------------------------------
    ws_ec = write_dataframe_sheet(wb, "EC acquisition", vat_out[vat_out["Tax Code"] == "ES_REG/ES_ICAG_21_RCR"])
    ws_dom21 = write_dataframe_sheet(wb, "Domestic sales (21% VAT)", vat_out[vat_out["Tax Code"] == "ES_REG/DSG_21"])
    ws_dom0 = write_dataframe_sheet(wb, "Domestic sales (0% VAT)", vat_out[vat_out["Tax Code"] == "ES_REG/ES_DSG_0_RC"])
    for breakout_ws in (ws_ec, ws_dom21, ws_dom0):
        apply_highlighting(breakout_ws)

    # --------------------------------------------------------------
    # Step 7: Sub-ledger invoices that never showed up in the VAT
    # Transaction file at all, restricted to the two VAT accounts (Sales VAT
    # payable and ICAG VAT receivable) - so this only flags invoices where
    # the missing GL posting is actually VAT-relevant. Checked against every
    # Spain invoice in the VAT file (not just the 4 tax codes handled above),
    # so an invoice posted under a different tax code isn't wrongly flagged.
    # Sub-ledger rows with no Invoice NO. are excluded (not invoice-driven).
    # --------------------------------------------------------------
    MISSED_INV_ACCOUNTS = [ACCOUNT_SALES, ACCOUNT_ICAG_RCR]
    spain_invoices = set(
        vat.loc[vat["Our Tax Registration No."] == TAX_REG_NO, "Invoice No."].astype(str).str.strip()
    )
    missed = sub[
        sub["Invoice NO."].notna()
        & (sub["Invoice NO."].astype(str).str.strip() != "")
        & (~sub["Invoice NO."].isin(spain_invoices))
        & (sub["Account Number"].isin(MISSED_INV_ACCOUNTS))
    ].copy()
    write_dataframe_sheet(wb, "Missed Inv in VAT Transaction", missed)

    # --------------------------------------------------------------
    # Step 7c: Build the VAT Validation sheet (unique VAT Registration No.
    # values checked against the official EU VIES service), then propagate
    # a "VAT Status" lookup column back onto the 4 comparison sheets.
    # --------------------------------------------------------------
    build_vat_validation_sheet(wb, vat_out)
    for status_ws in (ws, ws_ec, ws_dom21, ws_dom0):
        add_vat_status_column(status_ws)

    # --------------------------------------------------------------
    # Step 7b: Copy in the "Local Purchase" sheet from Local Purchase File.xlsx
    # (values/formulas, formatting, merged cells all preserved as-is)
    # --------------------------------------------------------------
    local_purchase_ws = add_local_purchase_sheet(wb)

    add_vat_summary_sheet(wb, ws.title, vat_out)

    # VAT Summary C19/I19 = live references to Local Purchase!H11/I11, so
    # they stay in sync with that sheet without needing to re-run this script.
    if local_purchase_ws is not None:
        summary_ws = wb[SUMMARY_SHEET_NAME]
        summary_ws["C19"] = f"='{LOCAL_PURCHASE_SHEET_NAME}'!H9"
        summary_ws["I19"] = f"='{LOCAL_PURCHASE_SHEET_NAME}'!I9"

    # --------------------------------------------------------------
    # Step 8: Set the final tab order
    # --------------------------------------------------------------
    SHEET_ORDER = [
        "VAT Summary",
        "Overall VAT Comparison",
        "EC acquisition",
        "Local Purchase",
        "Domestic sales (21% VAT)",
        "Domestic sales (0% VAT)",
        "Missed Inv in VAT Transaction",
        "VAT Validation",
    ]
    wb._sheets = [wb[name] for name in SHEET_ORDER if name in wb.sheetnames] + [
        wb[name] for name in wb.sheetnames if name not in SHEET_ORDER
    ]

    wb.save(OUTPUT_FILE)

    print(f"Done. Output saved to: {OUTPUT_FILE}")
    print(f"Rows in output: {len(vat_out)}")


if __name__ == "__main__":
    main()
