"""
VAT Transaction vs Sub-ledger Reconciliation - processing engine
================================================================
Pure processing logic, with no console or UI dependencies. The user interface
(see vat_reconciliation_tab.py) drives this module by calling
`run_reconciliation(...)` on a worker thread and passing callbacks for
progress reporting and cancellation.

Inputs supplied by the user at run time:
    1) The quarter's VAT Transaction report exports - several files, which this
       tool consolidates (see vat_source_loader.py)
    2) The Sub-ledger report export for the quarter - a single file

The blank "VAT Summary" layout is *not* an input - it is embedded in
vat_summary_template_data.py, so the tool has no external template dependency.

Output:
    A single .xlsx workbook containing the comparison sheets, tax-code
    breakouts, missed-invoice listing, VIES VAT validation and the populated
    VAT Summary.
"""

import copy
import datetime
import os
import random
import re
import time

import pandas as pd
import requests
import urllib3
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from vat_source_loader import (
    SUBLEDGER_KEY_COLUMNS,
    VAT_KEY_COLUMNS,
    SourceFileError,
    load_and_consolidate,
)
from vat_summary_template_data import SHEET_NAME as SUMMARY_SHEET_NAME
from vat_summary_template_data import load_template_worksheet

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ----------------------------------------------------------------------
# Reconciliation rules - business configuration
# ----------------------------------------------------------------------
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

# All dates in these reports are calendar days with no meaningful time-of-day
# (the source systems write midnight for every value), so cells are formatted
# to show just the date rather than Excel's default "2/9/2026 12:00:00 AM".
DATE_NUMBER_FORMAT = "M/D/YYYY"

DATA_SHEET_NAME = "Overall VAT Comparison"

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

SHEET_ORDER = [
    SUMMARY_SHEET_NAME,
    DATA_SHEET_NAME,
    "EC acquisition",
    LOCAL_PURCHASE_SHEET_NAME,
    "Domestic sales (21% VAT)",
    "Domestic sales (0% VAT)",
    "Missed Inv in VAT Transaction",
    "VAT Validation",
]

# Columns the engine requires from each input file. Checked up-front so a
# wrong file selection fails with a clear message instead of a KeyError
# halfway through processing.
REQUIRED_VAT_COLUMNS = [
    "Our Tax Registration No.", "Tax Code", "Invoice No.", "Inv Curr",
    "Taxable Amt (Tax Curr)", "Recoverable Amt (Tax Curr)",
]
REQUIRED_SUBLEDGER_COLUMNS = [
    "Invoice NO.", "Account Number", "Category", "Func Total",
    "Account Description",
]


class ReconciliationCancelled(Exception):
    """Raised internally when the caller's cancel check returns True."""


class ReconciliationError(Exception):
    """A problem with the inputs that the user can act on (bad file, missing
    columns, no matching rows). Carries a message safe to show in the UI."""


class ReconciliationResult:
    """Summary of a completed run, returned to the UI for display."""

    def __init__(self, output_file, row_count, vies_counts, notes):
        self.output_file = output_file
        self.row_count = row_count
        self.vies_counts = vies_counts
        self.notes = notes


# ----------------------------------------------------------------------
# Progress plumbing
# ----------------------------------------------------------------------
class _Reporter:
    """Wraps the caller's optional log/progress/cancel callbacks so the
    processing code below can call them unconditionally."""

    def __init__(self, log=None, progress=None, should_cancel=None):
        self._log = log
        self._progress = progress
        self._should_cancel = should_cancel
        self.notes = []

    def log(self, message, level="info"):
        if level == "note":
            self.notes.append(message)
        if self._log:
            self._log(message, level)

    def progress(self, fraction, message=None):
        """fraction is 0.0-1.0; message updates the status line if given."""
        if self._progress:
            self._progress(max(0.0, min(1.0, fraction)), message)

    def check_cancelled(self):
        if self._should_cancel and self._should_cancel():
            raise ReconciliationCancelled()


# ----------------------------------------------------------------------
# Input reading
# ----------------------------------------------------------------------
def _consolidate(paths, key_columns, label, reporter, progress_span):
    """Read and concatenate one set of quarterly report exports.

    Wraps the loader so its progress is folded into the overall progress bar
    and its errors surface as user-facing ReconciliationErrors.
    """
    start, end = progress_span

    def part_progress(done, total, message):
        reporter.progress(start + (end - start) * (done / max(total, 1)), message)

    reporter.log(f"Reading {len(paths)} {label} file(s):")
    try:
        return load_and_consolidate(
            paths,
            key_columns,
            label,
            log=reporter.log,
            progress=part_progress,
            check_cancelled=reporter.check_cancelled,
        )
    except SourceFileError as exc:
        raise ReconciliationError(str(exc))


def require_columns(df, required, file_label):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ReconciliationError(
            f"The {file_label} file is missing expected column(s):\n\n"
            + "\n".join(f"  - {c}" for c in missing)
            + "\n\nPlease check that the correct file was selected."
        )


# ----------------------------------------------------------------------
# Reconciliation helpers
# ----------------------------------------------------------------------
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


# ----------------------------------------------------------------------
# Workbook writing helpers
# ----------------------------------------------------------------------
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


def excel_safe(value):
    """Convert a dataframe value into something openpyxl can write.

    Missing values come in several flavours - float('nan'), pd.NA (from
    pandas' nullable dtypes) and pd.NaT - and openpyxl rejects the last two
    with "Cannot convert <NA> to Excel". Testing for float NaN alone is not
    enough; every one of them must become a blank cell.
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        # pd.isna returns an array for list/array-like values rather than a
        # single bool; those are not missing values, so pass them through.
        pass
    return value


def apply_date_number_format(ws, start_row=2):
    """Format every date/datetime cell to show a plain date, no time-of-day.

    Applied per-cell rather than per-column-dtype so it also catches columns
    like the sub-ledger's 'Period', which can hold a genuine datetime in one
    row and plain text in another once several report parts are concatenated.
    """
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if isinstance(cell.value, datetime.datetime):
                cell.number_format = DATE_NUMBER_FORMAT


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
        ws.append([excel_safe(v) for v in row])
    neutralize_formula_like_strings(ws)
    apply_date_number_format(ws)
    ws.freeze_panes = "A2"
    return ws


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


# ----------------------------------------------------------------------
# VIES VAT number validation
# ----------------------------------------------------------------------
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


def _interruptible_sleep(seconds, reporter):
    """Sleep in short slices so a cancel request is picked up promptly even
    during VIES's multi-second retry back-offs."""
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        reporter.check_cancelled()
        time.sleep(min(0.2, deadline - time.monotonic()))


def check_vat_via_vies(vat_reg_no, cache, reporter):
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
        reporter.check_cancelled()
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
                    _interruptible_sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5), reporter)
                    continue
                status = "ERROR: VIES service overloaded after retries"
                break

            resp.raise_for_status()
            data = resp.json()
            is_valid = data.get("isValid", False)
            err = (data.get("userError", "") or "").strip().upper()

            if err in VIES_RETRY_CODES:
                if attempt < VIES_MAX_RETRIES:
                    _interruptible_sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5), reporter)
                    continue
                status = f"ERROR: VIES busy ({err}) - please re-run"
                break

            if is_valid or err == "VALID":
                status = "VALID"
            elif err == "INVALID" and not is_valid:
                status = "INVALID"
            else:
                if attempt < VIES_MAX_RETRIES:
                    _interruptible_sleep(VIES_DELAY + random.uniform(0.3, 1.0), reporter)
                    continue
                status = f"INVALID - {err}" if err else "INVALID"
            break

        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt < VIES_MAX_RETRIES:
                _interruptible_sleep(min(2 ** attempt, 8) + random.uniform(0.5, 1.5), reporter)
                continue
            status = "ERROR: Connection/timeout after retries"
            break
        except ReconciliationCancelled:
            raise
        except Exception as e:
            status = f"ERROR: {e}"
            break

    cache[key] = status
    _interruptible_sleep(VIES_DELAY, reporter)
    return status


def build_vat_validation_sheet(wb, vat_out, reporter, progress_span=(0.35, 0.85)):
    """Create the 'VAT Validation' sheet: one row per unique VAT
    Registration No. found in the reconciliation data, with its VIES
    validation status. Returns (worksheet, status counts)."""
    # dropna() before astype(str): pandas 3 keeps NaN as NaN through
    # astype(str)/str.strip() (pandas 2 turned it into the string "nan"), so
    # blanks must be removed explicitly rather than filtered out by name.
    unique_vats = sorted(
        v for v in vat_out["VAT Registration No."].dropna().astype(str).str.strip().unique()
        if v and v.lower() != "nan"
    )

    total = len(unique_vats)
    start, end = progress_span
    reporter.log(f"Checking {total} unique VAT Registration No. value(s) against VIES ...")

    cache = {}
    rows = []
    counts = {}
    for i, vat_no in enumerate(unique_vats, start=1):
        reporter.check_cancelled()
        status = check_vat_via_vies(vat_no, cache, reporter)
        # Bucket the free-text statuses into VALID / INVALID / SKIPPED / ERROR
        # so the UI can show a one-line outcome summary.
        bucket = status.split(" - ")[0].split(":")[0].strip().upper()
        counts[bucket] = counts.get(bucket, 0) + 1
        reporter.log(f"  [{i}/{total}] {vat_no} -> {status}")
        reporter.progress(
            start + (end - start) * (i / total if total else 1),
            f"Validating VAT numbers against VIES ({i}/{total})",
        )
        rows.append({"VAT Registration No.": vat_no, "VAT Status": status})

    df = pd.DataFrame(rows, columns=["VAT Registration No.", "VAT Status"])
    return write_dataframe_sheet(wb, "VAT Validation", df), counts


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


# ----------------------------------------------------------------------
# Local Purchase sheet (optional)
# ----------------------------------------------------------------------
def insert_blank_column(ws, header_name, after_column_name):
    """Insert an empty column right after `after_column_name`, header only.

    For a column the team fills in by hand during review (no source data to
    populate it from), rather than one the tool computes. Existing columns
    shift right - safe to call before highlighting/VAT-status are added,
    since those locate columns by header name rather than position.
    """
    header = [c.value for c in ws[1]]
    if after_column_name not in header:
        return
    insert_idx = header.index(after_column_name) + 2  # 1-based, right after
    ws.insert_cols(insert_idx)
    ws.cell(row=1, column=insert_idx, value=header_name).font = Font(bold=True)


def add_local_purchase_sheet(wb, local_purchase_file, reporter):
    """Copy the 'Local Purchase' sheet from the given workbook into wb as-is
    (values/formulas, formatting, merged cells, column widths).

    Optional: the Local Purchase source is not one of the two required inputs,
    so a missing file simply means the sheet is omitted. Returns the new
    worksheet, or None if the file/sheet isn't available."""
    if not local_purchase_file:
        return None
    if not os.path.exists(local_purchase_file):
        reporter.log(
            f"Local Purchase file not found at '{local_purchase_file}' - skipping that sheet.",
            "note",
        )
        return None

    source_wb = load_workbook(local_purchase_file)
    if LOCAL_PURCHASE_SHEET_NAME not in source_wb.sheetnames:
        reporter.log(
            f"Sheet '{LOCAL_PURCHASE_SHEET_NAME}' not found in "
            f"'{os.path.basename(local_purchase_file)}' - skipping that sheet.",
            "note",
        )
        return None

    source_ws = source_wb[LOCAL_PURCHASE_SHEET_NAME]
    if LOCAL_PURCHASE_SHEET_NAME in wb.sheetnames:
        del wb[LOCAL_PURCHASE_SHEET_NAME]
    return copy_sheet(source_ws, wb, LOCAL_PURCHASE_SHEET_NAME)


# ----------------------------------------------------------------------
# VAT Summary sheet
# ----------------------------------------------------------------------
def add_vat_summary_sheet(wb, data_sheet_name, vat_out):
    """Copy the embedded VAT Summary template into wb and populate the
    required formula cells, referencing the reconciliation data on
    `data_sheet_name`."""
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        del wb[SUMMARY_SHEET_NAME]

    ws = copy_sheet(load_template_worksheet(), wb, SUMMARY_SHEET_NAME)

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
    return ws


# ----------------------------------------------------------------------
# Highlighting
# ----------------------------------------------------------------------
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


# ----------------------------------------------------------------------
# Main entry point
# ----------------------------------------------------------------------
def build_comparison_frame(vat, sub, reporter):
    """Steps 1-4: filter, aggregate, tag and trim down to the output columns.
    Returns (vat_out, sub_f, spain_invoices)."""
    # --------------------------------------------------------------
    # Data-quality fix: the source export shifts Curr..Journal Description
    # one column to the right whenever 'Journal Description' is blank,
    # spilling the true value into a stray 'Unnamed: 24' column. This
    # means 'Func Total' is wrong for affected rows - the real value is
    # sitting in 'Description' instead. Detect and correct it here so
    # every downstream calculation uses the right number.
    # --------------------------------------------------------------
    # The overflow column is whatever unnamed column the export spilled into.
    # Its index depends on how many columns that report version has (it was
    # "Unnamed: 24" on the 25-column layout), so find it rather than hardcode it.
    overflow_cols = [c for c in sub.columns if re.fullmatch(r"Unnamed: \d+", str(c))]
    overflow_col = overflow_cols[0] if overflow_cols else None

    if overflow_col and "Curr" in sub.columns and "Description" in sub.columns:
        shifted = sub["Curr"].isna() & sub[overflow_col].notna()
        n_shifted = int(shifted.sum())
        if n_shifted:
            reporter.log(
                f"{n_shifted} sub-ledger rows have a column-shift export defect "
                f"(spilled into '{overflow_col}') - corrected 'Func Total' from "
                f"the 'Description' column.",
                "note",
            )
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
        raise ReconciliationError(
            "No rows in the VAT Transaction file matched the filter conditions "
            f"(Our Tax Registration No. = {TAX_REG_NO} and one of the four "
            "configured tax codes).\n\nCheck that the correct VAT Transaction "
            "file and reporting period were selected."
        )

    reporter.log(f"{len(vat_f)} VAT transaction line(s) matched the filter.")
    reporter.progress(0.20, "Aggregating amounts per invoice ...")

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
    vat_f["Doc type"] = vat_f["Taxable Amt (Tax Curr)"].apply(
        lambda x: "Credit note" if x < 0 else "Invoice"
    )

    # --------------------------------------------------------------
    # Step 2: Filter File 2 (Sub-ledger Report)
    # --------------------------------------------------------------
    sub_f = sub[sub["Category"] != "Addition"].copy()

    # --------------------------------------------------------------
    # Step 3: Case-by-case comparison / tagging
    # --------------------------------------------------------------
    reporter.progress(0.25, "Comparing VAT transactions against the sub-ledger ...")
    comments = []
    mappings = []
    func_totals = []  # sum of 'Func Total' from File 2 for the applicable condition

    total_rows = len(vat_f)
    for i, (_, row) in enumerate(vat_f.iterrows(), start=1):
        if i % 200 == 0:
            reporter.check_cancelled()
            reporter.progress(
                0.25 + 0.10 * (i / total_rows),
                f"Comparing VAT transactions against the sub-ledger ({i}/{total_rows}) ...",
            )
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
    missing_out = [c for c in OUTPUT_COLUMNS if c not in vat_f.columns]
    if missing_out:
        raise ReconciliationError(
            "The VAT Transaction file is missing column(s) needed for the output:\n\n"
            + "\n".join(f"  - {c}" for c in missing_out)
        )
    vat_out = vat_f[OUTPUT_COLUMNS].copy()

    # Keep only unique records per Invoice No. + Tax Code + Inv Curr
    # (the source VAT file can contain multiple transaction lines for the
    # same invoice/tax code/currency combination; the aggregate amount,
    # Comments and Mapping are identical for all of them, so one row is kept)
    vat_out = vat_out.drop_duplicates(
        subset=["Invoice No.", "Tax Code", "Inv Curr"], keep="first"
    ).reset_index(drop=True)

    return vat_out, sub_f


def run_reconciliation(
    vat_files,
    subledger_file,
    output_file,
    local_purchase_file=None,
    log=None,
    progress=None,
    should_cancel=None,
):
    """Run the full reconciliation and write `output_file`.

    Args:
        vat_files: list of VAT Transaction report exports covering the quarter
            (typically three). Concatenated in the order given. A single path
            may be passed instead of a list.
        subledger_file: the Sub-ledger report export for the whole quarter -
            one file. A list is accepted too, and its parts are concatenated,
            for the case where the report has to be pulled in pieces again.
        output_file: path of the .xlsx workbook to write.
        local_purchase_file: optional Local Purchase workbook whose
            'Local Purchase' sheet is copied in verbatim. Omit to skip.
        log: callable(message, level) - level is "info", "note" or "error".
        progress: callable(fraction_0_to_1, status_message_or_None).
        should_cancel: callable() -> bool, polled between steps.

    Returns:
        ReconciliationResult

    Raises:
        ReconciliationError on actionable input problems,
        ReconciliationCancelled if the caller cancelled.
    """
    reporter = _Reporter(log, progress, should_cancel)

    # Accept a bare path as well as a list, so either report can be supplied
    # as one file or as several parts.
    if isinstance(vat_files, (str, os.PathLike)):
        vat_files = [vat_files]
    if isinstance(subledger_file, (str, os.PathLike)):
        subledger_files = [subledger_file]
    else:
        subledger_files = list(subledger_file)

    reporter.progress(0.02, "Reading VAT Transaction files ...")
    vat = _consolidate(
        vat_files, VAT_KEY_COLUMNS, "VAT Transaction", reporter, (0.02, 0.09)
    )
    require_columns(vat, REQUIRED_VAT_COLUMNS, "VAT Transaction")
    reporter.check_cancelled()

    reporter.progress(0.10, "Reading Sub-ledger file ...")
    sub = _consolidate(
        subledger_files, SUBLEDGER_KEY_COLUMNS, "Sub-ledger", reporter, (0.10, 0.14)
    )
    require_columns(sub, REQUIRED_SUBLEDGER_COLUMNS, "Sub-ledger")
    reporter.check_cancelled()

    reporter.progress(0.15, "Filtering VAT transactions ...")
    vat_out, sub_f = build_comparison_frame(vat, sub, reporter)
    reporter.log(f"{len(vat_out)} unique invoice/tax-code/currency record(s) in the output.")
    reporter.check_cancelled()

    # --------------------------------------------------------------
    # Step 5: Build the output workbook in memory
    # --------------------------------------------------------------
    reporter.progress(0.36, "Building the output workbook ...")
    wb = Workbook()
    wb.remove(wb.active)
    ws = write_dataframe_sheet(wb, DATA_SHEET_NAME, vat_out)
    apply_highlighting(ws)

    # --------------------------------------------------------------
    # Step 6: Tax-code specific breakout sheets (same columns/rows as the
    # main comparison sheet, just split out by tax code for easier review) -
    # with the same yellow/red highlighting applied
    # --------------------------------------------------------------
    ws_ec = write_dataframe_sheet(
        wb, "EC acquisition", vat_out[vat_out["Tax Code"] == "ES_REG/ES_ICAG_21_RCR"]
    )
    insert_blank_column(
        ws_ec, "Payment date", after_column_name="Registered Date / Invoice Creation Date:"
    )
    ws_dom21 = write_dataframe_sheet(
        wb, "Domestic sales (21% VAT)", vat_out[vat_out["Tax Code"] == "ES_REG/DSG_21"]
    )
    ws_dom0 = write_dataframe_sheet(
        wb, "Domestic sales (0% VAT)", vat_out[vat_out["Tax Code"] == "ES_REG/ES_DSG_0_RC"]
    )
    for breakout_ws in (ws_ec, ws_dom21, ws_dom0):
        apply_highlighting(breakout_ws)
    reporter.check_cancelled()

    # --------------------------------------------------------------
    # Step 7: Sub-ledger invoices that never showed up in the VAT
    # Transaction file at all, restricted to the two VAT accounts (Sales VAT
    # payable and ICAG VAT receivable) - so this only flags invoices where
    # the missing GL posting is actually VAT-relevant. Checked against every
    # Spain invoice in the VAT file (not just the 4 tax codes handled above),
    # so an invoice posted under a different tax code isn't wrongly flagged.
    # Sub-ledger rows with no Invoice NO. are excluded (not invoice-driven).
    # --------------------------------------------------------------
    reporter.progress(0.40, "Finding sub-ledger invoices missing from the VAT file ...")
    MISSED_INV_ACCOUNTS = [ACCOUNT_SALES, ACCOUNT_ICAG_RCR]
    spain_invoices = set(
        vat.loc[vat["Our Tax Registration No."] == TAX_REG_NO, "Invoice No."].astype(str).str.strip()
    )
    # The literal "nan"/"none" test matters because the Invoice NO. column has
    # already been through .astype(str): on pandas 2 that turns a blank cell
    # into the *string* "nan", which passes notna() and would wrongly flag
    # every blank-invoice row here. pandas 3 keeps it as NaN, so both spellings
    # of "blank" are excluded explicitly and the result is version-independent.
    sub_invoice_no = sub["Invoice NO."].astype(str).str.strip()
    missed = sub[
        sub["Invoice NO."].notna()
        & ~sub_invoice_no.str.lower().isin(["", "nan", "none", "nat"])
        & (~sub["Invoice NO."].isin(spain_invoices))
        & (sub["Account Number"].isin(MISSED_INV_ACCOUNTS))
    ].copy()
    write_dataframe_sheet(wb, "Missed Inv in VAT Transaction", missed)
    reporter.log(f"{len(missed)} sub-ledger row(s) flagged as missing from the VAT Transaction file.")
    reporter.check_cancelled()

    # --------------------------------------------------------------
    # Step 7c: Build the VAT Validation sheet (unique VAT Registration No.
    # values checked against the official EU VIES service), then propagate
    # a "VAT Status" lookup column back onto the 4 comparison sheets.
    # --------------------------------------------------------------
    _, vies_counts = build_vat_validation_sheet(wb, vat_out, reporter, progress_span=(0.42, 0.90))
    for status_ws in (ws, ws_ec, ws_dom21, ws_dom0):
        add_vat_status_column(status_ws)

    # --------------------------------------------------------------
    # Step 7b: Copy in the "Local Purchase" sheet, if one was supplied
    # (values/formulas, formatting, merged cells all preserved as-is)
    # --------------------------------------------------------------
    reporter.progress(0.92, "Building the VAT Summary sheet ...")
    local_purchase_ws = add_local_purchase_sheet(wb, local_purchase_file, reporter)

    add_vat_summary_sheet(wb, ws.title, vat_out)

    # VAT Summary C19/I19 = live references to Local Purchase!H9/I9, so they
    # stay in sync with that sheet without needing to re-run the tool.
    if local_purchase_ws is not None:
        summary_ws = wb[SUMMARY_SHEET_NAME]
        summary_ws["C19"] = f"='{LOCAL_PURCHASE_SHEET_NAME}'!H9"
        summary_ws["I19"] = f"='{LOCAL_PURCHASE_SHEET_NAME}'!I9"

    # --------------------------------------------------------------
    # Step 8: Set the final tab order and save
    # --------------------------------------------------------------
    wb._sheets = [wb[name] for name in SHEET_ORDER if name in wb.sheetnames] + [
        wb[name] for name in wb.sheetnames if name not in SHEET_ORDER
    ]

    reporter.progress(0.97, "Saving the output workbook ...")
    try:
        wb.save(output_file)
    except PermissionError:
        raise ReconciliationError(
            f"Could not write to:\n\n{output_file}\n\n"
            "The file is open in Excel, or the folder is read-only. Close it "
            "and run again, or choose a different output location."
        )

    reporter.progress(1.0, "Done.")
    reporter.log(f"Output saved to: {output_file}")
    return ReconciliationResult(output_file, len(vat_out), vies_counts, reporter.notes)
