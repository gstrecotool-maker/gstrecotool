import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
#  PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="GST 2B Reconciliation",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
#  CSS  — clean, light, minimal
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #f0f4f8; }

[data-testid="stSidebar"] { background: #1e293b !important; border-right: none !important; }
[data-testid="stSidebar"] * { color: #cbd5e1 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] h4 { color: #f1f5f9 !important; font-weight: 700 !important; }
[data-testid="stSidebar"] .stMarkdown p { color: #94a3b8 !important; font-size: 0.82rem !important; }
[data-testid="stSidebar"] hr { border-color: #334155 !important; }

.top-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 60%, #1e40af 100%);
    border-radius: 14px; padding: 28px 36px; margin-bottom: 24px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 20px rgba(29,78,216,0.2);
}
.header-left  { display: flex; align-items: center; gap: 16px; }
.header-icon  { width: 52px; height: 52px; background: rgba(255,255,255,0.15); border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 1.6rem; border: 1px solid rgba(255,255,255,0.2); }
.header-title { font-size: 1.5rem; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; margin: 0; }
.header-sub   { font-size: 0.82rem; color: rgba(255,255,255,0.7); margin: 3px 0 0 0; }
.header-badge { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25); color: #ffffff; font-size: 0.72rem; font-weight: 600; padding: 6px 14px; border-radius: 20px; letter-spacing: 0.5px; }

.upload-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px 22px; box-shadow: 0 1px 6px rgba(0,0,0,0.06); }
.box-label { font-size: 0.75rem; font-weight: 700; color: #374151; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 10px; display: block; }

.section-label { font-size: 0.7rem; font-weight: 700; color: #2563eb; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 4px; }
.section-title { font-size: 1.15rem; font-weight: 700; color: #111827; margin-bottom: 16px; letter-spacing: -0.3px; }

.legend-row { display: flex; gap: 10px; flex-wrap: wrap; margin: 12px 0 18px 0; }
.leg { display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px; border-radius: 20px; font-size: 0.76rem; font-weight: 600; }
.leg-green  { background: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }
.leg-yellow { background: #fef9c3; color: #854d0e; border: 1px solid #fde68a; }
.leg-red    { background: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }
.leg-blue   { background: #dbeafe; color: #1e40af; border: 1px solid #bfdbfe; }

.contact-wrap { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 14px; padding: 28px 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-top: 8px; }
.contact-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 18px; }
.contact-item { display: flex; align-items: flex-start; gap: 12px; padding: 14px 16px; background: #f8fafc; border: 1px solid #e9eef5; border-radius: 10px; }
.contact-icon { font-size: 1.3rem; margin-top: 1px; flex-shrink: 0; }
.contact-lbl  { font-size: 0.68rem; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px; }
.contact-val  { font-size: 0.9rem; font-weight: 600; color: #1d4ed8; margin-top: 2px; }
.contact-ph   { color: #9ca3af !important; font-style: italic; font-weight: 400 !important; font-size: 0.85rem; }

.page-footer { background: #1e293b; border-radius: 12px; padding: 18px 28px; margin-top: 32px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
.footer-left  { font-size: 0.8rem; color: #64748b; }
.footer-right { font-size: 0.76rem; color: #475569; }
.footer-brand { font-weight: 700; color: #94a3b8; }

.page-title { font-size: 1.6rem; font-weight: 700; color: #111827; margin-bottom: 2px; }
.page-sub   { font-size: 0.88rem; color: #6b7280; margin-bottom: 24px; }

[data-testid="metric-container"] {
    background: #ffffff !important; border: 1px solid #e5e7eb !important;
    border-radius: 12px !important; padding: 18px !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.05) !important;
}
[data-testid="stMetricValue"] { color: #111827 !important; font-weight: 800 !important; font-size: 1.6rem !important; }
[data-testid="stMetricLabel"] { color: #6b7280 !important; font-size: 0.78rem !important; font-weight: 600 !important; }
[data-testid="stMetricDelta"]  { font-size: 0.78rem !important; font-weight: 600 !important; }

.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: #ffffff !important; border: none !important;
    border-radius: 10px !important; font-weight: 700 !important;
    font-size: 0.95rem !important; padding: 13px 28px !important;
    width: 100% !important; letter-spacing: 0.2px !important;
    box-shadow: 0 4px 14px rgba(37,99,235,0.3) !important;
    transition: all 0.15s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #1e40af, #1d4ed8) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.4) !important;
    transform: translateY(-1px) !important;
}

[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #059669, #10b981) !important;
    color: #ffffff !important; border: none !important;
    border-radius: 10px !important; font-weight: 700 !important;
    width: 100% !important; box-shadow: 0 4px 14px rgba(16,185,129,0.25) !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(135deg, #047857, #059669) !important;
}

[data-baseweb="tab-list"] {
    background: #e9eef5 !important; border-radius: 10px !important;
    gap: 3px !important; padding: 4px !important;
}
[data-baseweb="tab"] { border-radius: 7px !important; color: #6b7280 !important; font-weight: 500 !important; font-size: 0.88rem !important; }
[aria-selected="true"] { background: #ffffff !important; color: #1d4ed8 !important; font-weight: 700 !important; box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important; }

div[data-testid="stExpander"] { background: #1e293b !important; border: 1px solid #334155 !important; border-radius: 8px !important; }
div[data-testid="stExpander"] * { color: #cbd5e1 !important; }

hr { border-color: #e2e8f0 !important; margin: 24px 0 !important; }
footer, #MainMenu { visibility: hidden !important; }
header[data-testid="stHeader"] { background: transparent !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)



# ============================================================
#  RECONCILIATION ENGINE v4.0
#  Built from actual GSTR-2B portal file analysis
#  ── GSTR-2B B2B sheet structure (confirmed from GST portal):
#     Row 1      : "Goods and Services Tax - GSTR-2B" (title, merged)
#     Rows 2-4   : blank / section title
#     Row 5      : Merged header row 1
#                   A=GSTIN of supplier, B=Trade/Legal name,
#                   C=Invoice Details (merged C:F),
#                   G=Place of supply, H=Rev Charge,
#                   I=Taxable Value (₹),
#                   J=Tax Amount (merged J:M),
#                   N=GSTR-1 Period, O=Filing Date,
#                   P=ITC Availability, Q=Reason,
#                   R=Tax Rate%, S=Source, T=IRN, U=IRN Date
#     Row 6      : Sub-header row 2
#                   C=Invoice number, D=Invoice type,
#                   E=Invoice Date, F=Invoice Value(₹),
#                   J=Integrated Tax(₹), K=Central Tax(₹),
#                   L=State/UT Tax(₹), M=Cess(₹)
#     Row 7+     : Actual data (GSTIN in col A)
# ============================================================

import openpyxl

# ── Exact B2B column positions (1-indexed, confirmed) ───────
B2B_COL_POS = {
    1:  "GSTIN",
    2:  "Vendor",
    3:  "Invoice No",
    4:  "Invoice Type",
    5:  "Invoice Date",
    6:  "Invoice Value",
    7:  "POS",
    8:  "Rev Charge",
    9:  "Taxable",
    10: "IGST",
    11: "CGST",
    12: "SGST",
    13: "Cess",
    14: "GSTR1 Period",
    15: "Filing Date",
    16: "ITC",
    17: "ITC Reason",
    18: "Tax Rate%",
    19: "Source",
    20: "IRN",
    21: "IRN Date",
}

# ── Purchase register keyword map — every possible col name ─
# Priority order: most specific first
BOOKS_COL_MAP = {
    "GSTIN": [
        "gstin of supplier", "gstin of party", "supplier gstin",
        "party gstin", "vendor gstin", "gstin no", "gst no",
        "gst number", "gst in", "gst registration no",
        "gst reg no", "gst reg number", "gstin",
    ],
    "Vendor": [
        "trade/legal name", "trade name", "legal name",
        "supplier name", "party name", "vendor name",
        "ledger name", "account name", "creditor name",
        "particulars", "name of supplier", "purchased from",
        "supplier", "vendor", "party",
    ],
    "Invoice No": [
        "invoice number", "invoice no.", "invoice no",
        "bill number", "bill no.", "bill no",
        "voucher number", "voucher no.", "voucher no",
        "reference number", "ref number", "ref no.", "ref no",
        "document number", "doc number", "doc no.", "doc no",
        "purchase invoice no", "purchase bill no",
        "challan number", "challan no", "receipt number",
        "entry number", "inv no", "inv no.", "inv number",
        "sr no", "serial no",
    ],
    "Invoice Date": [
        "invoice date", "bill date", "voucher date",
        "document date", "doc date", "purchase date",
        "transaction date", "entry date", "date of invoice",
        "challan date", "posting date", "ref date", "inv date",
        "date",
    ],
    "Taxable": [
        "taxable value", "taxable amount", "taxable amt",
        "assessable value", "assessable amount", "assessable amt",
        "basic amount", "basic value", "base amount", "basic amt",
        "value of supply", "taxable supply value",
        "net amount", "net value", "amount before tax",
        "purchase value", "taxable",
    ],
    "IGST": [
        "integrated tax", "integrated gst",
        "igst amount", "igst amt", "igst value",
        "igst paid", "igst tax",
        "i.g.s.t", "i.g.s.t.", "igst",
    ],
    "CGST": [
        "central tax", "central gst",
        "cgst amount", "cgst amt", "cgst value",
        "cgst paid", "cgst tax",
        "c.g.s.t", "c.g.s.t.", "cgst",
    ],
    "SGST": [
        "state/ut tax", "state tax", "ut tax", "utgst",
        "sgst/utgst", "sgst amount", "sgst amt", "sgst value",
        "sgst paid", "sgst tax",
        "s.g.s.t", "s.g.s.t.", "sgst",
    ],
}


# ── GSTR-2B B2B sheet loader ─────────────────────────────────
def _load_b2b_sheet(file_obj) -> tuple[pd.DataFrame, str]:
    """
    Load the B2B sheet from a GSTR-2B portal workbook.
    Uses openpyxl directly to read merged-cell structure correctly.
    Returns (clean DataFrame with standard column names, sheet name).
    """
    wb = openpyxl.load_workbook(file_obj, read_only=False, data_only=True)
    sheets = wb.sheetnames

    # ── Step 1: Find the B2B sheet ───────────────────────────
    # Priority 1: exact name 'B2B'
    b2b_name = None
    for s in sheets:
        if s.strip().upper() == "B2B":
            b2b_name = s
            break
    # Priority 2: starts with B2B
    if not b2b_name:
        for s in sheets:
            if re.match(r"^B2B", s.strip(), re.IGNORECASE) and "CDNR" not in s.upper() and "DNR" not in s.upper():
                b2b_name = s
                break
    # Priority 3: any sheet with 'b2b' in name
    if not b2b_name:
        for s in sheets:
            if "b2b" in s.lower() and "cdnr" not in s.lower():
                b2b_name = s
                break
    if not b2b_name:
        raise ValueError(
            f"Could not find B2B sheet.\nSheets in workbook: {sheets}"
        )

    ws = wb[b2b_name]

    # ── Step 2: Find data start row ──────────────────────────
    # Scan rows to find where GSTIN data starts (15-char alphanum in col A)
    data_start = None
    all_rows = list(ws.iter_rows(min_row=1, values_only=True))

    for i, row in enumerate(all_rows):
        val = row[0]  # Column A
        if val is not None:
            s = str(val).strip().upper().replace(" ", "")
            # Valid GSTIN: 15 chars, starts with 2 digits
            if len(s) == 15 and s[:2].isdigit():
                data_start = i  # 0-indexed
                break

    if data_start is None:
        # No data rows — return empty DataFrame with correct columns
        empty_df = pd.DataFrame(columns=list(B2B_COL_POS.values()))
        return empty_df, b2b_name

    # ── Step 3: Build DataFrame from exact column positions ──
    records = []
    for row in all_rows[data_start:]:
        # Skip empty rows
        if row[0] is None or str(row[0]).strip() in ("", "None"):
            continue
        record = {}
        for col_idx, col_name in B2B_COL_POS.items():
            raw = row[col_idx - 1] if col_idx - 1 < len(row) else None
            record[col_name] = raw
        records.append(record)

    df = pd.DataFrame(records)

    # ── Step 4: Clean and type-cast ──────────────────────────
    # GSTIN
    df["GSTIN"] = df["GSTIN"].astype(str).str.strip().str.upper().str.replace(r"\s+", "", regex=True)
    df = df[df["GSTIN"].str.match(r"^\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z][A-Z0-9]$", na=False)].copy()

    # Vendor
    df["Vendor"] = df["Vendor"].fillna("").astype(str).str.strip()

    # Invoice No — portal gives it as-is (may be numeric or string)
    df["Invoice No"] = df["Invoice No"].fillna("").astype(str).str.strip()

    # Invoice Date
    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors="coerce", dayfirst=True)

    # Numeric tax/value columns
    for col in ["Taxable", "IGST", "CGST", "SGST", "Cess", "Invoice Value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["Total Tax"] = df["IGST"] + df["CGST"] + df["SGST"]
    df = df.reset_index(drop=True)

    return df, b2b_name


# ── Books / Purchase Register column auto-detector ───────────
def _clean(s: str) -> str:
    """Lowercase, remove all non-alphanumeric except space."""
    return re.sub(r"[^a-z0-9 ]", " ", str(s).lower()).strip()


def _find_books_col(columns: list, field: str) -> str | None:
    """
    5-pass smart column finder for purchase register.
    Returns original column name or None.
    """
    keywords = BOOKS_COL_MAP.get(field, [field.lower()])
    lower_map = {c.lower().strip(): c for c in columns}
    clean_map = {_clean(c): c for c in columns}
    # Also build a dots-removed map for I.G.S.T style
    nodot_map = {re.sub(r'[^a-z0-9]', '', c.lower()): c for c in columns}

    # Pass 1: exact lowercase match
    for kw in keywords:
        if kw.lower() in lower_map:
            return lower_map[kw.lower()]

    # Pass 2: cleaned exact match (ignoring special chars)
    for kw in keywords:
        kw_c = _clean(kw)
        if kw_c in clean_map:
            return clean_map[kw_c]

    # Pass 3: keyword is substring of column name
    for kw in keywords:
        kw_c = _clean(kw)
        for col_c, col_orig in clean_map.items():
            if kw_c and kw_c in col_c:
                return col_orig

    # Pass 4: column name is substring of keyword (handles short col names)
    for kw in keywords:
        kw_c = _clean(kw)
        for col_c, col_orig in clean_map.items():
            if col_c and len(col_c) >= 3 and col_c in kw_c:
                return col_orig

    # Pass 5: dots-removed match (handles I.G.S.T, C.G.S.T, S.G.S.T)
    for kw in keywords:
        kw_nodot = re.sub(r'[^a-z0-9]', '', kw.lower())
        if kw_nodot in nodot_map:
            return nodot_map[kw_nodot]
        # Also try: nodot col contained in nodot keyword
        for col_nd, col_orig in nodot_map.items():
            if col_nd and len(col_nd) >= 3 and col_nd in kw_nodot:
                return col_orig

    return None


def _load_books(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Auto-detect columns in any purchase register format.
    Returns (normalized DataFrame, mapping dict for diagnostics).
    """
    # Strip all column names
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    mapping = {}  # standard_name -> original_col
    for field in BOOKS_COL_MAP:
        found = _find_books_col(list(df.columns), field)
        if found:
            mapping[field] = found

    # Rename detected columns
    rename = {v: k for k, v in mapping.items() if v in df.columns}
    df = df.rename(columns=rename)

    # Mandatory string columns
    for col in ["GSTIN", "Invoice No", "Vendor"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Clean GSTIN
    df["GSTIN"] = df["GSTIN"].str.upper().str.replace(r"\s+", "", regex=True)

    # Drop rows that are clearly junk — GSTIN too short (< 10 chars) or blank
    # Keep rows with length 10-15 to handle partial/legacy GSTINs
    valid_mask = df["GSTIN"].str.len().between(10, 15)
    df = df[valid_mask].copy()

    if df.empty:
        # Return with full column list to give better error message
        df["Invoice Date"] = pd.NaT
        for col in ["IGST", "CGST", "SGST", "Taxable"]:
            df[col] = 0.0
        df["Total Tax"] = 0.0
        return df.reset_index(drop=True), mapping

    # Invoice date
    if "Invoice Date" not in df.columns:
        df["Invoice Date"] = pd.NaT
    else:
        df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors="coerce", dayfirst=True)

    # Numeric columns
    for col in ["IGST", "CGST", "SGST", "Taxable"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["Total Tax"] = df["IGST"] + df["CGST"] + df["SGST"]
    df = df.reset_index(drop=True)

    return df, mapping


# ── Invoice number matching ──────────────────────────────────
def _norm_inv(inv) -> str:
    """Alphanumeric only, lowercase."""
    return re.sub(r"[^a-z0-9]", "", str(inv).lower().strip())

def _digits_only(inv) -> str:
    """Digits only."""
    return re.sub(r"\D", "", str(inv))

def _match_invoice(a: str, b: str) -> tuple[bool, str]:
    """
    6-strategy invoice matching. Returns (matched, method).
    Strategies ordered from strictest to most lenient.
    """
    if not a or not b or a in ("nan","None","") or b in ("nan","None",""):
        return False, ""

    na, nb = _norm_inv(a), _norm_inv(b)
    da, db = _digits_only(a), _digits_only(b)

    # S1 — Exact alphanumeric (case-insensitive, special-chars stripped)
    if na == nb and na:
        return True, "Exact"

    # S2 — Numeric part exact equality
    if da and db and da == db:
        return True, "Numeric Exact"

    # S3 — Numeric suffix: one ends with the other (handles prefix differences)
    #      e.g. "GST/2024/001" vs "001"  OR  "ABC001" vs "001"
    if da and db and len(da) >= 2 and len(db) >= 2:
        if da.endswith(db) or db.endswith(da):
            return True, "Numeric Suffix"

    # S4 — Alpha suffix: last min(len) chars match (handles year/prefix additions)
    #      e.g. "FY24INV001" vs "INV001"
    min_len = min(len(na), len(nb))
    if min_len >= 4:
        if na[-min_len:] == nb[-min_len:]:
            return True, "Alpha Suffix"
        if na[:min_len] == nb[:min_len]:
            return True, "Alpha Prefix"

    # S5 — Numeric contained (handles leading zeros, short refs)
    #      e.g. "00045" vs "45"
    if da and db and len(da) >= 3 and len(db) >= 3:
        if da.lstrip("0") == db.lstrip("0"):
            return True, "Leading Zero"
        if da in db or db in da:
            return True, "Numeric Contains"

    # S6 — Tally/Busy style: "Purchase-123" vs "123", strip common prefixes
    prefixes = ["purchase", "pur", "bill", "inv", "gst", "tax", "rcm", "import"]
    na_stripped = na
    nb_stripped = nb
    for pfx in prefixes:
        if na_stripped.startswith(pfx): na_stripped = na_stripped[len(pfx):]
        if nb_stripped.startswith(pfx): nb_stripped = nb_stripped[len(pfx):]
    if na_stripped and nb_stripped and na_stripped == nb_stripped and len(na_stripped) >= 3:
        return True, "Prefix Stripped"

    return False, ""


# ── Main reconciliation ──────────────────────────────────────
def run_reco(gstr2b_file, books_df_raw: pd.DataFrame, date_tol, amt_tol, taxable_tol):
    """
    Main entry point.
    gstr2b_file  : uploaded file object (GSTR-2B portal workbook)
    books_df_raw : DataFrame from any accounting software
    """

    # 1. Load GSTR-2B B2B sheet with exact column positions
    gstr2b, b2b_sheet = _load_b2b_sheet(gstr2b_file)

    # 2. Load and normalize books
    books, b_map = _load_books(books_df_raw)

    # 3. Validate
    if gstr2b.empty:
        raise ValueError(
            f"The B2B sheet ('{b2b_sheet}') has no data rows.\n"
            "Please ensure you uploaded the correct GSTR-2B workbook downloaded from the GST portal."
        )
    if books.empty:
        raise ValueError(
            "Purchase Register has no valid rows after processing.\n"
            f"Columns detected: {list(books_df_raw.columns[:15])}\n\n"
            "Possible reasons:\n"
            "• GSTIN column missing or values are not 10-15 characters\n"
            "• File has wrong sheet selected or is empty\n"
            "• Header row not in first 15 rows\n"
            "Please check your file and try again."
        )
    if "GSTIN" not in books.columns or books["GSTIN"].str.len().max() < 10:
        raise ValueError(
            f"Could not find GSTIN column in Purchase Register.\n"
            f"Columns found: {list(books_df_raw.columns[:20])}\n"
            "Please ensure your file has a GSTIN column."
        )
    if "Invoice No" not in books.columns:
        raise ValueError(
            f"Could not find Invoice Number column in Purchase Register.\n"
            f"Columns found: {list(books_df_raw.columns[:20])}"
        )

    # 4. Build GSTIN-indexed lookup on books (fast O(1) lookup)
    books_by_gstin: dict = {}
    for j, bk in books.iterrows():
        g = bk["GSTIN"]
        books_by_gstin.setdefault(g, []).append((j, bk))

    # Also build PAN-level index for mismatched state codes
    books_by_pan: dict = {}
    for g, entries in books_by_gstin.items():
        if len(g) == 15:
            pan = g[2:12]
            books_by_pan.setdefault(pan, []).extend(entries)

    results       = []
    matched_books = set()

    # 5. Match: GSTR-2B → Books
    for _, r2b in gstr2b.iterrows():
        gstin_2b = r2b["GSTIN"]
        inv_2b   = str(r2b["Invoice No"])

        best_j, best_bk, best_method = None, None, ""

        # Candidate pool: exact GSTIN first, then PAN fallback
        candidates = list(books_by_gstin.get(gstin_2b, []))
        if not candidates and len(gstin_2b) == 15:
            pan = gstin_2b[2:12]
            candidates = list(books_by_pan.get(pan, []))

        for j, bk in candidates:
            if j in matched_books:
                continue
            matched, method = _match_invoice(inv_2b, str(bk["Invoice No"]))
            if matched:
                best_j, best_bk, best_method = j, bk, method
                break

        if best_bk is not None:
            matched_books.add(best_j)
            bk = best_bk

            date_diff    = 0
            if pd.notna(r2b["Invoice Date"]) and pd.notna(bk["Invoice Date"]):
                date_diff = abs((r2b["Invoice Date"] - bk["Invoice Date"]).days)

            igst_diff    = abs(r2b["IGST"]    - bk["IGST"])
            cgst_diff    = abs(r2b["CGST"]    - bk["CGST"])
            sgst_diff    = abs(r2b["SGST"]    - bk["SGST"])
            taxable_diff = abs(r2b["Taxable"] - bk["Taxable"])

            diffs = []
            if date_diff    > date_tol:    diffs.append(f"Date Diff ({date_diff}d)")
            if igst_diff    > amt_tol:     diffs.append(f"IGST Diff ₹{igst_diff:,.2f}")
            if cgst_diff    > amt_tol:     diffs.append(f"CGST Diff ₹{cgst_diff:,.2f}")
            if sgst_diff    > amt_tol:     diffs.append(f"SGST Diff ₹{sgst_diff:,.2f}")
            if taxable_diff > taxable_tol: diffs.append(f"Taxable Diff ₹{taxable_diff:,.2f}")

            results.append({
                "GSTIN":                gstin_2b,
                "Vendor (2B)":          r2b["Vendor"],
                "Vendor (Books)":       bk.get("Vendor", ""),
                "Invoice No (2B)":      r2b["Invoice No"],
                "Invoice No (Books)":   bk["Invoice No"],
                "Invoice Date (2B)":    r2b["Invoice Date"],
                "Invoice Date (Books)": bk["Invoice Date"],
                "Taxable (2B)":         r2b["Taxable"],
                "Taxable (Books)":      bk["Taxable"],
                "IGST (2B)":            r2b["IGST"],
                "IGST (Books)":         bk["IGST"],
                "CGST (2B)":            r2b["CGST"],
                "CGST (Books)":         bk["CGST"],
                "SGST (2B)":            r2b["SGST"],
                "SGST (Books)":         bk["SGST"],
                "Total Tax (2B)":       r2b["Total Tax"],
                "Total Tax (Books)":    bk["Total Tax"],
                "ITC":                  r2b.get("ITC", ""),
                "Match Method":         best_method,
                "Status":               "MATCHED WITH DIFF" if diffs else "MATCHED",
                "Remarks":              " | ".join(diffs) if diffs else "Exact Match",
            })
        else:
            results.append({
                "GSTIN":                gstin_2b,
                "Vendor (2B)":          r2b["Vendor"],
                "Vendor (Books)":       "",
                "Invoice No (2B)":      r2b["Invoice No"],
                "Invoice No (Books)":   "",
                "Invoice Date (2B)":    r2b["Invoice Date"],
                "Invoice Date (Books)": pd.NaT,
                "Taxable (2B)":         r2b["Taxable"],
                "Taxable (Books)":      None,
                "IGST (2B)":            r2b["IGST"],
                "IGST (Books)":         None,
                "CGST (2B)":            r2b["CGST"],
                "CGST (Books)":         None,
                "SGST (2B)":            r2b["SGST"],
                "SGST (Books)":         None,
                "Total Tax (2B)":       r2b["Total Tax"],
                "Total Tax (Books)":    None,
                "ITC":                  r2b.get("ITC", ""),
                "Match Method":         "",
                "Status":               "IN 2B – NOT IN BOOKS",
                "Remarks":              "Not found in Purchase Register",
            })

    # 6. Books-only entries
    for j, bk in books.iterrows():
        if j not in matched_books:
            results.append({
                "GSTIN":                bk["GSTIN"],
                "Vendor (2B)":          "",
                "Vendor (Books)":       bk.get("Vendor", ""),
                "Invoice No (2B)":      "",
                "Invoice No (Books)":   bk["Invoice No"],
                "Invoice Date (2B)":    pd.NaT,
                "Invoice Date (Books)": bk["Invoice Date"],
                "Taxable (2B)":         None,
                "Taxable (Books)":      bk["Taxable"],
                "IGST (2B)":            None,
                "IGST (Books)":         bk["IGST"],
                "CGST (2B)":            None,
                "CGST (Books)":         bk["CGST"],
                "SGST (2B)":            None,
                "SGST (Books)":         bk["SGST"],
                "Total Tax (2B)":       None,
                "Total Tax (Books)":    bk["Total Tax"],
                "ITC":                  "",
                "Match Method":         "",
                "Status":               "IN BOOKS – NOT IN 2B",
                "Remarks":              "Not uploaded on GST Portal",
            })

    out = pd.DataFrame(results) if results else pd.DataFrame()

    meta = {
        "b2b_sheet":    b2b_sheet,
        "b2b_count":    len(gstr2b),
        "books_count":  len(books),
        "b_map":        b_map,
        "books_cols":   list(books_df_raw.columns[:25]),
    }
    return out, meta


# ============================================================
#  EXCEL OUTPUT
# ============================================================
def build_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    STATUS_STYLE = {
        "MATCHED":               ("E2EFDA","1E4620"),
        "MATCHED WITH DIFF":     ("FFF2CC","7B4F00"),
        "IN 2B – NOT IN BOOKS":  ("FCE4D6","843C0C"),
        "IN BOOKS – NOT IN 2B":  ("DAEEF3","0D3D5E"),
    }

    ws = wb.create_sheet("Reconciliation")
    ws.sheet_properties.tabColor = "1F3864"
    ws.merge_cells(f"A1:{get_column_letter(len(df.columns))}1")
    ws["A1"].value     = f"GST 2B vs Books — Reconciliation  |  {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A1"].font      = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    ws["A1"].fill      = PatternFill("solid", fgColor="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for c_i, col in enumerate(df.columns, 1):
        c = ws.cell(row=2, column=c_i, value=col)
        c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
    ws.row_dimensions[2].height = 22

    MONEY = {c for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST","Total Tax"])}

    for r_i, (_, row) in enumerate(df.iterrows()):
        r = 3 + r_i
        bg, fg = STATUS_STYLE.get(row.get("Status",""), ("FFFFFF","111111"))
        for c_i, val in enumerate(row, 1):
            col_name = df.columns[c_i-1]
            cell = ws.cell(row=r, column=c_i, value=val)
            cell.fill   = PatternFill("solid", fgColor=bg)
            cell.border = bdr
            cell.font   = Font(name="Calibri", size=10, bold=(col_name=="Status"), color=fg)
            if col_name in MONEY:
                cell.number_format = "##,##,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    CW = {"GSTIN":22,"Vendor":26,"Invoice No (2B)":18,"Invoice No (Books)":18,
          "Invoice Date (2B)":16,"Invoice Date (Books)":16,
          "Taxable (2B)":14,"Taxable (Books)":14,
          "IGST (2B)":12,"IGST (Books)":12,"CGST (2B)":12,"CGST (Books)":12,
          "SGST (2B)":12,"SGST (Books)":12,"Status":22,"Remarks":34}
    for c_i, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(c_i)].width = CW.get(col, 14)
    ws.freeze_panes = "A3"

    ws2 = wb.create_sheet("Summary")
    ws2.sheet_properties.tabColor = "375623"
    ws2.merge_cells("A1:B1")
    ws2["A1"].value     = "Reconciliation Summary"
    ws2["A1"].font      = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    ws2["A1"].fill      = PatternFill("solid", fgColor="1F3864")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    rows2 = [
        ("Total GSTR-2B Records",     (df["Invoice No (2B)"]    != "").sum(), "DAEEF3"),
        ("Total Books Records",        (df["Invoice No (Books)"] != "").sum(), "DAEEF3"),
        ("Matched — Exact",            (df["Status"]=="MATCHED").sum(),               "E2EFDA"),
        ("Matched — With Differences", (df["Status"]=="MATCHED WITH DIFF").sum(),     "FFF2CC"),
        ("In 2B — Not in Books",       (df["Status"]=="IN 2B – NOT IN BOOKS").sum(),  "FCE4D6"),
        ("In Books — Not in 2B",       (df["Status"]=="IN BOOKS – NOT IN 2B").sum(),  "DAEEF3"),
    ]
    for r_i, (label, count, color) in enumerate(rows2):
        r = r_i + 2
        c1 = ws2.cell(row=r, column=1, value=label)
        c2 = ws2.cell(row=r, column=2, value=count)
        for c in (c1, c2):
            c.fill = PatternFill("solid", fgColor=color); c.border = bdr
            c.font = Font(name="Calibri", size=11)
        c1.alignment = Alignment(horizontal="left",   vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.font = Font(name="Calibri", size=11, bold=True)
        ws2.row_dimensions[r].height = 20

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 14

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ============================================================
#  BOOKS FILE READER — handles any Excel format robustly
# ============================================================
def _read_books_file(file_obj) -> pd.DataFrame:
    """
    Reads a purchase register Excel file intelligently:
    - Tries every sheet, picks the one with the most GSTIN-like data
    - Scans first 15 rows to find the actual header row (skips blank/title rows)
    - Returns a clean DataFrame ready for _load_books()
    """
    file_obj.seek(0)
    raw_bytes = file_obj.read()

    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=False, data_only=True)
    sheets = wb.sheetnames

    best_df   = None
    best_score = -1
    best_sheet = ""

    for sheet_name in sheets:
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        # ── Find header row: first row with 3+ non-None values ──
        header_idx = None
        for i, row in enumerate(all_rows[:15]):
            non_none = [v for v in row if v is not None]
            if len(non_none) >= 3:
                header_idx = i
                break

        if header_idx is None:
            continue

        # ── Build DataFrame from header_idx ──────────────────
        header = [str(v).strip() if v is not None else f"Col_{j}"
                  for j, v in enumerate(all_rows[header_idx])]

        # Remove fully duplicate column names
        seen = {}
        clean_header = []
        for h in header:
            if h in seen:
                seen[h] += 1
                clean_header.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                clean_header.append(h)

        data_rows = []
        for row in all_rows[header_idx + 1:]:
            if any(v is not None for v in row):
                # Pad or trim row to header length
                row_list = list(row)
                while len(row_list) < len(clean_header):
                    row_list.append(None)
                data_rows.append(row_list[:len(clean_header)])

        if not data_rows:
            continue

        df = pd.DataFrame(data_rows, columns=clean_header)
        df = df.dropna(how="all").reset_index(drop=True)

        # ── Score: count rows that look like GST data ─────────
        # Heuristic: column contains 15-char values starting with digits
        score = 0
        for col in df.columns:
            vals = df[col].dropna().astype(str).str.strip()
            gstin_like = vals.str.match(r"^\d{2}[A-Z0-9]{13}$").sum()
            score += gstin_like * 10
            # Also score: column names look useful
            col_l = col.lower()
            if any(k in col_l for k in ["gstin","gst","invoice","bill","taxable","igst","cgst","sgst"]):
                score += 5

        if score > best_score:
            best_score  = score
            best_df     = df
            best_sheet  = sheet_name

    if best_df is None or best_df.empty:
        # Last resort: just read first sheet with header=0
        xl = pd.ExcelFile(io.BytesIO(raw_bytes))
        for sn in xl.sheet_names:
            try:
                df = xl.parse(sn)
                df = df.dropna(how="all").reset_index(drop=True)
                if not df.empty:
                    return df
            except Exception:
                continue
        raise ValueError(
            "Purchase Register file appears to be empty or unreadable.\n"
            "Please check the file and try again."
        )

    return best_df


# ============================================================
#  SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🧾 GST Reco Tool")
    st.markdown("---")

    st.markdown("#### ⚙️ Tolerance Settings")
    DATE_TOL    = st.slider("Date Tolerance (days)",        0,   15,   5)
    AMT_TOL     = st.slider("GST Amount Tolerance (₹)",    0.0, 50.0, 2.0,  step=0.5)
    TAXABLE_TOL = st.slider("Taxable Value Tolerance (₹)", 0.0,100.0, 10.0, step=1.0)

    st.markdown("---")
    st.markdown("#### 📋 Required Columns")
    with st.expander("GSTR-2B"):
        st.markdown("- `GSTIN`\n- `Trade/Legal name`\n- `Invoice No`\n- `Invoice Date`\n- `Taxable Value (₹)`\n- `IGST`, `CGST`, `SGST`")
    with st.expander("Books / Purchase Register"):
        st.markdown("- `GSTIN`\n- `Particulars` *(or Vendor)*\n- `Invoice No`\n- `Invoice Date`\n- `Taxable Value`\n- `IGST`, `CGST`, `SGST`")


# ============================================================
#  MAIN PAGE
# ============================================================
# ── Header banner ────────────────────────────────────────────
st.markdown("""
<div class="top-header">
  <div class="header-left">
    <div class="header-icon">🧾</div>
    <div>
      <div class="header-title">GST 2B vs Books Reconciliation</div>
      <div class="header-sub">Upload your GSTR-2B and Purchase Register — get a colour-coded Excel report instantly</div>
    </div>
  </div>
  <div class="header-badge">⚡ AI-Powered Matching</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-label">Step 1 — Upload Files</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Select Your Excel Files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="box-label">📥 GSTR-2B File</div>', unsafe_allow_html=True)
    gstr_file = st.file_uploader("GSTR-2B", type=["xlsx","xls"],
                                  label_visibility="collapsed", key="gstr")
    if gstr_file: st.success(f"✅ {gstr_file.name}")

with col2:
    st.markdown('<div class="box-label">📗 Books / Purchase Register</div>', unsafe_allow_html=True)
    books_file = st.file_uploader("Books", type=["xlsx","xls"],
                                   label_visibility="collapsed", key="books")
    if books_file: st.success(f"✅ {books_file.name}")

st.markdown("<br>", unsafe_allow_html=True)

btn_col, _ = st.columns([1, 3])
with btn_col:
    run = st.button("▶  Run Reconciliation", use_container_width=True)

# ── Process ──────────────────────────────────────────────────
if run:
    if not gstr_file or not books_file:
        st.error("Please upload both files before running.")
    else:
        with st.spinner("Reconciling…"):
            try:
                # ── Read books file robustly ──────────────────
                books_file.seek(0)
                b_df = _read_books_file(books_file)

                # ── Reset GSTR-2B file pointer ────────────────
                gstr_file.seek(0)

                result_df, meta = run_reco(
                    gstr_file, b_df,
                    DATE_TOL, AMT_TOL, TAXABLE_TOL,
                )
                st.session_state["result"] = result_df
                st.session_state["meta"]   = meta
                st.session_state["ran"]    = True

            except Exception as e:
                err = str(e)
                st.error(f"❌ {err}")

                # ── Diagnostic expander ───────────────────────
                try:
                    gstr_file.seek(0)
                    books_file.seek(0)
                    import openpyxl as _opx
                    wb_g = _opx.load_workbook(gstr_file, read_only=True)
                    wb_b = _opx.load_workbook(books_file, read_only=True)
                    with st.expander("🔍 Diagnostic — click to see detected sheets & columns"):
                        st.write("**GSTR-2B sheets:**", wb_g.sheetnames)
                        st.write("**Books sheets:**", wb_b.sheetnames)
                        for sn in wb_b.sheetnames:
                            ws = wb_b[sn]
                            rows = list(ws.iter_rows(max_row=10, values_only=True))
                            for r in rows:
                                vals = [v for v in r if v is not None]
                                if len(vals) >= 3:
                                    st.write(f"Books sheet *{sn}* first header row found:", list(r))
                                    break
                except Exception:
                    pass
                st.session_state["ran"] = False

# ── Results ──────────────────────────────────────────────────
if st.session_state.get("ran") and "result" in st.session_state:
    df = st.session_state["result"]

    matched      = (df["Status"] == "MATCHED").sum()
    matched_diff = (df["Status"] == "MATCHED WITH DIFF").sum()
    only_2b      = (df["Status"] == "IN 2B – NOT IN BOOKS").sum()
    only_books   = (df["Status"] == "IN BOOKS – NOT IN 2B").sum()
    total        = len(df)
    rate         = f"{matched/total*100:.1f}%" if total else "0%"

    st.markdown("---")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total",               total)
    m2.metric("✅ Matched",           matched,      rate)
    m3.metric("⚠️ With Differences",  matched_diff)
    m4.metric("❌ 2B Only",           only_2b)
    m5.metric("📘 Books Only",        only_books)

    # Legend
    st.markdown("""
    <div class="legend-row">
      <span class="leg leg-green">✅ Matched — Exact</span>
      <span class="leg leg-yellow">⚠️ Matched with Differences</span>
      <span class="leg leg-red">❌ In 2B — Not in Books</span>
      <span class="leg leg-blue">📘 In Books — Not in 2B</span>
    </div>
    """, unsafe_allow_html=True)

    # Column mapping diagnostics
    if "meta" in st.session_state:
        meta = st.session_state["meta"]
        with st.expander(f"🔍 Column Detection Report  |  B2B Sheet: **{meta['b2b_sheet']}**  |  2B Records: {meta['b2b_count']}  |  Books Records: {meta['books_count']}"):
            st.markdown("**Purchase Register — Columns Auto-Detected:**")
            if meta["b_map"]:
                cols_info = st.columns(3)
                items = list(meta["b_map"].items())
                chunk = max(1, len(items)//3 + 1)
                for i, col in enumerate(cols_info):
                    with col:
                        for field, orig in items[i*chunk:(i+1)*chunk]:
                            st.markdown(f"- `{orig}` → **{field}**")
            else:
                st.warning("No columns auto-detected. Check your Purchase Register format.")
            st.markdown("**GSTR-2B B2B** columns are read by fixed position (portal format) — no mapping needed.")

    tabs = st.tabs(["All", "✅ Matched", "⚠️ Differences", "❌ 2B Only", "📘 Books Only"])

    MONEY_COLS = {c: st.column_config.NumberColumn(format="₹%.2f")
                  for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST","Total Tax"])}
    DATE_COLS  = {c: st.column_config.DateColumn(c, format="DD-MMM-YYYY")
                  for c in ["Invoice Date (2B)","Invoice Date (Books)"] if c in df.columns}

    def show(data):
        if data.empty: st.info("No records in this category."); return
        st.dataframe(data, use_container_width=True, hide_index=True,
                     column_config={**MONEY_COLS, **DATE_COLS})
        st.caption(f"{len(data):,} records")

    for tab, flt in zip(tabs, [None,"MATCHED","MATCHED WITH DIFF","IN 2B – NOT IN BOOKS","IN BOOKS – NOT IN 2B"]):
        with tab:
            show(df if flt is None else df[df["Status"] == flt])

    st.markdown("---")
    dl_col, _ = st.columns([1, 3])
    with dl_col:
        st.download_button(
            "📥  Download Excel Report",
            data=build_excel(df),
            file_name=f"GST_Reco_{datetime.now().strftime('%d%b%Y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.caption(f"Report: Reconciliation Detail + Summary sheet  |  {datetime.now().strftime('%d-%b-%Y %H:%M')}")

# ── Contact Us ────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown('<div class="section-label">Get in Touch</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Contact Us</div>', unsafe_allow_html=True)

st.markdown("""
<!-- Email -->
<div class="contact-val">gstrecotool@gmail.com</div>

<!-- Mobile -->
<div class="contact-val">+91 8329194362</div>
""", unsafe_allow_html=True)
        
# ── Footer ────────────────────────────────────────────────────
st.markdown(f"""
<div class="page-footer">
  <div class="footer-left">
    <span class="footer-brand">GST 2B Reconciliation Tool</span><br>
    Simplifying GST compliance for businesses &amp; tax professionals
  </div>
  <div class="footer-right">
    © {datetime.now().year} &nbsp;·&nbsp; All rights reserved
  </div>
</div>
""", unsafe_allow_html=True)
