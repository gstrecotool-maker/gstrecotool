import streamlit as st
import pandas as pd
import re
import io
import json
import hashlib
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import openpyxl

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
#  CSS
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

/* ── Top header ── */
.top-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 60%, #1e40af 100%);
    border-radius: 14px; padding: 28px 36px; margin-bottom: 24px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 20px rgba(29,78,216,0.2);
}
.header-left  { display: flex; align-items: center; gap: 16px; }
.header-icon  { width: 52px; height: 52px; background: rgba(255,255,255,0.15);
                border-radius: 12px; display: flex; align-items: center;
                justify-content: center; font-size: 1.6rem; border: 1px solid rgba(255,255,255,0.2); }
.header-title { font-size: 1.5rem; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; margin: 0; }
.header-sub   { font-size: 0.82rem; color: rgba(255,255,255,0.7); margin: 3px 0 0 0; }
.header-badge { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25);
                color: #ffffff; font-size: 0.72rem; font-weight: 600; padding: 6px 14px;
                border-radius: 20px; letter-spacing: 0.5px; }

/* ── Labels ── */
.box-label    { font-size: 0.75rem; font-weight: 700; color: #374151;
                text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 10px; display: block; }
.section-label{ font-size: 0.7rem; font-weight: 700; color: #2563eb;
                text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 4px; }
.section-title{ font-size: 1.15rem; font-weight: 700; color: #111827;
                margin-bottom: 16px; letter-spacing: -0.3px; }

/* ── Legend ── */
.legend-row { display: flex; gap: 10px; flex-wrap: wrap; margin: 12px 0 18px 0; }
.leg        { display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px;
              border-radius: 20px; font-size: 0.76rem; font-weight: 600; }
.leg-green  { background: #dcfce7; color: #166534; border: 1px solid #bbf7d0; }
.leg-yellow { background: #fef9c3; color: #854d0e; border: 1px solid #fde68a; }
.leg-red    { background: #fee2e2; color: #991b1b; border: 1px solid #fecaca; }
.leg-blue   { background: #dbeafe; color: #1e40af; border: 1px solid #bfdbfe; }

/* ── Usage / payment ── */
.usage-bar-wrap { background: #e2e8f0; border-radius: 6px; height: 7px;
                  margin: 6px 0 3px 0; overflow: hidden; }
.usage-bar-fill { height: 7px; border-radius: 6px; transition: width 0.4s; }
.pay-card  { background: #fff7ed; border: 1.5px solid #fed7aa; border-radius: 14px;
             padding: 24px 28px; margin: 20px 0; }
.pay-title { font-size: 1.05rem; font-weight: 700; color: #c2410c; margin-bottom: 6px; }
.pay-desc  { font-size: 0.85rem; color: #78350f; margin-bottom: 14px; line-height: 1.6; }
.pay-price { font-size: 1.9rem; font-weight: 800; color: #1d4ed8; }
.pay-sub   { font-size: 0.78rem; color: #6b7280; margin-left: 5px; }

/* ── Login card ── */
.login-card { background: #eff6ff; border: 1.5px solid #bfdbfe; border-radius: 14px;
              padding: 24px 28px; margin: 16px 0; }
.login-title{ font-size: 1rem; font-weight: 700; color: #1e40af; margin-bottom: 6px; }
.login-desc { font-size: 0.85rem; color: #374151; margin-bottom: 16px; line-height: 1.6; }

/* ── User bar ── */
.user-bar { background: #f0fdf4; border: 1.5px solid #bbf7d0; border-radius: 10px;
            padding: 10px 16px; margin-bottom: 14px; font-size: 0.85rem;
            color: #166534; font-weight: 600; display: flex; gap: 10px;
            align-items: center; flex-wrap: wrap; }

/* ── Metrics ── */
[data-testid="metric-container"] {
    background: #ffffff !important; border: 1px solid #e5e7eb !important;
    border-radius: 12px !important; padding: 18px !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.05) !important;
}
[data-testid="stMetricValue"] { color: #111827 !important; font-weight: 800 !important; font-size: 1.6rem !important; }
[data-testid="stMetricLabel"] { color: #6b7280 !important; font-size: 0.78rem !important; font-weight: 600 !important; }
[data-testid="stMetricDelta"] { font-size: 0.78rem !important; font-weight: 600 !important; }

/* ── Buttons ── */
.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: #ffffff !important; border: none !important; border-radius: 10px !important;
    font-weight: 700 !important; font-size: 0.95rem !important;
    padding: 13px 28px !important; width: 100% !important;
    box-shadow: 0 4px 14px rgba(37,99,235,0.3) !important; transition: all 0.15s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #1e40af, #1d4ed8) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.4) !important; transform: translateY(-1px) !important;
}
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #059669, #10b981) !important;
    color: #ffffff !important; border: none !important; border-radius: 10px !important;
    font-weight: 700 !important; width: 100% !important;
    box-shadow: 0 4px 14px rgba(16,185,129,0.25) !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(135deg, #047857, #059669) !important;
}

/* ── Tabs ── */
[data-baseweb="tab-list"] {
    background: #e9eef5 !important; border-radius: 10px !important;
    gap: 3px !important; padding: 4px !important;
}
[data-baseweb="tab"] { border-radius: 7px !important; color: #6b7280 !important;
                        font-weight: 500 !important; font-size: 0.88rem !important; }
[aria-selected="true"] { background: #ffffff !important; color: #1d4ed8 !important;
                          font-weight: 700 !important; box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important; }

/* ── Expander ── */
div[data-testid="stExpander"] { background: #1e293b !important; border: 1px solid #334155 !important;
                                  border-radius: 8px !important; }
div[data-testid="stExpander"] * { color: #cbd5e1 !important; }

/* ── Contact ── */
.contact-wrap { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 14px;
                padding: 28px 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-top: 8px; }
.contact-item { display: flex; align-items: flex-start; gap: 12px; padding: 14px 16px;
                background: #f8fafc; border: 1px solid #e9eef5; border-radius: 10px; }
.contact-icon { font-size: 1.3rem; margin-top: 1px; flex-shrink: 0; }
.contact-lbl  { font-size: 0.68rem; font-weight: 700; color: #9ca3af;
                text-transform: uppercase; letter-spacing: 0.5px; }
.contact-val  { font-size: 0.9rem; font-weight: 600; color: #1d4ed8; margin-top: 2px; }
.contact-ph   { color: #9ca3af !important; font-style: italic;
                font-weight: 400 !important; font-size: 0.85rem; }

/* ── Footer ── */
.page-footer { background: #1e293b; border-radius: 12px; padding: 18px 28px; margin-top: 32px;
               display: flex; align-items: center; justify-content: space-between;
               flex-wrap: wrap; gap: 10px; }
.footer-left  { font-size: 0.8rem;  color: #64748b; }
.footer-right { font-size: 0.76rem; color: #475569; }
.footer-brand { font-weight: 700; color: #94a3b8; }

hr { border-color: #e2e8f0 !important; margin: 24px 0 !important; }
footer, #MainMenu { visibility: hidden !important; }
header[data-testid="stHeader"] { background: transparent !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
#  USER / USAGE MANAGEMENT
# ============================================================
FREE_LIMIT = 5          # free reconciliations per year
PRICE_INR  = 49         # ₹ per month
DATA_FILE  = "gst_users.json"

def _load_users() -> dict:
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f: return json.load(f)
        except Exception: pass
    return {}

def _save_users(db: dict):
    try:
        with open(DATA_FILE, "w") as f: json.dump(db, f, indent=2)
    except Exception: pass

def _year_key()  -> str: return datetime.now().strftime("%Y")
def _month_key() -> str: return datetime.now().strftime("%Y-%m")
def _uid(email: str) -> str: return hashlib.md5(email.strip().lower().encode()).hexdigest()

def _get_user(email: str) -> dict:
    db = _load_users(); uid = _uid(email)
    if uid not in db:
        db[uid] = {"email": email.strip().lower(), "name": "", "firm": "", "gstin": "",
                   "registered": datetime.now().strftime("%d-%b-%Y"),
                   "paid_until": "", "usage": {}}
        _save_users(db)
    return db[uid]

def _get_yearly_usage(email: str) -> int:
    return _get_user(email)["usage"].get(_year_key(), 0)

def _increment_usage(email: str):
    db = _load_users(); uid = _uid(email); yk = _year_key()
    db[uid]["usage"][yk] = db[uid]["usage"].get(yk, 0) + 1
    _save_users(db)

def _is_paid(email: str) -> bool:
    return _get_user(email).get("paid_until", "") >= _month_key()

def _mark_paid(email: str):
    db = _load_users(); uid = _uid(email)
    db[uid]["paid_until"] = _month_key(); _save_users(db)

def _update_profile(email: str, name: str, firm: str, gstin: str):
    db = _load_users(); uid = _uid(email)
    if uid in db: db[uid].update({"name": name, "firm": firm, "gstin": gstin}); _save_users(db)

def _can_run(email: str) -> tuple[bool, str]:
    if _is_paid(email): return True, "paid"
    used = _get_yearly_usage(email)
    if used < FREE_LIMIT: return True, f"free"
    return False, "limit_exceeded"


# ============================================================
#  GSTR-2B COLUMN KEYWORDS
# ============================================================
GSTR2B_COL_KEYWORDS = {
    "GSTIN":         ["gstin of supplier", "gstin/uin of supplier", "gstin", "gst no", "supplier gstin"],
    "Vendor":        ["trade/legal name", "trade name", "legal name", "supplier name", "name of supplier", "party name"],
    "Invoice No":    ["invoice number", "invoice no.", "invoice no", "inv no", "bill number", "bill no", "document number"],
    "Invoice Type":  ["invoice type", "document type", "type of invoice"],
    "Invoice Date":  ["invoice date", "bill date", "document date", "date of invoice", "inv date"],
    "Invoice Value": ["invoice value", "bill value", "total invoice value", "gross value", "total value"],
    "POS":           ["place of supply", "pos", "state of supply"],
    "Rev Charge":    ["reverse charge", "rev charge", "rcm"],
    "Taxable":       ["taxable value", "taxable amount", "value of supply", "assessable value", "taxable"],
    "IGST":          ["integrated tax", "integrated tax amount", "igst amount", "igst", "i.g.s.t"],
    "CGST":          ["central tax", "central tax amount", "cgst amount", "cgst", "c.g.s.t"],
    "SGST":          ["state/ut tax", "state tax", "sgst amount", "sgst", "utgst", "s.g.s.t"],
    "Cess":          ["cess amount", "cess", "compensation cess"],
    "GSTR1 Period":  ["gstr-1/isd period", "gstr1 period", "filing period", "return period"],
    "Filing Date":   ["filing date", "gstr-1 filing date", "date of filing"],
    "ITC":           ["itc availability", "itc available", "eligible for itc", "itc"],
    "ITC Reason":    ["reason", "itc reason", "reason for itc"],
    "Tax Rate%":     ["tax rate", "rate %", "gst rate", "rate of tax"],
    "Source":        ["source", "document source"],
    "IRN":           ["irn", "invoice reference number"],
    "IRN Date":      ["irn date", "irn generation date"],
}

B2B_COL_POS_FALLBACK = {
    1:"GSTIN", 2:"Vendor", 3:"Invoice No", 4:"Invoice Type",
    5:"Invoice Date", 6:"Invoice Value", 7:"POS", 8:"Rev Charge",
    9:"Taxable", 10:"IGST", 11:"CGST", 12:"SGST", 13:"Cess",
    14:"GSTR1 Period", 15:"Filing Date", 16:"ITC", 17:"ITC Reason",
    18:"Tax Rate%", 19:"Source", 20:"IRN", 21:"IRN Date",
}

BOOKS_COL_MAP = {
    "GSTIN":       ["gstin of supplier","gstin of party","supplier gstin","party gstin",
                    "vendor gstin","gstin no","gst no","gst number","gst in",
                    "gst registration no","gst reg no","gstin"],
    "Vendor":      ["trade/legal name","trade name","legal name","supplier name","party name",
                    "vendor name","ledger name","account name","creditor name","particulars",
                    "name of supplier","purchased from","supplier","vendor","party"],
    "Invoice No":  ["invoice number","invoice no.","invoice no","bill number","bill no.","bill no",
                    "voucher number","voucher no.","voucher no","reference number","ref number",
                    "ref no.","ref no","document number","doc number","doc no.","doc no",
                    "purchase invoice no","purchase bill no","challan number","challan no",
                    "entry number","inv no","inv no.","inv number","sr no","serial no"],
    "Invoice Date":["invoice date","bill date","voucher date","document date","doc date",
                    "purchase date","transaction date","entry date","date of invoice",
                    "challan date","posting date","ref date","inv date","date"],
    "Taxable":     ["taxable value","taxable amount","taxable amt","assessable value",
                    "assessable amount","basic amount","basic value","base amount",
                    "value of supply","net amount","net value","amount before tax",
                    "purchase value","taxable"],
    "IGST":        ["integrated tax","integrated gst","igst amount","igst amt",
                    "igst value","igst paid","igst tax","i.g.s.t","i.g.s.t.","igst"],
    "CGST":        ["central tax","central gst","cgst amount","cgst amt",
                    "cgst value","cgst paid","cgst tax","c.g.s.t","c.g.s.t.","cgst"],
    "SGST":        ["state/ut tax","state tax","ut tax","utgst","sgst/utgst",
                    "sgst amount","sgst amt","sgst value","sgst paid","sgst tax",
                    "s.g.s.t","s.g.s.t.","sgst"],
}


# ============================================================
#  COLUMN FINDER
# ============================================================
def _clean(s): return re.sub(r"[^a-z0-9 ]"," ",str(s).lower()).strip()
def _nodot(s):  return re.sub(r"[^a-z0-9]","",str(s).lower())

def _find_col(columns, field, kw_map):
    keywords  = kw_map.get(field, [field.lower()])
    lower_map = {c.lower().strip(): c for c in columns}
    clean_map = {_clean(c): c for c in columns}
    nodot_map = {_nodot(c): c for c in columns}
    for kw in keywords:
        if kw.lower() in lower_map: return lower_map[kw.lower()]
    for kw in keywords:
        if _clean(kw) in clean_map: return clean_map[_clean(kw)]
    for kw in keywords:
        kw_c = _clean(kw)
        for col_c, col_orig in clean_map.items():
            if kw_c and kw_c in col_c: return col_orig
    for kw in keywords:
        kw_c = _clean(kw)
        for col_c, col_orig in clean_map.items():
            if col_c and len(col_c) >= 3 and col_c in kw_c: return col_orig
    for kw in keywords:
        kw_nd = _nodot(kw)
        if kw_nd in nodot_map: return nodot_map[kw_nd]
        for col_nd, col_orig in nodot_map.items():
            if col_nd and len(col_nd) >= 3 and col_nd in kw_nd: return col_orig
    return None


# ============================================================
#  HEADER ROW FINDER
# ============================================================
def _find_header_row(all_rows, max_scan=20, col_keywords=None):
    if col_keywords is None:
        col_keywords = ["gstin","gst","invoice","bill","taxable","igst","cgst","sgst",
                        "date","voucher","supplier","vendor","party","value","amount"]
    best_idx = None; best_score = 0
    for i, row in enumerate(all_rows[:max_scan]):
        row_str  = " ".join(str(v).lower() for v in row if v is not None)
        score    = sum(1 for kw in col_keywords if kw in row_str)
        non_none = sum(1 for v in row if v is not None)
        if non_none >= 3:
            if score > best_score: best_score = score; best_idx = i
            elif best_idx is None: best_idx = i
    return best_idx if best_idx is not None else 0


# ============================================================
#  GSTR-2B LOADER
# ============================================================
def _load_b2b_sheet(file_obj):
    wb = openpyxl.load_workbook(file_obj, read_only=False, data_only=True)
    sheets = wb.sheetnames
    b2b_name = None
    for s in sheets:
        if s.strip().upper() == "B2B": b2b_name = s; break
    if not b2b_name:
        for s in sheets:
            if re.match(r"^B2B", s.strip(), re.IGNORECASE) and "CDNR" not in s.upper():
                b2b_name = s; break
    if not b2b_name:
        for s in sheets:
            if "b2b" in s.lower() and "cdnr" not in s.lower(): b2b_name = s; break
    if not b2b_name:
        raise ValueError(f"Could not find B2B sheet.\nSheets found: {sheets}")

    ws = wb[b2b_name]; all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return pd.DataFrame(columns=list(B2B_COL_POS_FALLBACK.values())), b2b_name, {}

    det = {"method": "", "header_row": 0, "col_map": {}}
    gstr2b_kws = ["gstin","invoice","taxable","igst","cgst","sgst","supplier","trade","legal","integrated","central"]
    header_idx = _find_header_row(all_rows, max_scan=20, col_keywords=gstr2b_kws)
    det["header_row"] = header_idx + 1

    hr = [str(v).strip() if v is not None else "" for v in all_rows[header_idx]]
    nr = [str(v).strip() if v is not None else "" for v in all_rows[header_idx+1]] \
         if header_idx+1 < len(all_rows) else []
    merged = [f"{hr[i]} {nr[i] if i < len(nr) else ''}".strip() or f"Col_{i}" for i in range(len(hr))]

    col_map = {}
    for field in GSTR2B_COL_KEYWORDS:
        found = _find_col(merged, field, GSTR2B_COL_KEYWORDS)
        if found and found not in col_map.values(): col_map[field] = found

    essential = ["GSTIN","Invoice No","Invoice Date","Taxable","IGST","CGST","SGST"]
    found_ess = sum(1 for e in essential if e in col_map)

    data_start = None
    for ds in [header_idx+1, header_idx+2, header_idx+3]:
        if ds >= len(all_rows): continue
        for v in all_rows[ds]:
            sv = str(v).strip().upper().replace(" ","")
            if len(sv) == 15 and sv[:2].isdigit(): data_start = ds; break
        if data_start is not None: break
    if data_start is None:
        for i, row in enumerate(all_rows):
            for v in row:
                sv = str(v).strip().upper().replace(" ","")
                if len(sv) == 15 and sv[:2].isdigit(): data_start = i; break
            if data_start is not None: break
    if data_start is None:
        return pd.DataFrame(columns=list(B2B_COL_POS_FALLBACK.values())), b2b_name, det

    if found_ess >= 4:
        det["method"] = f"Column-Name Detection ({found_ess}/7 essential found)"
        det["col_map"] = col_map
        col_idx_map = {}
        for field, orig in col_map.items():
            for i, h in enumerate(merged):
                if h == orig: col_idx_map[field] = i; break
        records = []
        for row in all_rows[data_start:]:
            if not any(v is not None for v in row): continue
            gi = col_idx_map.get("GSTIN", 0)
            gv = row[gi] if gi < len(row) else None
            if gv is None: continue
            sv = str(gv).strip().upper().replace(" ","")
            if not (8 <= len(sv) <= 15): continue
            rec = {}
            for field in B2B_COL_POS_FALLBACK.values():
                idx = col_idx_map.get(field)
                rec[field] = row[idx] if idx is not None and idx < len(row) else None
            records.append(rec)
    else:
        det["method"] = "Fixed-Position Fallback"
        det["col_map"] = {v: f"Col {k}" for k, v in B2B_COL_POS_FALLBACK.items()}
        records = []
        for row in all_rows[data_start:]:
            if not any(v is not None for v in row): continue
            col_a = row[0] if row else None
            if col_a is None: continue
            sv = str(col_a).strip().upper().replace(" ","")
            if not (8 <= len(sv) <= 15): continue
            rec = {}
            for ci, cn in B2B_COL_POS_FALLBACK.items():
                rec[cn] = row[ci-1] if ci-1 < len(row) else None
            records.append(rec)

    if not records:
        return pd.DataFrame(columns=list(B2B_COL_POS_FALLBACK.values())), b2b_name, det

    df = pd.DataFrame(records)
    df["GSTIN"] = df["GSTIN"].astype(str).str.strip().str.upper().str.replace(r"\s+","",regex=True)
    df = df[df["GSTIN"].str.match(r"^\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z][A-Z0-9]$", na=False)].copy()
    df["Vendor"]       = df.get("Vendor",     pd.Series([""]*len(df))).fillna("").astype(str).str.strip()
    df["Invoice No"]   = df.get("Invoice No", pd.Series([""]*len(df))).fillna("").astype(str).str.strip()
    df["Invoice Date"] = pd.to_datetime(df.get("Invoice Date"), errors="coerce", dayfirst=True)
    for col in ["Taxable","IGST","CGST","SGST","Cess","Invoice Value"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        else: df[col] = 0.0
    df["Total Tax"] = df["IGST"] + df["CGST"] + df["SGST"]
    return df.reset_index(drop=True), b2b_name, det


# ============================================================
#  BOOKS FILE READER
# ============================================================
def _read_books_file(file_obj, manual_header_row=None, manual_sheet=None):
    file_obj.seek(0); raw = file_obj.read()
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=False, data_only=True)
    sheets = wb.sheetnames
    best_df = None; best_score = -1; best_info = {}
    target  = [manual_sheet] if manual_sheet and manual_sheet in sheets else sheets

    for sn in target:
        ws = wb[sn]; all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows: continue
        hidx = (max(0, manual_header_row-1) if manual_header_row is not None
                else _find_header_row(all_rows, max_scan=20))
        if hidx >= len(all_rows): continue
        header = []; seen = {}
        for j, v in enumerate(all_rows[hidx]):
            h = str(v).strip() if v is not None else f"Col_{j}"
            if h in seen: seen[h] += 1; h = f"{h}_{seen[h]}"
            else: seen[h] = 0
            header.append(h)
        data_rows = []
        for row in all_rows[hidx+1:]:
            if any(v is not None for v in row):
                rl = list(row)
                while len(rl) < len(header): rl.append(None)
                data_rows.append(rl[:len(header)])
        if not data_rows: continue
        df = pd.DataFrame(data_rows, columns=header).dropna(how="all").reset_index(drop=True)
        score = 0
        for col in df.columns:
            vals = df[col].dropna().astype(str).str.strip()
            score += vals.str.match(r"^\d{2}[A-Z0-9]{13}$").sum() * 10
            if any(k in col.lower() for k in ["gstin","gst","invoice","bill","taxable","igst","cgst","sgst"]):
                score += 5
        info = {"sheet": sn, "header_row": hidx+1, "score": score, "warnings": []}
        if score > best_score: best_score = score; best_df = df; best_info = info

    if best_df is None or best_df.empty:
        xl = pd.ExcelFile(io.BytesIO(raw))
        for sn in xl.sheet_names:
            try:
                df = xl.parse(sn).dropna(how="all").reset_index(drop=True)
                if not df.empty:
                    return df, {"sheet": sn, "header_row": 1, "score": 0, "warnings": ["Last-resort reader"]}
            except Exception: continue
        raise ValueError("Purchase Register is empty or unreadable.\nUse 'Fix Purchase Register Format' in the sidebar.")
    return best_df, best_info


# ============================================================
#  BOOKS NORMALIZER
# ============================================================
def _load_books(df_raw, extra_hints=None):
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if extra_hints:
        for field, hint in extra_hints.items():
            if hint and field in BOOKS_COL_MAP:
                BOOKS_COL_MAP[field] = [hint.lower()] + BOOKS_COL_MAP[field]
    mapping = {field: found for field in BOOKS_COL_MAP
               if (found := _find_col(list(df.columns), field, BOOKS_COL_MAP))}
    df = df.rename(columns={v: k for k, v in mapping.items() if v in df.columns})
    for col in ["GSTIN","Invoice No","Vendor"]:
        if col not in df.columns: df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["GSTIN"] = df["GSTIN"].str.upper().str.replace(r"\s+","",regex=True)
    df = df[df["GSTIN"].str.len().between(10, 15)].copy()
    if df.empty:
        df["Invoice Date"] = pd.NaT
        for c in ["IGST","CGST","SGST","Taxable"]: df[c] = 0.0
        df["Total Tax"] = 0.0
        return df.reset_index(drop=True), mapping
    if "Invoice Date" not in df.columns: df["Invoice Date"] = pd.NaT
    else: df["Invoice Date"] = pd.to_datetime(df["Invoice Date"], errors="coerce", dayfirst=True)
    for c in ["IGST","CGST","SGST","Taxable"]:
        if c not in df.columns: df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Total Tax"] = df["IGST"] + df["CGST"] + df["SGST"]
    return df.reset_index(drop=True), mapping


# ============================================================
#  INVOICE MATCHING
# ============================================================
def _norm(s):   return re.sub(r"[^a-z0-9]","",str(s).lower().strip())
def _dig(s):    return re.sub(r"\D","",str(s))

def _match_invoice(a, b):
    if not a or not b or str(a) in ("nan","None","") or str(b) in ("nan","None",""): return False,""
    na,nb = _norm(a),_norm(b); da,db = _dig(a),_dig(b)
    if na==nb and na:                                       return True,"Exact"
    if da and db and da==db:                                return True,"Numeric Exact"
    if da and db and len(da)>=2 and len(db)>=2:
        if da.endswith(db) or db.endswith(da):              return True,"Numeric Suffix"
    ml = min(len(na),len(nb))
    if ml>=4:
        if na[-ml:]==nb[-ml:]:                              return True,"Alpha Suffix"
        if na[:ml]==nb[:ml]:                                return True,"Alpha Prefix"
    if da and db and len(da)>=3 and len(db)>=3:
        if da.lstrip("0")==db.lstrip("0"):                  return True,"Leading Zero"
        if da in db or db in da:                            return True,"Numeric Contains"
    na2,nb2 = na,nb
    for p in ["purchase","pur","bill","inv","gst","tax","rcm","import"]:
        if na2.startswith(p): na2=na2[len(p):]
        if nb2.startswith(p): nb2=nb2[len(p):]
    if na2 and nb2 and na2==nb2 and len(na2)>=3:            return True,"Prefix Stripped"
    return False,""


# ============================================================
#  RECONCILIATION ENGINE
# ============================================================
def run_reco(gstr2b_file, books_df_raw, date_tol, amt_tol, taxable_tol, extra_hints=None):
    gstr2b, b2b_sheet, gstr2b_det = _load_b2b_sheet(gstr2b_file)
    books,  b_map                 = _load_books(books_df_raw, extra_hints)
    if gstr2b.empty:
        raise ValueError(f"B2B sheet '{b2b_sheet}' has no valid data rows.")
    if books.empty:
        raise ValueError(
            "Purchase Register has no valid rows.\n"
            f"Columns found: {list(books_df_raw.columns[:15])}\n"
            "Use 'Fix Purchase Register Format' in the sidebar."
        )
    by_gstin = {}
    for j, bk in books.iterrows(): by_gstin.setdefault(bk["GSTIN"], []).append((j, bk))
    by_pan = {}
    for g, ents in by_gstin.items():
        if len(g)==15: by_pan.setdefault(g[2:12],[]).extend(ents)

    results=[]; matched=set()
    for _, r2b in gstr2b.iterrows():
        g2b=r2b["GSTIN"]; i2b=str(r2b["Invoice No"])
        bj=bbk=None; bm=""
        cands = list(by_gstin.get(g2b,[]))
        if not cands and len(g2b)==15: cands=list(by_pan.get(g2b[2:12],[]))
        for j,bk in cands:
            if j in matched: continue
            ok,mt = _match_invoice(i2b, str(bk["Invoice No"]))
            if ok: bj,bbk,bm=j,bk,mt; break

        if bbk is not None:
            matched.add(bj); bk=bbk
            dd=0
            if pd.notna(r2b["Invoice Date"]) and pd.notna(bk["Invoice Date"]):
                dd=abs((r2b["Invoice Date"]-bk["Invoice Date"]).days)
            diffs=[]
            if dd                                  > date_tol:    diffs.append(f"Date Diff ({dd}d)")
            if abs(r2b["IGST"]   -bk["IGST"])      > amt_tol:     diffs.append(f"IGST Diff ₹{abs(r2b['IGST']-bk['IGST']):,.2f}")
            if abs(r2b["CGST"]   -bk["CGST"])      > amt_tol:     diffs.append(f"CGST Diff ₹{abs(r2b['CGST']-bk['CGST']):,.2f}")
            if abs(r2b["SGST"]   -bk["SGST"])      > amt_tol:     diffs.append(f"SGST Diff ₹{abs(r2b['SGST']-bk['SGST']):,.2f}")
            if abs(r2b["Taxable"]-bk["Taxable"])   > taxable_tol: diffs.append(f"Taxable Diff ₹{abs(r2b['Taxable']-bk['Taxable']):,.2f}")
            results.append({"GSTIN":g2b,"Vendor (2B)":r2b["Vendor"],"Vendor (Books)":bk.get("Vendor",""),
                "Invoice No (2B)":r2b["Invoice No"],"Invoice No (Books)":bk["Invoice No"],
                "Invoice Date (2B)":r2b["Invoice Date"],"Invoice Date (Books)":bk["Invoice Date"],
                "Taxable (2B)":r2b["Taxable"],"Taxable (Books)":bk["Taxable"],
                "IGST (2B)":r2b["IGST"],"IGST (Books)":bk["IGST"],
                "CGST (2B)":r2b["CGST"],"CGST (Books)":bk["CGST"],
                "SGST (2B)":r2b["SGST"],"SGST (Books)":bk["SGST"],
                "Total Tax (2B)":r2b["Total Tax"],"Total Tax (Books)":bk["Total Tax"],
                "ITC":r2b.get("ITC",""),"Match Method":bm,
                "Status":"MATCHED WITH DIFF" if diffs else "MATCHED",
                "Remarks":" | ".join(diffs) if diffs else "Exact Match"})
        else:
            results.append({"GSTIN":g2b,"Vendor (2B)":r2b["Vendor"],"Vendor (Books)":"",
                "Invoice No (2B)":r2b["Invoice No"],"Invoice No (Books)":"",
                "Invoice Date (2B)":r2b["Invoice Date"],"Invoice Date (Books)":pd.NaT,
                "Taxable (2B)":r2b["Taxable"],"Taxable (Books)":None,
                "IGST (2B)":r2b["IGST"],"IGST (Books)":None,
                "CGST (2B)":r2b["CGST"],"CGST (Books)":None,
                "SGST (2B)":r2b["SGST"],"SGST (Books)":None,
                "Total Tax (2B)":r2b["Total Tax"],"Total Tax (Books)":None,
                "ITC":r2b.get("ITC",""),"Match Method":"",
                "Status":"IN 2B – NOT IN BOOKS","Remarks":"Not found in Purchase Register"})

    for j, bk in books.iterrows():
        if j not in matched:
            results.append({"GSTIN":bk["GSTIN"],"Vendor (2B)":"","Vendor (Books)":bk.get("Vendor",""),
                "Invoice No (2B)":"","Invoice No (Books)":bk["Invoice No"],
                "Invoice Date (2B)":pd.NaT,"Invoice Date (Books)":bk["Invoice Date"],
                "Taxable (2B)":None,"Taxable (Books)":bk["Taxable"],
                "IGST (2B)":None,"IGST (Books)":bk["IGST"],
                "CGST (2B)":None,"CGST (Books)":bk["CGST"],
                "SGST (2B)":None,"SGST (Books)":bk["SGST"],
                "Total Tax (2B)":None,"Total Tax (Books)":bk["Total Tax"],
                "ITC":"","Match Method":"",
                "Status":"IN BOOKS – NOT IN 2B","Remarks":"Not uploaded on GST Portal"})

    out  = pd.DataFrame(results) if results else pd.DataFrame()
    meta = {"b2b_sheet":b2b_sheet,"b2b_count":len(gstr2b),"books_count":len(books),
            "b_map":b_map,"books_cols":list(books_df_raw.columns[:25]),"gstr2b_det":gstr2b_det}
    return out, meta


# ============================================================
#  EXCEL OUTPUT BUILDER
# ============================================================
def build_excel(df):
    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    SS   = {"MATCHED":("E2EFDA","1E4620"),"MATCHED WITH DIFF":("FFF2CC","7B4F00"),
            "IN 2B – NOT IN BOOKS":("FCE4D6","843C0C"),"IN BOOKS – NOT IN 2B":("DAEEF3","0D3D5E")}
    ws = wb.create_sheet("Reconciliation"); ws.sheet_properties.tabColor="1F3864"
    ws.merge_cells(f"A1:{get_column_letter(len(df.columns))}1")
    ws["A1"].value     = f"GST 2B vs Books — Reconciliation  |  {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    ws["A1"].font      = Font(name="Calibri",bold=True,size=12,color="FFFFFF")
    ws["A1"].fill      = PatternFill("solid",fgColor="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height = 26
    for ci, col in enumerate(df.columns,1):
        c=ws.cell(row=2,column=ci,value=col)
        c.font=Font(name="Calibri",bold=True,size=10,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor="2E75B6")
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border=bdr
    ws.row_dimensions[2].height=22
    MONEY={c for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST","Total Tax"])}
    for ri,(_, row) in enumerate(df.iterrows()):
        r=3+ri; bg,fg=SS.get(row.get("Status",""),("FFFFFF","111111"))
        for ci, val in enumerate(row,1):
            cn=df.columns[ci-1]; cell=ws.cell(row=r,column=ci,value=val)
            cell.fill=PatternFill("solid",fgColor=bg); cell.border=bdr
            cell.font=Font(name="Calibri",size=10,bold=(cn=="Status"),color=fg)
            if cn in MONEY:
                cell.number_format="##,##,##0.00"
                cell.alignment=Alignment(horizontal="right",vertical="center")
            else: cell.alignment=Alignment(horizontal="left",vertical="center")
    CW={"GSTIN":22,"Vendor":26,"Invoice No (2B)":18,"Invoice No (Books)":18,
        "Invoice Date (2B)":16,"Invoice Date (Books)":16,
        "Taxable (2B)":14,"Taxable (Books)":14,
        "IGST (2B)":12,"IGST (Books)":12,"CGST (2B)":12,"CGST (Books)":12,
        "SGST (2B)":12,"SGST (Books)":12,"Status":22,"Remarks":34}
    for ci, col in enumerate(df.columns,1):
        ws.column_dimensions[get_column_letter(ci)].width=CW.get(col,14)
    ws.freeze_panes="A3"
    ws2=wb.create_sheet("Summary"); ws2.sheet_properties.tabColor="375623"
    ws2.merge_cells("A1:B1")
    ws2["A1"].value=  "Reconciliation Summary"
    ws2["A1"].font=   Font(name="Calibri",bold=True,size=12,color="FFFFFF")
    ws2["A1"].fill=   PatternFill("solid",fgColor="1F3864")
    ws2["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=26
    for ri,(lbl,cnt,clr) in enumerate([
        ("Total GSTR-2B Records", (df["Invoice No (2B)"]!="").sum(),"DAEEF3"),
        ("Total Books Records",   (df["Invoice No (Books)"]!="").sum(),"DAEEF3"),
        ("Matched — Exact",       (df["Status"]=="MATCHED").sum(),"E2EFDA"),
        ("Matched — With Diff",   (df["Status"]=="MATCHED WITH DIFF").sum(),"FFF2CC"),
        ("In 2B — Not in Books",  (df["Status"]=="IN 2B – NOT IN BOOKS").sum(),"FCE4D6"),
        ("In Books — Not in 2B",  (df["Status"]=="IN BOOKS – NOT IN 2B").sum(),"DAEEF3"),
    ]):
        r=ri+2; c1=ws2.cell(row=r,column=1,value=lbl); c2=ws2.cell(row=r,column=2,value=cnt)
        for c in (c1,c2):
            c.fill=PatternFill("solid",fgColor=clr); c.border=bdr; c.font=Font(name="Calibri",size=11)
        c1.alignment=Alignment(horizontal="left",vertical="center")
        c2.alignment=Alignment(horizontal="center",vertical="center")
        c2.font=Font(name="Calibri",size=11,bold=True); ws2.row_dimensions[r].height=20
    ws2.column_dimensions["A"].width=34; ws2.column_dimensions["B"].width=14
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ============================================================
#  SESSION STATE INIT
# ============================================================
if "ran" not in st.session_state:         st.session_state["ran"]         = False
if "logged_in_email" not in st.session_state: st.session_state["logged_in_email"] = ""
# track how many anonymous runs this session
if "anon_runs" not in st.session_state:   st.session_state["anon_runs"]   = 0


# ============================================================
#  SIDEBAR  — Tolerances + Format Fixer only
# ============================================================
with st.sidebar:
    st.markdown("## 🧾 GST Reco Tool")
    st.markdown("---")

    st.markdown("#### ⚙️ Tolerance Settings")
    DATE_TOL    = st.slider("Date Tolerance (days)",        0,   15,   5)
    AMT_TOL     = st.slider("GST Amount Tolerance (₹)",    0.0, 50.0, 2.0,  step=0.5)
    TAXABLE_TOL = st.slider("Taxable Value Tolerance (₹)", 0.0,100.0,10.0,  step=1.0)

    st.markdown("---")
    st.markdown("#### 🛠️ Fix Purchase Register Format")
    st.markdown("""<p style="font-size:0.78rem;color:#94a3b8;line-height:1.6;">
Use this <b>only if</b> results look wrong — zero matches, errors, or missing data.<br>
<b>Most files work automatically</b> without touching this.</p>""", unsafe_allow_html=True)

    with st.expander("▶  Open format fixer"):
        st.markdown("**Problem 1 — Columns start on wrong row?**")
        st.markdown("""<p style="font-size:0.76rem;color:#94a3b8;line-height:1.5;">
        Some files (Tally, Busy) have title rows before the actual column headings.
        Enter the row number where column names appear.</p>""", unsafe_allow_html=True)
        use_mhr = st.checkbox("Header is not on row 1", value=False, key="use_mhr")
        mhr_val = st.number_input("Header row number", min_value=2, max_value=50, value=7, step=1,
                                   key="mhr_val", help="e.g. if column names are on row 7, enter 7") \
                  if use_mhr else None

        st.markdown("---")
        st.markdown("**Problem 2 — Wrong sheet selected?**")
        st.markdown("""<p style="font-size:0.76rem;color:#94a3b8;line-height:1.5;">
        Specify the exact sheet name if auto-selection picked wrong one.</p>""", unsafe_allow_html=True)
        use_msn = st.checkbox("Pick sheet manually", value=False, key="use_msn")
        msn_val = st.text_input("Exact sheet name", placeholder="e.g. Purchase Register",
                                 key="msn_val") if use_msn else ""

        st.markdown("---")
        st.markdown("**Problem 3 — Column names not detected?**")
        st.markdown("""<p style="font-size:0.76rem;color:#94a3b8;line-height:1.5;">
        Type the exact column names from your file if auto-detection misses them.</p>""", unsafe_allow_html=True)
        hint_gstin = st.text_input("GSTIN column name",       placeholder="e.g.  Vendor GSTIN",  key="hint_gstin")
        hint_inv   = st.text_input("Invoice No column name",  placeholder="e.g.  Voucher No",    key="hint_inv")
        hint_date  = st.text_input("Invoice Date column name",placeholder="e.g.  Posting Date",  key="hint_date")

    st.markdown("---")
    st.markdown("#### 📋 Required Columns")
    with st.expander("GSTR-2B"):
        st.markdown("- `GSTIN`  - `Trade/Legal name`  - `Invoice No`\n- `Invoice Date`  - `Taxable Value`  - `IGST`, `CGST`, `SGST`")
    with st.expander("Purchase Register"):
        st.markdown("- `GSTIN`  - `Vendor / Party`  - `Invoice No`\n- `Invoice Date`  - `Taxable Value`  - `IGST`, `CGST`, `SGST`")

    # ── Account section (shown after login) ──────────────────
    _sid_email = st.session_state.get("logged_in_email","")
    if _sid_email:
        st.markdown("---")
        st.markdown("#### 👤 My Account")
        _su = _get_user(_sid_email)
        _sp = _is_paid(_sid_email)
        _su_used = _get_yearly_usage(_sid_email)
        st.markdown(f"**{_su.get('name') or _sid_email}**")
        if _sp:
            st.markdown("🟢 Subscribed — Unlimited")
        else:
            pct = min(_su_used/FREE_LIMIT, 1.0)
            bc  = "#ef4444" if pct>=1.0 else ("#f59e0b" if pct>=0.6 else "#22c55e")
            st.markdown(f"Free uses: **{_su_used} / {FREE_LIMIT}** this year")
            st.markdown(f"""<div class="usage-bar-wrap">
              <div class="usage-bar-fill" style="width:{pct*100:.0f}%;background:{bc};"></div>
            </div>""", unsafe_allow_html=True)
        with st.expander("✏️ Edit Profile"):
            un = st.text_input("Name",    value=_su.get("name",""),  key="sb_name")
            uf = st.text_input("Firm",    value=_su.get("firm",""),  key="sb_firm")
            ug = st.text_input("GSTIN",   value=_su.get("gstin",""), key="sb_gstin")
            if st.button("💾 Save", key="sb_save"):
                _update_profile(_sid_email, un, uf, ug)
                st.success("Saved!")
        if st.button("🔓 Sign Out", key="sb_signout"):
            st.session_state["logged_in_email"] = ""
            st.session_state["ran"] = False
            st.rerun()


# ============================================================
#  MAIN PAGE
# ============================================================
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


# ============================================================
#  STEP 1 — UPLOAD FILES  (always visible, no gating)
# ============================================================
st.markdown('<div class="section-label">Step 1 — Upload Files</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Select Your Excel Files</div>', unsafe_allow_html=True)

uc1, uc2 = st.columns(2)
with uc1:
    st.markdown('<div class="box-label">📥 GSTR-2B File</div>', unsafe_allow_html=True)
    gstr_file = st.file_uploader("GSTR-2B", type=["xlsx","xls"],
                                  label_visibility="collapsed", key="gstr")
    if gstr_file: st.success(f"✅ {gstr_file.name}")

with uc2:
    st.markdown('<div class="box-label">📗 Books / Purchase Register</div>', unsafe_allow_html=True)
    books_file = st.file_uploader("Books", type=["xlsx","xls"],
                                   label_visibility="collapsed", key="books")
    if books_file: st.success(f"✅ {books_file.name}")

st.markdown("<br>", unsafe_allow_html=True)
btn_col, _ = st.columns([1, 3])
with btn_col:
    run_btn = st.button("▶  Run Reconciliation", use_container_width=True)


# ============================================================
#  RUN LOGIC
# ============================================================
_email     = st.session_state.get("logged_in_email","")
_anon_runs = st.session_state.get("anon_runs", 0)

# Determine if this session is over the anonymous free limit
# Anonymous users: FREE_LIMIT runs per browser session (no persistence)
# Logged-in users: FREE_LIMIT runs per year (persisted), paid = unlimited
_logged_in  = bool(_email)
_paid_user  = _logged_in and _is_paid(_email)
_user_used  = _get_yearly_usage(_email) if _logged_in else 0

# Check if run is allowed
def _run_allowed():
    if not _logged_in:
        return _anon_runs < FREE_LIMIT          # anonymous: session-based limit
    ok, _ = _can_run(_email)
    return ok

if run_btn:
    if not gstr_file or not books_file:
        st.error("Please upload both files before running.")

    elif not _run_allowed():
        # ── LIMIT EXCEEDED: show login + payment wall ─────────
        st.markdown("""
        <div class="pay-card">
          <div class="pay-title">🔒 Free Usage Limit Reached</div>
          <div class="pay-desc">
            You've used all <strong>5 free reconciliations</strong>.<br>
            Log in to get 5 free runs per year tracked to your account, or subscribe for unlimited access.
          </div>
        </div>
        """, unsafe_allow_html=True)

        lc1, lc2, _ = st.columns([1, 1, 2])
        with lc1:
            st.markdown("""<div class="login-card">
              <div class="login-title">🔑 Log In with Gmail</div>
              <div class="login-desc">Sign in to continue with your free yearly allowance or subscribe.</div>
            </div>""", unsafe_allow_html=True)
            login_em = st.text_input("Gmail Address", placeholder="yourname@gmail.com", key="wall_email")
            if st.button("✅  Sign In", use_container_width=True, key="wall_signin"):
                em = login_em.strip().lower()
                if not em or "@" not in em:
                    st.error("Enter a valid Gmail address.")
                else:
                    _get_user(em)
                    st.session_state["logged_in_email"] = em
                    st.session_state["anon_runs"]       = 0
                    st.rerun()
        with lc2:
            st.markdown("""<div class="pay-card" style="height:100%">
              <div class="pay-title">💳 Subscribe</div>
              <div class="pay-desc">Unlimited reconciliations every month.</div>
              <div><span class="pay-price">₹49</span>
              <span class="pay-sub">/ month</span></div>
            </div>""", unsafe_allow_html=True)
            if st.button("💳  Pay ₹49 — Subscribe Now", use_container_width=True, key="wall_pay"):
                if _logged_in:
                    _mark_paid(_email)
                    st.success("✅ Subscribed! Unlimited access activated.")
                    st.rerun()
                else:
                    st.warning("Please log in first, then subscribe.")

    else:
        # ── RUN RECONCILIATION ────────────────────────────────
        # Check logged-in user's yearly limit specifically
        if _logged_in and not _paid_user and _user_used >= FREE_LIMIT:
            st.error("🔒 Yearly free limit reached. Please subscribe to continue.")
        else:
            with st.spinner("Reconciling…"):
                try:
                    _manual_hr = int(st.session_state.get("mhr_val", 7)) \
                                 if st.session_state.get("use_mhr") else None
                    _manual_sn = st.session_state.get("msn_val","").strip() \
                                 if st.session_state.get("use_msn") else None
                    _hints = {
                        "GSTIN":        st.session_state.get("hint_gstin","").strip(),
                        "Invoice No":   st.session_state.get("hint_inv","").strip(),
                        "Invoice Date": st.session_state.get("hint_date","").strip(),
                    }
                    _hints = {k: v for k, v in _hints.items() if v}

                    books_file.seek(0)
                    b_df, parse_info = _read_books_file(books_file,
                                                        manual_header_row=_manual_hr,
                                                        manual_sheet=_manual_sn or None)
                    gstr_file.seek(0)
                    result_df, meta  = run_reco(gstr_file, b_df,
                                                DATE_TOL, AMT_TOL, TAXABLE_TOL,
                                                extra_hints=_hints or None)

                    st.session_state["result"]     = result_df
                    st.session_state["meta"]       = meta
                    st.session_state["parse_info"] = parse_info
                    st.session_state["ran"]        = True

                    # Increment counters
                    if _logged_in:
                        _increment_usage(_email)
                    else:
                        st.session_state["anon_runs"] += 1

                except Exception as e:
                    st.error(f"❌ {e}")
                    try:
                        gstr_file.seek(0); books_file.seek(0)
                        wb_g = openpyxl.load_workbook(gstr_file,  read_only=True)
                        wb_b = openpyxl.load_workbook(books_file, read_only=True)
                        with st.expander("🔍 Diagnostic — sheets & rows"):
                            st.write("**GSTR-2B sheets:**", wb_g.sheetnames)
                            st.write("**Books sheets:**",   wb_b.sheetnames)
                            for sn in wb_b.sheetnames:
                                ws   = wb_b[sn]
                                rows = list(ws.iter_rows(max_row=12, values_only=True))
                                for i, r in enumerate(rows):
                                    if sum(1 for v in r if v is not None) >= 3:
                                        st.write(f"Sheet *{sn}* — Row {i+1}:", list(r)); break
                            st.markdown("💡 Open **Fix Purchase Register Format** in the sidebar and set the correct header row.")
                    except Exception: pass
                    st.session_state["ran"] = False


# ============================================================
#  RESULTS
# ============================================================
if st.session_state.get("ran") and "result" in st.session_state:
    df = st.session_state["result"]
    matched      = (df["Status"]=="MATCHED").sum()
    matched_diff = (df["Status"]=="MATCHED WITH DIFF").sum()
    only_2b      = (df["Status"]=="IN 2B – NOT IN BOOKS").sum()
    only_books   = (df["Status"]=="IN BOOKS – NOT IN 2B").sum()
    total        = len(df)
    rate         = f"{matched/total*100:.1f}%" if total else "0%"

    st.markdown("---")
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total",              total)
    m2.metric("✅ Matched",          matched,      rate)
    m3.metric("⚠️ With Differences", matched_diff)
    m4.metric("❌ 2B Only",          only_2b)
    m5.metric("📘 Books Only",       only_books)

    st.markdown("""<div class="legend-row">
      <span class="leg leg-green">✅ Matched — Exact</span>
      <span class="leg leg-yellow">⚠️ Matched with Differences</span>
      <span class="leg leg-red">❌ In 2B — Not in Books</span>
      <span class="leg leg-blue">📘 In Books — Not in 2B</span>
    </div>""", unsafe_allow_html=True)

    if "meta" in st.session_state:
        meta = st.session_state["meta"]; pi = st.session_state.get("parse_info",{})
        gd   = meta.get("gstr2b_det",{})
        with st.expander(f"🔍 Detection Report  |  2B Sheet: **{meta['b2b_sheet']}**  |  2B: {meta['b2b_count']} recs  |  Books: {meta['books_count']} recs"):
            dc1,dc2 = st.columns(2)
            with dc1:
                st.markdown(f"**GSTR-2B:** `{gd.get('method','Fixed Position')}` · Header row `{gd.get('header_row','N/A')}`")
            with dc2:
                st.markdown(f"**Purchase Register:** Sheet `{pi.get('sheet','N/A')}` · Header row `{pi.get('header_row','N/A')}`")
            st.markdown("---")
            st.markdown("**Purchase Register — Detected Column Mapping:**")
            if meta["b_map"]:
                ci_cols = st.columns(3); items = list(meta["b_map"].items())
                chunk   = max(1, len(items)//3+1)
                for i, col in enumerate(ci_cols):
                    with col:
                        for field,orig in items[i*chunk:(i+1)*chunk]:
                            st.markdown(f"- `{orig}` → **{field}**")
            else:
                st.warning("No columns detected. Use 'Fix Purchase Register Format' in sidebar.")

    tabs = st.tabs(["All","✅ Matched","⚠️ Differences","❌ 2B Only","📘 Books Only"])
    MC   = {c: st.column_config.NumberColumn(format="₹%.2f")
            for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST","Total Tax"])}
    DC   = {c: st.column_config.DateColumn(c, format="DD-MMM-YYYY")
            for c in ["Invoice Date (2B)","Invoice Date (Books)"] if c in df.columns}

    def show(data):
        if data.empty: st.info("No records in this category."); return
        st.dataframe(data, use_container_width=True, hide_index=True, column_config={**MC,**DC})
        st.caption(f"{len(data):,} records")

    for tab,flt in zip(tabs,[None,"MATCHED","MATCHED WITH DIFF","IN 2B – NOT IN BOOKS","IN BOOKS – NOT IN 2B"]):
        with tab: show(df if flt is None else df[df["Status"]==flt])

    st.markdown("---")
    dl_col, _ = st.columns([1,3])
    with dl_col:
        st.download_button("📥  Download Excel Report",
            data=build_excel(df),
            file_name=f"GST_Reco_{datetime.now().strftime('%d%b%Y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    st.caption(f"Report: Reconciliation Detail + Summary  |  {datetime.now().strftime('%d-%b-%Y %H:%M')}")


# ============================================================
#  CONTACT US
# ============================================================
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

# ============================================================
#  FOOTER
# ============================================================
st.markdown(f"""
<div class="page-footer">
  <div class="footer-left">
    <span class="footer-brand">GST 2B Reconciliation Tool</span><br>
    Simplifying GST compliance for businesses &amp; tax professionals
  </div>
  <div class="footer-right">© {datetime.now().year} &nbsp;·&nbsp; All rights reserved</div>
</div>
""", unsafe_allow_html=True)
