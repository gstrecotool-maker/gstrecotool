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
#  SAMPLE EXCEL BUILDERS
# ============================================================
def _hdr_style(cell, bg):
    cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="D0D0D0")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)


def make_sample_gstr2b() -> bytes:
    wb, ws = Workbook(), None
    wb.remove(wb.active) if wb.active else None
    ws = wb.create_sheet("GSTR-2B")
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["GSTIN", "Trade/Legal name", "Invoice No",
               "Invoice Date", "Taxable Value (₹)", "IGST", "CGST", "SGST"]
    for c_i, h in enumerate(headers, 1):
        _hdr_style(ws.cell(row=1, column=c_i, value=h), "1F3864")
    ws.row_dimensions[1].height = 20

    data = [
        ["27AABCU9603R1ZX", "ABC Traders",       "INV/2024/001", "05-04-2024", 50000,  9000,  0,    0   ],
        ["27AABCU9603R1ZX", "ABC Traders",       "INV/2024/002", "12-04-2024", 30000,  0,     2700, 2700],
        ["29AAKCS5153R1Z4", "XYZ Suppliers",     "PUR/001",      "18-04-2024", 75000,  13500, 0,    0   ],
        ["29AAKCS5153R1Z4", "XYZ Suppliers",     "PUR/002",      "22-04-2024", 20000,  0,     1800, 1800],
        ["07AAACR5055K1Z5", "Royal Enterprises", "RE/24-25/015", "30-04-2024", 45000,  8100,  0,    0   ],
        ["07AAACR5055K1Z5", "Royal Enterprises", "RE/24-25/016", "05-05-2024", 15000,  0,     1350, 1350],
        ["33AABCF1234G1ZK", "Fresh Goods Co",    "FG/001",       "10-05-2024", 60000,  10800, 0,    0   ],
        ["33AABCF1234G1ZK", "Fresh Goods Co",    "FG/002",       "15-05-2024", 25000,  0,     2250, 2250],
    ]
    alts = ["EEF4FF", "FFFFFF"]
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i+2, column=c_i, value=val)
            cell.border = bdr
            cell.font   = Font(name="Calibri", size=10)
            cell.fill   = PatternFill("solid", fgColor=alts[r_i % 2])
            if c_i == 4:   cell.alignment = Alignment(horizontal="center")
            elif c_i >= 5: cell.number_format = "##,##,##0.00"; cell.alignment = Alignment(horizontal="right")
            else:          cell.alignment = Alignment(horizontal="left")

    for i, w in enumerate([22,22,18,14,18,12,12,12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # Instructions
    wi = wb.create_sheet("Instructions")
    wi["A1"] = "Fill your actual GSTR-2B data in the GSTR-2B sheet following this column format."
    wi["A1"].font = Font(name="Calibri", bold=True, size=11, color="1F3864")
    wi.column_dimensions["A"].width = 60
    notes = [
        ("GSTIN",              "15-digit GST number — mandatory for matching"),
        ("Trade/Legal name",   "Supplier name as per GST portal"),
        ("Invoice No",         "Invoice number — alphanumeric"),
        ("Invoice Date",       "Format: DD-MM-YYYY"),
        ("Taxable Value (₹)",  "Amount before GST"),
        ("IGST",               "Put 0 for local/intra-state supplies"),
        ("CGST",               "Put 0 for inter-state supplies"),
        ("SGST",               "Put 0 for inter-state supplies"),
    ]
    for r_i, (col, desc) in enumerate(notes, 3):
        wi.cell(row=r_i, column=1, value=f"  {col}").font  = Font(name="Calibri", bold=True, size=10)
        wi.cell(row=r_i, column=2, value=desc).font        = Font(name="Calibri", size=10, color="555555")
    wi.column_dimensions["B"].width = 50

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


def make_sample_books() -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Purchase Register")
    thin = Side(style="thin", color="D0D0D0")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["GSTIN", "Particulars", "Invoice No",
               "Invoice Date", "Taxable Value", "IGST", "CGST", "SGST"]
    for c_i, h in enumerate(headers, 1):
        _hdr_style(ws.cell(row=1, column=c_i, value=h), "375623")
    ws.row_dimensions[1].height = 20

    data = [
        ["27AABCU9603R1ZX", "ABC Traders",       "INV/2024/001", "05-04-2024", 50000,  9000,  0,    0   ],  # exact
        ["27AABCU9603R1ZX", "ABC Traders",       "INV/2024/002", "12-04-2024", 31500,  0,     2700, 2700],  # taxable diff
        ["29AAKCS5153R1Z4", "XYZ Suppliers",     "PUR/001",      "19-04-2024", 75000,  13500, 0,    0   ],  # date diff
        ["29AAKCS5153R1Z4", "XYZ Suppliers",     "PUR/002",      "22-04-2024", 20000,  0,     1800, 1800],  # exact
        ["07AAACR5055K1Z5", "Royal Enterprises", "RE/24-25/015", "30-04-2024", 45000,  8100,  0,    0   ],  # exact
        # RE/24-25/016 intentionally missing → IN 2B NOT IN BOOKS
        ["33AABCF1234G1ZK", "Fresh Goods Co",    "FG/001",       "10-05-2024", 60000,  10800, 0,    0   ],  # exact
        ["33AABCF1234G1ZK", "Fresh Goods Co",    "FG/002",       "15-05-2024", 25000,  0,     2250, 2250],  # exact
        ["36AABCT1234H1ZP", "Techno Parts Ltd",  "TP/88",        "20-05-2024", 40000,  7200,  0,    0   ],  # books only
    ]
    alts = ["F0FAF0", "FFFFFF"]
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i+2, column=c_i, value=val)
            cell.border = bdr
            cell.font   = Font(name="Calibri", size=10)
            cell.fill   = PatternFill("solid", fgColor=alts[r_i % 2])
            if c_i == 4:   cell.alignment = Alignment(horizontal="center")
            elif c_i >= 5: cell.number_format = "##,##,##0.00"; cell.alignment = Alignment(horizontal="right")
            else:          cell.alignment = Alignment(horizontal="left")

    for i, w in enumerate([22,22,18,14,16,12,12,12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ============================================================
#  RECONCILIATION ENGINE
# ============================================================
def extract_numeric(inv):
    if pd.isna(inv): return ""
    return "".join(re.findall(r"\d+", str(inv)))


def normalize_columns(df):
    df.columns = [c.strip() for c in df.columns]
    return df


def run_reco(gstr2b_df, books_df, date_tol, amt_tol, taxable_tol):
    gstr2b = normalize_columns(gstr2b_df.copy())
    books  = normalize_columns(books_df.copy())

    gstr2b.rename(columns={"GSTIN ":"GSTIN","Trade/Legal name":"Vendor",
                            "Trade/ Legal name":"Vendor","Taxable Value (₹)":"Taxable",
                            "Taxable Value":"Taxable"}, inplace=True)
    books.rename(columns={"Particulars":"Vendor","Party Name":"Vendor",
                           "Supplier Name":"Vendor","Taxable Value":"Taxable",
                           "Taxable Amt":"Taxable","Taxable Amount":"Taxable"}, inplace=True)

    for col in ["IGST","CGST","SGST","Taxable"]:
        for df_ in [gstr2b, books]:
            if col not in df_.columns: df_[col] = 0.0
            df_[col] = pd.to_numeric(df_[col], errors="coerce").fillna(0.0)

    for df_ in [gstr2b, books]:
        df_["Invoice Date"] = pd.to_datetime(df_.get("Invoice Date"), errors="coerce", dayfirst=True)
        df_["INV_NUM"]      = df_["Invoice No"].apply(extract_numeric)

    results, matched_books = [], set()

    for _, r2b in gstr2b.iterrows():
        found = False
        for j, bk in books.iterrows():
            if r2b["GSTIN"] != bk["GSTIN"]: continue
            if not r2b["INV_NUM"] or not bk["INV_NUM"]: continue
            if not (r2b["INV_NUM"].endswith(bk["INV_NUM"]) or
                    bk["INV_NUM"].endswith(r2b["INV_NUM"])): continue

            date_diff = 0
            if pd.notna(r2b["Invoice Date"]) and pd.notna(bk["Invoice Date"]):
                date_diff = abs((r2b["Invoice Date"] - bk["Invoice Date"]).days)

            reasons = []
            if date_diff                            > date_tol:    reasons.append(f"Date Diff ({date_diff}d)")
            if abs(r2b["IGST"]    - bk["IGST"])    > amt_tol:     reasons.append("IGST Diff")
            if abs(r2b["CGST"]    - bk["CGST"])    > amt_tol:     reasons.append("CGST Diff")
            if abs(r2b["SGST"]    - bk["SGST"])    > amt_tol:     reasons.append("SGST Diff")
            if abs(r2b["Taxable"] - bk["Taxable"]) > taxable_tol: reasons.append("Taxable Diff")

            results.append({
                "GSTIN": r2b["GSTIN"], "Vendor": r2b.get("Vendor",""),
                "Invoice No (2B)": r2b["Invoice No"], "Invoice No (Books)": bk["Invoice No"],
                "Invoice Date (2B)": r2b["Invoice Date"], "Invoice Date (Books)": bk["Invoice Date"],
                "Taxable (2B)": r2b["Taxable"],  "Taxable (Books)": bk["Taxable"],
                "IGST (2B)": r2b["IGST"],  "IGST (Books)": bk["IGST"],
                "CGST (2B)": r2b["CGST"],  "CGST (Books)": bk["CGST"],
                "SGST (2B)": r2b["SGST"],  "SGST (Books)": bk["SGST"],
                "Status":  "MATCHED WITH DIFF" if reasons else "MATCHED",
                "Remarks": ", ".join(reasons) if reasons else "Exact Match",
            })
            matched_books.add(j); found = True; break

        if not found:
            results.append({
                "GSTIN": r2b["GSTIN"], "Vendor": r2b.get("Vendor",""),
                "Invoice No (2B)": r2b["Invoice No"], "Invoice No (Books)": "",
                "Invoice Date (2B)": r2b["Invoice Date"], "Invoice Date (Books)": pd.NaT,
                "Taxable (2B)": r2b["Taxable"],  "Taxable (Books)": None,
                "IGST (2B)": r2b["IGST"],  "IGST (Books)": None,
                "CGST (2B)": r2b["CGST"],  "CGST (Books)": None,
                "SGST (2B)": r2b["SGST"],  "SGST (Books)": None,
                "Status": "IN 2B – NOT IN BOOKS", "Remarks": "Not in Purchase Register",
            })

    for j, bk in books.iterrows():
        if j not in matched_books:
            results.append({
                "GSTIN": bk["GSTIN"], "Vendor": bk.get("Vendor",""),
                "Invoice No (2B)": "", "Invoice No (Books)": bk["Invoice No"],
                "Invoice Date (2B)": pd.NaT, "Invoice Date (Books)": bk["Invoice Date"],
                "Taxable (2B)": None, "Taxable (Books)": bk["Taxable"],
                "IGST (2B)": None, "IGST (Books)": bk["IGST"],
                "CGST (2B)": None, "CGST (Books)": bk["CGST"],
                "SGST (2B)": None, "SGST (Books)": bk["SGST"],
                "Status": "IN BOOKS – NOT IN 2B", "Remarks": "Not on GST Portal",
            })

    return pd.DataFrame(results)


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

    MONEY = {c for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST"])}

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
    st.markdown("#### 📥 Sample Files")
    st.caption("Download, fill your data, then upload above.")

    st.download_button("⬇ GSTR-2B Sample", data=make_sample_gstr2b(),
                       file_name="Sample_GSTR2B.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    st.download_button("⬇ Books Sample", data=make_sample_books(),
                       file_name="Sample_Books.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

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
                result_df = run_reco(
                    pd.read_excel(gstr_file),
                    pd.read_excel(books_file),
                    DATE_TOL, AMT_TOL, TAXABLE_TOL,
                )
                st.session_state["result"] = result_df
                st.session_state["ran"]    = True
            except Exception as e:
                err = str(e)
                st.error(f"Error: {err}")
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

    tabs = st.tabs(["All", "✅ Matched", "⚠️ Differences", "❌ 2B Only", "📘 Books Only"])

    MONEY_COLS = {c: st.column_config.NumberColumn(format="₹%.2f")
                  for c in df.columns if any(k in c for k in ["Taxable","IGST","CGST","SGST"])}
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
<div class="contact-val">+91 92700 12217</div>
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
