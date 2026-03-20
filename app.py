#!/usr/bin/env python3
"""
MO Commission Tools — Flask Web App

Step 1: Manager Split      — split masterlog by manager → per-manager xlsx + PDF
Step 2: Distributor Tabs    — one tab per Distrib Code + Summary → xlsx
Step 3: Commission Statements — Invoice List format → xlsx + PDF bundle
"""

import os
import re
import zipfile
import subprocess
import uuid
import shutil
import tempfile
from datetime import datetime
from copy import copy

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XlImage
from openpyxl.cell.cell import MergedCell

from process_commissions import (
    find_summary_header, scan_summary_meta, is_legacy_subtotal,
    build_surgeon_lookup, generate_distributor_tabs, create_summary_tab,
    insert_distributor_subtotals, apply_summary_alignment,
    collect_data_rows, copy_cell,
    get_sheet_ci, keep_sheets_ci,
    _build_group_summaries, _cleanup_tmp_images,
)

# ─── App Setup ────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "uploads")
app.config["OUTPUT_FOLDER"] = os.path.join(os.path.dirname(__file__), "outputs")

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "static", "maxx_logo.png")


# ─── Shared Helpers ───────────────────────────────────────────────────────────

def _make_job_dir():
    """Create and return a new job directory with an 8-char UUID."""
    job_id  = str(uuid.uuid4())[:8]
    job_dir = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
    os.makedirs(job_dir, exist_ok=True)
    return job_id, job_dir


def _save_upload(job_dir):
    """Validate and save the uploaded .xlsx file. Returns (file, input_path)."""
    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        return None, None
    path = os.path.join(job_dir, secure_filename(f.filename))
    f.save(path)
    return f, path


def _copy_sheet_data(src_ws, dst_ws):
    """Deep-copy cell values, styles, merges, and dimensions between sheets."""
    for merge in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merge))
    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width  = dim.width  or 8.43
        dst_ws.column_dimensions[col_letter].hidden = dim.hidden
    for row_num, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_num].height = dim.height or 15
        dst_ws.row_dimensions[row_num].hidden = getattr(dim, "hidden", False)
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row,
                                 max_col=src_ws.max_column):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            nc = dst_ws.cell(row=cell.row, column=cell.column)
            nc.value = cell.value
            if cell.has_style:
                nc.font          = copy(cell.font)
                nc.border        = copy(cell.border)
                nc.fill          = copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection    = copy(cell.protection)
                nc.alignment     = copy(cell.alignment)


def _convert_to_pdf_libreoffice(xlsx_path, out_dir, lo_home):
    """Run LibreOffice headless to convert xlsx → PDF."""
    try:
        subprocess.run(
            ["soffice", "--headless", "--norestore", "--calc",
             "--convert-to", "pdf", "--outdir", out_dir, xlsx_path],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, "HOME": lo_home},
        )
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — COMMISSION STATEMENT GENERATOR  (Invoice List format)
# ══════════════════════════════════════════════════════════════════════════════

# ── Style constants (Step 3 only) ─────────────────────────────────────────────
FONT_DATE     = Font(name="Arial", size=10)
FONT_CONTACT  = Font(name="Arial", size=10)
FONT_DIST_LBL = Font(name="Arial", size=12)
FONT_COMM_LBL = Font(name="Arial", size=10)
FONT_HDR      = Font(name="Arial", size=9, bold=True)
FONT_CODE     = Font(name="Arial", size=9, bold=True)
FONT_DATA     = Font(name="Arial", size=9)
FONT_TOT_LBL  = Font(name="Arial", size=9, bold=True)
FONT_TOT_NUM  = Font(name="Arial", size=9, bold=True)
FONT_FOOTER   = Font(name="Arial", size=12, bold=True, italic=True)

FONT_SUM_TITLE = Font(name="Arial", size=12, bold=True)
FONT_SUM_HDR   = Font(name="Calibri", size=11)
FONT_SUM_HDR_A = Font(name="Calibri", size=12, bold=True)
FONT_SUM_DATA  = Font(name="Arial", size=10)
FONT_SUM_BOLD  = Font(name="Arial", size=10, bold=True)

_W = True
ALIGN_HDR_C  = Alignment(horizontal="center", wrap_text=_W)
ALIGN_HDR_L  = Alignment(horizontal="left",   wrap_text=_W)
ALIGN_HDR_R  = Alignment(horizontal="right",  wrap_text=_W)
ALIGN_DATA_L = Alignment(horizontal="left",   wrap_text=_W)
ALIGN_DATA_C = Alignment(horizontal="center", wrap_text=_W)
ALIGN_DATA_R = Alignment(horizontal="right",  wrap_text=_W)
ALIGN_CODE_L = Alignment(horizontal="left",   wrap_text=_W)
ALIGN_RIGHT  = Alignment(horizontal="right")
ALIGN_CENTER = Alignment(horizontal="center")
ALIGN_CC     = Alignment(horizontal="center", vertical="center", wrap_text=_W)
ALIGN_RC     = Alignment(horizontal="right",  vertical="center")

_THIN = Side(style="thin")
_MED  = Side(style="medium")
BORDER_TOT_I   = Border(top=_THIN)
BORDER_TOT_J   = Border(top=_MED, bottom=_MED, left=_MED, right=_MED)
BORDER_THIN_B  = Border(bottom=_THIN)
BORDER_THIN_TB = Border(top=_THIN, bottom=_THIN)
BORDER_THIN_T  = Border(top=_THIN)

COL_WIDTHS  = {"A": 18.0, "B": 10.33, "C": 12.11, "D": 14.22, "E": 13.89,
               "F": 41.33, "G": 16.33, "H": 10.0, "I": 15.89, "J": 16.78}
ROW_HEIGHTS = {1: 42.6, 2: 41.4, 3: 34.2, 4: 27.6, 5: 51.0}
DATA_ROW_H  = 16.05
TOTAL_ROW_H = 18.6
SUM_COL_WIDTHS = {"A": 26.0, "B": 17.89, "C": 19.44, "D": 13.11,
                  "E": 13.0, "F": 15.89, "G": 16.66}


def detect_month_year(ws):
    """Auto-detect month/year from the first few rows of the Invoice List."""
    MONTHS = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    for r in range(1, 6):
        for c in range(1, 10):
            val = ws.cell(row=r, column=c).value
            if not val or not isinstance(val, str):
                continue
            m = re.match(r"(" + "|".join(MONTHS) + r")\s+(\d{4})", val)
            if m:
                name = m.group(1)
                return name, int(m.group(2)), MONTHS.index(name) + 1
    return None, None, None


def compute_pay_date(year, month_num):
    """Payment date = last day of the month following the sales month."""
    import calendar
    pay_m = month_num + 1 if month_num < 12 else 1
    pay_y = year if month_num < 12 else year + 1
    return datetime(pay_y, pay_m, calendar.monthrange(pay_y, pay_m)[1])


def load_lookup(wb):
    """Build distributor lookup from the 'Dist Lookup' sheet."""
    ws = wb["Dist Lookup"]
    lookup = {}
    for row in range(3, ws.max_row + 1):
        code = ws.cell(row=row, column=1).value
        if code:
            lookup[str(code).strip()] = {
                "name":    str(ws.cell(row=row, column=2).value or "").strip(),
                "contact": str(ws.cell(row=row, column=3).value or "").strip(),
            }
    return lookup


def parse_groups(ws):
    """Parse the Invoice List sheet into distributor groups."""
    groups = []
    current_code = current_name = None
    current_data = []
    for row_num in range(6, ws.max_row + 1):
        a = ws.cell(row=row_num, column=1).value
        b = ws.cell(row=row_num, column=2).value
        if a and str(a).strip().startswith("Total for"):
            groups.append({
                "code": current_code, "name": current_name, "data": current_data,
                "total_amount":     ws.cell(row=row_num, column=9).value or 0,
                "total_commission": ws.cell(row=row_num, column=10).value or 0,
            })
            current_code = current_name = None
            current_data = []
        elif a and b is None and not str(a).strip().startswith("Total"):
            current_code = str(a).strip()
        elif b is not None:
            row_data = {col: ws.cell(row=row_num, column=col).value for col in range(1, 12)}
            if a and current_name is None:
                current_name = str(a).strip()
            current_data.append(row_data)
    return groups


def create_tab(wb, tab_name, code, dist_name, contact, data_rows,
               total_amount, total_commission, pay_date, commission_label, logo_path):
    """Create a single distributor commission statement tab (Step 3)."""
    ws = wb.create_sheet(title=tab_name[:31])
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    for rn, h in ROW_HEIGHTS.items():
        ws.row_dimensions[rn].height = h

    # Logo
    if os.path.exists(logo_path):
        logo = XlImage(logo_path)
        logo.width, logo.height, logo.anchor = 271, 125, "B1"
        ws.add_image(logo)

    # Header cells
    c = ws.cell(row=1, column=10, value=pay_date)
    c.font = FONT_DATE; c.number_format = "m/d/yy;@"; c.alignment = ALIGN_RIGHT

    if contact:
        ws.cell(row=3, column=2, value=contact).font = FONT_CONTACT
    c = ws.cell(row=3, column=10, value=commission_label)
    c.font = FONT_COMM_LBL; c.alignment = ALIGN_RIGHT

    ws.cell(row=4, column=2, value=f"Distributor:  {dist_name}").font = FONT_DIST_LBL

    # Column headers
    for col, text, align in [
        (2, "Invoice Date", ALIGN_HDR_C), (3, "Invoice Number", ALIGN_HDR_C),
        (4, "P.O. Number", ALIGN_HDR_L), (5, "Surgeon", ALIGN_HDR_L),
        (6, "Name", ALIGN_HDR_L), (7, "Memo/ Description", ALIGN_HDR_L),
        (8, "Rate", ALIGN_HDR_C), (9, "Invoice Amount", ALIGN_HDR_C),
        (10, "Commission", ALIGN_HDR_R),
    ]:
        c = ws.cell(row=5, column=col, value=text)
        c.font = FONT_HDR; c.alignment = align
    ws.cell(row=5, column=8).number_format = "0%"

    # Distributor code row
    ws.row_dimensions[6].height = DATA_ROW_H
    ws.cell(row=6, column=1, value=code).font = FONT_CODE
    ws.cell(row=6, column=1).alignment = ALIGN_CODE_L

    # Data rows
    row = 7
    for i, d in enumerate(data_rows):
        ws.row_dimensions[row].height = DATA_ROW_H
        if i == 0 and d.get(1):
            ws.cell(row=row, column=1, value=d[1]).font = FONT_CODE
        for col in range(2, 11):
            val = d.get(col)
            if val is not None:
                c = ws.cell(row=row, column=col, value=val)
                c.font = FONT_DATA
                if col <= 7:
                    c.alignment = ALIGN_DATA_L
                elif col == 8:
                    c.alignment = ALIGN_DATA_C; c.number_format = "0%"
                else:
                    c.alignment = ALIGN_DATA_R; c.number_format = '#,##0.00\\ _€'
        row += 1

    # Total row
    ws.row_dimensions[row].height = TOTAL_ROW_H
    ws.cell(row=row, column=1, value=f"Total for {code}").font = FONT_TOT_LBL
    ws.cell(row=row, column=1).alignment = ALIGN_CODE_L
    for col, val, bdr in [(9, total_amount, BORDER_TOT_I), (10, total_commission, BORDER_TOT_J)]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = FONT_TOT_NUM; c.alignment = ALIGN_DATA_R
        c.number_format = '"$"* #,##0.00\\ _€'; c.border = bdr

    # Footer
    footer_row = row + 2
    ws.merge_cells(start_row=footer_row, start_column=2, end_row=footer_row, end_column=10)
    c = ws.cell(row=footer_row, column=2, value="Thank you for your continued support.")
    c.font = FONT_FOOTER; c.alignment = ALIGN_CENTER

    # Page setup
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"B1:J{footer_row}"
    return ws


def create_summary_step3(wb, groups, lookup, commission_label):
    """Create the Summary sheet for Step 3 (Commission Statement Generator)."""
    ws = wb.create_sheet(title="Summary", index=0)
    for col, w in SUM_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A1:G1"); ws.row_dimensions[1].height = 15.75
    ws.cell(row=1, column=1, value=2026).font = FONT_SUM_TITLE
    ws.cell(row=1, column=1).alignment = ALIGN_CENTER

    ws.merge_cells("A2:G2"); ws.row_dimensions[2].height = 15.75
    c = ws.cell(row=2, column=1, value=commission_label)
    c.font = FONT_SUM_TITLE; c.alignment = ALIGN_CENTER
    for col in range(1, 8):
        ws.cell(row=2, column=col).border = BORDER_THIN_B

    ws.row_dimensions[3].height = 35.4
    for col, text, font, align in [
        (1, "Distributor",                       FONT_SUM_HDR_A, None),
        (2, "Commission Earned",                 FONT_SUM_HDR,   ALIGN_CC),
        (3, "Chargeback to Maxx Orthopedics",    FONT_SUM_HDR,   ALIGN_CC),
        (4, "Expense report payments",           FONT_SUM_HDR,   ALIGN_CC),
        (5, "Other payments",                    FONT_SUM_HDR,   ALIGN_CC),
        (6, "Freight charges\ndeduction",        FONT_SUM_HDR,   ALIGN_CC),
        (7, "Total Commission\nPaid",            FONT_SUM_HDR,   ALIGN_CC),
    ]:
        c = ws.cell(row=3, column=col, value=text)
        c.font = font
        if align:
            c.alignment = align
        c.border = BORDER_THIN_TB

    sorted_groups = sorted(
        groups,
        key=lambda g: (lookup.get(g["code"], {}).get("name", g["name"] or "")).lower(),
    )
    row = 4
    total_all = 0
    for g in sorted_groups:
        ws.row_dimensions[row].height = 13.05
        info = lookup.get(g["code"], {})
        ws.cell(row=row, column=1, value=info.get("name", g["name"] or "")).font = FONT_SUM_DATA
        c = ws.cell(row=row, column=2, value=g["total_commission"])
        c.font = FONT_SUM_DATA; c.number_format = '"$"#,##0.00'
        for cc in range(3, 7):
            ws.cell(row=row, column=cc).font = FONT_SUM_DATA
            ws.cell(row=row, column=cc).number_format = '"$"#,##0.00'
        c = ws.cell(row=row, column=7, value=g["total_commission"])
        c.font = FONT_SUM_DATA; c.number_format = '"$"#,##0.00'; c.alignment = ALIGN_RC
        total_all += g["total_commission"]
        row += 1

    # Totals section
    tr = row
    ws.row_dimensions[tr].height = 13.05
    ws.cell(row=tr, column=1, value="Total Distributor Commission:").font = FONT_SUM_DATA
    for col in range(1, 8):
        ws.cell(row=tr, column=col).border = BORDER_THIN_T
    for col in [2, 7]:
        c = ws.cell(row=tr, column=col, value=total_all)
        c.font = FONT_SUM_DATA; c.number_format = '"$"#,##0.00'

    r = tr + 3
    ws.cell(row=r, column=1, value="Total Commission").font = FONT_SUM_BOLD
    ws.cell(row=r, column=2, value=total_all).font = FONT_SUM_DATA
    ws.cell(row=r, column=2).number_format = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'
    ws.cell(row=r, column=6, value="Total Payments").font = FONT_SUM_DATA
    ws.cell(row=r, column=7, value=total_all).font = FONT_SUM_DATA
    ws.cell(row=r, column=7).number_format = '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)'
    r += 2
    ws.cell(row=r, column=6, value="Total ACH").font = FONT_SUM_DATA
    ws.cell(row=r, column=7, value=total_all).font = FONT_SUM_DATA
    ws.cell(row=r, column=7).number_format = '"$"#,##0.00'
    r += 1
    ws.cell(row=r, column=6, value="Total Checks").font = FONT_SUM_DATA
    r += 1
    ws.cell(row=r, column=6, value="Total Payment").font = FONT_SUM_BOLD
    ws.cell(row=r, column=7, value=total_all).font = FONT_SUM_BOLD
    ws.cell(row=r, column=7).number_format = '"$"#,##0.00'
    return ws


def process_excel(input_path, job_dir):
    """Step 3: Build commission statement workbook from Invoice List format."""
    wb_src = openpyxl.load_workbook(input_path, data_only=True)

    required = {"Invoice List", "Dist Lookup"}
    missing  = required - set(wb_src.sheetnames)
    if missing:
        wb_src.close()
        raise ValueError(f"Missing required sheets: {', '.join(missing)}")

    month_name, year, month_num = detect_month_year(wb_src["Invoice List"])
    if not month_name:
        wb_src.close()
        raise ValueError("Could not detect month/year from Invoice List.")

    pay_date         = compute_pay_date(year, month_num)
    commission_label = f"Commission on {month_name} {year} Sales"
    lookup           = load_lookup(wb_src)
    groups           = parse_groups(wb_src["Invoice List"])

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    create_summary_step3(wb_out, groups, lookup, commission_label)

    # Copy source sheets
    for src_name in wb_src.sheetnames:
        _copy_sheet_data(wb_src[src_name], wb_out.create_sheet(title=src_name))
    wb_src.close()

    # Create distributor tabs (sorted by name)
    for g in sorted(groups, key=lambda x: (lookup.get(x["code"], {}).get("name", x["name"] or "")).lower()):
        code = g["code"]
        info = lookup.get(code, {})
        tab  = re.sub(r'[\\/*?\[\]:]', '', g["name"] or code)[:31]
        if tab in [s.title for s in wb_out.worksheets]:
            tab = f"{tab[:27]} {code}"[:31]
        create_tab(wb_out, tab, code, info.get("name", g["name"] or ""),
                   info.get("contact", ""), g["data"], g["total_amount"],
                   g["total_commission"], pay_date, commission_label, LOGO_PATH)

    xlsx_name = f"Commission_Statements_{month_name}_{year}.xlsx"
    wb_out.save(os.path.join(job_dir, xlsx_name))

    return {"xlsx_name": xlsx_name, "month": month_name, "year": year,
            "num_distributors": len(groups)}


def generate_pdfs(job_dir):
    """Step 3b: Convert each tab to PDF and bundle into a zip."""
    xlsx_path = next(
        (os.path.join(job_dir, f) for f in os.listdir(job_dir) if f.endswith(".xlsx")),
        None,
    )
    if not xlsx_path:
        raise ValueError("No Excel workbook found for this job.")

    xlsx_name = os.path.basename(xlsx_path)
    zip_name  = xlsx_name.replace(".xlsx", "_PDFs.zip")
    zip_path  = os.path.join(job_dir, zip_name)

    skip     = {"Invoice List", "Trauma", "Dist Lookup"}
    temp_dir = os.path.join(job_dir, "temp_sheets")
    pdf_dir  = os.path.join(job_dir, "pdfs")
    lo_home  = os.path.join(job_dir, "lo_home")
    os.makedirs(temp_dir, exist_ok=True)
    os.makedirs(pdf_dir,  exist_ok=True)
    os.makedirs(lo_home,  exist_ok=True)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    for name in wb.sheetnames:
        if name in skip:
            continue
        src_ws = wb[name]
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = name[:31]
        _copy_sheet_data(src_ws, new_ws)

        # Re-add logo images
        for img in src_ws._images:
            if os.path.exists(LOGO_PATH):
                ni = XlImage(LOGO_PATH)
                ni.width, ni.height, ni.anchor = img.width, img.height, "B1"
                new_ws.add_image(ni)

        if src_ws.print_area:
            new_ws.print_area = src_ws.print_area
        new_ws.page_setup.orientation = "landscape"
        new_ws.page_setup.fitToWidth = 1; new_ws.page_setup.fitToHeight = 0
        new_ws.sheet_properties.pageSetUpPr.fitToPage = True

        safe = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
        new_wb.save(os.path.join(temp_dir, f"{safe}.xlsx"))
        new_wb.close()
    wb.close()

    # Convert all temp xlsx to PDF
    for fname in sorted(os.listdir(temp_dir)):
        if fname.endswith(".xlsx"):
            _convert_to_pdf_libreoffice(os.path.join(temp_dir, fname), pdf_dir, lo_home)

    # Bundle into zip
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for pdf in sorted(os.listdir(pdf_dir)):
            if pdf.endswith(".pdf"):
                zf.write(os.path.join(pdf_dir, pdf), pdf)

    num_pdfs = sum(1 for f in os.listdir(pdf_dir) if f.endswith(".pdf"))
    shutil.rmtree(temp_dir, ignore_errors=True)
    shutil.rmtree(lo_home,  ignore_errors=True)

    return {"zip_name": zip_name, "num_pdfs": num_pdfs}


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — MANAGER SPLIT
# ══════════════════════════════════════════════════════════════════════════════

SPLIT_KEEP_SHEETS = {"masterlog"}


def process_manager_split(input_path, job_dir):
    """Split masterlog by manager into individual xlsx + PDF workbooks."""
    wb_src = openpyxl.load_workbook(input_path)

    src_sheet = get_sheet_ci(wb_src, "masterlog")
    if src_sheet is None:
        raise ValueError("No 'masterlog' sheet found in the workbook.")

    header_row, label_map = find_summary_header(src_sheet)
    if header_row is None:
        raise ValueError("Could not find header row with Manager, Hospital, "
                         "Comm $, and Distributor/Distrib Code.")

    mgr_idx  = label_map.get("manager")
    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    comm_idx = label_map.get("comm $")
    po_idx   = label_map.get("po")

    if any(x is None for x in [mgr_idx, hosp_idx, dist_idx, comm_idx]):
        raise ValueError("Missing required columns in masterlog header.")

    data_rows  = collect_data_rows(src_sheet, header_row, hosp_idx, po_idx)
    col_widths = {col: dim.width for col, dim in src_sheet.column_dimensions.items()}
    base_name  = os.path.splitext(os.path.basename(input_path))[0]

    # Detect managers (order of first appearance)
    managers, seen = [], set()
    for cells in data_rows:
        mgr = cells[mgr_idx].value if mgr_idx < len(cells) else None
        if mgr and str(mgr).strip() not in seen:
            seen.add(str(mgr).strip())
            managers.append(str(mgr).strip())
    if not managers:
        raise ValueError("No manager values found in data rows.")

    results = []
    lo_home = os.path.join(job_dir, ".lo_home")

    for manager in managers:
        out_wb = openpyxl.load_workbook(input_path)
        keep_sheets_ci(out_wb, SPLIT_KEEP_SHEETS)
        out_ws = get_sheet_ci(out_wb, "masterlog")

        # Clear data area
        for r in range(header_row + 1, out_ws.max_row + 1):
            for c in range(1, out_ws.max_column + 1):
                out_ws.cell(row=r, column=c).value = None

        # Write this manager's rows
        mgr_rows = [c for c in data_rows if str(c[mgr_idx].value).strip() == manager]
        wr = header_row + 1
        for cells in mgr_rows:
            for ci, src in enumerate(cells):
                copy_cell(src, out_ws.cell(row=wr, column=ci + 1))
            wr += 1

        for letter, w in col_widths.items():
            out_ws.column_dimensions[letter].width = w
        insert_distributor_subtotals(out_ws, header_row, header_row + 1, dist_idx, comm_idx)
        apply_summary_alignment(out_ws, label_map, header_row)

        safe      = re.sub(r'[<>:"/\\|?*]', "_", manager)
        xlsx_name = f"{safe}-{base_name}.xlsx"
        xlsx_path = os.path.join(job_dir, xlsx_name)
        out_wb.save(xlsx_path)

        # PDF conversion
        os.makedirs(lo_home, exist_ok=True)
        _convert_to_pdf_libreoffice(xlsx_path, job_dir, lo_home)
        pdf_path = os.path.splitext(xlsx_path)[0] + ".pdf"
        pdf_name = os.path.basename(pdf_path) if os.path.exists(pdf_path) else None

        results.append({
            "manager": manager, "xlsx_name": xlsx_name,
            "pdf_name": pdf_name, "num_rows": len(mgr_rows),
        })

    shutil.rmtree(lo_home, ignore_errors=True)

    # Bundle into zip
    zip_name = f"{base_name}_by_manager.zip"
    with zipfile.ZipFile(os.path.join(job_dir, zip_name), "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            zf.write(os.path.join(job_dir, r["xlsx_name"]), r["xlsx_name"])
            if r["pdf_name"]:
                zf.write(os.path.join(job_dir, r["pdf_name"]), r["pdf_name"])

    return {"managers": results, "zip_name": zip_name, "num_managers": len(managers)}


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — DISTRIBUTOR TAB GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

def process_distributor_tabs(input_path, job_dir):
    """
    Generate one tab per Distrib Code + a Summary sheet with totals.
    Returns dict with xlsx_name, num_tabs, tab_names.
    """
    wb_src = openpyxl.load_workbook(input_path)

    src_sheet = get_sheet_ci(wb_src, "masterlog")
    if src_sheet is None:
        raise ValueError("No 'masterlog' sheet found in the workbook.")

    header_row, label_map = find_summary_header(src_sheet)
    if header_row is None:
        raise ValueError("Could not find header row in masterlog.")

    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    comm_idx = label_map.get("comm $")
    po_idx   = label_map.get("po")

    if hosp_idx is None or dist_idx is None:
        raise ValueError("Missing required columns (Hospital, Distrib Code) in masterlog.")

    title_val, pay_date_val = scan_summary_meta(src_sheet, header_row)
    data_rows      = collect_data_rows(src_sheet, header_row, hosp_idx, po_idx)
    surgeon_lookup = build_surgeon_lookup(wb_src)

    out_wb = openpyxl.load_workbook(input_path)
    keep_sheets_ci(out_wb, {"masterlog", "Surgeon lookup", "template"})

    num_tabs = generate_distributor_tabs(
        out_wb, data_rows, dist_idx, label_map,
        title_val, pay_date_val, surgeon_lookup,
    )

    # Build and insert Summary tab at position 0
    group_summaries = _build_group_summaries(data_rows, dist_idx, comm_idx, surgeon_lookup)
    create_summary_tab(out_wb, group_summaries, title_val, pay_date_val)

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    xlsx_name = f"{base_name}_distributor_tabs.xlsx"
    out_wb.save(os.path.join(job_dir, xlsx_name))
    _cleanup_tmp_images(out_wb)

    # Collect distributor tab names (everything except infrastructure sheets)
    infra = {"masterlog", "summary", "surgeon lookup", "template"}
    tab_names = [s for s in out_wb.sheetnames if s.lower() not in infra]

    return {"xlsx_name": xlsx_name, "num_tabs": num_tabs, "tab_names": tab_names}


# ══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


# ── Step 1 routes ─────────────────────────────────────────────────────────────

@app.route("/split-upload", methods=["POST"])
def split_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    job_id, job_dir = _make_job_dir()
    input_path = os.path.join(job_dir, secure_filename(f.filename))
    f.save(input_path)
    try:
        result = process_manager_split(input_path, job_dir)
        return jsonify({"success": True, "job_id": job_id, **result})
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/download-split/<job_id>/zip")
def download_split_zip(job_id):
    job_dir = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
    if not os.path.exists(job_dir):
        return "Job not found", 404
    for f in os.listdir(job_dir):
        if f.endswith(".zip"):
            return send_file(os.path.join(job_dir, f), as_attachment=True, download_name=f)
    return "File not found", 404


@app.route("/download-split/<job_id>/file/<filename>")
def download_split_file(job_id, filename):
    job_dir = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
    safe    = os.path.basename(filename)
    path    = os.path.join(job_dir, safe)
    if not os.path.exists(path):
        return "File not found", 404
    return send_file(path, as_attachment=True, download_name=safe)


# ── Step 2 routes ─────────────────────────────────────────────────────────────

@app.route("/dist-tabs-upload", methods=["POST"])
def dist_tabs_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    job_id, job_dir = _make_job_dir()
    input_path = os.path.join(job_dir, secure_filename(f.filename))
    f.save(input_path)
    try:
        result = process_distributor_tabs(input_path, job_dir)
        return jsonify({"success": True, "job_id": job_id, **result})
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/download-dist-tabs/<job_id>/<filetype>")
def download_dist_tabs(job_id, filetype):
    job_dir = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
    if not os.path.exists(job_dir):
        return "Job not found", 404
    for f in os.listdir(job_dir):
        # xlsx: return only the processed output, never the uploaded input
        if filetype == "xlsx" and f.endswith("_distributor_tabs.xlsx"):
            return send_file(os.path.join(job_dir, f), as_attachment=True, download_name=f)
        if filetype == "zip" and f.endswith(".zip"):
            return send_file(os.path.join(job_dir, f), as_attachment=True, download_name=f)
    return "File not found", 404


# ── Step 3 routes ─────────────────────────────────────────────────────────────

def generate_distributor_tab_pdfs(job_dir):
    """Convert each distributor tab to an individual PDF and bundle into a zip."""
    xlsx_path = next(
        (os.path.join(job_dir, f) for f in os.listdir(job_dir) if f.endswith(".xlsx")),
        None,
    )
    if not xlsx_path:
        raise ValueError("No Excel workbook found for this job.")

    skip     = {"masterlog", "summary", "surgeon lookup", "template"}
    temp_dir = os.path.join(job_dir, "temp_sheets")
    pdf_dir  = os.path.join(job_dir, "pdfs")
    lo_home  = os.path.join(job_dir, "lo_home")
    os.makedirs(temp_dir, exist_ok=True)
    os.makedirs(pdf_dir,  exist_ok=True)
    os.makedirs(lo_home,  exist_ok=True)

    # Load without data_only so images are accessible
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    wb_imgs = openpyxl.load_workbook(xlsx_path)  # second load to access embedded images

    tmp_img_files = []  # track temp files for cleanup

    for name in wb.sheetnames:
        if name.lower() in skip:
            continue
        src_ws    = wb[name]
        src_ws_i  = wb_imgs[name]   # image-accessible version of the same sheet
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = name[:31]
        _copy_sheet_data(src_ws, new_ws)

        # Copy images (not carried by data_only load) using temp files
        for img in src_ws_i._images:
            try:
                raw    = img._data()
                suffix = ".png" if raw[:4] == b"\x89PNG" else ".jpg"
                tmp    = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
                tmp.write(raw)
                tmp.close()
                tmp_img_files.append(tmp.name)
                ni        = XlImage(tmp.name)
                ni.anchor = img.anchor
                new_ws.add_image(ni)
            except Exception as e:
                print(f"  Warning: could not copy image for '{name}': {e}")

        new_ws.page_setup.orientation = "landscape"
        new_ws.page_setup.fitToWidth  = 1
        new_ws.page_setup.fitToHeight = 0
        new_ws.sheet_properties.pageSetUpPr.fitToPage = True

        safe = re.sub(r'[<>:"/\\|?*]', "_", name).strip()
        new_wb.save(os.path.join(temp_dir, f"{safe}.xlsx"))
        new_wb.close()

    wb.close()
    wb_imgs.close()

    for p in tmp_img_files:
        try:
            os.unlink(p)
        except OSError:
            pass

    for fname in sorted(os.listdir(temp_dir)):
        if fname.endswith(".xlsx"):
            _convert_to_pdf_libreoffice(os.path.join(temp_dir, fname), pdf_dir, lo_home)

    zip_name = os.path.basename(xlsx_path).replace(".xlsx", "_PDFs.zip")
    zip_path = os.path.join(job_dir, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for pdf in sorted(os.listdir(pdf_dir)):
            if pdf.endswith(".pdf"):
                zf.write(os.path.join(pdf_dir, pdf), pdf)

    num_pdfs = sum(1 for f in os.listdir(pdf_dir) if f.endswith(".pdf"))
    shutil.rmtree(temp_dir, ignore_errors=True)
    shutil.rmtree(lo_home,  ignore_errors=True)

    return {"zip_name": zip_name, "num_pdfs": num_pdfs}


@app.route("/upload", methods=["POST"])
def upload():
    """Step 3: accept a distributor-tabs xlsx and immediately generate PDFs."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    job_id, job_dir = _make_job_dir()
    input_path = os.path.join(job_dir, secure_filename(f.filename))
    f.save(input_path)
    try:
        result = generate_distributor_tab_pdfs(job_dir)
        return jsonify({"success": True, "job_id": job_id, **result})
    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/download/<job_id>/zip")
def download(job_id):
    job_dir = os.path.join(app.config["OUTPUT_FOLDER"], job_id)
    if not os.path.exists(job_dir):
        return "File not found", 404
    for f in os.listdir(job_dir):
        if f.endswith(".zip"):
            return send_file(os.path.join(job_dir, f), as_attachment=True, download_name=f)
    return "File not found", 404


# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5001)), debug=True)
