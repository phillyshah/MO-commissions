"""
process_commissions.py — MO Commission Report Processor

Provides shared helpers for:
  - Step 1: Splitting a masterlog by manager into per-manager workbooks
  - Step 2: Generating one distributor tab per unique Distrib Code,
            plus a Summary sheet with per-distributor totals

Usage (CLI):
    python process_commissions.py <input.xlsx>
"""

import sys
import os
import copy
import tempfile
import subprocess
import shutil
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage
from openpyxl.cell.cell import MergedCell


# ══════════════════════════════════════════════════════════════════════════════
# WORKBOOK / SHEET HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def get_sheet_ci(wb, name):
    """Case-insensitive sheet lookup. Returns the worksheet or None."""
    target = name.lower()
    for sname in wb.sheetnames:
        if sname.lower() == target:
            return wb[sname]
    return None


def keep_sheets_ci(wb, keep_names):
    """Delete every sheet whose name is NOT in *keep_names* (case-insensitive)."""
    keep = {n.lower() for n in keep_names}
    for sname in list(wb.sheetnames):
        if sname.lower() not in keep:
            del wb[sname]


def copy_cell(src, dst):
    """Copy value + full style from *src* cell to *dst* cell."""
    dst.value = src.value
    if src.has_style:
        dst.font          = copy.copy(src.font)
        dst.fill          = copy.copy(src.fill)
        dst.border        = copy.copy(src.border)
        dst.alignment     = copy.copy(src.alignment)
        dst.number_format = src.number_format


# ── Reusable style constants ─────────────────────────────────────────────────

NO_FILL      = PatternFill(fill_type=None)
CLEAR_BORDER = Border()
THIN_SIDE    = Side(style="thin")

# Masterlog alignment rules (column-label → horizontal alignment)
LEFT_COLS  = {"po", "notes", "surgeon", "hospital", "manager"}
RIGHT_COLS = {"surgery date", "inv#", "inv date", "due date", "comm $"}

# Masterlog → template column-label aliases
SUMMARY_TO_TEMPLATE = {"hospital": "hospital or mc"}

# Masterlog columns that have no template equivalent
SUMMARY_SKIP = {"manager", "distrib code", "distributor", "status", "date pd"}

# Sheets kept in each per-manager workbook (Step 1)
KEEP_SHEETS = {"masterlog"}


# ══════════════════════════════════════════════════════════════════════════════
# MASTERLOG HEADER & META DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def find_summary_header(sheet):
    """
    Locate the masterlog header row by scanning for required column labels:
    Manager, Hospital, Comm $, and either Distributor or Distrib Code.

    Returns (1-based row number, {normalised_label: 0-based col index}),
    or (None, None) if not found.
    """
    REQUIRED = {"manager", "hospital", "comm $"}
    DIST_ALT = {"distributor", "distrib code"}

    for row in sheet.iter_rows():
        vals = [str(c.value).strip().lower() if c.value is not None else "" for c in row]
        s = set(vals)
        if REQUIRED.issubset(s) and DIST_ALT.intersection(s):
            return row[0].row, {v: i for i, v in enumerate(vals)}
    return None, None


def scan_summary_meta(sheet, header_row_num):
    """
    Read Title: and Pay Date: values from rows above the header.
    Returns (title_str | None, pay_date_value | None).
    """
    title = pay_date = None
    for row in sheet.iter_rows(min_row=1, max_row=header_row_num - 1):
        cells = list(row)
        for i, cell in enumerate(cells):
            if cell.value is None:
                continue
            label = str(cell.value).strip().lower()
            if label in ("title:", "pay date:"):
                # value is in the next non-empty cell to the right
                for j in range(i + 1, len(cells)):
                    if cells[j].value is not None:
                        if label == "title:":
                            title = str(cells[j].value).strip()
                        else:
                            pay_date = cells[j].value
                        break
    return title, pay_date


def is_legacy_subtotal(row_cells, po_idx):
    """Return True if this row is a legacy subtotal (bold 'Total…' in the PO cell)."""
    cell = row_cells[po_idx]
    if cell.value is None:
        return False
    text = str(cell.value).strip()
    if not text.lower().startswith("total"):
        return False
    cleaned = (text.replace(",", "").replace(".", "")
               .replace("Total", "").replace("total", "")
               .strip().lstrip("-"))
    return not cleaned.isnumeric() and cell.font and cell.font.bold


def collect_data_rows(sheet, header_row_num, hosp_idx, po_idx):
    """
    Return data rows below the header, skipping blanks and legacy subtotals.
    Each element is a list of openpyxl Cell objects for that row.
    """
    rows = []
    for row in list(sheet.iter_rows())[header_row_num:]:
        cells = list(row)
        hosp = cells[hosp_idx].value if hosp_idx < len(cells) else None
        if not hosp or not str(hosp).strip():
            continue
        if po_idx is not None and is_legacy_subtotal(cells, po_idx):
            continue
        rows.append(cells)
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# ALIGNMENT & SUBTOTALS (used by Step 1 manager-split)
# ══════════════════════════════════════════════════════════════════════════════

def apply_summary_alignment(ws, label_map, header_row_num):
    """Apply left/right alignment to every data cell based on column label."""
    col_align = {}
    for label, idx in label_map.items():
        if label in LEFT_COLS:
            col_align[idx + 1] = "left"
        elif label in RIGHT_COLS:
            col_align[idx + 1] = "right"

    for row in ws.iter_rows(min_row=header_row_num):
        for cell in row:
            if cell.column in col_align:
                prev = cell.alignment or Alignment()
                cell.alignment = Alignment(
                    horizontal=col_align[cell.column],
                    vertical=prev.vertical,
                    wrap_text=prev.wrap_text,
                )


def insert_distributor_subtotals(ws, header_row_num, data_start_row,
                                  dist_col_0idx, comm_col_0idx):
    """
    Group consecutive rows by distributor code and insert SUM subtotal rows.
    Returns the number of distinct groups written.
    """
    # Read existing data
    data = []
    for row in ws.iter_rows(min_row=data_start_row):
        data.append([cell.value for cell in row])

    # Build consecutive blocks
    blocks = []
    for row_data in data:
        key = str(row_data[dist_col_0idx]).strip() if dist_col_0idx < len(row_data) and row_data[dist_col_0idx] else ""
        if blocks and blocks[-1][0] == key:
            blocks[-1][1].append(row_data)
        else:
            blocks.append((key, [row_data]))

    # Clear and rewrite with subtotals
    for row in ws.iter_rows(min_row=data_start_row):
        for cell in row:
            cell.value = None

    cur = data_start_row
    comm_letter = get_column_letter(comm_col_0idx + 1)

    for code, rows in blocks:
        start = cur
        for rd in rows:
            for ci, val in enumerate(rd):
                ws.cell(row=cur, column=ci + 1).value = val
            cur += 1
        end = cur - 1
        cur += 1  # blank row

        ws.cell(row=cur, column=dist_col_0idx + 1, value=f"Total {code}").font = Font(bold=True)
        tc = ws.cell(row=cur, column=comm_col_0idx + 1)
        tc.value = f"=SUM({comm_letter}{start}:{comm_letter}{end})"
        tc.font = Font(bold=True)
        tc.number_format = "#,##0.00"
        cur += 2  # subtotal + trailing blank

    return len(blocks)


# ══════════════════════════════════════════════════════════════════════════════
# SURGEON LOOKUP
# ══════════════════════════════════════════════════════════════════════════════

def build_surgeon_lookup(wb):
    """
    Build {distrib_code: {"name": …, "contact": …}} from the Surgeon lookup sheet.
    First matching row per code wins.  Returns {} if the sheet is missing.
    """
    ws = get_sheet_ci(wb, "Surgeon lookup")
    if ws is None:
        return {}

    REQUIRED = {"distrib code", "distributor", "contact"}
    hdr_row = None
    lmap = {}
    for row in ws.iter_rows():
        vals = [str(c.value).strip().lower() if c.value else "" for c in row]
        if REQUIRED.issubset(set(vals)):
            hdr_row = row[0].row
            lmap = {v: i for i, v in enumerate(vals)}
            break
    if hdr_row is None:
        return {}

    ci, ni, coi = lmap["distrib code"], lmap["distributor"], lmap["contact"]
    lookup = {}
    for row in ws.iter_rows(min_row=hdr_row + 1, max_row=ws.max_row):
        cells = list(row)
        code = cells[ci].value if ci < len(cells) else None
        if not code:
            continue
        code = str(code).strip()
        if code in lookup:
            continue
        name    = cells[ni].value  if ni  < len(cells) else None
        contact = cells[coi].value if coi < len(cells) else None
        lookup[code] = {
            "name":    str(name).strip()    if name    else "",
            "contact": str(contact).strip() if contact else "",
        }
    return lookup


# ══════════════════════════════════════════════════════════════════════════════
# TEMPLATE SHEET HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def find_template_header(ws):
    """
    Find the template header row (must contain 'Inv #' and 'Comm $').
    Returns (1-based row, {label: 0-based col}) or (None, None).
    """
    REQUIRED = {"inv #", "comm $"}
    for row in ws.iter_rows():
        vals = [str(c.value).strip().lower() if c.value else "" for c in row]
        if REQUIRED.issubset(set(vals)):
            return row[0].row, {v: i for i, v in enumerate(vals)}
    return None, None


def scan_template_placeholders(ws, header_row_num):
    """
    Scan rows above the header for placeholder cells.
    Returns {lower_key: (row, 1-based col)} for known keys.
    """
    KEYS = {"distributor", "distrib code", "contact", "title", "=summary!b2"}
    result = {}
    for row in ws.iter_rows(min_row=1, max_row=header_row_num - 1):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            key = str(cell.value).strip().lower()
            if key in KEYS:
                result[key] = (cell.row, cell.column)
    return result


def get_template_preformatted_rows(ws, header_row_num):
    """
    Find bordered-but-empty placeholder rows below the header.
    Returns list of 1-based row numbers; stops at first unbordered row.
    """
    rows = []
    for row in ws.iter_rows(min_row=header_row_num + 1, max_row=ws.max_row):
        has_border = any(
            not isinstance(c, MergedCell) and c.border and
            (c.border.left.style or c.border.right.style or
             c.border.top.style  or c.border.bottom.style)
            for c in row
        )
        if has_border:
            rows.append(row[0].row)
        else:
            break
    return rows


def capture_row_borders(ws, row_num):
    """Return {1-based col: Border} for every non-merged cell in *row_num*."""
    return {
        cell.column: copy.copy(cell.border) if cell.border else CLEAR_BORDER
        for cell in next(ws.iter_rows(min_row=row_num, max_row=row_num))
        if not isinstance(cell, MergedCell)
    }


# ══════════════════════════════════════════════════════════════════════════════
# TAB NAME & COLUMN MAPPING
# ══════════════════════════════════════════════════════════════════════════════

def make_tab_name(code, dist_name, existing_names):
    """
    Build a sheet tab name <= 31 chars: '{code} ({name_prefix})'.
    Deduplicates against *existing_names*.
    """
    suffix    = dist_name[:7] if dist_name else code
    candidate = f"{code} ({suffix})"
    if len(candidate) > 31:
        room = max(31 - len(code) - 3, 0)   # 3 = len(" ()")
        candidate = f"{code} ({dist_name[:room]})"
    candidate = candidate[:31]

    base, n = candidate, 2
    while candidate in existing_names:
        ext       = f"_{n}"
        candidate = base[:31 - len(ext)] + ext
        n += 1
    return candidate


def build_column_map(summary_label_map, template_label_map):
    """
    Map masterlog columns to template columns.
    Returns [(sum_0idx, tmpl_0idx), …] for columns present in both sheets,
    applying label aliases and skipping SUMMARY_SKIP columns.
    """
    mapping = []
    for label, s_idx in summary_label_map.items():
        if label in SUMMARY_SKIP:
            continue
        t_label = SUMMARY_TO_TEMPLATE.get(label, label)
        if t_label in template_label_map:
            mapping.append((s_idx, template_label_map[t_label]))
    return mapping


# ══════════════════════════════════════════════════════════════════════════════
# DISTRIBUTOR TAB POPULATION
# ══════════════════════════════════════════════════════════════════════════════

def populate_distributor_tab(ws, tmpl_header_row, tmpl_label_map, placeholders,
                              preformatted_rows, first_row_borders,
                              dist_code, dist_name, contact, title, pay_date,
                              row_cells_list, summary_label_map):
    """
    Fill a copied template sheet with one distributor's data.
    Returns (n_data_rows, total_row_num).
    """
    # ── 1. Clear pre-formatted placeholder rows ──
    for r in preformatted_rows:
        for col in range(1, ws.max_column + 1):
            c = ws.cell(row=r, column=col)
            if not isinstance(c, MergedCell):
                c.value  = None
                c.border = CLEAR_BORDER
                c.font   = Font()

    # ── 2. Clear stale =SUM formulas below the header ──
    for row in ws.iter_rows(min_row=tmpl_header_row + 1, max_row=ws.max_row):
        for c in row:
            if (not isinstance(c, MergedCell) and isinstance(c.value, str)
                    and c.value.upper().startswith("=SUM")):
                c.value = None

    # ── 3. Fill header placeholders ──
    for key, value in [("distributor", dist_name), ("distrib code", dist_code),
                       ("contact", contact), ("title", title)]:
        if key in placeholders:
            r, col = placeholders[key]
            cell = ws.cell(row=r, column=col)
            cell.value = value
            cell.fill  = NO_FILL

    if "=summary!b2" in placeholders:
        r, col = placeholders["=summary!b2"]
        cell = ws.cell(row=r, column=col)
        cell.value = pay_date
        cell.number_format = "mm/dd/yyyy"
        cell.fill = NO_FILL

    # ── 4. Write data rows ──
    col_map        = build_column_map(summary_label_map, tmpl_label_map)
    idx_to_label   = {v: k for k, v in tmpl_label_map.items()}
    comm_0idx      = tmpl_label_map.get("comm $")
    hosp_0idx      = tmpl_label_map.get("hospital or mc")

    write_row      = tmpl_header_row + 1
    data_start     = write_row
    max_hosp_len   = len("Hospital or MC")

    for src_cells in row_cells_list:
        for s_idx, t_idx in col_map:
            if s_idx >= len(src_cells):
                continue
            src = src_cells[s_idx]
            dst = ws.cell(row=write_row, column=t_idx + 1)
            dst.value = src.value

            if src.has_style and src.number_format:
                dst.number_format = src.number_format
            if t_idx + 1 in first_row_borders:
                dst.border = copy.copy(first_row_borders[t_idx + 1])

            label = idx_to_label.get(t_idx, "")
            dst.alignment = Alignment(
                horizontal="center" if label in {"comm %", "comm $"} else "left"
            )

            if t_idx == hosp_0idx:
                max_hosp_len = max(max_hosp_len, len(str(src.value or "")))
        write_row += 1

    data_end    = write_row - 1
    n_data_rows = max(data_end - data_start + 1, 0)

    # ── 5. Auto-size Hospital column ──
    if hosp_0idx is not None:
        ws.column_dimensions[get_column_letter(hosp_0idx + 1)].width = max_hosp_len + 2

    # ── 6. Total row ──
    # Compute the actual sum from source rows so the value is present when the
    # workbook is read with data_only=True (e.g. for PDF generation).
    comm_sum = 0.0
    if comm_0idx is not None:
        src_comm_0idx = summary_label_map.get("comm $")
        if src_comm_0idx is not None:
            for src_cells in row_cells_list:
                v = src_cells[src_comm_0idx].value if src_comm_0idx < len(src_cells) else None
                if isinstance(v, (int, float)):
                    comm_sum += v

    total_row = data_end + 2
    if comm_0idx is not None:
        if comm_0idx >= 1:
            lbl = ws.cell(row=total_row, column=comm_0idx)
            lbl.value = "Total Comm $"
            lbl.font  = Font(bold=True)
        tc = ws.cell(row=total_row, column=comm_0idx + 1)
        tc.value         = comm_sum   # numeric value — readable without Excel evaluating formulas
        tc.font          = Font(bold=True)
        tc.number_format = "#,##0.00"

    # ── 7. Page setup ──
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return n_data_rows, total_row


# ══════════════════════════════════════════════════════════════════════════════
# GENERATE DISTRIBUTOR TABS
# ══════════════════════════════════════════════════════════════════════════════

def generate_distributor_tabs(out_wb, data_rows, dist_col_0idx,
                               summary_label_map, title, pay_date, surgeon_lookup):
    """
    Add one tab per unique Distrib Code to *out_wb* (must contain a 'template' sheet).
    Returns the number of tabs created.
    """
    tmpl_ws = get_sheet_ci(out_wb, "template")
    if tmpl_ws is None:
        print("  Warning: no 'template' sheet — skipping distributor tabs")
        return 0

    tmpl_header_row, tmpl_label_map = find_template_header(tmpl_ws)
    if tmpl_header_row is None:
        print("  Warning: could not find template header — skipping distributor tabs")
        return 0

    placeholders      = scan_template_placeholders(tmpl_ws, tmpl_header_row)
    preformatted_rows = get_template_preformatted_rows(tmpl_ws, tmpl_header_row)
    first_row_borders = capture_row_borders(tmpl_ws, preformatted_rows[0]) if preformatted_rows else {}

    # Group rows by Distrib Code, then sort alphabetically by distributor name
    groups = []
    seen   = {}
    for cells in data_rows:
        val  = cells[dist_col_0idx].value if dist_col_0idx < len(cells) else None
        code = str(val).strip() if val else ""
        if not code:
            continue
        if code in seen:
            groups[seen[code]][1].append(cells)
        else:
            seen[code] = len(groups)
            groups.append((code, [cells]))

    groups.sort(key=lambda g: surgeon_lookup.get(g[0], {}).get("name", g[0]).lower())

    # Pre-read template images to temp files (workaround for PIL closing BytesIO)
    template_images = []
    tmp_paths       = []
    for img in list(tmpl_ws._images):
        try:
            raw    = img._data()
            suffix = ".png" if raw[:4] == b"\x89PNG" else ".jpg"
            tmp    = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
            tmp.write(raw)
            tmp.close()
            template_images.append((tmp.name, img.anchor))
            tmp_paths.append(tmp.name)
            img.ref = tmp.name  # fix template's own image ref for save()
        except Exception as e:
            print(f"  Warning: could not extract template image: {e}")

    existing_names = set(out_wb.sheetnames)
    count = 0

    for dist_code, rows in groups:
        info      = surgeon_lookup.get(dist_code, {})
        dist_name = info.get("name", dist_code)
        contact   = info.get("contact", "")
        if not info:
            print(f"  Warning: no lookup entry for '{dist_code}'")

        tab_name = make_tab_name(dist_code, dist_name, existing_names)
        existing_names.add(tab_name)

        # Copy template sheet, then manually add images (copy_worksheet skips them)
        new_ws       = out_wb.copy_worksheet(tmpl_ws)
        new_ws.title = tab_name
        for path, anchor in template_images:
            try:
                ni        = XlImage(path)
                ni.anchor = anchor
                new_ws.add_image(ni)
            except Exception as e:
                print(f"  Warning: could not copy image for '{tab_name}': {e}")

        n, total_row = populate_distributor_tab(
            new_ws, tmpl_header_row, tmpl_label_map, placeholders,
            preformatted_rows, first_row_borders,
            dist_code, dist_name, contact, title, pay_date,
            rows, summary_label_map,
        )
        print(f"  Tab '{tab_name}': {n} rows, SUM at row {total_row}")
        count += 1

    # Store temp-file paths on the workbook for cleanup after save()
    if not hasattr(out_wb, "_tmp_image_files"):
        out_wb._tmp_image_files = []
    out_wb._tmp_image_files.extend(tmp_paths)

    return count


# ══════════════════════════════════════════════════════════════════════════════
# PDF CONVERSION
# ══════════════════════════════════════════════════════════════════════════════

def convert_to_pdf(xlsx_path):
    """Convert *xlsx_path* to PDF via LibreOffice headless.  Returns pdf path or None."""
    out_dir = os.path.dirname(xlsx_path)
    lo_home = os.path.join(out_dir, ".lo_home")
    os.makedirs(lo_home, exist_ok=True)
    try:
        subprocess.run(
            ["soffice", "--headless", "--norestore", "--calc",
             "--convert-to", "pdf", "--outdir", out_dir, xlsx_path],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, "HOME": lo_home},
        )
    except Exception:
        pass
    finally:
        shutil.rmtree(lo_home, ignore_errors=True)
    pdf = os.path.splitext(xlsx_path)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — SUMMARY TAB (per-distributor totals)
# ══════════════════════════════════════════════════════════════════════════════

# Shared font for the Summary sheet — 12-point for readability when printed
_SUMM_FONT      = Font(size=12)
_SUMM_FONT_BOLD = Font(size=12, bold=True)
_SUMM_FONT_HDR  = Font(size=12, bold=True)
_SUMM_FONT_TTL  = Font(size=14, bold=True)


def create_summary_tab(out_wb, group_summaries, title_val, pay_date_val):
    """
    Insert a 'Summary' sheet at position 0 with per-distributor commission totals.

    *group_summaries*: list of (dist_code, dist_name, n_surgeries, total_comm)
    Returns the new worksheet.
    """
    ws = out_wb.create_sheet("Summary", 0)

    # Column widths
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 18

    row = 1

    # Title
    if title_val:
        c = ws.cell(row=row, column=1, value=title_val)
        c.font = _SUMM_FONT_TTL
        ws.row_dimensions[row].height = 24
        row += 1

    # Pay date
    if pay_date_val:
        ws.cell(row=row, column=1, value="Pay Date:").font = _SUMM_FONT_BOLD
        pd = ws.cell(row=row, column=2, value=pay_date_val)
        pd.font = _SUMM_FONT
        if isinstance(pay_date_val, datetime):
            pd.number_format = "mm/dd/yyyy"
        row += 1

    row += 1  # blank separator

    # Header row
    hdr_row = row
    bottom_border = Border(bottom=THIN_SIDE)
    for col, label, align in [
        (1, "Distrib Code",    "left"),
        (2, "Distributor",     "left"),
        (3, "# of Surgeries",  "center"),
        (4, "Total Comm $",    "right"),
    ]:
        c = ws.cell(row=hdr_row, column=col, value=label)
        c.font      = _SUMM_FONT_HDR
        c.border    = bottom_border
        c.alignment = Alignment(horizontal=align)
    row += 1

    # Data rows
    grand_n   = 0
    grand_tot = 0.0
    for dist_code, dist_name, n_surg, total_comm in group_summaries:
        ws.cell(row=row, column=1, value=dist_code).font = _SUMM_FONT
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="left")

        ws.cell(row=row, column=2, value=dist_name).font = _SUMM_FONT
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="left")

        cn = ws.cell(row=row, column=3, value=n_surg)
        cn.font = _SUMM_FONT
        cn.alignment = Alignment(horizontal="center")

        ct = ws.cell(row=row, column=4, value=total_comm)
        ct.font = _SUMM_FONT
        ct.number_format = "#,##0.00"
        ct.alignment = Alignment(horizontal="right")

        grand_n   += n_surg
        grand_tot += total_comm
        row += 1

    # Grand total row
    row += 1
    top_border = Border(top=THIN_SIDE)
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = top_border

    ws.cell(row=row, column=2, value="Grand Total").font = _SUMM_FONT_BOLD

    gn = ws.cell(row=row, column=3, value=grand_n)
    gn.font      = _SUMM_FONT_BOLD
    gn.alignment = Alignment(horizontal="center")

    gt = ws.cell(row=row, column=4, value=grand_tot)
    gt.font          = _SUMM_FONT_BOLD
    gt.number_format = "#,##0.00"
    gt.alignment     = Alignment(horizontal="right")

    # Page setup — portrait, fit to one page wide
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 CLI — DISTRIBUTOR TABS
# ══════════════════════════════════════════════════════════════════════════════

def process_distributor_tabs(input_path):
    """
    CLI entry point: generate distributor tabs + Summary for one input file.
    Returns the output xlsx path.
    """
    wb_src = load_workbook(input_path)

    src_sheet = get_sheet_ci(wb_src, "masterlog")
    if src_sheet is None:
        sys.exit("Error: No 'masterlog' sheet found.")

    header_row_num, label_map = find_summary_header(src_sheet)
    if header_row_num is None:
        sys.exit("Error: Could not find header row in masterlog.")

    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    comm_idx = label_map.get("comm $")
    po_idx   = label_map.get("po")

    title_val, pay_date_val = scan_summary_meta(src_sheet, header_row_num)
    data_rows = collect_data_rows(src_sheet, header_row_num, hosp_idx, po_idx)
    surgeon_lookup = build_surgeon_lookup(wb_src)

    out_wb = load_workbook(input_path)
    keep_sheets_ci(out_wb, {"masterlog", "Surgeon lookup", "template"})

    num_tabs = generate_distributor_tabs(
        out_wb, data_rows, dist_idx, label_map,
        title_val, pay_date_val, surgeon_lookup,
    )

    # Build group summaries and create Summary tab
    group_summaries = _build_group_summaries(data_rows, dist_idx, comm_idx, surgeon_lookup)
    create_summary_tab(out_wb, group_summaries, title_val, pay_date_val)

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(input_path)),
        os.path.splitext(os.path.basename(input_path))[0] + "_distributor_tabs.xlsx",
    )
    out_wb.save(out_path)
    _cleanup_tmp_images(out_wb)

    pdf = convert_to_pdf(out_path)
    print(f"Saved {os.path.basename(out_path)} — {num_tabs} tabs"
          f"{', PDF saved' if pdf else ', PDF failed'}")
    return out_path


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 CLI — MANAGER SPLIT
# ══════════════════════════════════════════════════════════════════════════════

def process(input_path):
    """CLI entry point: split masterlog by manager into per-manager workbooks."""
    wb_src = load_workbook(input_path)

    src_sheet = get_sheet_ci(wb_src, "masterlog")
    if src_sheet is None:
        sys.exit("Error: No 'masterlog' sheet found in the workbook.")

    header_row_num, label_map = find_summary_header(src_sheet)
    if header_row_num is None:
        sys.exit("Error: Could not find header row in masterlog "
                 "(need Manager, Hospital, Comm $, and Distributor/Distrib Code).")

    mgr_idx  = label_map.get("manager")
    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    comm_idx = label_map.get("comm $")
    po_idx   = label_map.get("po")

    if any(x is None for x in [mgr_idx, hosp_idx, dist_idx, comm_idx]):
        sys.exit("Error: Missing required columns in masterlog header.")

    title_val, pay_date_val = scan_summary_meta(src_sheet, header_row_num)
    data_rows  = collect_data_rows(src_sheet, header_row_num, hosp_idx, po_idx)
    col_widths = {col: dim.width for col, dim in src_sheet.column_dimensions.items()}

    # Detect managers in order of first appearance
    managers, seen = [], set()
    for cells in data_rows:
        mgr = cells[mgr_idx].value if mgr_idx < len(cells) else None
        if mgr and str(mgr).strip() and str(mgr).strip() not in seen:
            seen.add(str(mgr).strip())
            managers.append(str(mgr).strip())
    if not managers:
        sys.exit("Error: No manager values found in data rows.")

    surgeon_lookup = build_surgeon_lookup(wb_src)
    input_dir = os.path.dirname(os.path.abspath(input_path))
    base_name = os.path.splitext(os.path.basename(input_path))[0]

    for manager in managers:
        out_wb = load_workbook(input_path)
        keep_sheets_ci(out_wb, KEEP_SHEETS)
        out_ws = get_sheet_ci(out_wb, "masterlog")

        # Clear data area
        for r in range(header_row_num + 1, out_ws.max_row + 1):
            for c in range(1, out_ws.max_column + 1):
                out_ws.cell(row=r, column=c).value = None

        # Write this manager's rows
        mgr_rows = [c for c in data_rows if str(c[mgr_idx].value).strip() == manager]
        wr = header_row_num + 1
        for cells in mgr_rows:
            for ci, src in enumerate(cells):
                copy_cell(src, out_ws.cell(row=wr, column=ci + 1))
            wr += 1

        # Restore column widths, insert subtotals, apply alignment
        for letter, w in col_widths.items():
            out_ws.column_dimensions[letter].width = w
        insert_distributor_subtotals(out_ws, header_row_num, header_row_num + 1, dist_idx, comm_idx)
        apply_summary_alignment(out_ws, label_map, header_row_num)

        safe   = manager.replace("/", "_").replace("\\", "_")
        out_path = os.path.join(input_dir, f"{safe}-{base_name}.xlsx")
        out_wb.save(out_path)

        pdf = convert_to_pdf(out_path)
        print(f"Saved {os.path.basename(out_path)} — {len(mgr_rows)} rows"
              f"{', PDF saved' if pdf else ', PDF failed'}")


# ══════════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _build_group_summaries(data_rows, dist_idx, comm_idx, surgeon_lookup):
    """
    Aggregate data rows into per-distributor summaries.
    Returns [(dist_code, dist_name, n_surgeries, total_comm), …]
    in order of first appearance.
    """
    groups = []
    seen   = {}
    for cells in data_rows:
        val  = cells[dist_idx].value if dist_idx < len(cells) else None
        code = str(val).strip() if val else ""
        if not code:
            continue
        cval = cells[comm_idx].value if comm_idx is not None and comm_idx < len(cells) else None
        comm = float(cval) if isinstance(cval, (int, float)) else 0.0

        if code in seen:
            i = seen[code]
            c, n, nr, tc = groups[i]
            groups[i] = (c, n, nr + 1, tc + comm)
        else:
            seen[code] = len(groups)
            info = surgeon_lookup.get(code, {})
            groups.append((code, info.get("name", code), 1, comm))

    groups.sort(key=lambda g: g[1].lower())
    return groups


def _cleanup_tmp_images(wb):
    """Remove temp image files stored on the workbook during tab generation."""
    for path in getattr(wb, "_tmp_image_files", []):
        try:
            os.unlink(path)
        except OSError:
            pass


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python process_commissions.py <input.xlsx>")
    process(sys.argv[1])
