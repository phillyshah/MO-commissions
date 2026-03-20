"""
process_commissions.py — MO Commission Report Processor

Usage:
    python process_commissions.py <input.xlsx>

Produces one output .xlsx per manager found in the Summary sheet.
Each workbook contains:
  - Summary sheet (manager's rows + distributor subtotals)
  - Surgeon lookup sheet (preserved)
  - template sheet (preserved)
  - One distributor tab per unique Distrib Code, built from the template sheet
"""

import sys
import os
import copy
import io
import subprocess
import shutil
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage
from openpyxl.cell.cell import MergedCell


# ── Constants ─────────────────────────────────────────────────────────────────

# Summary sheet alignment rules
LEFT_COLS  = {"po", "notes", "surgeon", "hospital", "manager"}
RIGHT_COLS = {"surgery date", "inv#", "inv date", "due date", "comm $"}

# Style constants
NO_FILL     = PatternFill(fill_type=None)
CLEAR_BORDER = Border()

# Summary → Template column label aliases
SUMMARY_TO_TEMPLATE = {
    "hospital": "hospital or mc",
}

# Summary columns with no template equivalent (skip when writing distributor tabs)
SUMMARY_SKIP = {"manager", "distrib code", "distributor", "status", "date pd"}

# Sheets to carry forward into each manager workbook (Step 1 — simple split)
KEEP_SHEETS = {"masterlog"}


# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY SHEET HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def find_summary_header(sheet):
    """
    Scan from top to find the Summary header row.
    Must contain: Manager, Hospital, Comm $, and either Distributor or Distrib Code.
    Returns (1-based row num, {normalized_label: 0-based col index}).
    """
    REQUIRED_FIXED = {"manager", "hospital", "comm $"}
    DIST_OPTIONS   = {"distributor", "distrib code"}

    for row in sheet.iter_rows():
        values = [str(c.value).strip().lower() if c.value is not None else "" for c in row]
        val_set = set(values)
        if REQUIRED_FIXED.issubset(val_set) and DIST_OPTIONS.intersection(val_set):
            return row[0].row, {v: i for i, v in enumerate(values)}
    return None, None


def scan_summary_meta(sheet, header_row_num):
    """
    Scan rows above the header for Title: and Pay Date: label/value pairs.
    Returns (title_str, pay_date_value).
    """
    title_val    = None
    pay_date_val = None

    for row in sheet.iter_rows(min_row=1, max_row=header_row_num - 1):
        cells = list(row)
        for i, cell in enumerate(cells):
            if cell.value is None:
                continue
            label = str(cell.value).strip().lower()
            if label == "title:":
                for j in range(i + 1, len(cells)):
                    if cells[j].value is not None:
                        title_val = str(cells[j].value).strip()
                        break
            elif label == "pay date:":
                for j in range(i + 1, len(cells)):
                    if cells[j].value is not None:
                        pay_date_val = cells[j].value
                        break

    return title_val, pay_date_val


def is_legacy_subtotal(row_cells, po_idx):
    """Return True if this row is a legacy subtotal (bold 'Total…' string in PO cell)."""
    cell = row_cells[po_idx]
    val  = cell.value
    if val is None:
        return False
    val_str = str(val).strip()
    if val_str.lower().startswith("total"):
        cleaned = (val_str.replace(",", "").replace(".", "")
                   .replace("Total", "").replace("total", "")
                   .strip().lstrip("-"))
        if not cleaned.isnumeric() and cell.font and cell.font.bold:
            return True
    return False


def copy_cell(src, dst):
    """Copy value and full style from src to dst cell."""
    dst.value = src.value
    if src.has_style:
        dst.font         = copy.copy(src.font)
        dst.fill         = copy.copy(src.fill)
        dst.border       = copy.copy(src.border)
        dst.alignment    = copy.copy(src.alignment)
        dst.number_format = src.number_format


def apply_summary_alignment(ws, label_map, header_row_num):
    """Apply left/right alignment rules to all rows including the header."""
    col_align = {}
    for label, idx in label_map.items():
        if label in LEFT_COLS:
            col_align[idx + 1] = "left"
        elif label in RIGHT_COLS:
            col_align[idx + 1] = "right"

    for row in ws.iter_rows(min_row=header_row_num):
        for cell in row:
            if cell.column in col_align:
                ex = cell.alignment or Alignment()
                cell.alignment = Alignment(
                    horizontal=col_align[cell.column],
                    vertical=ex.vertical,
                    wrap_text=ex.wrap_text,
                )


def insert_distributor_subtotals(ws, header_row_num, data_start_row,
                                  dist_col_0idx, comm_col_0idx):
    """
    Group data rows by consecutive distributor code and insert subtotal rows.
    Returns count of distinct groups.
    """
    # Collect existing data
    data_rows = []
    for row in ws.iter_rows(min_row=data_start_row):
        data_rows.append([cell.value for cell in row])

    # Build consecutive groups
    blocks = []
    for row_data in data_rows:
        dist_val = row_data[dist_col_0idx] if dist_col_0idx < len(row_data) else None
        dist_key = str(dist_val).strip() if dist_val else ""
        if blocks and blocks[-1][0] == dist_key:
            blocks[-1][1].append(row_data)
        else:
            blocks.append((dist_key, [row_data]))

    # Clear existing data area
    for row in ws.iter_rows(min_row=data_start_row):
        for cell in row:
            cell.value = None

    current_row      = data_start_row
    comm_col_letter  = get_column_letter(comm_col_0idx + 1)

    for dist_code, rows in blocks:
        group_start = current_row
        for row_data in rows:
            for ci, val in enumerate(row_data):
                ws.cell(row=current_row, column=ci + 1).value = val
            current_row += 1
        group_end = current_row - 1

        current_row += 1  # blank row before subtotal

        dist_c = ws.cell(row=current_row, column=dist_col_0idx + 1)
        dist_c.value = f"Total {dist_code}"
        dist_c.font  = Font(bold=True)

        comm_c = ws.cell(row=current_row, column=comm_col_0idx + 1)
        comm_c.value  = f"=SUM({comm_col_letter}{group_start}:{comm_col_letter}{group_end})"
        comm_c.font   = Font(bold=True)
        comm_c.number_format = "#,##0.00"

        current_row += 2  # subtotal row + trailing blank row

    return len(blocks)


# ══════════════════════════════════════════════════════════════════════════════
# SURGEON LOOKUP
# ══════════════════════════════════════════════════════════════════════════════

def build_surgeon_lookup(wb):
    """
    Build {distrib_code: {name, contact}} from the 'Surgeon lookup' sheet.
    Uses the first matching row for each code.
    """
    if "Surgeon lookup" not in wb.sheetnames:
        return {}

    ws = wb["Surgeon lookup"]
    REQUIRED = {"distrib code", "distributor", "contact"}
    header_row_num = None
    label_map = {}

    for row in ws.iter_rows():
        values = [str(c.value).strip().lower() if c.value else "" for c in row]
        if REQUIRED.issubset(set(values)):
            header_row_num = row[0].row
            label_map = {v: i for i, v in enumerate(values)}
            break

    if header_row_num is None:
        return {}

    ci  = label_map["distrib code"]
    ni  = label_map["distributor"]
    coi = label_map["contact"]

    lookup = {}
    for row in ws.iter_rows(min_row=header_row_num + 1, max_row=ws.max_row):
        cells = list(row)
        code  = cells[ci].value if ci < len(cells) else None
        if not code:
            continue
        code = str(code).strip()
        if code in lookup:
            continue  # first match only
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
    Returns (1-based row num, {normalized_label: 0-based col index}).
    """
    REQUIRED = {"inv #", "comm $"}
    for row in ws.iter_rows():
        values = [str(c.value).strip().lower() if c.value else "" for c in row]
        if REQUIRED.issubset(set(values)):
            return row[0].row, {v: i for i, v in enumerate(values)}
    return None, None


def scan_template_placeholders(ws, header_row_num):
    """
    Scan rows above the header for placeholder cells matching known keys.
    Returns {lower_key: (row, 1-based col)}.
    Keys: "distributor", "distrib code", "contact", "title", "=summary!b2"
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
    Find pre-formatted (bordered) empty data rows below the header.
    Returns list of 1-based row numbers; stops at first unbordered row.
    """
    rows = []
    for row in ws.iter_rows(min_row=header_row_num + 1, max_row=ws.max_row):
        has_border = any(
            (not isinstance(c, MergedCell) and c.border and
             (c.border.left.style or c.border.right.style or
              c.border.top.style  or c.border.bottom.style))
            for c in row
        )
        if has_border:
            rows.append(row[0].row)
        else:
            break
    return rows


def capture_row_borders(ws, row_num):
    """Return {1-based col: Border copy} for all cells in row_num."""
    result = {}
    for cell in next(ws.iter_rows(min_row=row_num, max_row=row_num)):
        if not isinstance(cell, MergedCell):
            result[cell.column] = copy.copy(cell.border) if cell.border else CLEAR_BORDER
    return result


# ══════════════════════════════════════════════════════════════════════════════
# TAB NAME LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def make_tab_name(code, dist_name, existing_names):
    """
    Build a sheet tab name ≤ 31 chars: '{code} ({first N chars of name})'.
    Always includes the full code; truncates name suffix to fit.
    Deduplicates against existing_names.
    """
    suffix    = dist_name[:7] if dist_name else code
    candidate = f"{code} ({suffix})"
    if len(candidate) > 31:
        max_suffix = 31 - len(code) - 3   # 3 = len(" ()")
        candidate  = f"{code} ({dist_name[:max(max_suffix, 0)]})"
    candidate = candidate[:31]

    base = candidate
    n    = 2
    while candidate in existing_names:
        ext       = f"_{n}"
        candidate = base[:31 - len(ext)] + ext
        n += 1

    return candidate


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN MAPPING (Summary → Template)
# ══════════════════════════════════════════════════════════════════════════════

def build_column_map(summary_label_map, template_label_map):
    """
    Returns list of (sum_0idx, tmpl_0idx) for columns present in both sheets.
    Applies label aliases; skips SUMMARY_SKIP columns.
    """
    mapping = []
    for sum_label, sum_0idx in summary_label_map.items():
        if sum_label in SUMMARY_SKIP:
            continue
        tmpl_label = SUMMARY_TO_TEMPLATE.get(sum_label, sum_label)
        if tmpl_label in template_label_map:
            mapping.append((sum_0idx, template_label_map[tmpl_label]))
    return mapping


# ══════════════════════════════════════════════════════════════════════════════
# DISTRIBUTOR TAB POPULATION
# ══════════════════════════════════════════════════════════════════════════════

def populate_distributor_tab(ws, tmpl_header_row, tmpl_label_map, placeholders,
                              preformatted_rows, first_row_borders,
                              dist_code, dist_name, contact, title, pay_date,
                              row_cells_list, summary_label_map):
    """
    Populate a copied template sheet with one distributor's data.
    Steps: clear template grid → fill placeholders → write rows → total row → page setup.
    Returns (n_data_rows, total_row_num).
    """
    # 1. Clear pre-formatted data rows (value, border, font)
    for r in preformatted_rows:
        for col in range(1, ws.max_column + 1):
            c = ws.cell(row=r, column=col)
            if isinstance(c, MergedCell):
                continue
            c.value  = None
            c.border = CLEAR_BORDER
            c.font   = Font()

    # 2. Clear any stale =SUM formulas below the header
    for row in ws.iter_rows(min_row=tmpl_header_row + 1, max_row=ws.max_row):
        for c in row:
            if (not isinstance(c, MergedCell) and c.value
                    and isinstance(c.value, str)
                    and c.value.upper().startswith("=SUM")):
                c.value = None

    # 3. Populate header placeholders
    def _set(key, value, num_fmt=None):
        if key not in placeholders:
            return
        r, col_ = placeholders[key]
        c = ws.cell(row=r, column=col_)
        c.value = value
        c.fill  = NO_FILL
        if num_fmt:
            c.number_format = num_fmt

    _set("distributor",  dist_name)
    _set("distrib code", dist_code)
    _set("contact",      contact)
    _set("title",        title)

    # Pay date — replace =Summary!B2 formula with actual value
    if "=summary!b2" in placeholders:
        r, col_ = placeholders["=summary!b2"]
        c = ws.cell(row=r, column=col_)
        if isinstance(pay_date, datetime):
            c.value = pay_date
        elif pay_date:
            c.value = pay_date
        c.number_format = "mm/dd/yyyy"
        c.fill = NO_FILL

    # 4. Write data rows
    col_map             = build_column_map(summary_label_map, tmpl_label_map)
    tmpl_0idx_to_label  = {v: k for k, v in tmpl_label_map.items()}
    comm_tmpl_0idx      = tmpl_label_map.get("comm $")
    hosp_tmpl_0idx      = tmpl_label_map.get("hospital or mc")

    write_row      = tmpl_header_row + 1
    data_start_row = write_row
    max_hosp_len   = len("Hospital or MC")

    for src_cells in row_cells_list:
        for sum_0idx, tmpl_0idx in col_map:
            if sum_0idx >= len(src_cells):
                continue
            src_cell = src_cells[sum_0idx]
            dst_c    = ws.cell(row=write_row, column=tmpl_0idx + 1)
            dst_c.value = src_cell.value

            # Preserve source number format (dates, currency, etc.)
            if src_cell.has_style and src_cell.number_format:
                dst_c.number_format = src_cell.number_format

            # Border from template's first data row
            if tmpl_0idx + 1 in first_row_borders:
                dst_c.border = copy.copy(first_row_borders[tmpl_0idx + 1])

            # Alignment: center Comm % / Comm $; left-align everything else
            tmpl_label = tmpl_0idx_to_label.get(tmpl_0idx, "")
            if tmpl_label in {"comm %", "comm $"}:
                dst_c.alignment = Alignment(horizontal="center")
            else:
                dst_c.alignment = Alignment(horizontal="left")

            # Track longest Hospital value for auto-sizing
            if hosp_tmpl_0idx is not None and tmpl_0idx == hosp_tmpl_0idx:
                val_len      = len(str(src_cell.value)) if src_cell.value else 0
                max_hosp_len = max(max_hosp_len, val_len)

        write_row += 1

    data_end_row = write_row - 1
    n_data_rows  = max(data_end_row - data_start_row + 1, 0)

    # 5. Auto-size Hospital column
    if hosp_tmpl_0idx is not None:
        col_letter = get_column_letter(hosp_tmpl_0idx + 1)
        ws.column_dimensions[col_letter].width = max_hosp_len + 2

    # 6. Total row (one blank row gap, then totals)
    total_row = data_end_row + 2
    if comm_tmpl_0idx is not None:
        comm_col_letter = get_column_letter(comm_tmpl_0idx + 1)

        # Label in the column immediately before Comm $
        label_col = comm_tmpl_0idx  # 1-based = comm_tmpl_0idx (== 0-based comm - 1 + 1)
        if label_col >= 1:
            lbl_c       = ws.cell(row=total_row, column=label_col)
            lbl_c.value = "Total Comm $"
            lbl_c.font  = Font(bold=True)

        # SUM formula in Comm $ column
        sum_c              = ws.cell(row=total_row, column=comm_tmpl_0idx + 1)
        sum_c.value        = f"=SUM({comm_col_letter}{data_start_row}:{comm_col_letter}{data_end_row})"
        sum_c.font         = Font(bold=True)
        sum_c.number_format = "#,##0.00"

    # 7. Page setup
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return n_data_rows, total_row


# ══════════════════════════════════════════════════════════════════════════════
# GENERATE DISTRIBUTOR TABS
# ══════════════════════════════════════════════════════════════════════════════

def generate_distributor_tabs(out_wb, manager_rows, dist_col_0idx,
                               summary_label_map, title, pay_date, surgeon_lookup):
    """
    Add distributor tabs to out_wb (which must contain a 'template' sheet).
    manager_rows: list of row-cell-lists for the current manager.
    Returns count of tabs generated.
    """
    if "template" not in out_wb.sheetnames:
        print("  Warning: no 'template' sheet — skipping distributor tabs")
        return 0

    tmpl_ws = out_wb["template"]
    tmpl_header_row, tmpl_label_map = find_template_header(tmpl_ws)
    if tmpl_header_row is None:
        print("  Warning: could not find template header row — skipping distributor tabs")
        return 0

    placeholders      = scan_template_placeholders(tmpl_ws, tmpl_header_row)
    preformatted_rows = get_template_preformatted_rows(tmpl_ws, tmpl_header_row)
    first_row_borders = (capture_row_borders(tmpl_ws, preformatted_rows[0])
                         if preformatted_rows else {})

    # Group rows by Distrib Code, preserving order of first appearance
    groups = []
    seen   = {}
    for row_cells in manager_rows:
        code_val = (row_cells[dist_col_0idx].value
                    if dist_col_0idx < len(row_cells) else None)
        code = str(code_val).strip() if code_val else ""
        if not code:
            continue
        if code in seen:
            groups[seen[code]][1].append(row_cells)
        else:
            seen[code] = len(groups)
            groups.append((code, [row_cells]))

    # Pre-read template image bytes and write to temp files.
    # PIL closes BytesIO objects after opening them (via __del__), so using
    # file paths lets PIL reliably reopen each image on every call to _data().
    import tempfile
    template_images   = []   # list of (tmp_file_path, anchor)
    tmpl_image_temps  = []   # keep file paths for cleanup

    for img in list(tmpl_ws._images):
        try:
            raw = img._data()
            suffix = ".png" if raw[:4] == b"\x89PNG" else ".jpg"
            tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
            tmp.write(raw)
            tmp.close()
            template_images.append((tmp.name, img.anchor))
            tmpl_image_temps.append(tmp.name)
            # Fix the template sheet's own image ref to use the file path
            # so it also writes correctly when the workbook is saved
            img.ref = tmp.name
        except Exception as e:
            print(f"  Warning: could not extract template image: {e}")

    existing_names = set(out_wb.sheetnames)
    count = 0

    for dist_code, rows in groups:
        info      = surgeon_lookup.get(dist_code, {})
        dist_name = info.get("name", dist_code)
        if not info:
            print(f"  Warning: no lookup entry for '{dist_code}' — using code as name")
        contact = info.get("contact", "")

        tab_name = make_tab_name(dist_code, dist_name, existing_names)
        existing_names.add(tab_name)

        # Create tab by copying the template (in-workbook copy)
        if tab_name in out_wb.sheetnames:
            new_ws = out_wb[tab_name]
        else:
            new_ws       = out_wb.copy_worksheet(tmpl_ws)
            new_ws.title = tab_name

            # copy_worksheet does not copy images — add from temp file paths
            for tmp_path, img_anchor in template_images:
                try:
                    new_img        = XlImage(tmp_path)
                    new_img.anchor = img_anchor
                    new_ws.add_image(new_img)
                except Exception as e:
                    print(f"  Warning: could not copy image for '{tab_name}': {e}")

        n_rows, total_row = populate_distributor_tab(
            new_ws, tmpl_header_row, tmpl_label_map, placeholders,
            preformatted_rows, first_row_borders,
            dist_code, dist_name, contact, title, pay_date,
            rows, summary_label_map,
        )
        print(f"  Tab '{tab_name}': {n_rows} data rows, SUM at row {total_row}")
        count += 1

    # Temp files are kept alive until after save() is called by the caller;
    # caller is responsible for cleanup, or they self-clean on process exit.
    # Store refs on the workbook object so they outlive this function.
    if not hasattr(out_wb, "_tmp_image_files"):
        out_wb._tmp_image_files = []
    out_wb._tmp_image_files.extend(tmpl_image_temps)

    return count


# ══════════════════════════════════════════════════════════════════════════════
# PDF CONVERSION
# ══════════════════════════════════════════════════════════════════════════════

def convert_to_pdf(xlsx_path):
    """Convert xlsx to PDF via LibreOffice headless. Returns pdf path or None."""
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
    pdf_path = os.path.splitext(xlsx_path)[0] + ".pdf"
    return pdf_path if os.path.exists(pdf_path) else None


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — SUMMARY TAB
# ══════════════════════════════════════════════════════════════════════════════

def create_summary_tab(out_wb, group_summaries, title_val, pay_date_val):
    """
    Insert a 'Summary' sheet at position 0 with per-distributor commission totals.

    group_summaries: list of (dist_code, dist_name, n_rows, total_comm)
    Returns the new worksheet.
    """
    ws = out_wb.create_sheet("Summary", 0)

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 16

    row = 1

    # Title
    if title_val:
        c = ws.cell(row=row, column=1, value=title_val)
        c.font = Font(bold=True, size=13)
        ws.row_dimensions[row].height = 22
        row += 1

    # Pay date
    if pay_date_val:
        lbl = ws.cell(row=row, column=1, value="Pay Date:")
        lbl.font = Font(bold=True, size=10)
        c = ws.cell(row=row, column=2, value=pay_date_val)
        from datetime import datetime as _dt
        if isinstance(pay_date_val, _dt):
            c.number_format = "mm/dd/yyyy"
        row += 1

    row += 1  # blank separator

    # Header row
    hdr_row = row
    thin = Side(style='thin')
    bottom_border = Border(bottom=thin)
    for col, label, align in [
        (1, "Distrib Code", "left"),
        (2, "Distributor",  "left"),
        (3, "# Lines",      "center"),
        (4, "Total Comm $", "right"),
    ]:
        c = ws.cell(row=hdr_row, column=col, value=label)
        c.font      = Font(bold=True, size=10)
        c.border    = bottom_border
        c.alignment = Alignment(horizontal=align)
    row += 1

    # Data rows
    grand_n   = 0
    grand_tot = 0.0
    for dist_code, dist_name, n_rows, total_comm in group_summaries:
        ws.cell(row=row, column=1, value=dist_code).alignment = Alignment(horizontal="left")
        ws.cell(row=row, column=2, value=dist_name).alignment = Alignment(horizontal="left")
        c_n = ws.cell(row=row, column=3, value=n_rows)
        c_n.alignment = Alignment(horizontal="center")
        c_t = ws.cell(row=row, column=4, value=total_comm)
        c_t.number_format = "#,##0.00"
        c_t.alignment = Alignment(horizontal="right")
        grand_n   += n_rows
        grand_tot += total_comm
        row += 1

    row += 1  # blank before grand total

    top_border = Border(top=thin)
    for col in range(1, 5):
        ws.cell(row=row, column=col).border = top_border

    lbl_c = ws.cell(row=row, column=2, value="Grand Total")
    lbl_c.font = Font(bold=True, size=10)

    c_gn = ws.cell(row=row, column=3, value=grand_n)
    c_gn.font      = Font(bold=True, size=10)
    c_gn.alignment = Alignment(horizontal="center")

    c_gt = ws.cell(row=row, column=4, value=grand_tot)
    c_gt.font         = Font(bold=True, size=10)
    c_gt.number_format = "#,##0.00"
    c_gt.alignment    = Alignment(horizontal="right")

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — DISTRIBUTOR TABS (all managers combined, one tab per Distrib Code)
# ══════════════════════════════════════════════════════════════════════════════

def process_distributor_tabs(input_path):
    """
    Generate one distributor tab per unique Distrib Code (across all managers).
    Returns the output xlsx path.
    """
    wb_src = load_workbook(input_path)

    if "masterlog" not in wb_src.sheetnames:
        sys.exit("Error: No 'masterlog' sheet found.")

    src_sheet                 = wb_src["masterlog"]
    header_row_num, label_map = find_summary_header(src_sheet)
    if header_row_num is None:
        sys.exit("Error: Could not find header row in Summary sheet.")

    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    po_idx   = label_map.get("po")

    title_val, pay_date_val = scan_summary_meta(src_sheet, header_row_num)

    all_rows  = list(src_sheet.iter_rows())
    data_rows = []
    for row in all_rows[header_row_num:]:          # header_row_num is 1-based, so index = header_row_num (skips header)
        cells    = list(row)
        hosp_val = cells[hosp_idx].value if hosp_idx < len(cells) else None
        if not hosp_val or not str(hosp_val).strip():
            continue
        if po_idx is not None and is_legacy_subtotal(cells, po_idx):
            continue
        data_rows.append(cells)

    surgeon_lookup = build_surgeon_lookup(wb_src)

    out_wb = load_workbook(input_path)
    for sname in list(out_wb.sheetnames):
        if sname not in {"masterlog", "Surgeon lookup", "template"}:
            del out_wb[sname]

    num_tabs = generate_distributor_tabs(
        out_wb, data_rows, dist_idx, label_map,
        title_val, pay_date_val, surgeon_lookup,
    )

    input_dir = os.path.dirname(os.path.abspath(input_path))
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    out_path  = os.path.join(input_dir, f"{base_name}_distributor_tabs.xlsx")
    out_wb.save(out_path)

    for tmp_path in getattr(out_wb, "_tmp_image_files", []):
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    pdf_path   = convert_to_pdf(out_path)
    pdf_status = ", PDF saved" if pdf_path else ", PDF: failed"
    print(f"Saved {os.path.basename(out_path)} — {num_tabs} distributor tabs{pdf_status}")
    return out_path


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def process(input_path):
    wb_src = load_workbook(input_path)

    if "masterlog" not in wb_src.sheetnames:
        sys.exit("Error: No 'masterlog' sheet found in the workbook.")

    src_sheet                    = wb_src["masterlog"]
    header_row_num, label_map    = find_summary_header(src_sheet)
    if header_row_num is None:
        sys.exit("Error: Could not find header row in masterlog sheet "
                 "(need Manager, Hospital, Comm $, and Distributor/Distrib Code).")

    mgr_idx  = label_map.get("manager")
    hosp_idx = label_map.get("hospital")
    dist_idx = label_map.get("distrib code", label_map.get("distributor"))
    comm_idx = label_map.get("comm $")
    po_idx   = label_map.get("po")

    if any(x is None for x in [mgr_idx, hosp_idx, dist_idx, comm_idx]):
        sys.exit("Error: Missing required columns in Summary header.")

    title_val, pay_date_val = scan_summary_meta(src_sheet, header_row_num)

    # Collect data rows
    all_rows  = list(src_sheet.iter_rows())
    header_0  = header_row_num - 1
    data_rows = []
    for row in all_rows[header_0 + 1:]:
        cells    = list(row)
        hosp_val = cells[hosp_idx].value if hosp_idx < len(cells) else None
        if not hosp_val or not str(hosp_val).strip():
            continue
        if po_idx is not None and is_legacy_subtotal(cells, po_idx):
            continue
        data_rows.append(cells)

    # Detect managers (order of first appearance)
    managers, seen = [], set()
    for cells in data_rows:
        mgr = cells[mgr_idx].value if mgr_idx < len(cells) else None
        if mgr and str(mgr).strip() and str(mgr).strip() not in seen:
            seen.add(str(mgr).strip())
            managers.append(str(mgr).strip())

    if not managers:
        sys.exit("Error: No manager values found in data rows.")

    surgeon_lookup = build_surgeon_lookup(wb_src)
    col_widths     = {col: dim.width for col, dim in src_sheet.column_dimensions.items()}
    input_dir      = os.path.dirname(os.path.abspath(input_path))
    base_name      = os.path.splitext(os.path.basename(input_path))[0]

    for manager in managers:
        out_wb = load_workbook(input_path)

        # Remove sheets not needed in the output
        for sname in list(out_wb.sheetnames):
            if sname not in KEEP_SHEETS:
                del out_wb[sname]

        out_ws  = out_wb["masterlog"]
        max_row = out_ws.max_row

        # Clear existing data rows
        for r in range(header_row_num + 1, max_row + 1):
            for c in range(1, out_ws.max_column + 1):
                out_ws.cell(row=r, column=c).value = None

        # Write this manager's rows
        manager_rows = [c for c in data_rows
                        if str(c[mgr_idx].value).strip() == manager]

        write_row = header_row_num + 1
        for cells in manager_rows:
            for col_0, src_cell in enumerate(cells):
                copy_cell(src_cell, out_ws.cell(row=write_row, column=col_0 + 1))
            write_row += 1

        # Restore column widths
        for col_letter, width in col_widths.items():
            out_ws.column_dimensions[col_letter].width = width

        # Insert distributor subtotals in Summary
        num_dists = insert_distributor_subtotals(
            out_ws, header_row_num, header_row_num + 1, dist_idx, comm_idx)

        # Apply column alignment
        apply_summary_alignment(out_ws, label_map, header_row_num)

        # Save xlsx
        safe_mgr  = manager.replace("/", "_").replace("\\", "_")
        out_name  = f"{safe_mgr}-{base_name}.xlsx"
        out_path  = os.path.join(input_dir, out_name)
        out_wb.save(out_path)

        # Convert to PDF
        pdf_path   = convert_to_pdf(out_path)
        pdf_status = ", PDF saved" if pdf_path else ", PDF: failed (LibreOffice not found?)"
        print(f"Saved {out_name} — {len(manager_rows)} data rows, "
              f"{num_dists} distributor groups{pdf_status}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python process_commissions.py <input.xlsx>")
    process(sys.argv[1])
