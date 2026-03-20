"""
process_commissions.py — MO Commission Report Processor

Usage:
    python process_commissions.py <input.xlsx>

Produces one output .xlsx per manager found in the Summary sheet,
with distributor subtotals inserted after each consecutive group.
"""

import sys
import os
import copy
import subprocess
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


LEFT_ALIGNED_COLS = {"po", "notes", "surgeon", "hospital", "manager"}
RIGHT_ALIGNED_COLS = {"surgery date", "inv#", "inv date", "due date", "comm $"}


REQUIRED_HEADERS = {"manager", "hospital", "distributor", "comm $"}


def find_header_row(sheet):
    """Scan from the top to find the row containing all required column headers."""
    for row in sheet.iter_rows():
        values = [str(cell.value).strip().lower() if cell.value is not None else "" for cell in row]
        if REQUIRED_HEADERS.issubset(set(values)):
            return row[0].row, {v: i for i, v in enumerate(values)}
    return None, None


def col_index(label_map, label):
    """Return 0-based index for a column label (case-insensitive)."""
    return label_map[label.lower().strip()]


def is_legacy_subtotal(row_cells, po_idx):
    """Return True if this row is a legacy subtotal (bold non-numeric 'Total...' in PO column)."""
    cell = row_cells[po_idx]
    val = cell.value
    if val is None:
        return False
    val_str = str(val).strip()
    if val_str.lower().startswith("total") and not val_str.replace(",", "").replace(".", "").replace("Total", "").strip().lstrip("-").isnumeric():
        if cell.font and cell.font.bold:
            return True
    return False


def copy_cell(src, dst):
    """Copy value and style from src cell to dst cell."""
    dst.value = src.value
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format


def apply_column_alignment(ws, label_map, header_row_num):
    """Apply left/right alignment to designated columns for all rows including header."""
    col_alignments = {}
    for label, idx in label_map.items():
        if label in LEFT_ALIGNED_COLS:
            col_alignments[idx + 1] = "left"   # 1-based column number
        elif label in RIGHT_ALIGNED_COLS:
            col_alignments[idx + 1] = "right"

    for row in ws.iter_rows(min_row=header_row_num):
        for cell in row:
            if cell.column in col_alignments:
                existing = cell.alignment or Alignment()
                cell.alignment = Alignment(
                    horizontal=col_alignments[cell.column],
                    vertical=existing.vertical,
                    wrap_text=existing.wrap_text,
                )


def insert_distributor_subtotals(ws, header_row_num, data_start_row, dist_col_idx, comm_col_idx):
    """
    Post-process the sheet to insert subtotal rows after each consecutive
    distributor group. Operates by collecting rows, rebuilding the sheet content
    after the header block.
    """
    # Collect all data rows currently written (from data_start_row onward)
    data_rows = []
    for row in ws.iter_rows(min_row=data_start_row):
        data_rows.append([cell.value for cell in row])

    # Build output blocks grouped by consecutive distributor
    blocks = []  # list of (distributor_code, [row_data, ...])
    for row_data in data_rows:
        dist_val = row_data[dist_col_idx] if dist_col_idx < len(row_data) else None
        dist_key = str(dist_val).strip() if dist_val else ""
        if blocks and blocks[-1][0] == dist_key:
            blocks[-1][1].append(row_data)
        else:
            blocks.append((dist_key, [row_data]))

    # Clear everything from data_start_row down
    for row in ws.iter_rows(min_row=data_start_row):
        for cell in row:
            cell.value = None

    # Delete excess rows (write from scratch by overwriting)
    current_row = data_start_row
    dist_col_letter = get_column_letter(dist_col_idx + 1)
    comm_col_letter = get_column_letter(comm_col_idx + 1)

    for dist_code, rows in blocks:
        group_start = current_row
        for row_data in rows:
            for col_idx_0, val in enumerate(row_data):
                ws.cell(row=current_row, column=col_idx_0 + 1).value = val
            current_row += 1
        group_end = current_row - 1

        # Blank row
        current_row += 1

        # Subtotal row
        subtotal_row = current_row
        dist_cell = ws.cell(row=subtotal_row, column=dist_col_idx + 1)
        dist_cell.value = f"Total {dist_code}"
        dist_cell.font = Font(bold=True)

        comm_cell = ws.cell(row=subtotal_row, column=comm_col_idx + 1)
        comm_cell.value = f"=SUM({comm_col_letter}{group_start}:{comm_col_letter}{group_end})"
        comm_cell.font = Font(bold=True)
        comm_cell.number_format = "#,##0.00"

        current_row += 1

        # Blank row
        current_row += 1

    return len(blocks)


def convert_to_pdf(xlsx_path):
    """Convert xlsx to PDF using LibreOffice headless. Returns pdf path or None."""
    out_dir = os.path.dirname(xlsx_path)
    lo_home = os.path.join(out_dir, '.lo_home')
    os.makedirs(lo_home, exist_ok=True)
    try:
        subprocess.run(
            ['soffice', '--headless', '--norestore', '--calc',
             '--convert-to', 'pdf', '--outdir', out_dir, xlsx_path],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, 'HOME': lo_home}
        )
    except Exception:
        pass
    finally:
        shutil.rmtree(lo_home, ignore_errors=True)
    pdf_path = os.path.splitext(xlsx_path)[0] + '.pdf'
    return pdf_path if os.path.exists(pdf_path) else None


def process(input_path):
    wb = load_workbook(input_path)

    if "Summary" not in wb.sheetnames:
        sys.exit("Error: No 'Summary' sheet found in the workbook.")

    src_sheet = wb["Summary"]

    # Step 1: Find header row dynamically
    header_row_num, label_map = find_header_row(src_sheet)
    if header_row_num is None:
        sys.exit(
            "Error: Could not find a header row containing Manager, Hospital, Distributor, and Comm $."
        )

    mgr_idx = col_index(label_map, "manager")
    hosp_idx = col_index(label_map, "hospital")
    dist_idx = col_index(label_map, "distributor")
    comm_idx = col_index(label_map, "comm $")

    # Try to find PO column; fall back gracefully if absent
    po_idx = label_map.get("po", None)

    # Step 2: Collect all data rows
    all_rows = list(src_sheet.iter_rows())
    header_row_0 = header_row_num - 1  # 0-based index into all_rows

    data_rows = []
    for row in all_rows[header_row_0 + 1:]:
        cells = list(row)
        # Exclude blank hospital rows
        hosp_val = cells[hosp_idx].value if hosp_idx < len(cells) else None
        if hosp_val is None or str(hosp_val).strip() == "":
            continue
        # Exclude legacy subtotal rows
        if po_idx is not None and is_legacy_subtotal(cells, po_idx):
            continue
        data_rows.append(cells)

    # Step 3: Detect unique managers
    managers = []
    seen = set()
    for cells in data_rows:
        mgr = cells[mgr_idx].value if mgr_idx < len(cells) else None
        if mgr and str(mgr).strip() and str(mgr).strip() not in seen:
            seen.add(str(mgr).strip())
            managers.append(str(mgr).strip())

    if not managers:
        sys.exit("Error: No manager values found in data rows.")

    # Preserve column widths from source
    col_widths = {}
    for col_letter, col_dim in src_sheet.column_dimensions.items():
        col_widths[col_letter] = col_dim.width

    input_dir = os.path.dirname(os.path.abspath(input_path))
    base_name = os.path.splitext(os.path.basename(input_path))[0]

    # Step 4: One workbook per manager
    for manager in managers:
        out_wb = load_workbook(input_path)  # fresh copy for style fidelity
        # Remove all sheets except Summary, then rebuild it
        for sname in out_wb.sheetnames:
            if sname != "Summary":
                del out_wb[sname]

        out_ws = out_wb["Summary"]

        # Clear everything after the header block; we'll rewrite data rows
        max_row = out_ws.max_row
        for r in range(header_row_num + 1, max_row + 1):
            for c in range(1, out_ws.max_column + 1):
                out_ws.cell(row=r, column=c).value = None

        # Write only this manager's data rows
        manager_rows = [cells for cells in data_rows if str(cells[mgr_idx].value).strip() == manager]

        write_row = header_row_num + 1
        for cells in manager_rows:
            for col_0, src_cell in enumerate(cells):
                dst_cell = out_ws.cell(row=write_row, column=col_0 + 1)
                copy_cell(src_cell, dst_cell)
            write_row += 1

        # Restore column widths
        for col_letter, width in col_widths.items():
            out_ws.column_dimensions[col_letter].width = width

        # Step 5: Insert distributor subtotals
        data_start = header_row_num + 1
        num_distributors = insert_distributor_subtotals(
            out_ws, header_row_num, data_start, dist_idx, comm_idx
        )

        # Step 6: Apply column alignments
        apply_column_alignment(out_ws, label_map, header_row_num)

        out_filename = f"{manager}-{base_name}.xlsx"
        out_path = os.path.join(input_dir, out_filename)
        out_wb.save(out_path)
        pdf_path = convert_to_pdf(out_path)
        pdf_status = f", PDF saved" if pdf_path else ", PDF: failed (LibreOffice not found?)"
        print(f"Saved {out_filename} — {len(manager_rows)} data rows, {num_distributors} distributors{pdf_status}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: python process_commissions.py <input.xlsx>")
    process(sys.argv[1])
