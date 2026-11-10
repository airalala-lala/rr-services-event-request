"""RR Services Event Request Excel export using the original template.

ONE worksheet for the whole request.
  * Sessions are grouped by event date. Sessions on the same date sit under one
    merged date cell; a new date starts a new block further down the same sheet.
  * Each block has at least the template's 3 menu rows and grows as needed, so
    every menu item is always on the sheet (no "part 2", no extra sheets).
"""

import math
from copy import copy
from io import BytesIO
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment

TEMPLATE = Path(__file__).resolve().parent / "event_request_template.xlsx"

# Template layout (rows of the original one-block template)
ROW_BAND, ROW_HEADER, ROW_ITEM_FIRST, ROW_ITEM_MID, ROW_ITEM_LAST = 14, 15, 16, 17, 18
ROW_TOTAL, ROW_REMARKS = 19, 20
BASE_ROWS = 3           # menu rows the template has
MENU_ROW_HEIGHT = 45
NCOLS = 12              # A..L
MONEY = '"RM" #,##0.00'


def put(ws, ref, value):
    """Write text typed by a customer without letting it become an Excel formula."""
    ws[ref] = value
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        ws[ref].data_type = "s"


def _copy_row_style(src, src_row, dst, dst_row):
    for c in range(1, NCOLS + 1):
        dst.cell(dst_row, c)._style = copy(src.cell(src_row, c)._style)
    dst.row_dimensions[dst_row].height = src.row_dimensions[src_row].height


def _replicate_merges(src, src_row, dst, dst_row):
    for rng in list(src.merged_cells.ranges):
        if rng.min_row == rng.max_row == src_row:
            dst.merge_cells(
                start_row=dst_row, end_row=dst_row,
                start_column=rng.min_col, end_column=rng.max_col,
            )


def _group_runs(keys):
    """[(start_index, end_index)] of consecutive equal keys."""
    runs, start = [], 0
    for i in range(1, len(keys) + 1):
        if i == len(keys) or keys[i] != keys[start]:
            runs.append((start, i - 1))
            start = i
    return runs


def _fmt_date(value):
    try:
        return date.fromisoformat(str(value)[:10]).strftime("%d %b %Y (%a)")
    except (ValueError, TypeError):
        return str(value or "")


def _blocks(sessions):
    """Group sessions by date (chronological) and order each block by time."""
    ordered = sorted(
        sessions,
        key=lambda s: (str(s.get("delivery_date") or "")[:10], str(s.get("delivery_time") or "")),
    )
    blocks, index = [], {}
    for s in ordered:
        key = str(s.get("delivery_date") or "")[:10]
        if key not in index:
            index[key] = len(blocks)
            blocks.append({"date": key, "sessions": []})
        blocks[index[key]]["sessions"].append(s)
    return blocks


def _emit_block(ws, pr, start, block):
    """Write one date block starting at row `start`. Returns (next_row, total_cell)."""
    sessions = block["sessions"]
    rows = [(s, it) for s in sessions for it in (s.get("items") or [])]
    n = max(len(rows), BASE_ROWS)

    # --- band: ORDERING TYPE - LUNCH / TEA
    _copy_row_style(pr, ROW_BAND, ws, start)
    _replicate_merges(pr, ROW_BAND, ws, start)
    labels = []
    for s in sessions:
        lab = str(s.get("session_label") or "EVENT").upper()[:30]
        if lab not in labels:
            labels.append(lab)
    ws.cell(start, 4).value = "ORDERING TYPE - " + " / ".join(labels)[:60]

    # --- column headings (copied from the template)
    hdr = start + 1
    _copy_row_style(pr, ROW_HEADER, ws, hdr)
    for c in range(1, NCOLS + 1):
        ws.cell(hdr, c).value = pr.cell(ROW_HEADER, c).value
    _replicate_merges(pr, ROW_HEADER, ws, hdr)

    # --- menu rows
    first = hdr + 1
    last = first + n - 1
    for i in range(n):
        r = first + i
        src = ROW_ITEM_FIRST if i == 0 else (ROW_ITEM_LAST if i == n - 1 else ROW_ITEM_MID)
        _copy_row_style(pr, src, ws, r)
        ws.row_dimensions[r].height = MENU_ROW_HEIGHT
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)   # E:H
        ws.cell(r, 5).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(r, 10).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 12).number_format = MONEY
        ws.cell(r, 12).alignment = Alignment(horizontal="right", vertical="center")

    for i, (s, item) in enumerate(rows):
        r = first + i
        put(ws, f"E{r}", item.get("menu_description") or "")
        # Only show verified billing codes.
        put(ws, f"I{r}", item.get("verified_billing_code") or "")

        qty = Decimal(str(item.get("actual_qty") or 0))
        agreed = item.get("agreed_unit_price")
        price = Decimal(str(agreed)) if agreed is not None else Decimal(str(item.get("requested_unit_price") or 0))
        ws[f"J{r}"] = float(price)
        ws[f"K{r}"] = float(qty)
        ws[f"L{r}"] = f"=J{r}*K{r}"

    # --- date (merged over the whole block), time and venue (merged per run)
    ws.merge_cells(start_row=first, end_row=last, start_column=1, end_column=1)
    put(ws, f"A{first}", _fmt_date(block["date"]))

    if rows:
        times = [str(s.get("delivery_time") or "")[:5] for s, _ in rows]
        venues = [str(s.get("venue") or "") for s, _ in rows]
    else:                                   # session with no menu lines yet
        times = [str(sessions[0].get("delivery_time") or "")[:5]]
        venues = [str(sessions[0].get("venue") or "")]
    times += [times[-1]] * (n - len(times))     # padding rows join the last run
    venues += [venues[-1]] * (n - len(venues))

    for a, b in _group_runs(times):
        ws.cell(first + a, 2)._style = copy(pr.cell(ROW_ITEM_FIRST, 2)._style)
        ws.merge_cells(start_row=first + a, end_row=first + b, start_column=2, end_column=2)
        put(ws, f"B{first + a}", times[a])
    for a, b in _group_runs(venues):
        ws.cell(first + a, 3)._style = copy(pr.cell(ROW_ITEM_FIRST, 3)._style)
        ws.merge_cells(start_row=first + a, end_row=first + b, start_column=3, end_column=4)
        put(ws, f"C{first + a}", venues[a])

    # --- block total
    tr = last + 1
    _copy_row_style(pr, ROW_TOTAL, ws, tr)
    _replicate_merges(pr, ROW_TOTAL, ws, tr)
    ws.cell(tr, 1).value = pr.cell(ROW_TOTAL, 1).value or "TOTAL AMOUNT"
    ws[f"L{tr}"] = f"=SUM(L{first}:L{last})"
    ws[f"L{tr}"].number_format = MONEY
    ws[f"L{tr}"].alignment = Alignment(horizontal="right", vertical="center")

    return tr + 1, f"L{tr}"


def excel_bytes(request, sessions):
    """Generate an Excel workbook (one sheet) from one event request.

    Billing codes are not automatically inferred.
    """

    wb = load_workbook(TEMPLATE)
    base = wb.worksheets[0]
    base.title = "_template_base"
    # Clean copy used only as a style source, so nothing leaks between blocks.
    pristine = wb.copy_worksheet(base)
    pristine.title = "_template_clean"

    ws = wb.copy_worksheet(pristine)
    ws.title = "Event Request"

    # Remove the template's single block (rows 14+); blocks are rebuilt below.
    for rng in [r for r in list(ws.merged_cells.ranges) if r.min_row >= ROW_BAND]:
        ws.unmerge_cells(str(rng))
    if ws.max_row >= ROW_BAND:
        ws.delete_rows(ROW_BAND, ws.max_row - ROW_BAND + 1)
    for r in [r for r in list(ws.row_dimensions.keys()) if r >= ROW_BAND]:
        del ws.row_dimensions[r]

    # Widen columns to prevent ####### for currency values.
    for col, width in {"A": 20, "B": 15, "I": 14, "J": 16, "K": 13, "L": 19}.items():
        ws.column_dimensions[col].width = width

    # Part 1: Requestor details.
    put(ws, "D4", request.get("quotation_reference") or "")

    created = request.get("created_at") or date.today().isoformat()
    ws["D6"] = str(created)[:10]

    put(ws, "D7", request.get("requestor") or "")
    put(ws, "D8", request.get("email") or "")
    ws["D8"].alignment = Alignment(horizontal="left", vertical="center", indent=0)   # keep email left
    put(ws, "D9", request.get("phone") or "")
    put(ws, "D10", request.get("department") or "")
    put(ws, "D11", request.get("event_name") or "")

    # Part 2: delivery blocks, one per date, all on this sheet.
    blocks = _blocks(sessions or []) or [{"date": "", "sessions": [{"session_label": "Event", "items": []}]}]
    row, totals = ROW_BAND, []
    for i, block in enumerate(blocks):
        row, total_cell = _emit_block(ws, pristine, row, block)
        totals.append(total_cell)
        if i < len(blocks) - 1:
            ws.row_dimensions[row].height = 12      # spacer row between dates
            row += 1

    if len(totals) > 1:                       # several dates: show an overall total too
        _copy_row_style(pristine, ROW_TOTAL, ws, row)
        _replicate_merges(pristine, ROW_TOTAL, ws, row)
        ws.cell(row, 1).value = "GRAND TOTAL (ALL DATES)"
        ws[f"L{row}"] = "=" + "+".join(totals)
        ws[f"L{row}"].number_format = MONEY
        ws[f"L{row}"].alignment = Alignment(horizontal="right", vertical="center")
        row += 1

    # Remarks (once, at the bottom).
    remarks = request.get("notes") or ""
    item_remarks = "; ".join(
        str(item.get("remarks") or "")
        for s in (sessions or [])
        for item in (s.get("items") or [])
        if item.get("remarks")
    )
    if item_remarks:
        remarks = (remarks + " | " if remarks else "") + item_remarks

    _copy_row_style(pristine, ROW_REMARKS, ws, row)
    _replicate_merges(pristine, ROW_REMARKS, ws, row)
    put(ws, f"A{row}", "Remarks : " + remarks if remarks else "Remarks : ")
    ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
    lines = sum(max(1, math.ceil(len(part) / 120)) for part in ("Remarks : " + remarks).split("\n"))
    ws.row_dimensions[row].height = max(18, 18 * lines)

    wb.remove(base)
    wb.remove(pristine)

    output = BytesIO()
    wb.save(output)

    return output.getvalue()
