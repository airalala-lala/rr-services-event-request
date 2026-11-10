"""RR Services Event Request Excel export using the original template.

One worksheet per delivery session. The template has room for 3 menu lines;
when a session has more, extra rows are inserted so EVERY item stays on the
same sheet (no more "Session 1 part 2").
"""

from copy import copy
from io import BytesIO
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment

TEMPLATE = Path(__file__).resolve().parent / "event_request_template.xlsx"

FIRST_ROW = 16          # first menu row in the template
BASE_ROWS = 3           # menu rows the template has (16, 17, 18)
MENU_ROW_HEIGHT = 45


def put(ws, ref, value):
    """Write text typed by a customer without letting it become an Excel formula."""
    ws[ref] = value
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        ws[ref].data_type = "s"


def expand_menu_rows(ws, n_items):
    """Insert extra menu rows (copying the look of row 18) when a session has
    more than BASE_ROWS items. Moves the total and remarks rows down and keeps
    the merged cells (date / time / venue block and the menu text) intact."""
    extra = max(0, n_items - BASE_ROWS)
    if not extra:
        return 0

    last = FIRST_ROW + BASE_ROWS - 1                      # 18
    ranges = list(ws.merged_cells.ranges)
    below = [r for r in ranges if r.min_row > last]       # total + remarks rows
    spanning = [r for r in ranges if r.min_row <= FIRST_ROW and r.max_row == last]

    for r in below + spanning:
        ws.unmerge_cells(str(r))

    heights = {
        row: dim.height
        for row, dim in list(ws.row_dimensions.items())
        if row > last and dim.height
    }
    for row in heights:
        ws.row_dimensions[row].height = None

    ws.insert_rows(last + 1, extra)

    for row, h in heights.items():
        ws.row_dimensions[row + extra].height = h

    for r in range(last + 1, last + extra + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c)._style = copy(ws.cell(last, c)._style)

    for rng in below:
        ws.merge_cells(
            start_row=rng.min_row + extra, end_row=rng.max_row + extra,
            start_column=rng.min_col, end_column=rng.max_col,
        )
    for rng in spanning:      # date, time and venue now cover every item row
        ws.merge_cells(
            start_row=rng.min_row, end_row=last + extra,
            start_column=rng.min_col, end_column=rng.max_col,
        )
    for r in range(last + 1, last + extra + 1):
        ws.merge_cells(f"E{r}:H{r}")

    return extra


def excel_bytes(request, sessions):
    """Generate an Excel workbook from one event request.

    Multiple delivery sessions are retained in one workbook, one sheet per
    session, with all of its menu items on that sheet.
    Billing codes are not automatically inferred.
    """

    wb = load_workbook(TEMPLATE)
    base = wb.worksheets[0]
    base.title = "_template_base"
    # Every sheet is copied from this CLEAN copy, so rows never leak between sessions.
    pristine = wb.copy_worksheet(base)
    pristine.title = "_template_clean"

    if not sessions:
        sessions = [{"session_label": "Event", "items": []}]

    for sn, session in enumerate(sessions, 1):
        items = session.get("items") or []

        ws = wb.copy_worksheet(pristine)
        ws.title = f"Session {sn}"[:31]

        extra = expand_menu_rows(ws, len(items))
        menu_rows = range(FIRST_ROW, FIRST_ROW + max(len(items), BASE_ROWS))
        first_item_row, last_item_row = FIRST_ROW, FIRST_ROW + max(len(items), BASE_ROWS) - 1
        total_row = 19 + extra
        remarks_row = 20 + extra

        # Widen columns to prevent ####### for currency values.
        widths = {"A": 20, "B": 15, "I": 14, "J": 16, "K": 13, "L": 19}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        # Taller menu rows so long descriptions are not squeezed.
        for row in menu_rows:
            ws.row_dimensions[row].height = MENU_ROW_HEIGHT
            ws[f"E{row}"].alignment = Alignment(wrap_text=True, vertical="center")
            ws[f"J{row}"].alignment = Alignment(horizontal="center", vertical="center")

        # Format monetary values.
        for row in list(menu_rows) + [total_row]:
            cell = ws[f"L{row}"]
            cell.number_format = '"RM" #,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")

        # Part 1: Requestor details.
        put(ws, "D4", request.get("quotation_reference") or "")

        created = request.get("created_at") or date.today().isoformat()
        ws["D6"] = str(created)[:10]

        put(ws, "D7", request.get("requestor") or "")
        put(ws, "D8", request.get("email") or "")

        # Keep email aligned to the left.
        ws["D8"].alignment = Alignment(horizontal="left", vertical="center", indent=0)

        put(ws, "D9", request.get("phone") or "")
        put(ws, "D10", request.get("department") or "")
        put(ws, "D11", request.get("event_name") or "")

        # Part 2: Delivery session.
        label = str(session.get("session_label") or "EVENT")
        ws["D14"] = "ORDERING TYPE - " + label.upper()[:60]

        delivery_date = session.get("delivery_date") or ""
        try:
            delivery_date = date.fromisoformat(
                str(delivery_date)[:10]
            ).strftime("%d %b %Y (%a)")
        except (ValueError, TypeError):
            delivery_date = str(delivery_date)

        put(ws, "A16", delivery_date)
        put(ws, "B16", str(session.get("delivery_time") or "")[:5])
        put(ws, "C16", session.get("venue") or "")

        # Menu items: all of them, on this one sheet.
        for idx, item in enumerate(items):
            row = FIRST_ROW + idx

            put(ws, f"E{row}", item.get("menu_description") or "")

            # Only show verified billing codes.
            put(ws, f"I{row}", item.get("verified_billing_code") or "")

            qty = Decimal(str(item.get("actual_qty") or 0))

            agreed_price = item.get("agreed_unit_price")
            if agreed_price is not None:
                price = Decimal(str(agreed_price))
            else:
                price = Decimal(str(item.get("requested_unit_price") or 0))

            ws[f"J{row}"] = float(price)
            ws[f"K{row}"] = float(qty)

            # Calculate amount.
            ws[f"L{row}"] = f"=J{row}*K{row}"

        # Worksheet total (covers every menu row).
        ws[f"L{total_row}"] = f"=SUM(L{first_item_row}:L{last_item_row})"

        # Remarks.
        remarks = request.get("notes") or ""

        item_remarks = "; ".join(
            str(item.get("remarks") or "")
            for item in items
            if item.get("remarks")
        )
        if item_remarks:
            remarks = (remarks + " | " if remarks else "") + item_remarks

        put(ws, f"A{remarks_row}", remarks)
        ws[f"A{remarks_row}"].alignment = Alignment(wrap_text=True, vertical="top")

    wb.remove(base)
    wb.remove(pristine)

    output = BytesIO()
    wb.save(output)

    return output.getvalue()
