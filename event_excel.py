"""RR Services Event Request Excel export using the original template."""

from io import BytesIO
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment

TEMPLATE = Path(__file__).resolve().parent / "event_request_template.xlsx"


def put(ws, ref, value):
    """Write text typed by a customer without letting it become an Excel formula."""
    ws[ref] = value
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        ws[ref].data_type = "s"


def excel_bytes(request, sessions):
    """Generate an Excel workbook from one event request.

    Multiple delivery sessions are retained in one workbook.
    Each sheet contains up to three menu items, matching the template.
    Billing codes are not automatically inferred.
    """

    wb = load_workbook(TEMPLATE)
    base = wb.worksheets[0]
    base.title = "_template_base"
    # Every sheet is copied from this CLEAN copy. Before, sheets were copied from a sheet
    # that had already been filled, so old menu rows leaked into later sessions.
    pristine = wb.copy_worksheet(base)
    pristine.title = "_template_clean"

    if not sessions:
        sessions = [{"session_label": "Event", "items": []}]


    for sn, session in enumerate(sessions, 1):
        items = session.get("items") or []
        chunks = [
            items[i:i + 3]
            for i in range(0, len(items), 3)
        ] or [[]]

        for part, chunk in enumerate(chunks, 1):
            ws = wb.copy_worksheet(pristine)

            ws.title = (
                f"Session {sn}"
                + (f" part {part}" if len(chunks) > 1 else "")
            )[:31]

            # Widen columns to prevent ####### for currency values.
            widths = {
                "A": 20,
                "B": 15,
                "I": 14,
                "J": 16,
                "K": 13,
                "L": 19,
            }

            for col, width in widths.items():
                ws.column_dimensions[col].width = width

            # Format monetary values.
            for row in (16, 17, 18, 19):
                cell = ws[f"L{row}"]
                cell.number_format = '"RM" #,##0.00'
                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            # Part 1: Requestor details.
            put(ws, "D4", request.get("quotation_reference") or "")

            created = request.get("created_at") or date.today().isoformat()
            ws["D6"] = str(created)[:10]

            put(ws, "D7", request.get("requestor") or "")
            put(ws, "D8", request.get("email") or "")

            # Keep email aligned to the left.
            ws["D8"].alignment = Alignment(
                horizontal="left",
                vertical="center",
                indent=0
            )

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

            # Menu items.
            for idx, item in enumerate(chunk):
                row = 16 + idx

                put(ws, f"E{row}", item.get("menu_description") or "")

                # Only show verified billing codes.
                put(ws, f"I{row}", item.get("verified_billing_code") or "")

                qty = Decimal(
                    str(item.get("actual_qty") or 0)
                )

                agreed_price = item.get("agreed_unit_price")

                if agreed_price is not None:
                    price = Decimal(str(agreed_price))
                else:
                    price = Decimal(
                        str(item.get("requested_unit_price") or 0)
                    )

                ws[f"J{row}"] = float(price)
                ws[f"K{row}"] = float(qty)

                # Calculate amount.
                ws[f"L{row}"] = f"=J{row}*K{row}"

            # Worksheet total.
            ws["L19"] = "=SUM(L16:L18)"
            ws["L19"].number_format = '"RM" #,##0.00'

            # Remarks.
            remarks = request.get("notes") or ""

            extra = "; ".join(
                str(item.get("remarks") or "")
                for item in chunk
                if item.get("remarks")
            )

            if extra:
                remarks = (
                    remarks + " | " if remarks else ""
                ) + extra

            if len(chunks) > 1:
                page_note = (
                    f"Session {sn}, page {part}/{len(chunks)}. "
                    "Session total is across all its pages."
                )
                remarks = (
                    remarks + " | " if remarks else ""
                ) + page_note

            put(ws, "A20", remarks)
            ws["A20"].alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )

    wb.remove(base)
    wb.remove(pristine)

    output = BytesIO()
    wb.save(output)

    return output.getvalue()
