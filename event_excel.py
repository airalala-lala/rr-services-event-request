"""RR Services Event Request Excel export. Uses the actual supplied Excel template."""
from io import BytesIO
from datetime import date, datetime
from decimal import Decimal
from openpyxl import load_workbook
from openpyxl.styles import Alignment

TEMPLATE = __file__.rsplit('/', 1)[0] + '/event_request_template.xlsx'

def excel_bytes(request, sessions):
    """One workbook per parent request, one or more sheets per session.
    Each template sheet has 3 preformatted menu rows; overflow creates continuation sheets.
    Billing codes are never inferred from food prices.
    """
    wb = load_workbook(TEMPLATE)
    base = wb.worksheets[0]
    if not sessions:
        sessions = [{'session_label': 'Event', 'items': []}]
    first = True
    for sn, session in enumerate(sessions, 1):
        items = session.get('items') or []
        chunks = [items[i:i+3] for i in range(0, len(items), 3)] or [[]]
        for part, chunk in enumerate(chunks, 1):
            ws = base if first else wb.copy_worksheet(base)
            first = False
            ws.title = (f'Session {sn}' + (f' part {part}' if len(chunks)>1 else ''))[:31]
            ws['D4'] = request.get('quotation_reference') or ''
            created = request.get('created_at') or date.today().isoformat()
            ws['D6'] = str(created)[:10]
            ws['D7'] = request.get('requestor') or ''
            ws['D8'] = request.get('email') or ''
            ws['D9'] = request.get('phone') or ''
            ws['D10'] = request.get('department') or ''
            ws['D11'] = request.get('event_name') or ''
            ws['D14'] = 'ORDERING TYPE - ' + str(session.get('session_label') or 'EVENT').upper()[:60]
            dt = session.get('delivery_date') or ''
            try:
                dt = date.fromisoformat(str(dt)[:10]).strftime('%d %b %Y (%a)')
            except (ValueError, TypeError):
                dt = str(dt)
            ws['A16'] = dt
            ws['B16'] = str(session.get('delivery_time') or '')[:5]
            ws['C16'] = session.get('venue') or ''
            total = Decimal('0')
            for idx, item in enumerate(chunk):
                row = 16 + idx
                ws[f'E{row}'] = item.get('menu_description') or ''
                # Code only if verified and explicitly supplied; customer submissions leave blank.
                ws[f'I{row}'] = item.get('verified_billing_code') or ''
                qty = Decimal(str(item.get('actual_qty') or 0))
                price = Decimal(str(item.get('agreed_unit_price') if item.get('agreed_unit_price') is not None else item.get('requested_unit_price') or 0))
                ws[f'J{row}'] = float(price)
                ws[f'K{row}'] = float(qty)
                ws[f'L{row}'] = f'=J{row}*K{row}'
                total += qty*price
            ws['L19'] = '=SUM(L16:L18)'
            ws['L19'].number_format = '#,##0.00'
            remarks = request.get('notes') or ''
            extra = '; '.join(str(i.get('remarks') or '') for i in chunk if i.get('remarks'))
            if extra:
                remarks = (remarks + ' | ' if remarks else '') + extra
            if len(chunks) > 1:
                remarks = (remarks + ' | ' if remarks else '') + f'Session {sn}, page {part}/{len(chunks)}. Session total is across all its pages.'
            ws['A20'] = remarks
            ws['A20'].alignment = Alignment(wrap_text=True, vertical='top')
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
