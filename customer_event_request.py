"""RR Services public Event Request Form.

Streamlit secrets: SUPABASE_URL and SUPABASE_ANON_KEY.
Never put a service-role key in this public app.
"""
from datetime import date, time
import streamlit as st
from supabase import create_client
from event_excel import excel_bytes

st.set_page_config(page_title="RR Services | Event Request", page_icon="🍽️", layout="centered")
st.title("RR SERVICES")
st.subheader("Event Request Form")
st.caption("Please provide your event and food delivery requirements. RR Services will review your request.")

if not st.secrets.get("SUPABASE_URL") or not st.secrets.get("SUPABASE_ANON_KEY"):
    st.error("App not configured. Please contact RR Services.")
    st.stop()

@st.cache_resource
def client(url, key):
    return create_client(url, key)

db = client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_ANON_KEY"])

if "submitted_request" not in st.session_state:
    st.session_state.submitted_request = None

if st.session_state.submitted_request is not None:
    saved = st.session_state.submitted_request
    st.success(f"Request submitted! Reference: {saved['reference']}")
    st.info("Status: Pending Review. RR Services will confirm your event after reviewing the request.")
    if saved.get("excel"):
        st.download_button(
            "📥 Download Event Request Excel",
            data=saved["excel"],
            file_name=saved["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Request submitted, but the Excel file could not be generated.")
    if st.button("Start a New Request"):
        st.session_state.submitted_request = None
        st.rerun()
    st.stop()

with st.container():   # was st.form: a form does not refresh until Submit is pressed
    st.markdown("### Part 1 — Requestor Details")
    requester = st.text_input("Requestor Full Name *", max_chars=150)
    email = st.text_input("Email Address *", max_chars=200)
    phone = st.text_input("Contact Number", max_chars=80)
    dept = st.text_input("Dept / Cost Centre", max_chars=200)
    title = st.text_input("Event Title *", max_chars=200)
    po_hint = st.text_input("PO Number / Blanket PO (if known)", help="Optional. RR Services will verify the PO.")

    st.markdown("### Part 2 — Delivery Sessions")
    count = st.number_input("Number of delivery sessions", min_value=1, max_value=12, value=1, step=1)
    sessions = []
    serving_summaries = []

    for s in range(int(count)):
        with st.expander(f"Session {s + 1}", expanded=(s == 0)):
            label = st.text_input("Meal / Session", value="Lunch" if s == 0 else "Tea", key=f"label{s}")
            c1, c2 = st.columns(2)
            day = c1.date_input("Event Date *", value=date.today(), key=f"date{s}")
            tm = c2.time_input("Time (Food Ready Time) *", value=time(12, 0), key=f"time{s}")
            venue = st.text_input("Venue / Location *", key=f"venue{s}")

            serving_type = st.radio(
                "Serving Type *",
                options=["Pack", "Buffet"],
                horizontal=True,
                key=f"serving{s}",
                help="Pack = individually packed meals. Buffet = food served for sharing.",
            )
            st.caption("Pack: individually packed meals · Buffet: food served for sharing")
            serving_summaries.append(f"Session {s + 1} ({label}): Serving Type - {serving_type}")

            n = st.number_input("Number of menu lines", min_value=1, max_value=30, value=1, step=1, key=f"count{s}")
            items = []
            for j in range(int(n)):
                st.markdown(f"**Menu item {j + 1}**")
                desc = st.text_input("Description / Menu *", key=f"desc{s}_{j}", max_chars=500)
                a, b = st.columns(2)
                qty = a.number_input("Quantity / Pax *", min_value=1.0, value=10.0, step=1.0, key=f"qty{s}_{j}")
                price = b.number_input("Requested Unit Price (RM) *", min_value=0.0, value=0.0, step=0.5, format="%.2f", key=f"price{s}_{j}")
                remarks = st.text_input("Menu Item Remarks", placeholder="e.g. No chili, less sugar", key=f"remarks{s}_{j}")
                items.append({
                    "menu_description": desc.strip(),
                    "actual_qty": qty,
                    "requested_unit_price": price,
                    "remarks": remarks.strip(),
                })
            sessions.append({
                "session_label": label.strip(),
                "delivery_date": day.isoformat(),
                "delivery_time": tm.strftime("%H:%M"),
                "venue": venue.strip(),
                "items": items,
            })

    st.markdown("### Part 3 — Additional Instructions")
    notes = st.text_area("Additional Instructions", placeholder="e.g. Include cutlery, separate sauces, buffet setup required")
    st.caption("Billing codes are assigned internally by RR Services. Submission does not confirm your booking.")
    confirmed = st.checkbox("I confirm the information above is accurate.")
    submit = st.button("Submit Event Request", type="primary")

if submit:
    errors = []
    if len(requester.strip()) < 2 or "@" not in email or "." not in email.split("@")[-1] or len(title.strip()) < 2:
        errors.append("Complete your name, valid email and event title.")
    if not confirmed:
        errors.append("Please confirm the information.")
    for k, session in enumerate(sessions, 1):
        if len(session["venue"]) < 2 or not session["session_label"]:
            errors.append(f"Session {k}: meal/session and venue are required.")
        if any(len(item["menu_description"]) < 2 for item in session["items"]):
            errors.append(f"Session {k}: menu descriptions are required.")

    if errors:
        for error in errors:
            st.error(error)
    else:
        all_notes = "\n".join(serving_summaries)
        if notes.strip():
            all_notes += "\n\nAdditional Instructions:\n" + notes.strip()
        payload = {
            "requestor": requester.strip(),
            "email": email.strip(),
            "phone": phone.strip(),
            "department": dept.strip(),
            "event_name": title.strip(),
            "po_hint": po_hint.strip(),
            "notes": all_notes,
            "sessions": sessions,
        }
        try:
            response = db.rpc("rr_submit_event_request", {"payload": payload}).execute()
            reference = str(response.data)
        except Exception:
            st.error("Submission failed. Please contact RR Services and try again. Do not repeatedly resubmit if you already received a reference.")
        else:
            excel_data = None
            try:
                excel_data = excel_bytes(payload, sessions)
            except Exception:
                st.warning("Request submitted, but the Excel export could not be prepared.")
            st.session_state.submitted_request = {
                "reference": reference,
                "excel": excel_data,
                "filename": "RR_Event_Request_" + date.today().strftime("%Y%m%d") + ".xlsx",
            }
            st.rerun()
