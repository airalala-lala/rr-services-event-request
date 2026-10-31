"""Deploy as a SEPARATE public Streamlit app. Secrets: SUPABASE_URL, SUPABASE_ANON_KEY.
Never add the service-role key to this public app.
"""
from datetime import date, time
import streamlit as st
from supabase import create_client
from event_excel import excel_bytes

st.set_page_config(page_title="RR Services | Event Request",page_icon="🍽️",layout="centered")
st.title("RR SERVICES")
st.subheader("Event Request Form")
st.caption("Please provide your event and food delivery requirements. RR Services will review your request.")
if not st.secrets.get("SUPABASE_URL") or not st.secrets.get("SUPABASE_ANON_KEY"):
    st.error("App not configured. Please contact RR Services."); st.stop()
@st.cache_resource
def client(url,key): return create_client(url,key)
db=client(st.secrets["SUPABASE_URL"],st.secrets["SUPABASE_ANON_KEY"])

with st.form("event_form",clear_on_submit=False):
    st.markdown("### Part 1 — Requestor Details")
    requester=st.text_input("Requestor Full Name *",max_chars=150)
    email=st.text_input("Email Address *",max_chars=200)
    phone=st.text_input("Contact Number",max_chars=80)
    dept=st.text_input("Dept / Cost Centre",max_chars=200)
    title=st.text_input("Event Title *",max_chars=200)
    po_hint=st.text_input("PO Number / Blanket PO (if known)",help="Optional. RR Services will verify the PO.")
    st.markdown("### Part 2 — Delivery Sessions")
    count=st.number_input("Number of delivery sessions",min_value=1,max_value=12,value=1,step=1)
    sessions=[]
    for s in range(int(count)):
        with st.expander(f"Session {s+1}",expanded=s==0):
            label=st.text_input("Meal / Session",value="Lunch" if s==0 else "Tea",key=f"label{s}")
            c1,c2=st.columns(2)
            day=c1.date_input("Event Date *",value=date.today(),key=f"date{s}")
            tm=c2.time_input("Time (Food Ready Time) *",value=time(12,0),key=f"time{s}")
            venue=st.text_input("Venue / Location *",key=f"venue{s}")
            n=st.number_input("Number of menu lines",min_value=1,max_value=30,value=1,key=f"count{s}")
            items=[]
            for j in range(int(n)):
                st.markdown(f"**Menu item {j+1}**")
                desc=st.text_input("Description / Menu *",key=f"desc{s}_{j}",max_chars=500)
                a,b=st.columns(2)
                qty=a.number_input("Quantity / Pax *",min_value=1.0,value=10.0,step=1.0,key=f"qty{s}_{j}")
                price=b.number_input("Requested Unit Price (RM) *",min_value=0.0,value=0.0,step=0.5,format="%.2f",key=f"price{s}_{j}")
                remarks=st.text_input("Serving details / Remarks",key=f"remarks{s}_{j}")
                items.append(dict(menu_description=desc.strip(),actual_qty=qty,requested_unit_price=price,remarks=remarks))
            sessions.append(dict(session_label=label,delivery_date=day.isoformat(),delivery_time=tm.strftime("%H:%M"),venue=venue.strip(),items=items))
    notes=st.text_area("Additional Instructions")
    st.caption("Billing codes are assigned internally by RR Services. Submission does not confirm your booking.")
    confirmed=st.checkbox("I confirm the information above is accurate.")
    submit=st.form_submit_button("Submit Event Request",type="primary")
if submit:
    errors=[]
    if len(requester.strip())<2 or "@" not in email or "." not in email.split("@")[-1] or len(title.strip())<2: errors.append("Complete your name, valid email and event title.")
    if not confirmed: errors.append("Please confirm the information.")
    for k,s in enumerate(sessions,1):
        if len(s["venue"])<2 or any(len(i["menu_description"])<2 for i in s["items"]): errors.append(f"Session {k}: venue and menu descriptions are required.")
    if errors:
        for err in errors: st.error(err)
    else:
        payload=dict(requestor=requester.strip(),email=email.strip(),phone=phone,department=dept,event_name=title.strip(),po_hint=po_hint,notes=notes,sessions=sessions)
        try:
            res=db.rpc("rr_submit_event_request",{"payload":payload}).execute()
            st.success(f"Request submitted! Reference: {res.data}")
            try:
                file_data = excel_bytes(dict(requestor=requester.strip(),email=email.strip(),phone=phone,department=dept,event_name=title.strip(),notes=notes), sessions)
                st.download_button("📥 Download Event Request Excel", file_data, file_name="RR_Event_Request_" + date.today().strftime("%Y%m%d") + ".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", on_click="ignore")
            except Exception:
                st.warning("Request submitted, but Excel download could not be prepared. Please contact RR Services.")
            st.info("Status: Pending Review. RR Services will confirm your event after reviewing the request.")
        except Exception:
            st.error("Submission failed. Please contact RR Services and try again. Do not repeatedly resubmit if you already received a reference.")
