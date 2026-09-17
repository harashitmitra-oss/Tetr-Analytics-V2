from __future__ import annotations

from datetime import date
import re
import pandas as pd
import streamlit as st

from tetr_data import GoogleStore, CYCLE, clean, normalize_batch, batch_sheet_name
from attendance_parsers import read_uploaded_table, standardize_attendance, detect_format


st.set_page_config(page_title="Tetr Community 2026-27 Admin", layout="wide")

GREEN = "#0b3d2e"
GREEN2 = "#1f7a56"

st.markdown(
    """
    <style>
      .stApp {background: linear-gradient(180deg,#ffffff 0%,#f5fbf7 100%);}
      div[data-testid="stMetric"] {
        background:#fff;border:1px solid #dbeee0;border-radius:16px;padding:10px 12px;
      }
      h1,h2,h3 {color:#12372a !important;}
      .admin-card {
        background:#fff;border:1px solid #dbeee0;border-radius:18px;padding:16px 18px;
        box-shadow:0 4px 14px rgba(11,61,46,.05);margin-bottom:12px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def get_store():
    if "GOOGLE_SERVICE_ACCOUNT" not in st.secrets:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT is missing from Streamlit secrets.")
    sid = st.secrets.get("COMMUNITY_MASTER_SPREADSHEET_ID", "")
    if not sid:
        raise RuntimeError("COMMUNITY_MASTER_SPREADSHEET_ID is missing from Streamlit secrets.")
    admin_name = st.session_state.get("admin_name") or st.secrets.get("ADMIN_DEFAULT_NAME", "Community Admin")
    return GoogleStore(dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"]), sid, admin_name)


@st.cache_resource(show_spinner=False)
def cached_store(service_key: str, spreadsheet_id: str, admin_name: str):
    # Resource cache is keyed by non-secret identifiers; credentials are fetched above.
    return get_store()


def refresh():
    st.cache_data.clear()
    st.rerun()


def admin_login():
    expected = st.secrets.get("ADMIN_PASSWORD", "")
    if not expected:
        st.warning("ADMIN_PASSWORD is not configured. Add it before production deployment.")
        st.session_state["admin_ok"] = True
        return True
    if st.session_state.get("admin_ok"):
        return True
    st.title("Tetr Community Admin · 2026–27")
    pwd = st.text_input("Admin password", type="password")
    name = st.text_input("Your name", value=st.secrets.get("ADMIN_DEFAULT_NAME", ""))
    if st.button("Unlock Admin"):
        if pwd == expected:
            st.session_state["admin_ok"] = True
            st.session_state["admin_name"] = name or "Community Admin"
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def overview(store: GoogleStore):
    st.title("Community Operations · 2026–27")
    st.caption("UG and PG restart from Batch 1. Admissions/payment status is mirrored from the external admissions trackers and is not editable here.")

    students = store.master_df()
    batches = store.list_batches()
    act_ws = store.ensure_tab("Activity_Master", [
        "Activity ID","Cycle","Activity Name","Activity Date","Activity Type","Program","Batches","Source","Created At","Created By"
    ])
    activities = pd.DataFrame(act_ws.get_all_records())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Students", len(students))
    c2.metric("UG", int((students.get("Program", pd.Series(dtype=str)) == "UG").sum()) if not students.empty else 0)
    c3.metric("PG", int((students.get("Program", pd.Series(dtype=str)) == "PG").sum()) if not students.empty else 0)
    c4.metric("Activities", len(activities))

    if not students.empty:
        st.subheader("Current students")
        cols = [c for c in ["Student ID","Student Name","Email","Program","Batch","Community Status","Admissions Status","Payment Date","Offer Date","Deadline"] if c in students.columns]
        st.dataframe(students[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No students have been added yet.")


def setup_schema(store):
    st.title("Initialise 2026–27")
    st.write("Creates the new-cycle core/admin sheets plus **UG B1** and **PG B1**. Safe to run again; existing sheets are retained.")
    if st.button("Create / verify 2026–27 structure", type="primary"):
        store.ensure_schema()
        st.success("2026–27 structure is ready.")


def students_page(store):
    st.title("Students")
    tab1, tab2, tab3 = st.tabs(["Add student", "Bulk upload", "Edit community-owned fields"])

    with tab1:
        with st.form("add_student"):
            c1, c2, c3 = st.columns(3)
            program = c1.selectbox("Program", ["UG", "PG", "GY"])
            batch = c2.text_input("Batch", "B1")
            community = c3.selectbox("Community Status", ["", "In", "Tetr X", "Added to Term 0", "Left", "Out"])
            name = st.text_input("Student Name *")
            c1, c2, c3 = st.columns(3)
            email = c1.text_input("Email")
            phone = c2.text_input("Mobile")
            country = c3.text_input("Country")
            c1, c2, c3 = st.columns(3)
            income = c1.text_input("Income")
            counsellor = c2.text_input("Counsellor Name")
            offer = c3.date_input("Offer Date", value=None)
            deadline = st.date_input("Deadline", value=None)
            submitted = st.form_submit_button("Add student", type="primary")
        if submitted:
            try:
                sid = store.add_student(
                    program=program, name=name, email=email, phone=phone, country=country,
                    income=income, batch=batch, community_status=community, counsellor=counsellor,
                    offer_date=offer.isoformat() if offer else "",
                    deadline=deadline.isoformat() if deadline else "",
                )
                st.success(f"Student added: {sid}")
            except Exception as e:
                st.error(str(e))

    with tab2:
        f = st.file_uploader("Student file (.xlsx/.csv)", type=["xlsx","xls","csv"], key="bulk_students")
        c1, c2 = st.columns(2)
        default_program = c1.selectbox("Default Program", ["UG","PG","GY"], key="bulk_program")
        default_batch = c2.text_input("Default Batch", "B1", key="bulk_batch")
        if f:
            try:
                df = read_uploaded_table(f)
                st.dataframe(df.head(25), use_container_width=True)
                if st.button("Validate and add students", type="primary"):
                    result = store.bulk_add_students(df, default_program=default_program, default_batch=default_batch)
                    st.dataframe(result, use_container_width=True, hide_index=True)
                    st.success(f"Added {(result['Result'] == 'Added').sum()} students; skipped {(result['Result'] == 'Skipped').sum()}.")
            except Exception as e:
                st.error(str(e))

    with tab3:
        df = store.master_df()
        if df.empty:
            st.info("No students available.")
        else:
            options = {f"{r['Student Name']} · {r['Email']} · {r['Program']} · {r['Batch']}": r for _, r in df.iterrows()}
            label = st.selectbox("Student", list(options))
            r = options[label]
            st.caption(f"Admissions Status: **{r.get('Admissions Status','')}** · Payment Date: **{r.get('Payment Date','')}** (read-only)")
            field = st.selectbox("Field to edit", ["Community Status", "Batch", "Counsellor Name", "Offer Date", "Deadline", "Country", "Income"])
            new = st.text_input("New value", value=clean(r.get(field, "")))
            if st.button("Save change"):
                try:
                    if field == "Batch":
                        new = normalize_batch(new)
                        store.create_batch(r["Program"], new, log_if_exists=False)
                    store.update_student_field(r["Student ID"], r["Program"], field, new)
                    st.success("Updated.")
                except Exception as e:
                    st.error(str(e))


def batches_page(store):
    st.title("Batch Management")
    c1, c2 = st.columns(2)
    program = c1.selectbox("Program", ["UG","PG","GY"], key="batch_program")
    batch = c2.text_input("New batch", "B2")
    if st.button("Create batch", type="primary"):
        try:
            ws, created = store.create_batch(program, batch)
            st.success(f"{ws.title} {'created' if created else 'already exists'}.")
        except Exception as e:
            st.error(str(e))
    st.subheader("Batch register")
    st.dataframe(store.list_batches(), use_container_width=True, hide_index=True)


def activities_page(store):
    st.title("Activities")
    st.caption("Create the activity first. The app writes its metadata into row 1/2/3 and attendance into the batch activity column, preserving the existing dashboard parser structure.")
    c1, c2, c3 = st.columns(3)
    program = c1.selectbox("Program", ["UG","PG","GY"])
    name = c2.text_input("Activity Name")
    activity_date = c3.date_input("Activity Date", value=date.today())
    activity_type = st.selectbox("Activity Type", [
        "Online Event", "Masterclass", "Competition", "Hackathon",
        "General", "Fun", "Fun Task", "Poll", "Quiz", "Other"
    ])
    bdf = store.list_batches(program)
    batches = bdf["Batch"].astype(str).tolist() if not bdf.empty else ["B1"]
    selected = st.multiselect("Applicable batches", batches, default=batches[:1])
    if st.button("Create activity", type="primary"):
        try:
            aid = store.create_activity(
                name=name, activity_date=activity_date.isoformat(), activity_type=activity_type,
                program=program, batches=selected
            )
            st.success(f"Created {aid}.")
        except Exception as e:
            st.error(str(e))

    aws = store.ensure_tab("Activity_Master", [
        "Activity ID","Cycle","Activity Name","Activity Date","Activity Type","Program","Batches","Source","Created At","Created By"
    ])
    adf = pd.DataFrame(aws.get_all_records())
    if not adf.empty:
        st.subheader("Activity register")
        st.dataframe(adf, use_container_width=True, hide_index=True)


def attendance_page(store):
    st.title("Attendance Upload")
    st.caption("Supports Zoom automatically and flexible Dinero attendee exports. Matching order: **Email → exact full name → phone last 8 digits**.")

    aws = store.ensure_tab("Activity_Master", [
        "Activity ID","Cycle","Activity Name","Activity Date","Activity Type","Program","Batches","Source","Created At","Created By"
    ])
    adf = pd.DataFrame(aws.get_all_records())
    if adf.empty:
        st.info("Create an activity first.")
        return

    adf["_label"] = adf.apply(lambda r: f"{r.get('Activity Date','')} · {r.get('Activity Name','')} · {r.get('Program','')} · {r.get('Activity ID','')}", axis=1)
    chosen = st.selectbox("Activity", adf["_label"].tolist())
    ar = adf[adf["_label"].eq(chosen)].iloc[0]

    f = st.file_uploader("Zoom / Dinero attendance file", type=["xlsx","xls","csv"])
    if not f:
        return

    try:
        raw = read_uploaded_table(f)
        detected = detect_format(raw)
        st.info(f"Detected format: **{detected}** · {len(raw):,} source rows")
        standardized, mapping = standardize_attendance(raw)
        st.write("Detected column mapping:", mapping)
        st.dataframe(standardized.head(30), use_container_width=True, hide_index=True)

        matched = store.match_attendance(standardized)
        counts = matched["Match Status"].value_counts().to_dict()
        c1, c2, c3 = st.columns(3)
        c1.metric("Unique attendees", len(standardized))
        c2.metric("Matched", counts.get("Matched", 0))
        c3.metric("Unmatched", counts.get("Unmatched", 0))

        st.subheader("Matching preview")
        show_cols = [c for c in [
            "Uploaded Name","Uploaded Email","Uploaded Phone","Match Status","Matched By",
            "Student Name","Program","Batch","Reason"
        ] if c in matched.columns]
        st.dataframe(matched[show_cols], use_container_width=True, hide_index=True)

        st.session_state["attendance_preview"] = matched
        st.session_state["attendance_mapping"] = mapping
        st.session_state["attendance_filename"] = f.name

        if st.button("Confirm & write attendance", type="primary"):
            result = store.commit_attendance(
                matched,
                activity_id=clean(ar["Activity ID"]),
                activity_name=clean(ar["Activity Name"]),
                activity_date=clean(ar["Activity Date"]),
                activity_type=clean(ar["Activity Type"]),
                file_name=f.name,
                detected_format=detected,
                input_rows=mapping["input_rows"],
                unique_attendees=mapping["unique_attendees"],
            )
            st.success(
                f"Committed {result['committed']} attendance records. "
                f"{result['unmatched']} unmatched rows were saved for review. Upload ID: {result['upload_id']}"
            )
    except Exception as e:
        st.error(str(e))


def admissions_page(store):
    st.title("Admissions Status Sync")
    st.warning("These statuses are **read-only in this app**. The external UG/PG/GY tracker remains the source of truth.")

    config = {
        "UG": (st.secrets.get("UG_ADMISSIONS_SPREADSHEET_ID",""), st.secrets.get("UG_ADMISSIONS_SHEET_NAME","")),
        "PG": (st.secrets.get("PG_ADMISSIONS_SPREADSHEET_ID",""), st.secrets.get("PG_ADMISSIONS_SHEET_NAME","")),
        "GY": (st.secrets.get("GY_ADMISSIONS_SPREADSHEET_ID",""), st.secrets.get("GY_ADMISSIONS_SHEET_NAME","")),
    }
    for program, (sid, sname) in config.items():
        with st.expander(f"{program} tracker", expanded=(program=="UG")):
            st.write("Spreadsheet configured:", "Yes" if sid else "No")
            st.write("Sheet:", sname or "Not configured")
            if st.button(f"Sync {program}", key=f"sync_{program}", disabled=not (sid and sname)):
                try:
                    result = store.sync_admissions_source(program=program, spreadsheet_id=sid, sheet_name=sname)
                    st.success(f"{program}: {result['matched']} matched from {result['source_rows']} tracker rows.")
                except Exception as e:
                    st.error(str(e))

    cws = store.ensure_tab("Admissions_Status_Cache", [
        "Student ID","Program","Student Name","Email","Mobile","Admissions Status",
        "Payment Date","Source Spreadsheet","Source Sheet","Matched By","Synced At"
    ])
    cdf = pd.DataFrame(cws.get_all_records())
    if not cdf.empty:
        st.subheader("Latest synced status cache")
        st.dataframe(cdf, use_container_width=True, hide_index=True)


def audit_page(store):
    st.title("Audit & Exceptions")
    tabs = st.tabs(["Audit Log", "Upload Log", "Unmatched Attendance"])
    for tab, title in zip(tabs, ["Admin_Audit_Log","Upload_Log","Unmatched_Attendance"]):
        with tab:
            ws = store.worksheet_or_create(title)
            df = pd.DataFrame(ws.get_all_records())
            if df.empty:
                st.info("No records yet.")
            else:
                st.dataframe(df.iloc[::-1].reset_index(drop=True), use_container_width=True, hide_index=True)


def connection_page(store):
    st.title("Connection Health")
    st.success(f"Connected to master workbook: {store.book.title}")
    st.write("Service account:", st.secrets["GOOGLE_SERVICE_ACCOUNT"].get("client_email",""))
    st.write("Workbook ID configured:", bool(st.secrets.get("COMMUNITY_MASTER_SPREADSHEET_ID","")))
    st.write("Sheets visible:", len(store.list_sheet_names()))
    st.dataframe(pd.DataFrame({"Sheet": store.list_sheet_names()}), use_container_width=True, hide_index=True)


def main():
    if not admin_login():
        return

    try:
        store = get_store()
    except Exception as e:
        st.error(f"Connection failed: {e}")
        st.stop()

    with st.sidebar:
        st.markdown("## Tetr Community")
        st.caption("2026–27 · Admin & Ingestion")
        page = st.radio(
            "Section",
            ["Overview","Initialise 2026–27","Students","Batches","Activities",
             "Attendance Upload","Admissions Sync","Audit & Exceptions","Connection Health"],
            label_visibility="collapsed",
        )
        st.divider()
        st.caption("Admissions/payment status is external and read-only here.")
        if st.button("Lock admin"):
            st.session_state["admin_ok"] = False
            st.rerun()

    if page == "Overview":
        overview(store)
    elif page == "Initialise 2026–27":
        setup_schema(store)
    elif page == "Students":
        students_page(store)
    elif page == "Batches":
        batches_page(store)
    elif page == "Activities":
        activities_page(store)
    elif page == "Attendance Upload":
        attendance_page(store)
    elif page == "Admissions Sync":
        admissions_page(store)
    elif page == "Audit & Exceptions":
        audit_page(store)
    elif page == "Connection Health":
        connection_page(store)


if __name__ == "__main__":
    main()
