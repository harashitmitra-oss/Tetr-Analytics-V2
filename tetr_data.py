from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import hashlib
import re
import uuid
from typing import Iterable

import pandas as pd
import gspread
from gspread.exceptions import WorksheetNotFound
from google.oauth2.service_account import Credentials


GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CYCLE = "2026-27"

MASTER_HEADERS = [
    "Student ID", "Student Name", "Email", "Mobile", "Country", "Income",
    "Batch", "Community Status", "Counsellor Name", "Offer Date", "Deadline",
    "Admissions Status", "Payment Date", "Admissions Source", "Last Synced",
]

DATES_HEADERS = [
    "Student ID", "Student Name", "Email", "Program", "Batch",
    "Offered Date", "Deadline", "Deadline Version", "Last Updated"
]

BATCH_BASE_HEADERS = [
    "Student ID", "Student Name", "Email", "Mobile", "Batch", "Country", "Income",
    "Counsellor Name", "Community Status", "Term Zero Group", "Admissions Status",
    "Payment Date", "Offer Date", "Deadline", "Overall Engagement Score",
    "Overall Engagement %", "Notes", "Source", "Last Updated"
]

BATCH_MASTER_HEADERS = [
    "Cycle", "Program", "Batch", "Sheet Name", "Created At", "Created By", "Active"
]
ACTIVITY_MASTER_HEADERS = [
    "Activity ID", "Cycle", "Activity Name", "Activity Date", "Activity Type",
    "Program", "Batches", "Source", "Created At", "Created By"
]
AUDIT_HEADERS = [
    "Timestamp", "Admin", "Action", "Entity", "Entity ID", "Program", "Batch",
    "Sheet", "Field", "Old Value", "New Value", "Details"
]
UPLOAD_LOG_HEADERS = [
    "Upload ID", "Timestamp", "Admin", "File Name", "Detected Format",
    "Activity ID", "Activity Name", "Input Rows", "Unique Attendees",
    "Matched", "Unmatched", "Committed"
]
UNMATCHED_HEADERS = [
    "Upload ID", "Activity ID", "Uploaded Name", "Uploaded Email", "Uploaded Phone",
    "Reason", "Suggested Program", "Suggested Batch", "Source Row", "Logged At"
]
STATUS_CACHE_HEADERS = [
    "Student ID", "Program", "Student Name", "Email", "Mobile", "Admissions Status",
    "Payment Date", "Source Spreadsheet", "Source Sheet", "Matched By", "Synced At"
]
WINNER_HEADERS = [
    "Student ID", "Student Name", "Email", "Program", "Activity ID", "Activity",
    "Activity Date", "Amount Won", "Recorded At"
]

CONVERSION_TRACKER_HEADERS = [
    "Student Name", "Email", "Mobile", "Admissions Status", "Payment Date",
    "Offer Date", "Deadline", "Counsellor Name", "Last Updated"
]

TETRAPP_HEADERS = [
    "Student ID", "Name", "Email", "Phone", "Program", "Registered At",
    "Activity", "Source", "Matched By"
]


def clean(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).replace("\n", " ").replace("\r", " ").replace("\xa0", " ").strip()


def norm_email(x) -> str:
    return clean(x).lower()


def norm_name(x) -> str:
    s = clean(x).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_phone8(x) -> str:
    s = re.sub(r"\D+", "", clean(x))
    return s[-8:] if len(s) >= 8 else s


def now_iso() -> str:
    return pd.Timestamp.now(tz="Asia/Kolkata").strftime("%Y-%m-%d %H:%M:%S")


def normalize_batch(batch) -> str:
    s = clean(batch).upper().replace(" ", "")
    m = re.search(r"B?(\d+)", s)
    return f"B{int(m.group(1))}" if m else s


def batch_sheet_name(program: str, batch: str) -> str:
    p = clean(program).upper()
    b = normalize_batch(batch)
    if p not in {"UG", "PG", "GY"}:
        raise ValueError("Program must be UG, PG, or GY")
    prefix = "Gap Year" if p == "GY" else p
    return f"{prefix} {b}"


def _unique_id(prefix: str) -> str:
    stamp = pd.Timestamp.now(tz="Asia/Kolkata").strftime("%y%m%d%H%M%S")
    return f"{prefix}-{stamp}-{uuid.uuid4().hex[:4].upper()}"


def _col_letter(col_num: int) -> str:
    """1-based spreadsheet column number -> A1 column letters."""
    col_num = int(col_num)
    if col_num < 1:
        raise ValueError("Column number must be >= 1")
    out = ""
    while col_num:
        col_num, rem = divmod(col_num - 1, 26)
        out = chr(65 + rem) + out
    return out


def student_id_from_identity(program: str, email: str, name: str, phone: str) -> str:
    material = f"{program.upper()}|{norm_email(email)}|{norm_name(name)}|{norm_phone8(phone)}"
    digest = hashlib.sha1(material.encode("utf-8")).hexdigest()[:10].upper()
    return f"STU-{CYCLE.replace('-', '')}-{digest}"


def _find_header_index(rows: list[list[str]], aliases: set[str], scan=30) -> int:
    for i, row in enumerate(rows[:scan]):
        normed = {norm_name(v) for v in row if clean(v)}
        if normed.intersection(aliases):
            return i
    return 0


def _header_map(headers: Iterable[str]) -> dict[str, int]:
    return {norm_name(h): i for i, h in enumerate(headers) if clean(h)}


def _best_header(headers, aliases):
    hm = _header_map(headers)
    for a in aliases:
        na = norm_name(a)
        if na in hm:
            return hm[na]
    for a in aliases:
        na = norm_name(a)
        for k, idx in hm.items():
            if na and na in k:
                return idx
    return None


@dataclass
class GoogleStore:
    service_account: dict
    master_spreadsheet_id: str
    admin_user: str = "Admin"

    def __post_init__(self):
        creds = Credentials.from_service_account_info(self.service_account, scopes=GSHEETS_SCOPES)
        self.gc = gspread.authorize(creds)
        self.book = self.gc.open_by_key(self.master_spreadsheet_id)

    def worksheet(self, title: str):
        return self.book.worksheet(title)

    def worksheet_or_create(self, title: str, rows=1000, cols=30):
        try:
            return self.book.worksheet(title)
        except WorksheetNotFound:
            return self.book.add_worksheet(title=title, rows=rows, cols=cols)

    def ensure_tab(self, title: str, headers: list[str]):
        ws = self.worksheet_or_create(title, rows=max(1000, len(headers) + 20), cols=max(30, len(headers) + 5))
        values = ws.get_all_values()
        if not values or not any(clean(x) for x in values[0]):
            ws.update("A1", [headers], value_input_option="USER_ENTERED")
            ws.freeze(rows=1)
        return ws

    def ensure_schema(self):
        for title, headers in [
            ("Master UG", MASTER_HEADERS),
            ("Master PG", MASTER_HEADERS),
            ("Gap Year", MASTER_HEADERS),
            ("Dates", DATES_HEADERS),
            ("Batch_Master", BATCH_MASTER_HEADERS),
            ("Activity_Master", ACTIVITY_MASTER_HEADERS),
            ("Admin_Audit_Log", AUDIT_HEADERS),
            ("Upload_Log", UPLOAD_LOG_HEADERS),
            ("Unmatched_Attendance", UNMATCHED_HEADERS),
            ("Admissions_Status_Cache", STATUS_CACHE_HEADERS),
            ("Winner", WINNER_HEADERS),
            ("TetrApp_Competitions", TETRAPP_HEADERS),
            ("TetrApp_Quizzes", TETRAPP_HEADERS),
            ("UG Conversion Tracker", CONVERSION_TRACKER_HEADERS),
            ("PG Conversion Tracker", CONVERSION_TRACKER_HEADERS),
            ("Gap Year Conversion Tracker", CONVERSION_TRACKER_HEADERS),
            ("Tetr-X-UG", MASTER_HEADERS),
            ("Tetr-X-PG", MASTER_HEADERS),
        ]:
            self.ensure_tab(title, headers)

        self.create_batch("UG", "B1", created_by=self.admin_user, log_if_exists=False)
        self.create_batch("PG", "B1", created_by=self.admin_user, log_if_exists=False)

    def log(self, action, entity="", entity_id="", program="", batch="", sheet="", field="", old="", new="", details=""):
        ws = self.ensure_tab("Admin_Audit_Log", AUDIT_HEADERS)
        ws.append_row([
            now_iso(), self.admin_user, action, entity, entity_id, program, batch,
            sheet, field, clean(old), clean(new), clean(details)
        ], value_input_option="USER_ENTERED")

    def list_sheet_names(self):
        return [w.title for w in self.book.worksheets()]

    def list_batches(self, program: str | None = None) -> pd.DataFrame:
        ws = self.ensure_tab("Batch_Master", BATCH_MASTER_HEADERS)
        rows = ws.get_all_records()
        df = pd.DataFrame(rows)
        if df.empty:
            return pd.DataFrame(columns=BATCH_MASTER_HEADERS)
        if program:
            df = df[df["Program"].astype(str).str.upper().eq(program.upper())]
        return df

    def create_batch(self, program: str, batch: str, created_by: str | None = None, log_if_exists=True):
        program = clean(program).upper()
        batch = normalize_batch(batch)
        title = batch_sheet_name(program, batch)
        if title in self.list_sheet_names():
            if log_if_exists:
                self.log("Batch exists", "Batch", title, program, batch, title)
            return self.book.worksheet(title), False

        ws = self.book.add_worksheet(title=title, rows=2500, cols=180)
        # Preserve the legacy activity-sheet structure:
        # row 1 activity type, row 2 activity name, row 3 date, rows 4-5 reserved,
        # row 6 student headers. Event columns begin at column T (index 19).
        matrix = [
            [""] * len(BATCH_BASE_HEADERS),
            [""] * len(BATCH_BASE_HEADERS),
            [""] * len(BATCH_BASE_HEADERS),
            [""] * len(BATCH_BASE_HEADERS),
            [""] * len(BATCH_BASE_HEADERS),
            BATCH_BASE_HEADERS,
        ]
        ws.update("A1", matrix, value_input_option="USER_ENTERED")
        ws.freeze(rows=6, cols=2)

        bws = self.ensure_tab("Batch_Master", BATCH_MASTER_HEADERS)
        bws.append_row([
            CYCLE, program, batch, title, now_iso(), created_by or self.admin_user, "Yes"
        ], value_input_option="USER_ENTERED")
        self.log("Create batch", "Batch", title, program, batch, title, details=f"{CYCLE} batch created")
        return ws, True

    def _master_title(self, program: str) -> str:
        p = clean(program).upper()
        return {"UG": "Master UG", "PG": "Master PG", "GY": "Gap Year"}.get(p, "")

    def master_df(self, program: str | None = None) -> pd.DataFrame:
        frames = []
        programs = [program.upper()] if program else ["UG", "PG", "GY"]
        for p in programs:
            title = self._master_title(p)
            if not title:
                continue
            ws = self.ensure_tab(title, MASTER_HEADERS)
            recs = ws.get_all_records()
            df = pd.DataFrame(recs)
            if df.empty:
                continue
            df["Program"] = p
            df["_MasterSheet"] = title
            frames.append(df)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=MASTER_HEADERS + ["Program", "_MasterSheet"])

    def _find_duplicate(self, program, name, email, phone):
        df = self.master_df(program)
        if df.empty:
            return None, ""
        e = norm_email(email)
        n = norm_name(name)
        p = norm_phone8(phone)
        if e and "Email" in df:
            m = df[df["Email"].map(norm_email).eq(e)]
            if not m.empty:
                return m.iloc[0].to_dict(), "Email"
        if n and "Student Name" in df:
            m = df[df["Student Name"].map(norm_name).eq(n)]
            if len(m) == 1:
                return m.iloc[0].to_dict(), "Full Name"
        if p and "Mobile" in df:
            m = df[df["Mobile"].map(norm_phone8).eq(p)]
            if len(m) == 1:
                return m.iloc[0].to_dict(), "Phone last 8"
        return None, ""

    def add_student(self, *, program, name, email="", phone="", country="", income="",
                    batch="B1", community_status="", counsellor="", offer_date="",
                    deadline="", source="Admin"):
        program = clean(program).upper()
        batch = normalize_batch(batch)
        if not clean(name):
            raise ValueError("Student name is required.")
        existing, matched_by = self._find_duplicate(program, name, email, phone)
        if existing:
            raise ValueError(f"Possible duplicate student matched by {matched_by}: {existing.get('Student Name', '')} / {existing.get('Email', '')}")

        sid = student_id_from_identity(program, email, name, phone)
        title = self._master_title(program)
        if not title:
            raise ValueError("Program must be UG, PG, or GY")

        self.create_batch(program, batch, log_if_exists=False)
        master = self.ensure_tab(title, MASTER_HEADERS)
        row = [
            sid, clean(name), norm_email(email), clean(phone), clean(country), clean(income),
            batch, clean(community_status), clean(counsellor), clean(offer_date), clean(deadline),
            "", "", "", now_iso(),
        ]
        master.append_row(row, value_input_option="USER_ENTERED")

        dates = self.ensure_tab("Dates", DATES_HEADERS)
        dates.append_row([
            sid, clean(name), norm_email(email), program, batch, clean(offer_date), clean(deadline),
            1, now_iso()
        ], value_input_option="USER_ENTERED")

        self._append_student_to_batch(program, batch, row)
        self.log("Add student", "Student", sid, program, batch, title, details=f"{name} added")
        return sid

    def _append_student_to_batch(self, program: str, batch: str, master_row: list):
        ws, _ = self.create_batch(program, batch, log_if_exists=False)
        rows = ws.get_all_values()
        header = rows[5] if len(rows) > 5 else BATCH_BASE_HEADERS
        hm = _header_map(header)
        existing_ids = {clean(r[hm.get("student id", 0)]) for r in rows[6:] if r}
        sid = clean(master_row[0])
        if sid in existing_ids:
            return

        data = {
            "Student ID": master_row[0],
            "Student Name": master_row[1],
            "Email": master_row[2],
            "Mobile": master_row[3],
            "Batch": master_row[6],
            "Country": master_row[4],
            "Income": master_row[5],
            "Counsellor Name": master_row[8],
            "Community Status": master_row[7],
            "Term Zero Group": "",
            "Admissions Status": master_row[11],
            "Payment Date": master_row[12],
            "Offer Date": master_row[9],
            "Deadline": master_row[10],
            "Overall Engagement Score": "",
            "Overall Engagement %": "",
            "Notes": "",
            "Source": "Master",
            "Last Updated": now_iso(),
        }
        values = [data.get(h, "") for h in header]
        # pad for any event columns already in the batch
        if len(values) < len(header):
            values += [""] * (len(header) - len(values))
        ws.append_row(values, value_input_option="USER_ENTERED")

    def bulk_add_students(self, df: pd.DataFrame, default_program=None, default_batch="B1"):
        aliases = {norm_name(c): c for c in df.columns}
        def col(*names):
            for n in names:
                if norm_name(n) in aliases:
                    return aliases[norm_name(n)]
            return None

        c_name = col("Student Name", "Name", "Full Name")
        c_email = col("Email", "Email Address")
        c_phone = col("Mobile", "Phone", "Phone Number", "Contact")
        c_country = col("Country")
        c_income = col("Income", "Income Distribution")
        c_program = col("Program", "UG/PG")
        c_batch = col("Batch")
        c_offer = col("Offer Date", "Offered Date")
        c_deadline = col("Deadline")
        c_community = col("Community Status", "Master Status")
        c_counsellor = col("Counsellor", "Counsellor Name")

        if not c_name:
            raise ValueError("Bulk student file needs a Name / Student Name column.")

        results = []
        for i, r in df.iterrows():
            program = clean(r.get(c_program, "")) if c_program else clean(default_program)
            batch = clean(r.get(c_batch, "")) if c_batch else default_batch
            try:
                sid = self.add_student(
                    program=program or default_program,
                    name=r.get(c_name, ""),
                    email=r.get(c_email, "") if c_email else "",
                    phone=r.get(c_phone, "") if c_phone else "",
                    country=r.get(c_country, "") if c_country else "",
                    income=r.get(c_income, "") if c_income else "",
                    batch=batch or default_batch,
                    community_status=r.get(c_community, "") if c_community else "",
                    counsellor=r.get(c_counsellor, "") if c_counsellor else "",
                    offer_date=r.get(c_offer, "") if c_offer else "",
                    deadline=r.get(c_deadline, "") if c_deadline else "",
                    source="Bulk Upload",
                )
                results.append({"Row": i + 2, "Result": "Added", "Student ID": sid, "Message": ""})
            except Exception as e:
                results.append({"Row": i + 2, "Result": "Skipped", "Student ID": "", "Message": str(e)})
        return pd.DataFrame(results)

    def update_student_field(self, student_id: str, program: str, field: str, new_value):
        title = self._master_title(program)
        ws = self.ensure_tab(title, MASTER_HEADERS)
        values = ws.get_all_values()
        if not values:
            raise ValueError("Master sheet is empty.")
        header = values[0]
        if field not in header:
            raise ValueError(f"Field '{field}' not found in {title}.")
        id_idx = header.index("Student ID")
        col_idx = header.index(field) + 1
        for row_idx, row in enumerate(values[1:], start=2):
            sid = row[id_idx] if id_idx < len(row) else ""
            if clean(sid) == clean(student_id):
                old = row[col_idx - 1] if col_idx - 1 < len(row) else ""
                ws.update_cell(row_idx, col_idx, clean(new_value))
                self.log("Edit student", "Student", student_id, program, sheet=title, field=field, old=old, new=new_value)
                return
        raise ValueError("Student ID not found.")

    def create_activity(self, *, name, activity_date, activity_type, program=None, batches=None, targets=None, source="Admin"):
        """Create one activity across one or many program/batch targets.

        Backward compatible with the older ``program + batches`` call.  The new
        Admin UI can instead pass ``targets=[("UG", "B1"), ("PG", "B1"), ...]``.
        Every target sheet gets the same Activity ID so one mixed attendance file
        can later write to the correct UG/PG/GY batch automatically.
        """
        if not clean(name):
            raise ValueError("Activity name is required.")

        if targets is None:
            p = clean(program).upper()
            targets = [(p, normalize_batch(b)) for b in (batches or [])]

        normalized_targets = []
        seen = set()
        for item in targets or []:
            if isinstance(item, dict):
                p = clean(item.get("Program", item.get("program", ""))).upper()
                b = normalize_batch(item.get("Batch", item.get("batch", "")))
            else:
                try:
                    p, b = item
                except Exception:
                    raise ValueError("Each activity target must contain Program and Batch.")
                p = clean(p).upper()
                b = normalize_batch(b)
            if p not in {"UG", "PG", "GY"} or not b:
                continue
            key = (p, b)
            if key not in seen:
                seen.add(key)
                normalized_targets.append(key)

        if not normalized_targets:
            raise ValueError("Select at least one UG / PG / GY batch for the activity.")

        # Stable program order in Activity_Master.
        program_order = {"UG": 0, "PG": 1, "GY": 2}
        normalized_targets.sort(key=lambda x: (program_order.get(x[0], 9), int(re.sub(r"\\D", "", x[1]) or 999999)))

        activity_id = _unique_id("ACT")
        target_sheet_names = [batch_sheet_name(p, b) for p, b in normalized_targets]
        programs_label = ", ".join(dict.fromkeys([p for p, _ in normalized_targets]))

        am = self.ensure_tab("Activity_Master", ACTIVITY_MASTER_HEADERS)
        am.append_row([
            activity_id, CYCLE, clean(name), clean(activity_date), clean(activity_type),
            programs_label, ", ".join(target_sheet_names), clean(source), now_iso(), self.admin_user
        ], value_input_option="USER_ENTERED")

        for p, b in normalized_targets:
            self._ensure_activity_column(p, b, activity_id, name, activity_type, activity_date)

        self.log(
            "Create activity", "Activity", activity_id, programs_label,
            ", ".join(target_sheet_names), "Activity_Master",
            details=f"{name} | {activity_type} | {activity_date}"
        )
        return activity_id

    def _unmerge_activity_metadata(self, ws, end_col=None):
        """Remove accidental merges in activity metadata rows (1-3) from column T onward.

        The 2026-27 structure requires exactly one activity type/name/date cell per
        activity column.  This also repairs old template merges that made one event
        type visually span several new columns.
        """
        end_col = max(int(end_col or ws.col_count or 20), 20)
        try:
            self.book.batch_update({
                "requests": [{
                    "unmergeCells": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 3,
                            "startColumnIndex": 19,
                            "endColumnIndex": end_col,
                        }
                    }
                }]
            })
        except Exception:
            # No merge / older gspread backend: safe to continue.
            pass

    @staticmethod
    def _activity_date_value(value):
        dt = pd.to_datetime(clean(value), errors="coerce", dayfirst=False)
        if pd.isna(dt):
            return pd.NaT
        try:
            return dt.normalize()
        except Exception:
            return dt

    def repair_activity_layout(self, program, batch):
        """Unmerge, de-ghost and date-sort the activity columns in one batch sheet.

        Student data in A:S is untouched.  From T onward, every real activity is
        compacted to one column and sorted by Activity Date ascending. Attendance
        values move together with their activity metadata and Activity ID.
        """
        program = clean(program).upper()
        batch = normalize_batch(batch)
        ws, _ = self.create_batch(program, batch, log_if_exists=False)
        rows = ws.get_all_values()
        if len(rows) < 6:
            return {"Sheet": ws.title, "Activities": 0, "Changed": False}

        max_cols = max([len(r) for r in rows] + [20])
        self._unmerge_activity_metadata(ws, max(max_cols, ws.col_count))

        def cell(ridx, cidx):
            if ridx < len(rows) and cidx < len(rows[ridx]):
                return clean(rows[ridx][cidx])
            return ""

        # Activity_Master is the canonical metadata fallback. This is important for
        # repairing older merged type headers: after unmerge, only the top-left
        # cell of a merged range keeps its value, so a Competition column could
        # otherwise lose its own type. Activity ID lets us restore it exactly.
        activity_meta = {}
        try:
            am = self.ensure_tab("Activity_Master", ACTIVITY_MASTER_HEADERS)
            for rec in am.get_all_records():
                aid = clean(rec.get("Activity ID", ""))
                if aid:
                    activity_meta[aid] = {
                        "type": clean(rec.get("Activity Type", "")),
                        "name": clean(rec.get("Activity Name", "")),
                        "date": clean(rec.get("Activity Date", "")),
                    }
        except Exception:
            activity_meta = {}

        event_cols = []
        orphan_cols = []
        max_touched = 19
        for cidx in range(19, max_cols):
            ev_type = cell(0, cidx)
            ev_name = cell(1, cidx)
            ev_date = cell(2, cidx)
            ev_id = cell(5, cidx)
            if ev_type or ev_name or ev_date or ev_id:
                max_touched = max(max_touched, cidx)
            # A real activity needs an ID, name, or date. A type by itself is a
            # ghost left by an old merged header and must not occupy a column.
            if ev_id or ev_name or ev_date:
                event_cols.append(cidx)
            elif ev_type:
                orphan_cols.append(cidx)

        def effective_meta(cidx):
            ev_id = cell(5, cidx)
            master = activity_meta.get(ev_id, {})
            return {
                "id": ev_id,
                "type": master.get("type") or cell(0, cidx),
                "name": master.get("name") or cell(1, cidx),
                "date": master.get("date") or cell(2, cidx),
            }

        # Stable date sort; same-day events preserve their existing order.
        def sort_key(cidx):
            dt = self._activity_date_value(effective_meta(cidx)["date"])
            if pd.isna(dt):
                return (1, pd.Timestamp.max, cidx)
            return (0, dt, cidx)

        sorted_cols = sorted(event_cols, key=sort_key)
        target_cols = list(range(19, 19 + len(sorted_cols)))
        changed = sorted_cols != target_cols or bool(orphan_cols)

        if max_touched >= 19:
            last_row = max(len(rows), 6)
            clear_range = f"T1:{_col_letter(max_touched + 1)}{last_row}"

            # Snapshot the full values of every activity column BEFORE clearing.
            # Rows 1/2/3 are rebuilt from Activity_Master when available so each
            # column gets its own correct Type / Name / Date after unmerging.
            matrix = []
            for ridx in range(last_row):
                row_values = []
                for cidx in sorted_cols:
                    if ridx == 0:
                        row_values.append(effective_meta(cidx)["type"])
                    elif ridx == 1:
                        row_values.append(effective_meta(cidx)["name"])
                    elif ridx == 2:
                        row_values.append(effective_meta(cidx)["date"])
                    else:
                        row_values.append(cell(ridx, cidx))
                matrix.append(row_values)

            try:
                ws.batch_clear([clear_range])
            except Exception:
                # Fallback for older gspread versions.
                blank_width = max_touched - 19 + 1
                ws.update(
                    f"T1:{_col_letter(max_touched + 1)}{last_row}",
                    [[""] * blank_width for _ in range(last_row)],
                    value_input_option="USER_ENTERED",
                )

            if sorted_cols:
                end_col = 19 + len(sorted_cols)  # 1-based column count for A1 helper below
                ws.update(
                    f"T1:{_col_letter(end_col)}{last_row}",
                    matrix,
                    value_input_option="USER_ENTERED",
                )

        return {"Sheet": ws.title, "Activities": len(sorted_cols), "Changed": changed}

    def repair_all_activity_layouts(self, targets=None):
        """Repair/sort all registered batch sheets or a supplied target list."""
        if targets is None:
            bdf = self.list_batches()
            targets = []
            if bdf is not None and not bdf.empty:
                for _, r in bdf.iterrows():
                    p = clean(r.get("Program", "")).upper()
                    b = normalize_batch(r.get("Batch", ""))
                    active = clean(r.get("Active", "Yes")).lower()
                    if p in {"UG", "PG", "GY"} and b and active not in {"no", "false", "0"}:
                        targets.append((p, b))

        results = []
        seen = set()
        for p, b in targets or []:
            p = clean(p).upper()
            b = normalize_batch(b)
            if (p, b) in seen or p not in {"UG", "PG", "GY"} or not b:
                continue
            seen.add((p, b))
            try:
                results.append(self.repair_activity_layout(p, b))
            except Exception as e:
                results.append({"Sheet": batch_sheet_name(p, b), "Activities": 0, "Changed": False, "Error": str(e)})
        return pd.DataFrame(results)

    def _insert_activity_column(self, ws, col_num):
        """Insert one physical column before col_num, preserving attendance columns."""
        col_num = int(col_num)
        if col_num > ws.col_count:
            ws.add_cols(col_num - ws.col_count)
            return
        try:
            self.book.batch_update({
                "requests": [{
                    "insertDimension": {
                        "range": {
                            "sheetId": ws.id,
                            "dimension": "COLUMNS",
                            "startIndex": col_num - 1,
                            "endIndex": col_num,
                        },
                        "inheritFromBefore": True,
                    }
                }]
            })
        except Exception:
            # Public gspread fallback.
            ws.insert_cols([[""]], col=col_num, value_input_option="USER_ENTERED", inherit_from_before=True)

    def _ensure_activity_column(self, program, batch, activity_id, name, activity_type, activity_date):
        # Repair old merged headers and put all existing activities in date order first.
        self.repair_activity_layout(program, batch)

        ws, _ = self.create_batch(program, batch, log_if_exists=False)
        rows = ws.get_all_values()
        if len(rows) < 6:
            raise ValueError(f"Batch sheet {ws.title} does not have the expected six-row header.")

        max_cols = max([len(r) for r in rows] + [20])
        self._unmerge_activity_metadata(ws, max(max_cols, ws.col_count))

        def cell(ridx, cidx):
            if ridx < len(rows) and cidx < len(rows[ridx]):
                return clean(rows[ridx][cidx])
            return ""

        # Existing Activity ID -> nothing more to create.
        for cidx in range(19, max_cols):
            if cell(5, cidx) == activity_id:
                return cidx + 1

        existing = []
        for cidx in range(19, max_cols):
            ev_type = cell(0, cidx)
            ev_name = cell(1, cidx)
            ev_date_raw = cell(2, cidx)
            ev_id = cell(5, cidx)
            if ev_id or ev_name or ev_date_raw:
                existing.append((cidx + 1, self._activity_date_value(ev_date_raw)))
            elif ev_type:
                # Orphan type-only cell from an old merged header: clear it.
                try:
                    ws.update_cell(1, cidx + 1, "")
                except Exception:
                    pass

        new_dt = self._activity_date_value(activity_date)
        target_col = 20
        if existing:
            # Insert before the first strictly later event. Same-day events stay together
            # and the newly-created activity goes after existing same-day activities.
            later_col = None
            if pd.notna(new_dt):
                for col_num, old_dt in existing:
                    if pd.notna(old_dt) and old_dt > new_dt:
                        later_col = col_num
                        break
            if later_col is not None:
                target_col = later_col
                self._insert_activity_column(ws, target_col)
            else:
                target_col = max(col for col, _ in existing) + 1
                if target_col > ws.col_count:
                    ws.add_cols(target_col - ws.col_count)
        elif ws.col_count < 20:
            ws.add_cols(20 - ws.col_count)

        # One activity = one metadata cell in each row; no merging.
        ws.update_cell(1, target_col, clean(activity_type))
        ws.update_cell(2, target_col, clean(name))
        ws.update_cell(3, target_col, clean(activity_date))
        ws.update_cell(6, target_col, activity_id)
        return target_col

    def students_for_matching(self):
        df = self.master_df()
        if df.empty:
            return df
        for c in ["Student Name", "Email", "Mobile", "Batch", "Program", "Student ID"]:
            if c not in df:
                df[c] = ""
        return df

    def match_attendance(self, attendance_df: pd.DataFrame) -> pd.DataFrame:
        students = self.students_for_matching().copy()
        if students.empty:
            raise ValueError("No students exist in the 2026-27 master yet.")

        students["_email"] = students["Email"].map(norm_email)
        students["_name"] = students["Student Name"].map(norm_name)
        students["_phone"] = students["Mobile"].map(norm_phone8)

        email_map = {}
        name_map = {}
        phone_map = {}
        for idx, r in students.iterrows():
            if r["_email"]:
                email_map.setdefault(r["_email"], []).append(idx)
            if r["_name"]:
                name_map.setdefault(r["_name"], []).append(idx)
            if r["_phone"]:
                phone_map.setdefault(r["_phone"], []).append(idx)

        out = []
        for _, r in attendance_df.iterrows():
            candidates = []
            matched_by = ""
            e = norm_email(r.get("Uploaded Email", ""))
            n = norm_name(r.get("Uploaded Name", ""))
            p = norm_phone8(r.get("Uploaded Phone", ""))

            if e and len(email_map.get(e, [])) == 1:
                candidates = email_map[e]
                matched_by = "Email"
            elif n and len(name_map.get(n, [])) == 1:
                candidates = name_map[n]
                matched_by = "Full Name"
            elif p and len(phone_map.get(p, [])) == 1:
                candidates = phone_map[p]
                matched_by = "Phone last 8"

            base = r.to_dict()
            if candidates:
                s = students.loc[candidates[0]]
                base.update({
                    "Match Status": "Matched",
                    "Matched By": matched_by,
                    "Student ID": s.get("Student ID", ""),
                    "Student Name": s.get("Student Name", ""),
                    "Program": s.get("Program", ""),
                    "Batch": s.get("Batch", ""),
                    "Master Email": s.get("Email", ""),
                })
            else:
                reason = "No unique match"
                if e and len(email_map.get(e, [])) > 1:
                    reason = "Ambiguous email"
                elif n and len(name_map.get(n, [])) > 1:
                    reason = "Ambiguous name"
                elif p and len(phone_map.get(p, [])) > 1:
                    reason = "Ambiguous phone"
                base.update({
                    "Match Status": "Unmatched",
                    "Matched By": "",
                    "Student ID": "",
                    "Student Name": "",
                    "Program": "",
                    "Batch": "",
                    "Master Email": "",
                    "Reason": reason,
                })
            out.append(base)
        return pd.DataFrame(out)

    def commit_attendance(self, matched_df, *, activity_id, activity_name, activity_date,
                          activity_type, file_name, detected_format, input_rows, unique_attendees):
        upload_id = _unique_id("UPL")
        matched = matched_df[matched_df["Match Status"].eq("Matched")].copy()
        unmatched = matched_df[~matched_df["Match Status"].eq("Matched")].copy()

        activity_master = self.ensure_tab("Activity_Master", ACTIVITY_MASTER_HEADERS)
        activities = pd.DataFrame(activity_master.get_all_records())
        if activities.empty or activity_id not in activities.get("Activity ID", pd.Series(dtype=str)).astype(str).tolist():
            # Create activity across actual matched batches if it has not already been created.
            combos = matched[["Program", "Batch"]].drop_duplicates()
            for program, sub in combos.groupby("Program"):
                batches = [b for b in sub["Batch"].astype(str).tolist() if clean(b)]
                if batches:
                    # one shared ID needs manual insertion rather than create_activity generating a new id
                    activity_master.append_row([
                        activity_id, CYCLE, activity_name, activity_date, activity_type,
                        program, ", ".join(batches), detected_format, now_iso(), self.admin_user
                    ], value_input_option="USER_ENTERED")
                    for b in batches:
                        self._ensure_activity_column(program, b, activity_id, activity_name, activity_type, activity_date)

        committed = 0
        for _, r in matched.iterrows():
            program = clean(r.get("Program", "")).upper()
            batch = normalize_batch(r.get("Batch", ""))
            sid = clean(r.get("Student ID", ""))
            if not program or not batch or not sid:
                continue
            ws, _ = self.create_batch(program, batch, log_if_exists=False)
            col = self._ensure_activity_column(program, batch, activity_id, activity_name, activity_type, activity_date)
            rows = ws.get_all_values()
            if len(rows) < 7:
                continue
            header = rows[5]
            try:
                sid_idx = header.index("Student ID")
            except ValueError:
                sid_idx = 0

            target_row = None
            for row_num, row in enumerate(rows[6:], start=7):
                if sid_idx < len(row) and clean(row[sid_idx]) == sid:
                    target_row = row_num
                    break

            if target_row is None:
                # Rehydrate the student into the batch sheet if needed.
                master = self.master_df(program)
                hit = master[master["Student ID"].astype(str).eq(sid)]
                if not hit.empty:
                    m = hit.iloc[0]
                    mrow = [
                        m.get("Student ID",""), m.get("Student Name",""), m.get("Email",""),
                        m.get("Mobile",""), m.get("Country",""), m.get("Income",""),
                        m.get("Batch",""), m.get("Community Status",""), m.get("Counsellor Name",""),
                        m.get("Offer Date",""), m.get("Deadline",""), m.get("Admissions Status",""),
                        m.get("Payment Date",""), m.get("Admissions Source",""), m.get("Last Synced",""),
                    ]
                    self._append_student_to_batch(program, batch, mrow)
                    rows = ws.get_all_values()
                    for row_num, row in enumerate(rows[6:], start=7):
                        if sid_idx < len(row) and clean(row[sid_idx]) == sid:
                            target_row = row_num
                            break

            if target_row:
                ws.update_cell(target_row, col, 1)
                committed += 1

        if not unmatched.empty:
            uws = self.ensure_tab("Unmatched_Attendance", UNMATCHED_HEADERS)
            rows_to_append = []
            for _, r in unmatched.iterrows():
                rows_to_append.append([
                    upload_id, activity_id, clean(r.get("Uploaded Name","")),
                    clean(r.get("Uploaded Email","")), clean(r.get("Uploaded Phone","")),
                    clean(r.get("Reason","No unique match")), clean(r.get("Program","")),
                    clean(r.get("Batch","")), clean(r.get("Source Row","")), now_iso()
                ])
            if rows_to_append:
                uws.append_rows(rows_to_append, value_input_option="USER_ENTERED")

        lws = self.ensure_tab("Upload_Log", UPLOAD_LOG_HEADERS)
        lws.append_row([
            upload_id, now_iso(), self.admin_user, file_name, detected_format,
            activity_id, activity_name, int(input_rows), int(unique_attendees),
            int(len(matched)), int(len(unmatched)), int(committed)
        ], value_input_option="USER_ENTERED")
        self.log("Attendance upload", "Upload", upload_id, sheet="Upload_Log",
                 details=f"{activity_name}: {committed} attendance records written; {len(unmatched)} unmatched")
        return {"upload_id": upload_id, "matched": len(matched), "unmatched": len(unmatched), "committed": committed}

    def update_community_status(self, student_id, program, new_status):
        self.update_student_field(student_id, program, "Community Status", new_status)

    def sync_admissions_source(self, *, program, spreadsheet_id, sheet_name,
                               status_aliases=None, payment_date_aliases=None):
        program = clean(program).upper()
        if not spreadsheet_id or not sheet_name:
            raise ValueError(f"{program} admissions spreadsheet ID and sheet name are required.")

        source_book = self.gc.open_by_key(spreadsheet_id)
        ws = source_book.worksheet(sheet_name)
        rows = ws.get_all_values()
        if not rows:
            return {"source_rows": 0, "matched": 0, "unmatched": 0}

        aliases = {"name", "student name", "full name", "email", "email address"}
        header_idx = _find_header_index(rows, aliases)
        headers = rows[header_idx]
        data = rows[header_idx + 1:]
        df = pd.DataFrame(data, columns=headers)
        df = df.loc[:, ~df.columns.duplicated()].copy()

        name_i = _best_header(headers, ["Student Name", "Name", "Full Name"])
        email_i = _best_header(headers, ["Email", "Email Address"])
        phone_i = _best_header(headers, ["Mobile", "Phone", "Phone Number", "Contact"])
        status_i = _best_header(headers, status_aliases or ["Status", "Payment Status", "Admissions Status", "Stage"])
        payment_i = _best_header(headers, payment_date_aliases or ["Payment Date", "Deposit Date", "Paid Date"])

        if status_i is None:
            raise ValueError(f"Could not find an admissions/payment status column in {sheet_name}.")

        colname = lambda i: headers[i] if i is not None and i < len(headers) else None
        c_name, c_email, c_phone, c_status, c_payment = map(colname, [name_i, email_i, phone_i, status_i, payment_i])

        master = self.master_df(program).copy()
        if master.empty:
            raise ValueError(f"No {program} students exist in the 2026-27 master.")
        master["_email"] = master["Email"].map(norm_email)
        master["_name"] = master["Student Name"].map(norm_name)
        master["_phone"] = master["Mobile"].map(norm_phone8)

        def unique_map(col):
            d = {}
            for idx, val in master[col].items():
                if val:
                    d.setdefault(val, []).append(idx)
            return d
        email_map, name_map, phone_map = unique_map("_email"), unique_map("_name"), unique_map("_phone")

        cache_rows = []
        matched_count = 0
        for _, r in df.iterrows():
            name = clean(r.get(c_name, "")) if c_name else ""
            email = clean(r.get(c_email, "")) if c_email else ""
            phone = clean(r.get(c_phone, "")) if c_phone else ""
            status = clean(r.get(c_status, "")) if c_status else ""
            payment = clean(r.get(c_payment, "")) if c_payment else ""
            if not any([name, email, phone, status]):
                continue

            idxs, matched_by = [], ""
            e, n, p = norm_email(email), norm_name(name), norm_phone8(phone)
            if e and len(email_map.get(e, [])) == 1:
                idxs, matched_by = email_map[e], "Email"
            elif n and len(name_map.get(n, [])) == 1:
                idxs, matched_by = name_map[n], "Full Name"
            elif p and len(phone_map.get(p, [])) == 1:
                idxs, matched_by = phone_map[p], "Phone last 8"
            if not idxs:
                continue
            s = master.loc[idxs[0]]
            cache_rows.append([
                s.get("Student ID",""), program, s.get("Student Name",""), s.get("Email",""),
                s.get("Mobile",""), status, payment, spreadsheet_id, sheet_name, matched_by, now_iso()
            ])
            matched_count += 1

        # Replace only this program's rows in cache, preserve other programs.
        cws = self.ensure_tab("Admissions_Status_Cache", STATUS_CACHE_HEADERS)
        current = pd.DataFrame(cws.get_all_records())
        if not current.empty and "Program" in current:
            current = current[~current["Program"].astype(str).str.upper().eq(program)].copy()
        combined = []
        if not current.empty:
            combined.extend(current.reindex(columns=STATUS_CACHE_HEADERS, fill_value="").values.tolist())
        combined.extend(cache_rows)
        cws.clear()
        cws.update("A1", [STATUS_CACHE_HEADERS] + combined, value_input_option="USER_ENTERED")
        cws.freeze(rows=1)

        self._apply_status_cache_to_master(program, cache_rows)
        self.log("Sync admissions", "Admissions Sync", program, program=program, sheet=sheet_name,
                 details=f"{matched_count} tracker rows matched")
        return {"source_rows": len(df), "matched": matched_count, "unmatched": max(0, len(df) - matched_count)}

    def _apply_status_cache_to_master(self, program: str, cache_rows: list[list]):
        """Write a read-only mirrored copy of external status/payment to master for
        legacy analytics compatibility. The source of truth remains the admissions tracker.
        """
        if not cache_rows:
            return
        title = self._master_title(program)
        ws = self.ensure_tab(title, MASTER_HEADERS)
        values = ws.get_all_values()
        if not values:
            return
        header = values[0]
        id_idx = header.index("Student ID")
        status_col = header.index("Admissions Status") + 1
        pay_col = header.index("Payment Date") + 1
        source_col = header.index("Admissions Source") + 1
        synced_col = header.index("Last Synced") + 1
        row_by_sid = {}
        for row_num, row in enumerate(values[1:], start=2):
            if id_idx < len(row):
                row_by_sid[clean(row[id_idx])] = row_num

        for cr in cache_rows:
            sid, _, _, _, _, status, payment, source_id, source_sheet, _, synced = cr
            rnum = row_by_sid.get(clean(sid))
            if not rnum:
                continue
            ws.update_cell(rnum, status_col, status)
            ws.update_cell(rnum, pay_col, payment)
            ws.update_cell(rnum, source_col, f"{source_id}:{source_sheet}")
            ws.update_cell(rnum, synced_col, synced)

        # Also mirror into batch sheets so the legacy activity parser sees the live status.
        master = self.master_df(program)
        for _, s in master.iterrows():
            sid = clean(s.get("Student ID",""))
            batch = normalize_batch(s.get("Batch",""))
            if not sid or not batch:
                continue
            try:
                bws = self.book.worksheet(batch_sheet_name(program, batch))
            except WorksheetNotFound:
                continue
            rows = bws.get_all_values()
            if len(rows) < 7:
                continue
            bh = rows[5]
            if "Student ID" not in bh:
                continue
            sid_idx = bh.index("Student ID")
            st_idx = bh.index("Admissions Status") + 1 if "Admissions Status" in bh else None
            pd_idx = bh.index("Payment Date") + 1 if "Payment Date" in bh else None
            for rnum, row in enumerate(rows[6:], start=7):
                if sid_idx < len(row) and clean(row[sid_idx]) == sid:
                    if st_idx:
                        bws.update_cell(rnum, st_idx, clean(s.get("Admissions Status","")))
                    if pd_idx:
                        bws.update_cell(rnum, pd_idx, clean(s.get("Payment Date","")))
                    break
