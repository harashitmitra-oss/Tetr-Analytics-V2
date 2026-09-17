from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import pandas as pd


NAME_ALIASES = [
    "name", "full name", "student name", "attendee name", "participant name",
    "name (original name)", "display name", "registrant name"
]
EMAIL_ALIASES = [
    "email", "email address", "user email", "registrant email", "participant email",
    "attendee email"
]
PHONE_ALIASES = [
    "phone", "mobile", "mobile number", "phone number", "contact", "contact number"
]
STATUS_ALIASES = [
    "attendance", "attended", "status", "attendance status", "participation status",
    "joined"
]
DURATION_ALIASES = [
    "duration", "duration (minutes)", "total duration (minutes)", "time in session",
    "minutes attended"
]


def clean(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).replace("\n", " ").replace("\r", " ").strip()


def normalize_col(x) -> str:
    s = clean(x).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _find_col(columns, aliases):
    norm = {c: normalize_col(c) for c in columns}
    alias_norm = [normalize_col(a) for a in aliases]
    for alias in alias_norm:
        for c, nc in norm.items():
            if nc == alias:
                return c
    for alias in alias_norm:
        for c, nc in norm.items():
            if alias and alias in nc:
                return c
    return None


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    name = getattr(uploaded_file, "name", "upload.xlsx").lower()
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    bio = BytesIO(raw)
    if name.endswith(".csv"):
        try:
            return pd.read_csv(bio)
        except UnicodeDecodeError:
            bio.seek(0)
            return pd.read_csv(bio, encoding="latin-1")
    if name.endswith(".xls") or name.endswith(".xlsx"):
        return pd.read_excel(bio)
    raise ValueError("Upload must be .csv, .xls, or .xlsx")


def detect_format(df: pd.DataFrame) -> str:
    cols = {normalize_col(c) for c in df.columns}
    zoom_markers = {
        "name original name", "user email", "total duration minutes",
        "join time", "leave time"
    }
    if len(cols.intersection(zoom_markers)) >= 2:
        return "Zoom"

    # Dinero exports can vary. Treat a table with attendee identity + attendance/status
    # as Dinero-like when it is not Zoom. The UI still shows the detected mappings.
    has_name = _find_col(df.columns, NAME_ALIASES) is not None
    has_email = _find_col(df.columns, EMAIL_ALIASES) is not None
    has_status = _find_col(df.columns, STATUS_ALIASES) is not None
    if (has_name or has_email) and has_status:
        return "Dinero"
    if has_name or has_email:
        return "Generic attendee list"
    return "Unknown"


def standardize_attendance(
    df: pd.DataFrame,
    source_format: str = "Auto",
    name_col: str | None = None,
    email_col: str | None = None,
    phone_col: str | None = None,
    status_col: str | None = None,
    duration_col: str | None = None,
) -> tuple[pd.DataFrame, dict]:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Uploaded Name", "Uploaded Email", "Uploaded Phone", "Attended", "Duration", "Source Row"]), {}

    detected = detect_format(df)
    fmt = detected if source_format in {"", "Auto", None} else source_format

    name_col = name_col or _find_col(df.columns, NAME_ALIASES)
    email_col = email_col or _find_col(df.columns, EMAIL_ALIASES)
    phone_col = phone_col or _find_col(df.columns, PHONE_ALIASES)
    status_col = status_col or _find_col(df.columns, STATUS_ALIASES)
    duration_col = duration_col or _find_col(df.columns, DURATION_ALIASES)

    if not name_col and not email_col and not phone_col:
        raise ValueError("Could not find a Name, Email, or Phone column in the attendance file.")

    out = pd.DataFrame(index=df.index)
    out["Uploaded Name"] = df[name_col].map(clean) if name_col else ""
    out["Uploaded Email"] = df[email_col].map(clean) if email_col else ""
    out["Uploaded Phone"] = df[phone_col].map(clean) if phone_col else ""

    if status_col:
        status = df[status_col].map(lambda x: clean(x).lower())
        negative = status.str.contains(r"\b(no|absent|did not attend|not attended|cancelled|canceled)\b", regex=True, na=False)
        out["Attended"] = ~negative
    else:
        # Zoom/Dinero attendance exports normally only contain attendees.
        out["Attended"] = True

    if duration_col:
        out["Duration"] = pd.to_numeric(df[duration_col], errors="coerce")
    else:
        out["Duration"] = pd.NA

    out["Source Row"] = [int(i) + 2 for i in range(len(out))]
    out = out[out["Attended"]].copy()

    # Remove completely blank identity rows, then deduplicate.
    has_identity = (
        out["Uploaded Name"].astype(str).str.strip().ne("")
        | out["Uploaded Email"].astype(str).str.strip().ne("")
        | out["Uploaded Phone"].astype(str).str.strip().ne("")
    )
    out = out[has_identity].copy()

    def key(r):
        e = clean(r["Uploaded Email"]).lower()
        if e:
            return "e:" + e
        p = re.sub(r"\D+", "", clean(r["Uploaded Phone"]))[-8:]
        if p:
            return "p:" + p
        n = re.sub(r"[^a-z0-9]+", " ", clean(r["Uploaded Name"]).lower()).strip()
        return "n:" + n

    out["_dedupe"] = out.apply(key, axis=1)
    if "Duration" in out:
        out = out.sort_values("Duration", ascending=False, na_position="last")
    out = out.drop_duplicates("_dedupe", keep="first").drop(columns="_dedupe").reset_index(drop=True)

    mapping = {
        "format": fmt,
        "detected_format": detected,
        "name_col": name_col,
        "email_col": email_col,
        "phone_col": phone_col,
        "status_col": status_col,
        "duration_col": duration_col,
        "input_rows": int(len(df)),
        "unique_attendees": int(len(out)),
    }
    return out, mapping
