"""
Ituran Report Analyzer – Streamlit UI
Phase 1: Ituran GPS report → classified Excel output
"""

import streamlit as st
import pandas as pd
import requests
from datetime import date, datetime
from collections import defaultdict
import os

from ituran_analyzer import (
    analyze_to_buffer,
    TYPE_UNLOAD, TYPE_TRANSPORT, TYPE_PARKING, TYPE_ANOMALY,
)

# ─── page config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ניתוח דוחות איתוראן",
    page_icon="🚗",
    layout="wide",
)

# ─── password gate ───────────────────────────────────────────────────────────

APP_PASSWORD = os.environ.get("APP_PASSWORD", "Elul2026")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 כניסה למערכת")
    pwd = st.text_input("סיסמה", type="password", placeholder="הכנס סיסמה...")
    if st.button("כניסה"):
        if pwd == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("סיסמה שגויה. נסה שנית.")
    st.stop()


# ─── RTL + styles ────────────────────────────────────────────────────────────

st.markdown("""
<style>
    body, .stApp { direction: rtl; text-align: right; }
    .stDownloadButton button { width: 100%; }
    /* sidebar spacing – target ~75% height coverage */
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.55rem !important; }
    section[data-testid="stSidebar"] [data-testid="element-container"] { margin-bottom: 0 !important; }
    section[data-testid="stSidebar"] label { font-size: 0.85rem !important; margin-bottom: 2px !important; }
    section[data-testid="stSidebar"] p  { margin: 2px 0 !important; font-size: 0.85rem !important; }
    section[data-testid="stSidebar"] small { font-size: 0.78rem !important; }
    section[data-testid="stSidebar"] hr { margin: 8px 0 !important; }
    section[data-testid="stSidebar"] [data-testid="stNumberInput"] input { height: 34px !important; font-size: 0.9rem !important; }
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] { margin: 4px 0 !important; }
    /* widen dataframe */
    [data-testid="stDataFrame"] { width: 100% !important; }
    /* date picker RTL fix */
    [data-testid="stDateInput"] { direction: ltr; }
    [data-testid="stDateInput"] label { direction: rtl; text-align: right; width: 100%; }
    /* shrink + style the calendar popup using correct baseweb selectors */
    /* calendar: open UPWARD by shifting the popover up */
    section[data-testid="stSidebar"] [data-baseweb="popover"] {
        margin-top: -445px !important;
        left: 4px !important;
        right: auto !important;
        overflow: visible !important;
    }
    [data-baseweb="calendar"] {
        transform: scale(0.82) !important;
        transform-origin: bottom left !important;
        border: 2px solid #4a90d9 !important;
        border-radius: 12px !important;
        box-shadow: 0 -4px 20px rgba(0,0,0,0.20) !important;
        overflow: hidden !important;
    }
    /* file uploader compact */
    [data-testid="stFileUploader"] { max-width: 420px; }
    /* number inputs compact */
    [data-testid="stNumberInput"] input { text-align: center; }
</style>
""", unsafe_allow_html=True)

# ─── header ──────────────────────────────────────────────────────────────────

st.title("🚗 ניתוח דוחות איתוראן")
st.markdown("**הפרדה בין שעות עבודה בשטח לבין שעות נסיעה**")
st.divider()

# ─── helpers: holidays ───────────────────────────────────────────────────────

@st.cache_data(ttl=86_400 * 30, show_spinner=False)
def fetch_israel_holidays(year: int):
    """
    Returns (full_holiday_dates, eve_dates) as frozensets of date objects.
    full_holiday_dates: days that are full holidays (Shabbat-like, all day off)
    eve_dates: holiday eves where Friday-like cutoff applies
    """
    url = (
        f"https://www.hebcal.com/hebcal?v=1&cfg=json&maj=on&min=off"
        f"&mod=on&nx=on&year={year}&month=x&c=off&i=on&s=off"
    )
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
    except Exception:
        return frozenset(), frozenset()

    full_holidays = set()
    eve_holidays  = set()

    for item in data.get("items", []):
        cat      = item.get("category", "")
        subcat   = item.get("subcat",   "")
        date_str = item.get("date", "")
        if not date_str:
            continue
        try:
            d = datetime.strptime(date_str[:10], "%Y-%m-%d").date()
        except ValueError:
            continue

        if cat == "holiday" and subcat != "modern":
            # Check if it's an eve (erev) event
            title = item.get("title", "").lower()
            if "erev" in title or cat == "erev":
                eve_holidays.add(d)
            else:
                full_holidays.add(d)
        elif cat == "erev":
            eve_holidays.add(d)

    return frozenset(full_holidays), frozenset(eve_holidays)


def parse_vacation_file(file_buffer) -> frozenset:
    """Read dates from column A starting row 2 of an Excel file."""
    try:
        df = pd.read_excel(file_buffer, header=None, usecols=[0])
        dates = set()
        for val in df.iloc[1:, 0]:
            if pd.isna(val):
                continue
            if isinstance(val, (datetime, date)):
                d = val.date() if isinstance(val, datetime) else val
                dates.add(d)
            else:
                try:
                    d = pd.to_datetime(val).date()
                    dates.add(d)
                except Exception:
                    pass
        return frozenset(dates)
    except Exception:
        return frozenset()


def get_special_label(
    item_date: date,
    item_start: datetime,
    vacation_dates: frozenset,
    full_holidays: frozenset,
    eve_holidays: frozenset,
    friday_end_h: int,
) -> str | None:
    """
    Returns a special label if this date/time falls in a non-work period,
    or None if it's a normal working event.
    Priority: חופשה > חג > סוף שבוע
    """
    if item_date in vacation_dates:
        return "חופשה"

    weekday = item_date.weekday()  # 0=Mon … 4=Fri, 5=Sat, 6=Sun

    # Saturday (weekday 5) → always סוף שבוע
    if weekday == 5:
        return "סוף שבוע"

    # Full holiday (whole day off)
    if item_date in full_holidays:
        return "חג"

    # Friday after cutoff hour OR holiday eve after cutoff hour
    is_friday = (weekday == 4)
    is_eve    = (item_date in eve_holidays)
    if (is_friday or is_eve) and item_start.hour >= friday_end_h:
        return "סוף שבוע" if is_friday else "חג"

    return None


# ─── sidebar – parameters ────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ פרמטרים")

    c1, c2 = st.columns(2)
    work_start = c1.number_input(
        "תחילת עבודה", min_value=0, max_value=12, value=5, step=1, format="%d"
    )
    work_end = c2.number_input(
        "סיום עבודה", min_value=12, max_value=23, value=20, step=1, format="%d"
    )

    c1, c2 = st.columns(2)
    friday_end = c1.number_input(
        "סיום שישי/ערב חג", min_value=12, max_value=20, value=17, step=1, format="%d",
        help="שעה שממנה ואילך שישי וערבי חג מסווגים כסוף שבוע / חג"
    )
    commute_deduction = c2.number_input(
        "קיזוז נסיעה (דק')", min_value=0, max_value=60, value=0, step=5, format="%d",
        help="דקות הלוך + דקות חזור מנוכות מהנסיעה הראשונה/אחרונה של כל יום"
    )

    threshold = st.number_input(
        "סף פריקת מכולה (דק')", min_value=30, max_value=360, value=120, step=15,
        help="עצירה ארוכה יותר = 'עבודה'"
    )
    h, m = divmod(threshold, 60)
    st.caption(f"חלון עבודה: {work_start:02d}:00–{work_end:02d}:00  |  פריקה ≥ {h}ש' {m:02d}′  |  "
               f"סיום שישי: {friday_end:02d}:00  |  קיזוז: {commute_deduction}′ הלוך+חזור")

    include_holidays = st.checkbox("כלול חגי ישראל אוטומטית", value=True)

    st.divider()
    st.markdown("**🗓️ ימי חופשה**")
    if "vacation_dates_list" not in st.session_state:
        st.session_state.vacation_dates_list = []

    # בחירה מרובה: גרור טווח או Ctrl+לחיצה
    picked = st.date_input(
        "בחר תאריך/ים (טווח או יחיד)",
        value=(),
        format="DD/MM/YYYY",
        key="vac_date_picker",
        help="בחר תאריך בודד או גרור כדי לבחור טווח, ואז לחץ הוסף",
    )
    if st.button("➕ הוסף לרשימה", use_container_width=True):
        # picked יכול להיות date בודד, tuple של (start,end), או tuple ריק
        if picked:
            if isinstance(picked, tuple) and len(picked) == 2:
                # טווח — הוסף כל יום בין start ל-end
                from datetime import timedelta as _td
                start_d, end_d = picked
                cur = start_d
                while cur <= end_d:
                    if cur not in st.session_state.vacation_dates_list:
                        st.session_state.vacation_dates_list.append(cur)
                    cur += _td(days=1)
            else:
                d = picked[0] if isinstance(picked, tuple) else picked
                if d not in st.session_state.vacation_dates_list:
                    st.session_state.vacation_dates_list.append(d)
            st.rerun()

    if st.session_state.vacation_dates_list:
        for i, d in enumerate(sorted(st.session_state.vacation_dates_list)):
            c1, c2 = st.columns([4, 1])
            c1.caption(d.strftime("%d/%m/%Y"))
            if c2.button("✕", key=f"rm_vac_{i}"):
                st.session_state.vacation_dates_list.remove(d)
                st.rerun()
        if st.button("🗑️ נקה הכל", use_container_width=True):
            st.session_state.vacation_dates_list = []
            st.rerun()
    else:
        st.caption("לא נבחרו ימי חופשה")

    st.divider()
    st.caption("גרסה 1.3 | שלב 1")

# ─── file upload + analysis button ───────────────────────────────────────────

uploaded_ituran = st.file_uploader(
    "📂 דוח הנעה וכיבוי מאיתוראן",
    type=["xlsx", "xls"],
    help="ייצא מאיתוראן Online ← דוח הנעה וכיבוי ← Excel",
    key="ituran_file",
)

btn_analyze = st.button(
    "🔍 בצע ניתוח",
    type="primary",
    disabled=(uploaded_ituran is None),
)

# Clear cached results when a new file is uploaded
if uploaded_ituran is None:
    st.session_state.pop("analysis_result", None)

if btn_analyze and uploaded_ituran is not None:
    with st.spinner("מנתח את הדוח…"):
        try:
            buf, summary, stops, drives = analyze_to_buffer(
                uploaded_ituran, uploaded_ituran.name, threshold, work_start, work_end
            )
            st.session_state["analysis_result"] = {
                "buf": buf, "summary": summary,
                "stops": stops, "drives": drives,
                "filename": uploaded_ituran.name,
            }
        except Exception as e:
            st.error(f"שגיאה בניתוח הקובץ: {e}")
            st.session_state.pop("analysis_result", None)

if "analysis_result" not in st.session_state:
    if uploaded_ituran is None:
        st.info("העלה קובץ איתוראן כדי להתחיל בניתוח.")
    st.stop()

# ─── unpack results ───────────────────────────────────────────────────────────

res     = st.session_state["analysis_result"]
buf     = res["buf"]
stops   = res["stops"]
drives  = res["drives"]
filename = res["filename"]

st.success("הניתוח הושלם!")
st.divider()

# ─── build special-date sets ─────────────────────────────────────────────────

vacation_dates: frozenset = frozenset(
    st.session_state.get("vacation_dates_list", [])
)

# Collect years present in the data
all_dates = [s["date"] for s in stops] + [d["date"] for d in drives]
years = set(d.year for d in all_dates) if all_dates else {date.today().year}

full_holidays: frozenset = frozenset()
eve_holidays:  frozenset = frozenset()
if include_holidays:
    for yr in years:
        fh, eh = fetch_israel_holidays(yr)
        full_holidays = full_holidays | fh
        eve_holidays  = eve_holidays  | eh

SKIP_LABELS = {"חופשה", "חג", "סוף שבוע"}

def special_label(item_date, item_start):
    return get_special_label(
        item_date, item_start,
        vacation_dates, full_holidays, eve_holidays, friday_end
    )

HEBREW_DAYS = {0: "שני", 1: "שלישי", 2: "רביעי", 3: "חמישי", 4: "שישי", 5: "שבת", 6: "ראשון"}

# ─── classification labels & colors ──────────────────────────────────────────

LABEL_MAP = {
    ("stop",  TYPE_UNLOAD):    "עבודה",
    ("stop",  TYPE_TRANSPORT): "עצירת ביניים",
    ("stop",  TYPE_PARKING):   "חניה / שבת",
    ("stop",  TYPE_ANOMALY):   "חריג",
    ("drive", TYPE_TRANSPORT): "נסיעה",
    ("drive", TYPE_ANOMALY):   "נסיעה",
}
ROW_COLORS = {
    "עבודה":          "#dce8f5",
    "עצירת ביניים":   "#e8f5e8",
    "נסיעה":          "#e8f5e8",
    "חניה / שבת":     "#fafadc",
    "חריג":           "#ffe0e0",
    "סוף שבוע":       "#f0e8ff",
    "חג":             "#f0e8ff",
    "חופשה":          "#e8fff0",
}

# ─── commute deduction (skip special days) ───────────────────────────────────

def calc_drive_deductions(drives_list, commute_min):
    """Returns dict (date, start) -> deduction_hours, skipping special days."""
    if commute_min == 0:
        return {}
    # Only deduct on normal working days
    normal_drives = [
        d for d in drives_list
        if special_label(d["date"], d["start"]) not in SKIP_LABELS
    ]
    day_map = defaultdict(list)
    for d in normal_drives:
        day_map[d["date"]].append(d)
    result = {}
    for day_date, day_drives in day_map.items():
        sorted_d = sorted(day_drives, key=lambda x: x["start"])
        n    = len(sorted_d)
        durs = [d["duration"].total_seconds() / 60 for d in sorted_d]
        deds = [0.0] * n
        # Morning commute – deduct from first drive(s) forward
        rem = float(commute_min)
        for i in range(n):
            if rem <= 0: break
            take = min(rem, durs[i])
            deds[i] += take
            rem -= take
        # Evening commute – deduct from last drive(s) backward
        rem = float(commute_min)
        for i in range(n - 1, -1, -1):
            if rem <= 0: break
            avail = durs[i] - deds[i]
            take  = min(rem, max(0.0, avail))
            deds[i] += take
            rem -= take
        for i, d in enumerate(sorted_d):
            result[(day_date, d["start"])] = round(deds[i] / 60, 2)
    return result

drive_ded_map = calc_drive_deductions(drives, commute_deduction)

# ─── summary metrics ─────────────────────────────────────────────────────────

st.subheader("📊 סיכום")

def is_normal(item_date, item_start):
    return special_label(item_date, item_start) not in SKIP_LABELS

unload_stops    = [s for s in stops if s["type"] == TYPE_UNLOAD   and is_normal(s["date"], s["start"])]
transport_stops = [s for s in stops if s["type"] == TYPE_TRANSPORT and is_normal(s["date"], s["start"])]
anomaly_stops   = [s for s in stops if s["type"] == TYPE_ANOMALY]
anomaly_drives  = [d for d in drives if d["type"] == TYPE_ANOMALY]

total_unload_h    = sum(s["duration"].total_seconds() for s in unload_stops)    / 3600
total_transport_h = sum(s["duration"].total_seconds() for s in transport_stops) / 3600
total_drive_h     = sum(
    d["duration"].total_seconds() for d in drives
    if d["type"] == TYPE_TRANSPORT and is_normal(d["date"], d["start"])
) / 3600

gross_transport_h = total_transport_h + total_drive_h
total_deduction_h = sum(drive_ded_map.values())
net_transport_h   = max(0.0, gross_transport_h - total_deduction_h)

work_days = len(set(
    s["date"] for s in stops
    if s["type"] in (TYPE_UNLOAD, TYPE_TRANSPORT) and is_normal(s["date"], s["start"])
))

col1, col2, col3, col4 = st.columns(4)
col1.metric("שעות עבודה",  f"{total_unload_h:.2f} ש'",
            help="סה\"כ שעות עצירות ≥ סף הגדרה (ימי עבודה בלבד)")
col2.metric("שעות נסיעה",  f"{gross_transport_h:.2f} ש'",
            help="נסיעות + עצירות ביניים — לפני קיזוז (ימי עבודה בלבד)")
col3.metric("חריגים",
            f"{len(anomaly_stops) + len(anomaly_drives)}",
            help="פעילויות מחוץ לשעות העבודה",
            delta=None if not anomaly_stops else "לבדיקה",
            delta_color="inverse")
col4.metric("ימי עבודה", work_days)

if commute_deduction > 0:
    st.info(
        f"✂️ נסיעה נטו אחרי קיזוז: **{net_transport_h:.2f} ש'** "
        f"(קיזוז {commute_deduction} דק' הלוך + {commute_deduction} דק' חזור = "
        f"{total_deduction_h:.2f} ש' סה\"כ)"
    )

# ─── build detail rows ────────────────────────────────────────────────────────

all_items = [("stop", s) for s in stops] + [("drive", d) for d in drives]
all_items.sort(key=lambda x: x[1]["start"])

detail_rows = []
for kind, item in all_items:
    duration_h = round(item["duration"].total_seconds() / 3600, 2)
    sp_label   = special_label(item["date"], item["start"])
    label      = sp_label if sp_label else LABEL_MAP.get((kind, item["type"]), item["type"])

    if kind == "drive":
        frm  = item.get("from_address", "") or ""
        to   = item.get("to_address",   "") or ""
        addr = f"{frm} ← {to}" if frm or to else ""
    else:
        addr = item.get("address", "") or ""

    ded   = drive_ded_map.get((item["date"], item["start"]), 0.0) if kind == "drive" else 0.0
    net_h = round(max(0.0, duration_h - ded), 2)

    detail_rows.append({
        "שם עובד":       item["driver"],
        "תאריך":         item["date"].strftime("%d/%m/%Y"),
        "יום":           HEBREW_DAYS[item["date"].weekday()],
        "שעת התחלה":     item["start"].strftime("%H:%M"),
        "שעת סיום":      item["end"].strftime("%H:%M"),
        "משך (ש')":      duration_h,
        "קיזוז (ש')":    ded,
        "נטו (ש')":      net_h,
        "סיווג":         label,
        "כתובת / מסלול": addr,
        "_date_obj":     item["date"],   # for sorting (hidden)
    })

DETAIL_COLS = [
    "כתובת / מסלול", "נטו (ש')", "קיזוז (ש')", "משך (ש')",
    "סיווג", "שעת סיום", "שעת התחלה", "יום", "תאריך", "שם עובד",
]


def highlight_detail(row):
    if row.get("שם עובד") == 'סה"כ':
        return ["font-weight: bold; background-color: #e8f0fe"] * len(row)
    color = ROW_COLORS.get(row.get("סיווג", ""), "")
    return [f"background-color: {color}" if color else ""] * len(row)


# ─── daily summary table (default view) ───────────────────────────────────────

st.divider()
st.subheader("📅 סיכום יומי")

day_map_s: dict = {}
for row in detail_rows:
    key = (row["שם עובד"], row["תאריך"])
    if key not in day_map_s:
        day_map_s[key] = {
            "שם עובד":       row["שם עובד"],
            "תאריך":         row["תאריך"],
            "יום":           row["יום"],
            "עבודה (ש')":    0.0,
            "נסיעה נטו (ש')": 0.0,
            "סוג יום":       "",
            "_date_obj":     row["_date_obj"],
        }
    lbl = row["סיווג"]
    if lbl in SKIP_LABELS:
        day_map_s[key]["סוג יום"] = lbl
    elif lbl == "עבודה":
        day_map_s[key]["עבודה (ש')"]    = round(day_map_s[key]["עבודה (ש')"]    + row["נטו (ש')"], 2)
    elif lbl in ("נסיעה", "עצירת ביניים"):
        day_map_s[key]["נסיעה נטו (ש')"] = round(day_map_s[key]["נסיעה נטו (ש')"] + row["נטו (ש')"], 2)
    if not day_map_s[key]["סוג יום"]:
        day_map_s[key]["סוג יום"] = "יום עבודה"

summary_rows = sorted(day_map_s.values(), key=lambda x: x["_date_obj"])
SUMMARY_COLS = ["שם עובד", "תאריך", "יום", "עבודה (ש')", "נסיעה נטו (ש')", "סוג יום"]
df_day = pd.DataFrame(summary_rows)[SUMMARY_COLS]

def highlight_day(row):
    color = ROW_COLORS.get(row["סוג יום"], "")
    if row["סוג יום"] == "יום עבודה":
        color = "#dce8f5"
    return [f"background-color: {color}" if color else ""] * len(row)

st.dataframe(
    df_day.style.apply(highlight_day, axis=1)
               .format({"עבודה (ש')": "{:.2f}", "נסיעה נטו (ש')": "{:.2f}"}),
    use_container_width=True, hide_index=True,
)

# ─── detail table (collapsible) ───────────────────────────────────────────────

st.divider()
with st.expander("📋 פירוט מלא נסיעות ועצירות", expanded=False):
    # ── filters ──
    fc1, fc2, fc3 = st.columns(3)
    all_labels  = sorted(set(r["סיווג"]   for r in detail_rows))
    all_drivers = sorted(set(r["שם עובד"] for r in detail_rows))
    all_days    = ["ראשון","שני","שלישי","רביעי","חמישי","שישי","שבת"]

    sel_labels  = fc1.multiselect("סיווג",  all_labels,  default=all_labels,  key="flt_lbl")
    sel_drivers = fc2.multiselect("עובד",   all_drivers, default=all_drivers, key="flt_drv")
    sel_days    = fc3.multiselect("יום",    all_days,    default=all_days,    key="flt_day")

    filtered = [r for r in detail_rows
                if r["סיווג"]   in sel_labels
                and r["שם עובד"] in sel_drivers
                and r["יום"]     in sel_days]

    # Total row for filtered set
    ftotal = {
        "שם עובד": 'סה"כ', "תאריך": "", "יום": "", "שעת התחלה": "", "שעת סיום": "",
        "משך (ש')":   round(sum(r["משך (ש')"]   for r in filtered), 2),
        "קיזוז (ש')": round(sum(r["קיזוז (ש')"] for r in filtered), 2),
        "נטו (ש')":   round(sum(r["נטו (ש')"]   for r in filtered), 2),
        "סיווג": "", "כתובת / מסלול": "", "_date_obj": None,
    }
    filtered.append(ftotal)

    df_detail = pd.DataFrame(filtered)[DETAIL_COLS]
    st.dataframe(
        df_detail.style
                 .apply(highlight_detail, axis=1)
                 .format({"משך (ש')": "{:.2f}", "קיזוז (ש')": "{:.2f}", "נטו (ש')": "{:.2f}"}),
        use_container_width=True, hide_index=True,
    )
    st.caption(
        "🔵 כחול=עבודה  🟢 ירוק=נסיעה/עצירת ביניים  🟡 צהוב=חניה/שבת  "
        "🔴 אדום=חריג  🟣 סגול=סוף שבוע/חג  🟩 ירוק בהיר=חופשה"
    )

# ─── anomalies callout ───────────────────────────────────────────────────────

if anomaly_stops or anomaly_drives:
    st.divider()
    st.subheader("⚠️ חריגים שדורשים בירור")

    all_anomalies = sorted(
        [("עצירה", s) for s in anomaly_stops] +
        [("נסיעה", d) for d in anomaly_drives],
        key=lambda x: x[1]["start"]
    )

    anom_rows = []
    for kind, item in all_anomalies:
        mins = int(item["duration"].total_seconds() / 60)
        h2, m2 = divmod(mins, 60)
        if "from_address" in item:
            frm  = item.get("from_address", "") or ""
            to   = item.get("to_address",   "") or ""
            addr = f"{frm} ← {to}" if frm or to else "(נסיעה)"
        else:
            addr = item.get("address", "") or ""
        anom_rows.append({
            "כתובת / מסלול": addr,
            "סוג":           kind,
            "משך":           f"{h2}ש' {m2:02d}′",
            "שעת סיום":      item["end"].strftime("%H:%M"),
            "שעת התחלה":     item["start"].strftime("%H:%M"),
            "תאריך":         item["date"].strftime("%d/%m/%Y"),
        })

    st.dataframe(pd.DataFrame(anom_rows), use_container_width=True, hide_index=True)
    st.caption("הפעילויות הנ\"ל התרחשו מחוץ לשעות העבודה. יש לבדוק מול העובד.")

# ─── download ────────────────────────────────────────────────────────────────

st.divider()

base_name = filename.rsplit(".", 1)[0]
out_name  = f"{base_name}_ניתוח.xlsx"

st.download_button(
    label="📥 הורד קובץ ניתוח מלא (Excel)",
    data=buf,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "הקובץ כולל 4 גיליונות: סיכום · פירוט עצירות · חריגים לבירור · פרמטרים"
)
