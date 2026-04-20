"""
Ituran Report Analyzer – Streamlit UI
Phase 1: Ituran GPS report → classified Excel output
"""

import streamlit as st
import pandas as pd
from datetime import date
from ituran_analyzer import (
    analyze_to_buffer,
    TYPE_UNLOAD, TYPE_TRANSPORT, TYPE_PARKING, TYPE_ANOMALY,
)

# ─── page config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ניתוח דוחות איתוראן",
    page_icon="🚗",
    layout="centered",
)

# ─── password gate ───────────────────────────────────────────────────────────

import os
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

# Force RTL for the whole page
st.markdown("""
<style>
    body, .stApp { direction: rtl; text-align: right; }
    .stDownloadButton button { width: 100%; }
    .metric-box {
        background: #f0f4ff;
        border-radius: 8px;
        padding: 12px 18px;
        margin: 4px 0;
    }
    .anomaly-box {
        background: #fff0f0;
        border: 1px solid #ffb3b3;
        border-radius: 8px;
        padding: 12px 18px;
        margin: 8px 0;
    }
</style>
""", unsafe_allow_html=True)

# ─── header ──────────────────────────────────────────────────────────────────

st.title("🚗 ניתוח דוחות איתוראן")
st.markdown("**הפרדה בין שעות עבודה בשטח לבין שעות נסיעה**")
st.divider()

# ─── sidebar – parameters ────────────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ פרמטרים")

    st.subheader("שעות עבודה")
    work_start = st.number_input(
        "שעת תחילת עבודה",
        min_value=0, max_value=12, value=5, step=1,
        format="%d",
        help="שעה שלמה (לדוגמה: 5 = 05:00)"
    )
    work_end = st.number_input(
        "שעת סיום עבודה",
        min_value=12, max_value=23, value=20, step=1,
        format="%d",
        help="שעה שלמה (לדוגמה: 20 = 20:00)"
    )
    st.caption(f"חלון עבודה: {work_start:02d}:00 – {work_end:02d}:00")

    st.divider()
    st.subheader("פריקת מכולות")
    threshold = st.number_input(
        "זמן תקן לפריקת מכולה (דקות)",
        min_value=30, max_value=360, value=120, step=15,
        help="עצירה ארוכה יותר מערך זה תסווג כ'פריקת מכולות'"
    )
    h, m = divmod(threshold, 60)
    st.caption(f"= {h} שעות {m:02d} דקות")

    st.divider()
    st.subheader("קיזוז נסיעה יומי")
    commute_deduction = st.number_input(
        "קיזוז לכיוון (דקות)",
        min_value=0, max_value=60, value=0, step=5,
        help="דקות שמקזזים בגין יציאה לעבודה וחזרה הביתה. הקיזוז הכולל ליום = פי 2 מהערך הזה"
    )
    if commute_deduction > 0:
        st.caption(f"קיזוז יומי: {commute_deduction * 2} דקות ({commute_deduction} הלוך + {commute_deduction} חזור)")

    st.divider()
    st.markdown("**חניה/שבת:** עצירה >12ש' או לילה >5ש'")
    st.markdown("גרסה 1.1 | שלב 1")

# ─── file upload ─────────────────────────────────────────────────────────────

st.subheader("📂 העלאת קובץ")
uploaded = st.file_uploader(
    "העלה דוח הנעה וכיבוי מאיתוראן (קובץ Excel)",
    type=["xlsx", "xls"],
    help="ייצא את הדוח מאיתוראן Online ← דוח הנעה וכיבוי ← Excel"
)

# ─── analysis ────────────────────────────────────────────────────────────────

if uploaded is None:
    st.info("העלה קובץ איתוראן כדי להתחיל בניתוח.")
    st.stop()

with st.spinner("מנתח את הדוח…"):
    try:
        buf, summary, stops, drives = analyze_to_buffer(
            uploaded, uploaded.name, threshold, work_start, work_end
        )
    except Exception as e:
        st.error(f"שגיאה בניתוח הקובץ: {e}")
        st.stop()

st.success("הניתוח הושלם!")
st.divider()

# ─── summary metrics ─────────────────────────────────────────────────────────

st.subheader("📊 סיכום")

unload_stops  = [s for s in stops if s["type"] == TYPE_UNLOAD]
transport_stops = [s for s in stops if s["type"] == TYPE_TRANSPORT]
parking_stops = [s for s in stops if s["type"] == TYPE_PARKING]
anomaly_stops = [s for s in stops if s["type"] == TYPE_ANOMALY]
anomaly_drives = [d for d in drives if d["type"] == TYPE_ANOMALY]

total_unload_h    = sum(s["duration"].total_seconds() for s in unload_stops)   / 3600
total_transport_h = sum(s["duration"].total_seconds() for s in transport_stops) / 3600
total_drive_h     = sum(d["duration"].total_seconds() for d in drives
                        if d["type"] == TYPE_TRANSPORT) / 3600

gross_transport_h = total_transport_h + total_drive_h
# ימים שיש בהם נסיעות בפועל (לא רק עצירות קצרות) — רק הם מקבלים קיזוז
days_with_drives  = set(d["date"] for d in drives if d["type"] == TYPE_TRANSPORT)
total_deduction_h = (commute_deduction * 2 / 60) * len(days_with_drives)
net_transport_h   = max(0.0, gross_transport_h - total_deduction_h)

col1, col2, col3, col4 = st.columns(4)
col1.metric("פריקת מכולות", f"{total_unload_h:.2f} ש'",
            help="סה\"כ שעות בעצירות ≥ סף הגדרה")
col2.metric("הסעות עובדים", f"{gross_transport_h:.2f} ש'",
            help="נסיעות + עצירות קצרות בשעות עבודה — לפני קיזוז")
col3.metric("חריגים", f"{len(anomaly_stops) + len(anomaly_drives)}",
            help="פעילויות מחוץ לשעות 05:00-20:00",
            delta=None if not anomaly_stops else "לבדיקה",
            delta_color="inverse")
col4.metric("ימי עבודה", len(set(s["date"] for s in stops
                                  if s["type"] in (TYPE_UNLOAD, TYPE_TRANSPORT))))
if commute_deduction > 0:
    st.info(f"✂️ הסעות נטו אחרי קיזוז: **{net_transport_h:.2f} ש'** "
            f"(קיזוז {commute_deduction} דק' הלוך + {commute_deduction} דק' חזור × "
            f"{len(days_with_drives)} ימי נסיעה = {total_deduction_h:.2f} ש' סה\"כ)")

# ─── summary table ───────────────────────────────────────────────────────────

st.divider()
st.subheader("📋 פירוט נסיעות ועצירות")

# נסיעה ראשונה ואחרונה לכל יום — מכל הסיווגים (כולל חריגים)
from collections import defaultdict
_day_starts = defaultdict(list)
for d in drives:
    _day_starts[d["date"]].append(d["start"])
day_first = {dt: min(v) for dt, v in _day_starts.items()}
day_last  = {dt: max(v) for dt, v in _day_starts.items()}

# צבעי רקע לפי סיווג
ROW_COLORS = {
    TYPE_UNLOAD:    "#dce8f5",
    TYPE_TRANSPORT: "#e8f5e8",
    TYPE_PARKING:   "#fafadc",
    TYPE_ANOMALY:   "#ffe0e0",
}

# בניית שורות — שורה לכל נסיעה/עצירה
all_items = [("stop", s) for s in stops] + [("drive", d) for d in drives]
all_items.sort(key=lambda x: x[1]["start"])

detail_rows = []
for kind, item in all_items:
    duration_h = round(item["duration"].total_seconds() / 3600, 2)

    if kind == "drive":
        frm  = item.get("from_address", "") or ""
        to   = item.get("to_address",   "") or ""
        addr = f"{frm} ← {to}" if frm or to else ""
    else:
        addr = item.get("address", "") or ""

    # קיזוז — נסיעה ראשונה/אחרונה ביום מכל סוג
    ded = 0.0
    if kind == "drive" and commute_deduction > 0:
        is_first = item["start"] == day_first.get(item["date"])
        is_last  = item["start"] == day_last.get(item["date"])
        if is_first and is_last:          # יום עם נסיעה אחת בלבד
            ded = round(commute_deduction * 2 / 60, 2)
        elif is_first or is_last:
            ded = round(commute_deduction / 60, 2)

    net_h = round(max(0.0, duration_h - ded), 2)

    detail_rows.append({
        "שם עובד":       item["driver"],
        "תאריך":         item["date"].strftime("%d/%m/%Y"),
        "שעת התחלה":     item["start"].strftime("%H:%M"),
        "שעת סיום":      item["end"].strftime("%H:%M"),
        "משך (ש')":      duration_h,
        "קיזוז (ש')":    ded,
        "נטו (ש')":      net_h,
        "סיווג":         item["type"],
        "כתובת / מסלול": addr,
    })

# שורת סה"כ
total_row = {
    "שם עובד": 'סה"כ', "תאריך": "", "שעת התחלה": "", "שעת סיום": "",
    "משך (ש')":   round(sum(r["משך (ש')"]   for r in detail_rows), 2),
    "קיזוז (ש')": round(sum(r["קיזוז (ש')"] for r in detail_rows), 2),
    "נטו (ש')":   round(sum(r["נטו (ש')"]   for r in detail_rows), 2),
    "סיווג": "", "כתובת / מסלול": "",
}
detail_rows.append(total_row)

DETAIL_COLS = [
    "כתובת / מסלול", "נטו (ש')", "קיזוז (ש')", "משך (ש')",
    "סיווג", "שעת סיום", "שעת התחלה", "תאריך", "שם עובד",
]
df_detail = pd.DataFrame(detail_rows)[DETAIL_COLS]


def highlight_detail(row):
    if row["שם עובד"] == 'סה"כ':
        return ["font-weight: bold; background-color: #e8f0fe"] * len(row)
    color = ROW_COLORS.get(row["סיווג"], "")
    return [f"background-color: {color}" if color else ""] * len(row)


st.dataframe(
    df_detail.style.apply(highlight_detail, axis=1)
                   .format({"משך (ש')": "{:.2f}", "קיזוז (ש')": "{:.2f}", "נטו (ש')": "{:.2f}"}),
    width="stretch",
    hide_index=True,
)
st.caption("🔴 אדום=חריג  🔵 כחול=פריקת מכולות  🟢 ירוק=נסיעה  🟡 צהוב=חניה/שבת")

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
        h, m = divmod(mins, 60)
        if "from_address" in item:
            frm = item.get("from_address", "") or ""
            to  = item.get("to_address",   "") or ""
            addr = f"{frm} ← {to}" if frm or to else "(נסיעה)"
        else:
            addr = item.get("address", "") or ""
        anom_rows.append({
            "כתובת / מסלול": addr,
            "סוג":           kind,
            "משך":           f"{h}ש' {m:02d}′",
            "שעת סיום":      item["end"].strftime("%H:%M"),
            "שעת התחלה":     item["start"].strftime("%H:%M"),
            "תאריך":         item["date"].strftime("%d/%m/%Y"),
        })

    st.dataframe(pd.DataFrame(anom_rows), width="stretch", hide_index=True)
    st.caption("הפעילויות הנ\"ל התרחשו מחוץ לשעות 05:00-20:00. יש לבדוק מול העובד.")

# ─── download ────────────────────────────────────────────────────────────────

st.divider()

base_name = uploaded.name.rsplit(".", 1)[0]
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
