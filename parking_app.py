"""
parking_app.py — אפליקציית ניהול חניות לדיירי הבניין
אימות: טלפון + סיסמה | פאנל אדמין | Turso cloud DB
"""

import streamlit as st
import urllib.parse
import os
from datetime import date, timedelta, time

import parking_db as db

# ─── הגדרות ──────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="חניות הבניין",
    page_icon="🅿️",
    layout="centered",
    initial_sidebar_state="expanded",
)

APP_URL = os.environ.get("APP_URL", "http://localhost:8502")

# ─── CSS בסיסי ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
  body, .stApp, [class*="css"] { direction: rtl; text-align: right; }
  .stButton > button { width: 100%; }
  .privacy-note {
    background:#f3e5f5; border-radius:8px;
    padding:6px 12px; font-size:0.85em; color:#6a1b9a;
    margin-top:4px; display:inline-block;
  }
  .offer-box {
    background:#e8f5e9; border-radius:8px; padding:10px 14px; margin-top:6px;
  }
</style>
""", unsafe_allow_html=True)

db.init_db()
db.expire_old_records()

# ─── עזרים ───────────────────────────────────────────────────────────────────

DAYS_HE = ["שני","שלישי","רביעי","חמישי","שישי","שבת","ראשון"]

def fmt_date(d):
    try:
        dt = date.fromisoformat(d)
        return f"{dt.day}/{dt.month}/{dt.year} (יום {DAYS_HE[dt.weekday()]})"
    except Exception:
        return d

def fmt_time(tf, tt):
    return f"{tf}–{tt}"

def wa_share_url(name, apt, req_date, time_from, time_to, spots_needed):
    spots_txt = "חניה אחת" if spots_needed == 1 else f"{spots_needed} חניות"
    lines = [
        "🅿️ *בקשת חניה חדשה בבניין*",
        f"🏠 דירה {apt} — {name}",
        f"📅 {fmt_date(req_date)}",
        f"🕐 {fmt_time(time_from, time_to)}  ·  {spots_txt}",
        "",
        "━━━━━━━━━━━━━━━",
        "👆 מי שפנויה לו החניה — לחץ כאן לתת:",
        APP_URL,
    ]
    return f"https://wa.me/?text={urllib.parse.quote(chr(10).join(lines))}"


# ════════════════════════════════════════════════════════════════════════════
# מסך כניסה
# ════════════════════════════════════════════════════════════════════════════

if "user" not in st.session_state:
    st.session_state.user = None
if "login_mode" not in st.session_state:
    st.session_state.login_mode = None   # None | "resident" | "admin"

if not st.session_state.user:

    st.markdown("""
    <style>
      [data-testid="stSidebar"]  { display:none !important; }
      [data-testid="stHeader"]   { display:none !important; }
      [data-testid="stToolbar"]  { display:none !important; }
      .block-container { padding:0 !important; max-width:100% !important; }

      /* ─── כותרת כחולה ─── */
      .login-top {
        background:#1a73e8; padding:52px 24px 48px;
        text-align:center; color:white;
      }
      .login-top .icon { font-size:60px; margin-bottom:10px; }
      .login-top h1   { font-size:28px; font-weight:700; margin-bottom:4px; }
      .login-top .sub { font-size:16px; opacity:0.88; }

      /* ─── כרטיס לבן ─── */
      .login-card {
        background:white; border-radius:24px 24px 0 0;
        padding:32px 24px 40px; margin-top:-18px; min-height:58vh;
      }

      /* ─── כפתורי בחירה גדולים ─── */
      div[data-testid="stVerticalBlock"] .stButton.login-btn > button {
        background:white !important; color:#1a73e8 !important;
        font-size:18px !important; font-weight:700 !important;
        padding:18px 12px !important; border-radius:14px !important;
        border:2px solid #1a73e8 !important;
        margin-bottom:14px !important;
      }
      div[data-testid="stVerticalBlock"] .stButton.login-btn > button:hover {
        background:#e8f0fe !important;
      }

      /* ─── כפתור כניסה ראשי ─── */
      .login-card .stButton.primary-btn > button {
        background:#1a73e8 !important; color:white !important;
        font-size:18px !important; font-weight:700 !important;
        padding:18px !important; border-radius:12px !important;
        border:none !important; margin-top:8px !important;
      }
      .login-card .stButton.primary-btn > button:active { background:#1558b0 !important; }

      /* ─── כפתור חזרה ─── */
      .login-card .stButton.back-btn > button {
        background:transparent !important; color:#555 !important;
        font-size:15px !important; border:1px solid #ccc !important;
        border-radius:10px !important; padding:10px !important;
        margin-top:4px !important;
      }

      /* ─── שדות קלט ─── */
      .login-card input {
        font-size:16px !important; border-radius:12px !important;
        padding:14px 12px !important;
      }
    </style>

    <div class="login-top">
      <div class="icon">🅿️</div>
      <h1>חניות הבניין</h1>
      <p class="sub">כניסה למערכת</p>
    </div>
    <div class="login-card">
    """, unsafe_allow_html=True)

    mode = st.session_state.login_mode

    # ── בחירת סוג כניסה ──────────────────────────────────────────────────────
    if mode is None:
        st.markdown("#### בחר סוג כניסה")
        st.markdown("")

        col_r, col_a = st.columns(2)
        with col_r:
            if st.button("🏠 כניסת דיירים", key="btn_resident", use_container_width=True):
                st.session_state.login_mode = "resident"
                st.rerun()
        with col_a:
            if st.button("⚙️ כניסת מנהל", key="btn_admin", use_container_width=True):
                st.session_state.login_mode = "admin"
                st.rerun()

    # ── טופס כניסת דיירים ────────────────────────────────────────────────────
    elif mode == "resident":
        st.markdown("#### 🏠 כניסת דיירים")
        with st.form("resident_login"):
            phone = st.text_input("מספר טלפון", placeholder="05X-XXXXXXX")
            pwd   = st.text_input("סיסמה", type="password", placeholder="••••••")

            submitted = st.form_submit_button("כניסה", type="primary", use_container_width=True)
            if submitted:
                if not phone.strip() or not pwd.strip():
                    st.error("נא למלא טלפון וסיסמה")
                else:
                    norm = db.normalize_phone(phone)
                    user = db.authenticate(norm, pwd)
                    if user:
                        if user["role"] == "resident" and not db.is_whitelisted(norm):
                            st.error("מספר הטלפון אינו מורשה — פנה למנהל הבניין")
                        else:
                            st.session_state.user = user
                            st.session_state.login_mode = None
                            st.rerun()
                    else:
                        st.error("טלפון או סיסמה שגויים")

        if st.button("← חזרה", key="back_resident", use_container_width=True):
            st.session_state.login_mode = None
            st.rerun()

    # ── טופס כניסת מנהל ──────────────────────────────────────────────────────
    elif mode == "admin":
        st.markdown("#### ⚙️ כניסת מנהל")
        with st.form("admin_login"):
            email = st.text_input("שם משתמש", placeholder="admin@example.com")
            pwd   = st.text_input("סיסמה", type="password", placeholder="••••••")

            submitted = st.form_submit_button("כניסה", type="primary", use_container_width=True)
            if submitted:
                if not email.strip() or not pwd.strip():
                    st.error("נא למלא שם משתמש וסיסמה")
                else:
                    user = db.authenticate(email.strip(), pwd)
                    if user and user["role"] == "admin":
                        st.session_state.user = user
                        st.session_state.login_mode = None
                        st.rerun()
                    else:
                        st.error("שם משתמש או סיסמה שגויים")

        if st.button("← חזרה", key="back_admin", use_container_width=True):
            st.session_state.login_mode = None
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# ─── משתמש מחובר ─────────────────────────────────────────────────────────────

user = st.session_state.user
me   = user["phone"]
role = user["role"]

# ─── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("🅿️ חניות הבניין")
    display_name = user.get("name") or me
    st.success(f"👋 שלום, {display_name}")

    if st.button("יציאה"):
        st.session_state.user = None
        st.session_state.login_mode = None
        if "last_request" in st.session_state:
            del st.session_state["last_request"]
        st.rerun()

    st.divider()

    if role == "admin":
        page = st.radio("ניווט", [
            "🏠 לוח בקשות", "📝 הבקשות שלי", "🤝 החניות שנתתי", "⚙️ ניהול"
        ], label_visibility="collapsed")
    else:
        page = st.radio("ניווט", [
            "🏠 לוח בקשות", "📝 הבקשות שלי", "🤝 החניות שנתתי"
        ], label_visibility="collapsed")


# ════════════════════════════════════════════════════════════════════════════
# דף 1 — לוח בקשות
# ════════════════════════════════════════════════════════════════════════════

if page == "🏠 לוח בקשות":
    st.header("🏠 לוח בקשות חניה")
    st.caption("בקשות שעדיין לא קיבלו מענה מלא. לחץ על בקשה כדי לתת חניה.")

    summary = db.get_summary()
    c1, c2  = st.columns(2)
    c1.metric("בקשות פתוחות", summary["open_requests"])
    c2.metric('נענו עד כה', summary["fulfilled_total"])
    st.divider()

    # ── טופס פרסום בקשה ──────────────────────────────────────────────────────
    with st.expander("➕ אני צריך חניה — פרסם בקשה", expanded=False):
        with st.form("new_request"):
            default_name = st.session_state.get("my_name", user.get("name",""))
            default_apt  = st.session_state.get("my_apt",  user.get("apartment",""))

            col_n, col_a = st.columns(2)
            with col_n:
                req_name = st.text_input("שם מלא *", value=default_name, placeholder="משפחת כהן")
            with col_a:
                req_apt  = st.text_input("מספר דירה *", value=default_apt, placeholder="42")

            col1, col2 = st.columns(2)
            with col1:
                req_date = st.date_input("תאריך",
                    value=date.today()+timedelta(days=1),
                    min_value=date.today(),
                    max_value=date.today()+timedelta(days=60),
                    format="DD/MM/YYYY")
            with col2:
                spots = st.number_input("כמה חניות?", min_value=1, max_value=5, value=1, step=1)

            col3, col4 = st.columns(2)
            with col3:
                t_from = st.time_input("משעה", value=time(9,0), step=1800)
            with col4:
                t_to   = st.time_input("עד שעה", value=time(13,0), step=1800)

            time_ok = t_to > t_from
            if not time_ok:
                st.error("⚠️ שעת הסיום חייבת להיות אחרי שעת ההתחלה")

            if st.form_submit_button("📢 פרסם בקשה", type="primary", disabled=not time_ok):
                ok, msg = db.add_request(
                    me, req_name, req_apt, req_date.isoformat(),
                    t_from.strftime("%H:%M"), t_to.strftime("%H:%M"), int(spots)
                )
                if ok:
                    st.session_state["my_name"] = req_name.strip()
                    st.session_state["my_apt"]  = req_apt.strip()
                    st.session_state["last_request"] = {
                        "name": req_name.strip(), "apt": req_apt.strip(),
                        "date": req_date.isoformat(),
                        "time_from": t_from.strftime("%H:%M"),
                        "time_to":   t_to.strftime("%H:%M"),
                        "spots": int(spots),
                    }
                    st.rerun()
                else:
                    st.error(msg)

    # כפתור שיתוף WhatsApp אחרי פרסום
    if "last_request" in st.session_state:
        lr = st.session_state["last_request"]
        wa_url = wa_share_url(lr["name"], lr["apt"], lr["date"],
                              lr["time_from"], lr["time_to"], lr["spots"])
        st.success(
            f"✅ הבקשה פורסמה!  📅 {fmt_date(lr['date'])}  ·  "
            f"🕐 {fmt_time(lr['time_from'], lr['time_to'])}"
        )
        st.link_button("📲 שתף עכשיו בקבוצת הוואטסאפ", wa_url, type="primary")
        st.caption("לחיצה תפתח את WhatsApp עם הודעה מוכנה — שלח לקבוצה.")
        if st.button("סגור", key="close_share"):
            del st.session_state["last_request"]
            st.rerun()
        st.divider()

    # ── רשימת בקשות ──────────────────────────────────────────────────────────
    my_active_rids = {o["request_id"] for o in db.get_my_offers(me)
                      if o["offer_status"] == "active"}
    open_reqs = db.get_open_requests()

    if not open_reqs:
        st.info("🎉 אין בקשות פתוחות כרגע!")
    else:
        default_owner_name = st.session_state.get("my_name", user.get("name",""))
        default_owner_apt  = st.session_state.get("my_apt",  user.get("apartment",""))

        for req in open_reqs:
            is_mine   = req["owner_phone"] == me
            i_offered = req["id"] in my_active_rids
            needed    = req["spots_needed"]
            filled    = req["spots_filled"]
            left      = needed - filled

            label = (
                f"דירה {req['requester_apt']} — {req['requester_name']}  ·  "
                f"📅 {fmt_date(req['date'])}  ·  "
                f"🕐 {fmt_time(req['time_from'], req['time_to'])}"
                + (f"  ·  {filled}/{needed} נענו" if needed > 1 else "")
            )

            with st.expander(label, expanded=False):
                if is_mine:
                    st.caption("זו הבקשה שלך")
                    if st.button("❌ בטל בקשה", key=f"cancel_b_{req['id']}", type="secondary"):
                        db.cancel_request(req["id"], me)
                        st.rerun()

                elif i_offered:
                    st.success("✔ כבר הצעת חניה לבקשה זו")
                    if st.button("בטל הצעה", key=f"wd_b_{req['id']}", type="secondary"):
                        ok, msg = db.withdraw_offer(req["id"], me)
                        if ok: st.rerun()
                        else:  st.error(msg)

                elif left > 0:
                    st.markdown(f"**רוצה לתת חניה ל{req['requester_name']} (דירה {req['requester_apt']})?**")
                    with st.form(key=f"offer_{req['id']}"):
                        oc1, oc2 = st.columns(2)
                        with oc1:
                            o_name = st.text_input("שמך", value=default_owner_name, placeholder="ישראל ישראלי")
                        with oc2:
                            o_apt  = st.text_input("דירה שלך", value=default_owner_apt, placeholder="15")
                        spot_num = st.text_input("מספר החניה שלך", placeholder="15 / B12 / מינוס 2 שמאל")
                        cond     = st.text_input("הערות (אופציונלי)", placeholder='לדוגמה: "חסום אותי", "פנה עד 23:00"')
                        if st.form_submit_button("✅ תן חניה", type="primary"):
                            ok, msg = db.offer_spot(req["id"], me, o_name, o_apt, spot_num, cond)
                            if ok:
                                st.session_state["my_name"] = o_name.strip()
                                st.session_state["my_apt"]  = o_apt.strip()
                                st.success("✅ נרשמת!")
                                st.rerun()
                            else:
                                st.error(msg)
                else:
                    st.success("✅ הבקשה מולאה במלואה")


# ════════════════════════════════════════════════════════════════════════════
# דף 2 — הבקשות שלי
# ════════════════════════════════════════════════════════════════════════════

elif page == "📝 הבקשות שלי":
    st.header("📝 הבקשות שלי")
    my_reqs = db.get_my_requests(me)

    if not my_reqs:
        st.info("לא פרסמת בקשות עדיין.")
    else:
        for req in my_reqs:
            status = req["status"]
            needed = req["spots_needed"]
            filled = req["spots_filled"]
            icons  = {"open":"⏳","fulfilled":"✅","expired":"🕐","cancelled":"❌"}

            with st.expander(
                f"{icons.get(status,'')}  {fmt_date(req['date'])}  ·  "
                f"{fmt_time(req['time_from'], req['time_to'])}  ·  {filled}/{needed} נענו",
                expanded=(status in ("open","fulfilled"))
            ):
                st.caption(f"🏠 דירה {req['requester_apt']} — {req['requester_name']}")

                offers = db.get_offers_for_request(req["id"])
                if offers:
                    st.divider()
                    st.markdown("**פרטי החניות שקיבלת:**")
                    for o in offers:
                        st.markdown(
                            f'<div class="offer-box">'
                            f'🏠 <b>{o["owner_name"]}</b> דירה {o["owner_apt"]} '
                            f'&nbsp;·&nbsp; 🅿️ חניה <b>{o["spot_number"]}</b>'
                            + (f'<br>📌 {o["offer_note"]}' if o.get("offer_note") else "")
                            + f'<br>📞 <a href="tel:{o["owner_phone"]}">{o["owner_phone"]}</a>'
                            + "</div>",
                            unsafe_allow_html=True
                        )
                    st.markdown(
                        '<div class="privacy-note">'
                        '🔒 פרטים אלו גלויים לך בלבד'
                        '</div>', unsafe_allow_html=True
                    )
                elif status == "open":
                    st.info("⏳ ממתין — עדיין אין מי שהציע")

                if status == "open":
                    wa_url = wa_share_url(
                        req["requester_name"], req["requester_apt"],
                        req["date"], req["time_from"], req["time_to"], req["spots_needed"]
                    )
                    st.link_button("📲 שתף שוב בוואטסאפ", wa_url)
                    st.divider()
                    if st.button("❌ בטל בקשה", key=f"cancel_{req['id']}", type="secondary"):
                        db.cancel_request(req["id"], me)
                        st.rerun()


# ════════════════════════════════════════════════════════════════════════════
# דף 3 — החניות שנתתי
# ════════════════════════════════════════════════════════════════════════════

elif page == "🤝 החניות שנתתי":
    st.header("🤝 החניות שנתתי")
    my_offers = db.get_my_offers(me)

    if not my_offers:
        st.info("לא הצעת חניה לאף אחד עדיין.")
    else:
        active    = [o for o in my_offers if o["offer_status"] == "active"]
        withdrawn = [o for o in my_offers if o["offer_status"] == "withdrawn"]

        if active:
            for o in active:
                with st.container(border=True):
                    c1, c2 = st.columns([4,1])
                    with c1:
                        st.markdown(
                            f"**נתת חניה {o['spot_number']} ל{o['requester_name']} "
                            f"(דירה {o['requester_apt']})**"
                        )
                        st.caption(f"📅 {fmt_date(o['date'])}  ·  🕐 {fmt_time(o['time_from'], o['time_to'])}")
                        if o.get("offer_note"):
                            st.caption(f"📌 {o['offer_note']}")
                        st.markdown(
                            f'📞 <a href="tel:{o["requester_phone"]}">{o["requester_phone"]}</a>',
                            unsafe_allow_html=True
                        )
                        if o["req_status"] == "fulfilled":
                            st.success("✅ הבקשה מולאה")
                        else:
                            st.info(f"⏳ {o['spots_filled']}/{o['spots_needed']} נענו")
                    with c2:
                        if o["req_status"] == "open":
                            if st.button("בטל", key=f"wd_{o['offer_id']}", type="secondary"):
                                ok, msg = db.withdraw_offer(o["request_id"], me)
                                if ok: st.rerun()
                                else:  st.error(msg)

        if withdrawn:
            with st.expander(f"הצעות שבוטלו ({len(withdrawn)})"):
                for o in withdrawn:
                    st.caption(f"{o['requester_name']} · {fmt_date(o['date'])} · {fmt_time(o['time_from'],o['time_to'])}")


# ════════════════════════════════════════════════════════════════════════════
# דף 4 — ניהול (אדמין בלבד)
# ════════════════════════════════════════════════════════════════════════════

elif page == "⚙️ ניהול" and role == "admin":
    st.header("⚙️ פאנל ניהול")

    tab1, tab2, tab3 = st.tabs(["👥 דיירים", "🔑 סיסמאות", "📋 רשימת טלפונים"])

    # ── דיירים ──────────────────────────────────────────────────────────────
    with tab1:
        st.subheader("דיירים רשומים")
        users = db.get_all_users()
        if not users:
            st.info("אין דיירים רשומים עדיין.")
        else:
            for u in users:
                with st.container(border=True):
                    c1, c2 = st.columns([4,1])
                    with c1:
                        label = u["name"] or "(ללא שם)"
                        st.markdown(f"**{label}**  ·  דירה {u['apartment'] or '?'}")
                        st.caption(f"📞 {u['phone']}")
                    with c2:
                        if st.button("איפוס סיסמה", key=f"rst_{u['phone']}", type="secondary"):
                            db.reset_password_to_default(u["phone"])
                            st.success("סיסמה אופסה לברירת המחדל")

    # ── סיסמאות ─────────────────────────────────────────────────────────────
    with tab2:
        st.subheader("סיסמת ברירת מחדל לדיירים חדשים")
        with st.form("change_default"):
            new_def  = st.text_input("סיסמה חדשה", type="password")
            new_def2 = st.text_input("אישור סיסמה", type="password")
            if st.form_submit_button("עדכן", type="primary"):
                if not new_def:
                    st.error("נא למלא סיסמה")
                elif new_def != new_def2:
                    st.error("הסיסמאות אינן תואמות")
                else:
                    db.change_default_password(new_def)
                    st.success("✅ סיסמת ברירת המחדל עודכנה")

        st.divider()
        st.subheader("שינוי סיסמת דייר ספציפי")
        with st.form("change_user_pwd"):
            phone_to_change = st.text_input("טלפון הדייר")
            new_pwd  = st.text_input("סיסמה חדשה", type="password")
            new_pwd2 = st.text_input("אישור", type="password")
            if st.form_submit_button("עדכן", type="primary"):
                if not phone_to_change or not new_pwd:
                    st.error("נא למלא את כל השדות")
                elif new_pwd != new_pwd2:
                    st.error("הסיסמאות אינן תואמות")
                else:
                    db.change_password(phone_to_change, new_pwd)
                    st.success(f"✅ סיסמה עודכנה לדייר {phone_to_change}")

    # ── רשימת טלפונים מורשים ────────────────────────────────────────────────
    with tab3:
        st.subheader("מצב גישה למערכת")

        current_mode = db.get_whitelist_mode()
        new_mode = st.radio(
            "מי יכול להיכנס למערכת?",
            options=["open", "restricted"],
            format_func=lambda m: "🌐 כל אחד יכול להשתמש" if m == "open" else "🔒 רק מספרים ברשימה",
            index=0 if current_mode == "open" else 1,
            horizontal=True,
        )
        if new_mode != current_mode:
            db.set_whitelist_mode(new_mode)
            st.success("✅ ההגדרה עודכנה")
            st.rerun()

        if new_mode == "open":
            st.info("כרגע **כל אחד** עם מספר טלפון וסיסמה נכונה יכול להיכנס.")
        else:
            st.warning("כניסה מוגבלת — **רק** מספרים הרשומים ברשימה למטה.")

        st.divider()
        st.subheader("רשימת טלפונים מורשים")

        # ── הטענת קובץ (bulk) ────────────────────────────────────────────────
        with st.expander("📂 ייבוא רשימה מקובץ"):
            st.caption("קובץ טקסט — מספר טלפון בכל שורה (מקפים מותרים).")
            uploaded = st.file_uploader("בחר קובץ .txt", type=["txt","csv"])
            if uploaded:
                content = uploaded.read().decode("utf-8", errors="ignore")
                phones  = [ln.strip() for ln in content.splitlines() if ln.strip()]
                if phones:
                    st.write(f"נמצאו **{len(phones)}** מספרים בקובץ.")
                    if st.button("ייבא לרשימה", type="primary"):
                        added, skipped = db.bulk_add_whitelist(phones)
                        st.success(f"✅ נוספו {added} מספרים. {skipped} כבר היו ברשימה.")
                        st.rerun()

        # ── רשימה קיימת ──────────────────────────────────────────────────────
        whitelist = db.get_whitelist()
        if whitelist:
            for w in whitelist:
                c1, c2 = st.columns([4,1])
                c1.write("📞 " + w["phone"] + (f"  —  {w['note']}" if w.get("note") else ""))
                if c2.button("הסר", key=f"rm_{w['phone']}"):
                    db.remove_from_whitelist(w["phone"])
                    st.rerun()
        else:
            st.caption("הרשימה ריקה")

        # ── הוספה ידנית ──────────────────────────────────────────────────────
        with st.form("add_whitelist"):
            col1, col2 = st.columns(2)
            with col1:
                new_phone = st.text_input("טלפון")
            with col2:
                note = st.text_input("הערה (אופציונלי)", placeholder="דירה 42")
            if st.form_submit_button("הוסף"):
                if new_phone.strip():
                    db.add_to_whitelist(db.normalize_phone(new_phone), note.strip())
                    st.rerun()
