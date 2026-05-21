"""
parking_db.py — שכבת הנתונים לאפליקציית ניהול חניות
Turso HTTP API — ללא קימפול
"""

import hashlib, requests as _req, streamlit as st
from datetime import date

# ─── Turso HTTP client ────────────────────────────────────────────────────────

class _Turso:
    def __init__(self, url, token):
        self.url  = url.replace("libsql://","https://").rstrip("/")
        self.hdrs = {"Authorization":f"Bearer {token}","Content-Type":"application/json"}

    def _arg(self, v):
        if v is None:          return {"type":"null"}
        if isinstance(v,int):  return {"type":"integer","value":str(v)}
        if isinstance(v,float):return {"type":"float",  "value":str(v)}
        return {"type":"text","value":str(v)}

    def run(self, stmts):
        payload = {"requests": stmts + [{"type":"close"}]}
        r = _req.post(f"{self.url}/v2/pipeline", json=payload,
                      headers=self.hdrs, timeout=15)
        r.raise_for_status()
        return r.json().get("results", [])

    def execute(self, sql, params=()):
        stmt = {"type":"execute","stmt":{"sql":sql,"args":[self._arg(p) for p in params]}}
        res  = self.run([stmt])
        return _Rows(res[0] if res else {})

    def script(self, sql):
        stmts = [{"type":"execute","stmt":{"sql":s.strip()}}
                 for s in sql.split(";") if s.strip()]
        if stmts: self.run(stmts)

    def commit(self): pass


class _Rows:
    def __init__(self, result):
        rs   = result.get("response",{}).get("result",{})
        cols = [c["name"] for c in rs.get("cols",[])]
        self.rows = [
            {cols[i]: (cell.get("value") if cell.get("type")!="null" else None)
             for i,cell in enumerate(row)}
            for row in rs.get("rows",[])
        ]
        self.rowcount = rs.get("affected_row_count", 0)

    def fetchone(self): return self.rows[0] if self.rows else None
    def fetchall(self): return self.rows
    def __iter__(self): return iter(self.rows)


@st.cache_resource
def _db():
    s = st.secrets["turso"]
    return _Turso(s["url"], s["token"])

# ─── עזרים ────────────────────────────────────────────────────────────────────

def _hash(pwd): return hashlib.sha256(pwd.strip().encode()).hexdigest()

def normalize_phone(p):
    """מסיר מקפים ורווחים ממספר טלפון."""
    return p.replace("-","").replace(" ","").strip()

ADMIN_EMAIL      = "dror.amitai@gmail.com"
ADMIN_PASS       = "DA1405"
DEFAULT_RES_PASS = "0258"

# ─── אתחול ────────────────────────────────────────────────────────────────────

def init_db():
    db = _db()
    db.script("""
        CREATE TABLE IF NOT EXISTS users (
            phone      TEXT PRIMARY KEY,
            password   TEXT NOT NULL,
            role       TEXT NOT NULL DEFAULT 'resident',
            name       TEXT NOT NULL DEFAULT '',
            apartment  TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS phone_whitelist (
            phone TEXT PRIMARY KEY,
            note  TEXT
        );
        CREATE TABLE IF NOT EXISTS requests (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_phone    TEXT NOT NULL,
            requester_name TEXT NOT NULL DEFAULT '',
            requester_apt  TEXT NOT NULL DEFAULT '',
            date           TEXT NOT NULL,
            time_from      TEXT NOT NULL,
            time_to        TEXT NOT NULL,
            spots_needed   INTEGER NOT NULL DEFAULT 1,
            spots_filled   INTEGER NOT NULL DEFAULT 0,
            status         TEXT NOT NULL DEFAULT 'open',
            created_at     TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS offers (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            request_id   INTEGER NOT NULL,
            owner_phone  TEXT NOT NULL,
            owner_name   TEXT NOT NULL DEFAULT '',
            owner_apt    TEXT NOT NULL DEFAULT '',
            spot_number  TEXT NOT NULL,
            offer_note   TEXT,
            status       TEXT NOT NULL DEFAULT 'active',
            created_at   TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        )
    """)
    # אדמין ראשוני
    db.execute(
        "INSERT OR IGNORE INTO users (phone,password,role,name) VALUES (?,?,?,?)",
        (ADMIN_EMAIL, _hash(ADMIN_PASS), "admin", "מנהל מערכת")
    )
    db.execute(
        "INSERT OR IGNORE INTO settings (key,value) VALUES ('default_password',?)",
        (_hash(DEFAULT_RES_PASS),)
    )
    db.execute(
        "INSERT OR IGNORE INTO settings (key,value) VALUES ('whitelist_mode','open')",
    )

# ─── הגדרות ───────────────────────────────────────────────────────────────────

def get_setting(key):
    row = _db().execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else ""

def set_setting(key, value):
    _db().execute("INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)", (key, value))

def get_whitelist_mode():
    return get_setting("whitelist_mode") or "open"

def set_whitelist_mode(mode):
    set_setting("whitelist_mode", mode)

# ─── אימות ────────────────────────────────────────────────────────────────────

def authenticate(identifier: str, password: str):
    identifier = normalize_phone(identifier) if "@" not in identifier else identifier.strip()
    h  = _hash(password)
    db = _db()
    row = db.execute(
        "SELECT phone,role,name,apartment FROM users WHERE phone=? AND password=?",
        (identifier, h)
    ).fetchone()
    if row:
        return {"phone":row["phone"],"role":row["role"],
                "name":row["name"],"apartment":row["apartment"]}
    # דייר חדש — רק אם סיסמת ברירת מחדל
    if "@" not in identifier:
        default_h = get_setting("default_password")
        if h == default_h:
            db.execute(
                "INSERT OR IGNORE INTO users (phone,password,role) VALUES (?,?,?)",
                (identifier, h, "resident")
            )
            return {"phone":identifier,"role":"resident","name":"","apartment":""}
    return None

def is_whitelisted(phone: str) -> bool:
    if get_whitelist_mode() == "open":
        return True
    row = _db().execute(
        "SELECT phone FROM phone_whitelist WHERE phone=?", (phone,)
    ).fetchone()
    return row is not None

def change_password(phone, new_password):
    _db().execute("UPDATE users SET password=? WHERE phone=?", (_hash(new_password), phone))

def reset_password_to_default(phone):
    _db().execute("UPDATE users SET password=? WHERE phone=?",
                  (get_setting("default_password"), phone))

def change_default_password(new_password):
    set_setting("default_password", _hash(new_password))

# ─── דיירים ───────────────────────────────────────────────────────────────────

def get_all_users():
    rows = _db().execute(
        "SELECT phone,role,name,apartment,created_at FROM users "
        "WHERE role='resident' ORDER BY created_at"
    ).fetchall()
    return rows   # already list of dicts

def get_whitelist():
    return _db().execute("SELECT phone,note FROM phone_whitelist ORDER BY phone").fetchall()

def add_to_whitelist(phone, note=""):
    _db().execute("INSERT OR IGNORE INTO phone_whitelist (phone,note) VALUES (?,?)", (phone, note))

def remove_from_whitelist(phone):
    _db().execute("DELETE FROM phone_whitelist WHERE phone=?", (phone,))

def bulk_add_whitelist(phones: list) -> tuple:
    db = _db()
    added = skipped = 0
    for p in phones:
        p = normalize_phone(p)
        if not p: continue
        ex = db.execute("SELECT phone FROM phone_whitelist WHERE phone=?", (p,)).fetchone()
        if ex: skipped += 1
        else:
            db.execute("INSERT INTO phone_whitelist (phone) VALUES (?)", (p,))
            added += 1
    return added, skipped

# ─── auto-expiry ──────────────────────────────────────────────────────────────

def expire_old_records():
    _db().execute(
        "UPDATE requests SET status='expired' WHERE date<? AND status='open'",
        (date.today().isoformat(),)
    )

# ─── בקשות ───────────────────────────────────────────────────────────────────

def add_request(phone, name, apt, req_date, t_from, t_to, spots):
    if t_from >= t_to:    return False, "שעת הסיום חייבת להיות אחרי שעת ההתחלה"
    if not name.strip():  return False, "נא למלא שם"
    if not apt.strip():   return False, "נא למלא מספר דירה"
    db = _db()
    if db.execute("SELECT id FROM requests WHERE owner_phone=? AND date=? AND status='open'",
                  (phone, req_date)).fetchone():
        return False, "כבר פרסמת בקשה פתוחה לתאריך זה"
    db.execute(
        "INSERT INTO requests (owner_phone,requester_name,requester_apt,"
        "date,time_from,time_to,spots_needed) VALUES (?,?,?,?,?,?,?)",
        (phone, name.strip(), apt.strip(), req_date, t_from, t_to, spots)
    )
    return True, ""

def get_open_requests():
    return _db().execute(
        "SELECT id,owner_phone,requester_name,requester_apt,date,"
        "time_from,time_to,spots_needed,spots_filled,status,created_at "
        "FROM requests WHERE status='open' AND date>=? ORDER BY date,time_from",
        (date.today().isoformat(),)
    ).fetchall()

def get_my_requests(phone):
    return _db().execute(
        "SELECT id,owner_phone,requester_name,requester_apt,date,"
        "time_from,time_to,spots_needed,spots_filled,status,created_at "
        "FROM requests WHERE owner_phone=? AND date>=? ORDER BY date",
        (phone, date.today().isoformat())
    ).fetchall()

def cancel_request(req_id, phone):
    _db().execute(
        "UPDATE requests SET status='cancelled' WHERE id=? AND owner_phone=? AND status='open'",
        (req_id, phone)
    )

# ─── הצעות ───────────────────────────────────────────────────────────────────

def _overlap(f1,t1,f2,t2): return f1<t2 and f2<t1

def offer_spot(req_id, owner_phone, owner_name, owner_apt, spot_num, note=""):
    if not spot_num.strip():   return False, "חובה למלא מספר חניה"
    if not owner_name.strip(): return False, "חובה למלא שם"
    if not owner_apt.strip():  return False, "חובה למלא מספר דירה"
    db = _db()
    req = db.execute(
        "SELECT id,owner_phone,date,time_from,time_to,spots_needed,spots_filled "
        "FROM requests WHERE id=? AND status='open'", (req_id,)
    ).fetchone()
    if not req:                           return False, "הבקשה כבר טופלה"
    if req["owner_phone"] == owner_phone: return False, "אי אפשר להציע חניה לעצמך"
    if db.execute("SELECT id FROM offers WHERE request_id=? AND owner_phone=? AND status='active'",
                  (req_id, owner_phone)).fetchone():
        return False, "כבר הצעת לבקשה זו"
    for o in db.execute(
        "SELECT r.time_from,r.time_to FROM offers o JOIN requests r ON o.request_id=r.id "
        "WHERE o.owner_phone=? AND o.status='active' AND r.date=? AND r.id!=?",
        (owner_phone, req["date"], req_id)
    ).fetchall():
        if _overlap(req["time_from"],req["time_to"],o["time_from"],o["time_to"]):
            return False, f"כבר הצעת חניה לבקשה אחרת בשעות {o['time_from']}–{o['time_to']}"
    db.execute(
        "INSERT INTO offers (request_id,owner_phone,owner_name,owner_apt,spot_number,offer_note) "
        "VALUES (?,?,?,?,?,?)",
        (req_id, owner_phone, owner_name.strip(), owner_apt.strip(), spot_num.strip(), note)
    )
    db.execute("UPDATE requests SET spots_filled=spots_filled+1 WHERE id=?", (req_id,))
    updated = db.execute(
        "SELECT spots_needed,spots_filled FROM requests WHERE id=?", (req_id,)
    ).fetchone()
    if int(updated["spots_filled"]) >= int(updated["spots_needed"]):
        db.execute("UPDATE requests SET status='fulfilled' WHERE id=?", (req_id,))
    return True, ""

def get_offers_for_request(req_id):
    return _db().execute(
        "SELECT id,request_id,owner_phone,owner_name,owner_apt,"
        "spot_number,offer_note,status,created_at "
        "FROM offers WHERE request_id=? AND status='active' ORDER BY created_at",
        (req_id,)
    ).fetchall()

def withdraw_offer(req_id, owner_phone):
    db = _db()
    req = db.execute("SELECT status FROM requests WHERE id=?", (req_id,)).fetchone()
    if not req or req["status"] != "open":
        return False, "לא ניתן לבטל — הבקשה כבר לא פתוחה"
    db.execute("UPDATE offers SET status='withdrawn' "
               "WHERE request_id=? AND owner_phone=? AND status='active'",
               (req_id, owner_phone))
    db.execute("UPDATE requests SET spots_filled=spots_filled-1 WHERE id=?", (req_id,))
    return True, ""

def get_my_offers(phone):
    return _db().execute(
        "SELECT o.id AS offer_id, o.request_id, o.status AS offer_status, "
        "o.spot_number, o.offer_note, "
        "r.owner_phone AS requester_phone, r.requester_name, r.requester_apt, "
        "r.date, r.time_from, r.time_to, r.spots_needed, r.spots_filled, "
        "r.status AS req_status "
        "FROM offers o JOIN requests r ON o.request_id=r.id "
        "WHERE o.owner_phone=? AND r.date>=? ORDER BY r.date,r.time_from",
        (phone, date.today().isoformat())
    ).fetchall()

# ─── סטטיסטיקות ──────────────────────────────────────────────────────────────

def get_summary():
    today = date.today().isoformat()
    db = _db()
    o = db.execute("SELECT COUNT(*) AS cnt FROM requests WHERE status='open' AND date>=?",
                   (today,)).fetchone()
    f = db.execute("SELECT COUNT(*) AS cnt FROM requests WHERE status='fulfilled'").fetchone()
    return {"open_requests": int(o["cnt"]) if o else 0,
            "fulfilled_total": int(f["cnt"]) if f else 0}
