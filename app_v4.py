#!/usr/bin/env python3
"""
Store Mission Lookup v4
New in v4:
  - Excel (.xlsx / .xls) file support
  - File type validation with clear error messages
  - Missing column / data validation
  - Resizable columns in Activity History
  - Dark / Light mode toggle
  - VACUUM on delete (DB file shrinks after deletions)
  - Import DB file from another machine
  - Fuzzy search / match
"""

import os, csv, sqlite3, threading, webbrowser, time, json, socket, shutil
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

# ── Optional Excel support ────────────────────────────────────────────────────
try:
    import openpyxl
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    import xlrd
    XLS_OK = True
except ImportError:
    XLS_OK = False

# ── Config ────────────────────────────────────────────────────────────────────
APPDATA       = os.environ.get("APPDATA", os.path.expanduser("~"))
CONFIG_DIR    = os.path.join(APPDATA, "StoreMissionLookup")
SETTINGS_FILE = os.path.join(CONFIG_DIR, "settings.json")
DEFAULT_DB    = os.path.join(CONFIG_DIR, "data.db")
VERSION       = "4.0"
MACHINE       = socket.gethostname()
UNDO_TTL      = 30
HEARTBEAT_S   = 20
POLL_S        = 25
STALE_S       = 60
FUZZY_THRESH  = 55      # minimum fuzzy score (0-100) to show a result

SUPPORTED_EXT = {".csv", ".txt", ".xlsx", ".xls"}

os.makedirs(CONFIG_DIR, exist_ok=True)

# ── Settings ──────────────────────────────────────────────────────────────────
def load_settings():
    try:
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    except Exception:
        return {"db_path": DEFAULT_DB,
                "user_name": os.environ.get("USERNAME", "User"),
                "auto_refresh": 0,
                "appearance": "dark",
                "fuzzy": False}

def save_settings(s):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(s, f, indent=2)

def is_valid_db_path(path):
    if not path: return False
    if path.startswith("http://") or path.startswith("https://"): return False
    if not path.endswith(".db"): return False
    return os.path.isabs(path) or path.startswith("\\\\")

def get_db_path():
    path = load_settings().get("db_path", DEFAULT_DB)
    if not is_valid_db_path(path):
        s = load_settings(); s["db_path"] = DEFAULT_DB; save_settings(s)
        return DEFAULT_DB
    return path

def get_user():         return load_settings().get("user_name", os.environ.get("USERNAME", "User"))
def get_auto_refresh(): return load_settings().get("auto_refresh", 0)
def get_fuzzy():        return load_settings().get("fuzzy", False)

# ── Database ──────────────────────────────────────────────────────────────────
def get_conn():
    path = get_db_path()
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    except Exception:
        path = DEFAULT_DB
        os.makedirs(os.path.dirname(path), exist_ok=True)
    conn = sqlite3.connect(path, timeout=15)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=10000")
    return conn

def init_db():
    with get_conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS entries (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                store    TEXT NOT NULL,
                store_lc TEXT NOT NULL,
                mission  TEXT NOT NULL,
                link     TEXT DEFAULT '',
                dedup    TEXT UNIQUE NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_slc ON entries(store_lc);
            CREATE INDEX IF NOT EXISTS idx_mis ON entries(mission);
            CREATE TABLE IF NOT EXISTS import_history (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                filename     TEXT NOT NULL,
                imported_at  TEXT NOT NULL,
                imported_by  TEXT NOT NULL,
                rows_added   INTEGER DEFAULT 0,
                rows_skipped INTEGER DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS activity_log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                action       TEXT NOT NULL,
                detail       TEXT NOT NULL,
                performed_at TEXT NOT NULL,
                performed_by TEXT NOT NULL,
                count        INTEGER DEFAULT 0,
                extra        TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS online_users (
                username  TEXT NOT NULL,
                machine   TEXT NOT NULL,
                last_seen TEXT NOT NULL,
                PRIMARY KEY (username, machine)
            );
            CREATE TABLE IF NOT EXISTS db_locks (
                id        INTEGER PRIMARY KEY,
                locked_by TEXT NOT NULL,
                locked_at TEXT NOT NULL,
                action    TEXT NOT NULL
            );
        """)

# ── Fuzzy search ──────────────────────────────────────────────────────────────
def fuzzy_score(query: str, text: str) -> int:
    """
    Returns 0-100. 100 = exact substring match.
    Uses sequential character matching for typo tolerance.
    """
    q = query.lower().strip()
    t = text.lower()
    if not q: return 0
    if q in t: return 100
    # Sequential char match — tolerates missing / transposed chars
    qi = 0
    consecutive = 0
    bonus = 0
    for ch in t:
        if qi < len(q) and ch == q[qi]:
            qi += 1
            consecutive += 1
            bonus += consecutive  # reward consecutive matches
        else:
            consecutive = 0
    if qi == 0: return 0
    base = int((qi / len(q)) * 75)
    return min(99, base + min(bonus, 24))

def db_get_stores_fuzzy(q: str, mf: str = ""):
    """Return stores sorted by fuzzy score descending."""
    if not q:
        return db_get_stores("", mf)
    with get_conn() as c:
        if mf:
            rows = c.execute(
                "SELECT DISTINCT store FROM entries WHERE mission=? ORDER BY store COLLATE NOCASE", (mf,)
            ).fetchall()
        else:
            rows = c.execute(
                "SELECT DISTINCT store FROM entries ORDER BY store COLLATE NOCASE"
            ).fetchall()
    scored = []
    for (store,) in rows:
        s = fuzzy_score(q, store)
        if s >= FUZZY_THRESH:
            scored.append((s, store))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [store for _, store in scored]

# ── Core DB queries ───────────────────────────────────────────────────────────
def db_get_stores(q="", mf=""):
    with get_conn() as c:
        if mf:
            return [r[0] for r in c.execute(
                "SELECT DISTINCT store FROM entries WHERE store_lc LIKE ? AND mission=? ORDER BY store COLLATE NOCASE",
                (f"%{q.lower()}%", mf)).fetchall()]
        return [r[0] for r in c.execute(
            "SELECT DISTINCT store FROM entries WHERE store_lc LIKE ? ORDER BY store COLLATE NOCASE",
            (f"%{q.lower()}%",)).fetchall()]

def db_get_missions(store):
    with get_conn() as c:
        return c.execute(
            "SELECT mission, link FROM entries WHERE store_lc=? ORDER BY mission",
            (store.lower(),)).fetchall()

def db_get_rows_for_stores(store_names):
    if not store_names: return []
    lc = [s.lower() for s in store_names]
    rows = []
    conn = get_conn()
    try:
        for i in range(0, len(lc), 900):
            ch = lc[i:i+900]
            ph = ",".join("?"*len(ch))
            rows += conn.execute(
                f"SELECT store,store_lc,mission,link,dedup FROM entries WHERE store_lc IN ({ph})", ch
            ).fetchall()
    finally:
        conn.close()
    return rows

def db_stats():
    with get_conn() as c:
        t = c.execute("SELECT COUNT(*) FROM entries").fetchone()[0]
        s = c.execute("SELECT COUNT(DISTINCT store_lc) FROM entries").fetchone()[0]
        m = c.execute("SELECT COUNT(DISTINCT mission) FROM entries").fetchone()[0]
    return s, m, t

def db_all_missions():
    with get_conn() as c:
        return [r[0] for r in c.execute(
            "SELECT DISTINCT mission FROM entries ORDER BY mission").fetchall()]

def db_delete_stores(store_names):
    if not store_names: return 0
    lc = [s.lower() for s in store_names]
    total = 0
    conn = get_conn()
    try:
        cur = conn.cursor()
        for i in range(0, len(lc), 900):
            ch = lc[i:i+900]
            ph = ",".join("?"*len(ch))
            cur.execute(f"DELETE FROM entries WHERE store_lc IN ({ph})", ch)
            total += cur.rowcount
        conn.commit()
    finally:
        conn.close()
    # VACUUM: reclaim freed space so the .db file actually shrinks
    _vacuum_db()
    return total

def _vacuum_db():
    """Run VACUUM in a separate connection to shrink the DB file after deletions."""
    try:
        path = get_db_path()
        conn = sqlite3.connect(path, timeout=30)
        conn.execute("VACUUM")
        conn.close()
    except Exception:
        pass

def db_restore_rows(rows):
    if not rows: return 0
    restored = 0
    conn = get_conn()
    try:
        cur = conn.cursor()
        for store, store_lc, mission, link, dedup in rows:
            try:
                cur.execute(
                    "INSERT INTO entries (store,store_lc,mission,link,dedup) VALUES (?,?,?,?,?)",
                    (store, store_lc, mission, link, dedup))
                restored += 1
            except sqlite3.IntegrityError:
                pass
        conn.commit()
    finally:
        conn.close()
    return restored

def db_insert_batch(rows):
    added = skipped = 0
    conn = get_conn()
    try:
        cur = conn.cursor()
        for store, mission, link in rows:
            try:
                cur.execute(
                    "INSERT INTO entries (store,store_lc,mission,link,dedup) VALUES (?,?,?,?,?)",
                    (store, store.lower(), mission, link,
                     store.lower() + "|||" + mission.lower()))
                added += 1
            except sqlite3.IntegrityError:
                skipped += 1
        conn.commit()
    finally:
        conn.close()
    return added, skipped

# ── Activity log ──────────────────────────────────────────────────────────────
def db_log_import(filename, added, skipped):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = get_user()
    with get_conn() as c:
        c.execute(
            "INSERT INTO import_history (filename,imported_at,imported_by,rows_added,rows_skipped) VALUES (?,?,?,?,?)",
            (filename, now, user, added, skipped))
        c.execute(
            "INSERT INTO activity_log (action,detail,performed_at,performed_by,count,extra) VALUES (?,?,?,?,?,?)",
            ("Import", filename, now, user, added, f"{skipped} duplicates skipped"))

def db_log_delete(store_names, rows_deleted):
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = get_user()
    n    = len(store_names)
    detail = store_names[0] if n == 1 else f"{n:,} stores"
    with get_conn() as c:
        c.execute(
            "INSERT INTO activity_log (action,detail,performed_at,performed_by,count,extra) VALUES (?,?,?,?,?,?)",
            ("Delete", detail, now, user, rows_deleted, f"{n} store{'s' if n>1 else ''} removed"))

def db_log_undo(store_names, rows_restored):
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user = get_user()
    n    = len(store_names)
    detail = store_names[0] if n == 1 else f"{n:,} stores"
    with get_conn() as c:
        c.execute(
            "INSERT INTO activity_log (action,detail,performed_at,performed_by,count,extra) VALUES (?,?,?,?,?,?)",
            ("Undo", detail, now, user, rows_restored, f"{n} store{'s' if n>1 else ''} restored"))

def db_get_history(order_col="performed_at", order_dir="DESC"):
    safe_cols = {"performed_at","action","detail","performed_by","count","extra"}
    safe_dirs = {"ASC","DESC"}
    if order_col not in safe_cols: order_col = "performed_at"
    if order_dir not in safe_dirs: order_dir = "DESC"
    with get_conn() as c:
        return c.execute(
            f"SELECT action,detail,performed_at,performed_by,count,extra "
            f"FROM activity_log ORDER BY {order_col} {order_dir} LIMIT 500"
        ).fetchall()

def db_get_latest_activity_id():
    with get_conn() as c:
        r = c.execute("SELECT MAX(id) FROM activity_log").fetchone()
    return r[0] or 0

def db_get_new_activity(since_id):
    with get_conn() as c:
        return c.execute(
            "SELECT action,detail,performed_at,performed_by,count,extra "
            "FROM activity_log WHERE id > ? ORDER BY id ASC", (since_id,)).fetchall()

# ── Online users ──────────────────────────────────────────────────────────────
def db_heartbeat():
    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_conn() as c:
            c.execute(
                "INSERT OR REPLACE INTO online_users (username,machine,last_seen) VALUES (?,?,?)",
                (get_user(), MACHINE, now))
    except Exception:
        pass

def db_get_online_users():
    try:
        cutoff = (datetime.now() - timedelta(seconds=STALE_S)).strftime("%Y-%m-%d %H:%M:%S")
        with get_conn() as c:
            return [r[0] for r in c.execute(
                "SELECT username FROM online_users WHERE last_seen > ? ORDER BY username",
                (cutoff,)).fetchall()]
    except Exception:
        return []

def db_go_offline():
    try:
        with get_conn() as c:
            c.execute("DELETE FROM online_users WHERE username=? AND machine=?",
                      (get_user(), MACHINE))
    except Exception:
        pass

# ── Import lock ───────────────────────────────────────────────────────────────
def db_acquire_lock(action="Import"):
    try:
        stale = (datetime.now() - timedelta(minutes=10)).strftime("%Y-%m-%d %H:%M:%S")
        conn = get_conn()
        conn.execute("DELETE FROM db_locks WHERE locked_at < ?", (stale,))
        conn.commit()
        existing = conn.execute("SELECT locked_by, action FROM db_locks").fetchone()
        if existing:
            conn.close()
            return False, f"{existing[0]} ({existing[1]})"
        conn.execute(
            "INSERT INTO db_locks (locked_by,locked_at,action) VALUES (?,?,?)",
            (get_user(), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), action))
        conn.commit()
        conn.close()
        return True, None
    except Exception:
        return True, None

def db_release_my_lock():
    try:
        with get_conn() as c:
            c.execute("DELETE FROM db_locks WHERE locked_by=?", (get_user(),))
    except Exception:
        pass

# ── File reading helpers ──────────────────────────────────────────────────────
ALLOWED_EXT = {".csv", ".txt", ".xlsx", ".xls"}

def validate_file_type(path: str) -> tuple[bool, str]:
    """Returns (ok, error_message)."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in ALLOWED_EXT:
        return False, (
            f"File type '{ext}' is not supported.\n\n"
            "Please upload a CSV (.csv) or Excel (.xlsx / .xls) file."
        )
    if ext in (".xlsx", ".xls") and not EXCEL_OK:
        return False, (
            "Excel support requires the openpyxl package.\n\n"
            "Install it by running:\n  pip install openpyxl\n\nThen restart the app."
        )
    return True, ""

def read_file_rows(path: str) -> tuple[list[str], list[list[str]], str]:
    """
    Returns (headers, data_rows, error).
    data_rows is a list of string lists (all values as strings).
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".csv", ".txt"):
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                reader = csv.reader(f)
                headers = [h.strip() for h in next(reader, [])]
                rows    = [[str(c).strip() for c in row] for row in reader]
            return headers, rows, ""

        elif ext == ".xlsx":
            wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws   = wb.active
            data = list(ws.iter_rows(values_only=True))
            wb.close()
            if not data:
                return [], [], "The Excel file appears to be empty."
            headers = [str(c).strip() if c is not None else "" for c in data[0]]
            rows    = [[str(c).strip() if c is not None else "" for c in row]
                       for row in data[1:]]
            return headers, rows, ""

        elif ext == ".xls":
            import xlrd
            wb   = xlrd.open_workbook(path)
            ws   = wb.sheet_by_index(0)
            if ws.nrows == 0:
                return [], [], "The Excel file appears to be empty."
            headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
            rows    = [[str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)]
                       for r in range(1, ws.nrows)]
            return headers, rows, ""

    except Exception as e:
        return [], [], f"Could not read file: {e}"

    return [], [], "Unsupported file type."

def detect_columns(headers: list[str]) -> tuple[int, int, int]:
    """Auto-detect store / mission / link column indices. Returns (-1,-1,-1) if not found."""
    lh = [h.lower() for h in headers]
    sc = next((i for i, h in enumerate(lh) if "store"   in h), -1)
    mc = next((i for i, h in enumerate(lh) if "mission" in h
               and "link" not in h and "url" not in h), -1)
    lc = next((i for i, h in enumerate(lh) if "link"    in h or "url" in h), -1)
    return sc, mc, lc

def validate_required_columns(sc, mc, lc, headers) -> str:
    """Returns an error message if any required column is missing, else ''."""
    missing = []
    if sc == -1: missing.append("Store Details")
    if mc == -1: missing.append("Mission Name")
    if lc == -1: missing.append("Mission Link")
    if not missing:
        return ""
    # Build shared values used in both branches
    cols_str = ", ".join(headers) if headers else "(none)"
    if len(missing) == 1:
        return (f"Required column not found: \"{missing[0]}\"\n\n"
                f"Columns detected in your file:\n  {cols_str}\n\n"
                "Please map the columns manually using the dropdowns below.")
    quoted = ", ".join('"' + m + '"' for m in missing)
    return (f"Required columns not found: {quoted}\n\n"
            f"Columns detected in your file:\n  {cols_str}\n\n"
            "Please map the columns manually using the dropdowns below.")

# ── Main Application ──────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Store Mission Lookup  v{VERSION}")
        # Centre window on screen at launch
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = min(1440, sw - 80), 720
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(900, 560)
        self._search_after    = None
        self._stores          = []
        self._sel_store       = None
        self._undo_buffer     = None
        self._undo_after      = None
        self._undo_tick_after = None
        self._last_activity_id = 0
        self._auto_after      = None
        self._stop_threads    = threading.Event()
        self._child_windows   = {}
        self._build_ui()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(200, self._startup)

    # ── Startup ───────────────────────────────────────────────────────────────
    def _startup(self):
        init_db()
        db_heartbeat()
        try:
            self._last_activity_id = db_get_latest_activity_id()
        except Exception:
            pass
        self._refresh()
        self._start_background_threads()
        self._schedule_auto_refresh()

    def _on_close(self):
        self._stop_threads.set()
        db_go_offline()
        db_release_my_lock()
        self.destroy()

    # ── Background threads ────────────────────────────────────────────────────
    def _start_background_threads(self):
        threading.Thread(target=self._heartbeat_loop, daemon=True).start()
        threading.Thread(target=self._poll_loop,       daemon=True).start()

    def _heartbeat_loop(self):
        while not self._stop_threads.wait(HEARTBEAT_S):
            db_heartbeat()
            self.after(0, self._update_online_bar)

    def _poll_loop(self):
        while not self._stop_threads.wait(POLL_S):
            self.after(0, self._check_for_changes)

    def _check_for_changes(self):
        try:
            new_rows = db_get_new_activity(self._last_activity_id)
            for action, detail, performed_at, performed_by, count, extra in new_rows:
                if performed_by != get_user():
                    icon  = "⬆" if action == "Import" else ("↩" if action == "Undo" else "🗑")
                    short = detail if len(detail) <= 40 else detail[:38] + "…"
                    self._notify(f"{icon} {performed_by}  {action.lower()}ed: {short} — click Refresh")
            if new_rows:
                self._last_activity_id = db_get_latest_activity_id()
            self._update_online_bar()
        except Exception:
            pass

    def _schedule_auto_refresh(self):
        if self._auto_after:
            self.after_cancel(self._auto_after)
            self._auto_after = None
        mins = get_auto_refresh()
        if mins and mins > 0:
            self._auto_after = self.after(mins * 60 * 1000, self._auto_refresh_tick)

    def _auto_refresh_tick(self):
        self._refresh()
        self._schedule_auto_refresh()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = ctk.CTkFrame(self, corner_radius=0, fg_color="#111111")
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text=f"STORE MISSION LOOKUP  v{VERSION}",
                     font=ctk.CTkFont(family="Courier New", size=13, weight="bold"),
                     text_color="#c8f060").pack(side="left", padx=16, pady=10)
        self.lbl_stats = ctk.CTkLabel(hdr, text="",
                                       font=ctk.CTkFont(family="Courier New", size=11),
                                       text_color="#666")
        self.lbl_stats.pack(side="right", padx=8)

        # Online bar
        self.online_bar = ctk.CTkFrame(self, corner_radius=0, fg_color="#0d1a00", height=26)
        self.online_bar.pack(fill="x")
        self.online_bar.pack_propagate(False)
        self.lbl_online = ctk.CTkLabel(self.online_bar, text="",
                                        font=ctk.CTkFont(family="Courier New", size=11),
                                        text_color="#4a7a20")
        self.lbl_online.pack(side="left", padx=12)
        self.lbl_lock = ctk.CTkLabel(self.online_bar, text="",
                                      font=ctk.CTkFont(family="Courier New", size=11),
                                      text_color="#f0c060")
        self.lbl_lock.pack(side="right", padx=12)

        # Notification bar
        self.notif_bar = ctk.CTkFrame(self, corner_radius=0, fg_color="#1a1200", height=28)
        self.notif_bar.pack_propagate(False)
        self.lbl_notif = ctk.CTkLabel(self.notif_bar, text="",
                                       font=ctk.CTkFont(size=12), text_color="#f0c060")
        self.lbl_notif.pack(side="left", padx=12, pady=2)
        _btn(self.notif_bar, "✕", self._dismiss_notif, h=20, w=28, fs=11
             ).pack(side="right", padx=8, pady=4)

        # Toolbar
        tb = ctk.CTkFrame(self, corner_radius=0, fg_color="#161616")
        tb.pack(fill="x")
        self.var_search = ctk.StringVar()
        self.var_search.trace_add("write", self._schedule_search)
        self._search_entry = ctk.CTkEntry(
            tb, textvariable=self.var_search, width=300,
            placeholder_text="🔍  Search store…  (Ctrl+F)",
            font=ctk.CTkFont(size=13))
        self._search_entry.pack(side="left", padx=(12,6), pady=9)

        # Fuzzy toggle
        self.var_fuzzy = ctk.BooleanVar(value=get_fuzzy())
        self.chk_fuzzy = ctk.CTkCheckBox(
            tb, text="Fuzzy", variable=self.var_fuzzy,
            command=self._on_fuzzy_toggle,
            font=ctk.CTkFont(size=12), width=70,
            text_color="#888", fg_color="#2a4a10", hover_color="#3a5a18",
            checkmark_color="#c8f060")
        self.chk_fuzzy.pack(side="left", padx=(0,6))

        self.var_mf = ctk.StringVar(value="All missions")
        self.mf_menu = ctk.CTkOptionMenu(
            tb, variable=self.var_mf, values=["All missions"],
            command=self._on_filter, width=200, font=ctk.CTkFont(size=13))
        self.mf_menu.pack(side="left", padx=6, pady=9)

        _btn(tb, "⬆  Import",    self._import,
             fg="#1a2208", border="#8aab30", txt="#c8f060", hov="#1f2d0a", w=110
             ).pack(side="left", padx=6, pady=9)
        _btn(tb, "↻  Refresh",   self._do_refresh,    w=95
             ).pack(side="left", padx=6, pady=9)
        _btn(tb, "⬇  Export",    self._export,        w=90
             ).pack(side="left", padx=6, pady=9)
        _btn(tb, "📋  History",  self._open_history,  w=100
             ).pack(side="left", padx=6, pady=9)
        _btn(tb, "⌨  Shortcuts", self._open_shortcuts, w=105
             ).pack(side="left", padx=6, pady=9)
        _btn(tb, "⚙  Settings", self._open_settings, w=100
             ).pack(side="left", padx=6, pady=9)

        self.lbl_result = ctk.CTkLabel(tb, text="",
                                        font=ctk.CTkFont(family="Courier New", size=11),
                                        text_color="#555")
        self.lbl_result.pack(side="right", padx=14)

        # Split pane
        pane = ctk.CTkFrame(self, corner_radius=0, fg_color="#0f0f0f")
        pane.pack(fill="both", expand=True)

        # Left panel
        left = ctk.CTkFrame(pane, corner_radius=0, fg_color="#111111", width=410)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        ltb = ctk.CTkFrame(left, corner_radius=0, fg_color="#181818")
        ltb.pack(fill="x")
        _btn(ltb, "Select All",   self._select_all,   h=28, fs=12, w=90
             ).pack(side="left", padx=(8,4), pady=6)
        _btn(ltb, "Deselect All", self._deselect_all, h=28, fs=12, w=95
             ).pack(side="left", padx=4, pady=6)
        self.btn_del_sel = _btn(ltb, "Delete Selected", self._delete_selected,
                                 txt="#f06060", hov="#2e0808", h=28, fs=12)
        self.btn_del_sel.pack(side="left", padx=4, pady=6)
        self.btn_del_sel.configure(state="disabled")

        lf = ctk.CTkFrame(left, corner_radius=0, fg_color="#111111")
        lf.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            lf, bg="#111111", fg="#c8c6c0",
            selectbackground="#1e2e0a", selectforeground="#c8f060",
            borderwidth=0, highlightthickness=0,
            font=("Segoe UI", 11), activestyle="none", relief="flat",
            selectmode=tk.EXTENDED)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(6,0), pady=6)
        self.listbox.bind("<<ListboxSelect>>", self._on_lb_select)
        sb_l = ctk.CTkScrollbar(lf, command=self.listbox.yview)
        sb_l.pack(side="right", fill="y", pady=6)
        self.listbox.configure(yscrollcommand=sb_l.set)

        # Undo bar
        self.undo_bar = ctk.CTkFrame(left, corner_radius=0, fg_color="#1a2e08", height=40)
        self.undo_bar.pack_propagate(False)
        self.lbl_undo = ctk.CTkLabel(self.undo_bar, text="",
                                      font=ctk.CTkFont(size=12), text_color="#c8f060")
        self.lbl_undo.pack(side="left", padx=10)
        self.btn_undo = _btn(self.undo_bar, "↩  Undo", self._do_undo,
                              fg="#2a4a10", border="#8aab30", txt="#c8f060",
                              hov="#3a5a18", h=28, fs=12, w=80)
        self.btn_undo.pack(side="right", padx=(4,8))
        self.lbl_undo_tick = ctk.CTkLabel(self.undo_bar, text="",
                                           font=ctk.CTkFont(family="Courier New", size=11),
                                           text_color="#4a7a20", width=32)
        self.lbl_undo_tick.pack(side="right", padx=2)

        # Divider
        ctk.CTkFrame(pane, width=1, corner_radius=0, fg_color="#2a2a2a"
                     ).pack(side="left", fill="y")

        # Right panel
        right = ctk.CTkFrame(pane, corner_radius=0, fg_color="#0f0f0f")
        right.pack(side="left", fill="both", expand=True)
        rtop = ctk.CTkFrame(right, corner_radius=0, fg_color="#0f0f0f")
        rtop.pack(fill="x", padx=16, pady=(14,6))
        self.lbl_store = ctk.CTkLabel(rtop, text="← Select a store",
                                       font=ctk.CTkFont(size=13),
                                       text_color="#444", anchor="w", wraplength=640)
        self.lbl_store.pack(side="left", fill="x", expand=True)
        self.btn_del_store = _btn(rtop, "🗑  Delete Store", self._delete_current_store,
                                   txt="#f06060", hov="#2e0808", h=30, fs=12)
        self.btn_del_store.pack(side="right")
        self.btn_del_store.configure(state="disabled")
        self.detail_frame = ctk.CTkScrollableFrame(right, fg_color="#0f0f0f", corner_radius=0)
        self.detail_frame.pack(fill="both", expand=True, padx=6, pady=(0,6))

        # Status bar
        sbar = ctk.CTkFrame(self, corner_radius=0, fg_color="#0d0d0d", height=24)
        sbar.pack(fill="x", side="bottom")
        sbar.pack_propagate(False)
        self.lbl_db = ctk.CTkLabel(sbar, text="",
                                    font=ctk.CTkFont(family="Courier New", size=10),
                                    text_color="#3a3a3a")
        self.lbl_db.pack(side="left", padx=10)
        self.lbl_autoref = ctk.CTkLabel(sbar, text="",
                                         font=ctk.CTkFont(family="Courier New", size=10),
                                         text_color="#3a3a3a")
        self.lbl_autoref.pack(side="right", padx=10)




    # ── Fuzzy toggle ──────────────────────────────────────────────────────────
    def _on_fuzzy_toggle(self):
        s = load_settings()
        s["fuzzy"] = self.var_fuzzy.get()
        save_settings(s)
        self._do_search()

    # ── Online bar ────────────────────────────────────────────────────────────
    def _update_online_bar(self):
        try:
            users = db_get_online_users()
            self.lbl_online.configure(
                text=("🟢 Online: " + ",  ".join(users)) if users else "⚫ No users online")
            with get_conn() as c:
                lock = c.execute("SELECT locked_by, action FROM db_locks").fetchone()
            self.lbl_lock.configure(
                text=f"🔒 {lock[0]} is importing — please wait" if lock else "")
        except Exception:
            pass

    # ── Notification bar ──────────────────────────────────────────────────────
    def _notify(self, msg):
        self.lbl_notif.configure(text=f"🔔  {msg}")
        self.notif_bar.pack(fill="x", after=self.online_bar)
        self.after(12000, self._dismiss_notif)

    def _dismiss_notif(self):
        self.notif_bar.pack_forget()

    # ── Keyboard shortcuts ────────────────────────────────────────────────────
    def _bind_shortcuts(self):
        self.bind("<Control-f>",      lambda e: (self._focus_search(), "break"))
        self.bind("<Control-F>",      lambda e: (self._focus_search(), "break"))
        self.bind("<Control-r>",      lambda e: self._do_refresh())
        self.bind("<Control-R>",      lambda e: self._do_refresh())
        self.bind("<Control-h>",      lambda e: self._open_history())
        self.bind("<Control-H>",      lambda e: self._open_history())
        self.bind("<Control-s>",      lambda e: self._open_settings())
        self.bind("<Control-S>",      lambda e: self._open_settings())
        self.bind("<Control-i>",      lambda e: self._import())
        self.bind("<Control-I>",      lambda e: self._import())
        self.bind("<Control-e>",      lambda e: self._export())
        self.bind("<Control-E>",      lambda e: self._export())
        self.bind("<Control-z>",      lambda e: self._do_undo())
        self.bind("<Control-Z>",      lambda e: self._do_undo())
        self.bind("<Delete>",         lambda e: self._delete_selected())
        self.bind("<Escape>",         lambda e: self._clear_search())
        self.bind("<F1>",             lambda e: self._open_shortcuts())
        self.bind("<Control-slash>",  lambda e: self._open_shortcuts())
        self.bind("<Control-grave>",  lambda e: self._cycle_window())

    def _focus_search(self):
        self._search_entry.focus_set()
        self._search_entry.select_range(0, "end")

    def _clear_search(self):
        self.var_search.set("")
        self._search_entry.focus_set()

    # ── Window registry & cycling ─────────────────────────────────────────────
    def _register_window(self, name, win):
        self._child_windows[name] = win
        win.protocol("WM_DELETE_WINDOW",
                     lambda: (self._child_windows.pop(name, None), win.destroy()))

    def _unregister_window(self, name):
        self._child_windows.pop(name, None)

    def _get_open_windows(self):
        order = ["history", "shortcuts", "settings"]
        alive = []
        for key in order:
            win = self._child_windows.get(key)
            if win and win.winfo_exists():
                alive.append((key, win))
            elif win:
                self._child_windows.pop(key, None)
        return alive

    def _cycle_window(self):
        open_wins = self._get_open_windows()
        if not open_wins:
            self._open_history(); return
        if len(open_wins) == 1:
            _, win = open_wins[0]
            win.lift(); win.focus_force(); return
        focused = self.focus_displayof()
        current_idx = -1
        for i, (key, win) in enumerate(open_wins):
            try:
                if str(focused).startswith(str(win)):
                    current_idx = i; break
            except Exception:
                pass
        next_idx = (current_idx + 1) % len(open_wins)
        _, next_win = open_wins[next_idx]
        next_win.lift(); next_win.focus_force()

    # ── Refresh ───────────────────────────────────────────────────────────────
    def _refresh(self):
        try:
            init_db()
            s, m, t = db_stats()
            self.lbl_stats.configure(
                text=f"Stores: {s:,}   Missions: {m:,}   Rows: {t:,}")
            missions = db_all_missions()
            self.mf_menu.configure(values=["All missions"] + missions)
            db_path  = get_db_path()
            is_shared = db_path != DEFAULT_DB
            self.lbl_db.configure(
                text=f"{'🔗 Shared' if is_shared else '💾 Local'}  •  {db_path}",
                text_color="#4a7a20" if is_shared else "#3a3a3a")
            mins = get_auto_refresh()
            self.lbl_autoref.configure(
                text=f"↻ Auto: {'every ' + str(mins) + ' min' if mins else 'off'}")
            self._update_online_bar()
        except Exception as e:
            messagebox.showerror("Database Error", f"Cannot connect:\n{e}\n\nCheck Settings.")
        self._do_search()

    def _do_refresh(self):
        self.var_search.set("")
        self.var_mf.set("All missions")
        self.listbox.selection_clear(0, tk.END)
        self.btn_del_sel.configure(text="Delete Selected", state="disabled")
        self._sel_store = None
        self.btn_del_store.configure(state="disabled")
        self.lbl_store.configure(text="← Select a store", text_color="#444",
                                  font=ctk.CTkFont(size=13))
        self._clear_detail()
        self._search_entry.focus_set()
        self.lbl_result.configure(text="Refreshing…", text_color="#c8f060")
        self.after(50, self._refresh)

    # ── Search ────────────────────────────────────────────────────────────────
    def _schedule_search(self, *_):
        if self._search_after:
            self.after_cancel(self._search_after)
        self._search_after = self.after(130, self._do_search)

    def _on_filter(self, *_):
        self._do_search()

    def _do_search(self):
        t0  = time.perf_counter()
        q   = self.var_search.get().strip()
        mf  = self.var_mf.get()
        if mf == "All missions": mf = ""
        use_fuzzy = self.var_fuzzy.get()
        try:
            if use_fuzzy and q:
                self._stores = db_get_stores_fuzzy(q, mf)
                mode_label   = " (fuzzy)"
            else:
                self._stores = db_get_stores(q, mf)
                mode_label   = ""
        except Exception:
            self._stores = []
            mode_label   = ""
        self.listbox.delete(0, tk.END)
        for name in self._stores:
            self.listbox.insert(tk.END, "  " + name)
        ms  = round((time.perf_counter() - t0) * 1000)
        cnt = len(self._stores)
        self.lbl_result.configure(
            text=f"{cnt:,} store{'s' if cnt!=1 else ''}{mode_label}  •  {ms} ms",
            text_color="#555")
        self.btn_del_sel.configure(text="Delete Selected", state="disabled")
        if self._sel_store and self._sel_store not in self._stores:
            self._sel_store = None
            self.lbl_store.configure(text="← Select a store", text_color="#444",
                                      font=ctk.CTkFont(size=13))
            self.btn_del_store.configure(state="disabled")
            self._clear_detail()

    # ── Listbox ───────────────────────────────────────────────────────────────
    def _on_lb_select(self, _event):
        sel = self.listbox.curselection()
        cnt = len(sel)
        self.btn_del_sel.configure(
            text=f"Delete Selected ({cnt:,})" if cnt else "Delete Selected",
            state="normal" if cnt else "disabled")
        if cnt == 1:
            idx = sel[0]
            if idx < len(self._stores):
                store = self._stores[idx]
                if store != self._sel_store:
                    self._sel_store = store
                    self._show_missions(store)
        elif cnt == 0:
            self._sel_store = None
            self.lbl_store.configure(text="← Select a store", text_color="#444",
                                      font=ctk.CTkFont(size=13))
            self.btn_del_store.configure(state="disabled")
            self._clear_detail()

    def _select_all(self):
        self.listbox.select_set(0, tk.END)
        cnt = self.listbox.size()
        self.btn_del_sel.configure(
            text=f"Delete Selected ({cnt:,})" if cnt else "Delete Selected",
            state="normal" if cnt else "disabled")

    def _deselect_all(self):
        self.listbox.selection_clear(0, tk.END)
        self.btn_del_sel.configure(text="Delete Selected", state="disabled")

    # ── Mission detail ────────────────────────────────────────────────────────
    def _show_missions(self, store):
        try:
            missions = db_get_missions(store)
        except Exception:
            missions = []
        self.lbl_store.configure(text=store, text_color="#e8e6e0",
                                  font=ctk.CTkFont(size=13, weight="bold"))
        self.btn_del_store.configure(state="normal")
        self._clear_detail()
        if not missions:
            ctk.CTkLabel(self.detail_frame, text="No missions found.",
                         text_color="#555", font=ctk.CTkFont(size=13)).pack(pady=24)
            return
        for mission, link in missions:
            card = ctk.CTkFrame(self.detail_frame, fg_color="#1a1a1a",
                                 border_color="#2e2e2e", border_width=1, corner_radius=8)
            card.pack(fill="x", padx=4, pady=4)
            top = ctk.CTkFrame(card, fg_color="transparent")
            top.pack(fill="x", padx=12, pady=(10,4))
            ctk.CTkLabel(top, text=mission,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color="#e8e6e0", anchor="w", wraplength=560
                         ).pack(side="left", fill="x", expand=True)
            if link:
                bot = ctk.CTkFrame(card, fg_color="transparent")
                bot.pack(fill="x", padx=12, pady=(0,10))
                lbl = ctk.CTkLabel(bot, text=link,
                                   font=ctk.CTkFont(family="Courier New", size=10),
                                   text_color="#5599cc", anchor="w",
                                   wraplength=500, cursor="hand2")
                lbl.pack(side="left", fill="x", expand=True, padx=(0,8))
                lbl.bind("<Button-1>", lambda e, u=link: webbrowser.open(u))
                _btn(bot, "Open", lambda u=link: webbrowser.open(u),
                     fg="#081c2e", border="#185fa5", txt="#60b4f0", hov="#0c2d4a",
                     h=26, w=64, fs=11).pack(side="right", padx=(3,0))
                _btn(bot, "Copy", lambda u=link: self._copy(u),
                     h=26, w=64, fs=11).pack(side="right", padx=(0,3))
            else:
                ctk.CTkLabel(card, text="No link provided",
                             font=ctk.CTkFont(size=11), text_color="#444", anchor="w"
                             ).pack(padx=12, pady=(0,10), anchor="w")

    def _clear_detail(self):
        for w in self.detail_frame.winfo_children():
            w.destroy()

    # ── Delete & Undo ─────────────────────────────────────────────────────────
    def _delete_current_store(self):
        store = self._sel_store
        if not store: return
        if not messagebox.askyesno("Delete Store",
            f"Remove this store from ALL missions?\n\n{store}\n\nUndoable within 30s."):
            return
        self._execute_delete([store])

    def _delete_selected(self):
        sel = self.listbox.curselection()
        if not sel: return
        stores = [self._stores[i] for i in sel if i < len(self._stores)]
        if not stores: return
        n = len(stores)
        if not messagebox.askyesno("Delete Stores",
            f"Remove {n:,} store{'s' if n>1 else ''} from ALL missions?\n\nUndoable within 30s."):
            return
        self._execute_delete(stores)

    def _execute_delete(self, stores):
        saved_rows = db_get_rows_for_stores(stores)
        deleted    = db_delete_stores(stores)
        db_log_delete(stores, deleted)
        self._undo_buffer = {"store_names": stores, "rows": saved_rows}
        self._show_undo_bar(stores, deleted)
        if self._sel_store in stores:
            self._sel_store = None
            self.btn_del_store.configure(state="disabled")
            self.lbl_store.configure(text="← Select a store", text_color="#444",
                                      font=ctk.CTkFont(size=13))
            self._clear_detail()
        self._refresh()

    def _show_undo_bar(self, stores, deleted):
        if self._undo_after:     self.after_cancel(self._undo_after)
        if self._undo_tick_after: self.after_cancel(self._undo_tick_after)
        n   = len(stores)
        msg = f"🗑  {n} store{'s' if n>1 else ''} deleted  ({deleted} rows)"
        self.lbl_undo.configure(text=msg)
        self.undo_bar.pack(fill="x", before=self.listbox.master)
        self._undo_start    = time.time()
        self._tick_undo_countdown()
        self._undo_after    = self.after(UNDO_TTL * 1000, self._expire_undo)

    def _tick_undo_countdown(self):
        remaining = UNDO_TTL - int(time.time() - self._undo_start)
        if remaining > 0:
            self.lbl_undo_tick.configure(text=f"{remaining}s")
            self._undo_tick_after = self.after(1000, self._tick_undo_countdown)
        else:
            self.lbl_undo_tick.configure(text="")

    def _expire_undo(self):
        self._undo_buffer = None
        self.undo_bar.pack_forget()
        self.lbl_undo_tick.configure(text="")

    def _do_undo(self):
        if not self._undo_buffer: return
        buf = self._undo_buffer; self._undo_buffer = None
        if self._undo_after:      self.after_cancel(self._undo_after)
        if self._undo_tick_after: self.after_cancel(self._undo_tick_after)
        self.undo_bar.pack_forget()
        restored = db_restore_rows(buf["rows"])
        db_log_undo(buf["store_names"], restored)
        self._refresh()
        self._toast(f"↩  Undo successful — {restored} rows restored")

    # ── Import ────────────────────────────────────────────────────────────────
    def _import(self):
        ok, locker = db_acquire_lock("Import")
        if not ok:
            messagebox.showwarning("Import Locked",
                f"🔒  {locker} is currently importing data.\n\nPlease wait and try again shortly.")
            return
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[
                ("Supported files", "*.csv *.txt *.xlsx *.xls"),
                ("CSV files",       "*.csv *.txt"),
                ("Excel files",     "*.xlsx *.xls"),
                ("All files",       "*.*"),
            ])
        if not path:
            db_release_my_lock(); return
        self._open_import_for_path(path)

    def _open_import_for_path(self, path):
        """Validate and open ImportWindow for a given path."""
        ok, err = validate_file_type(path)
        if not ok:
            db_release_my_lock()
            messagebox.showerror("File Type Not Supported", err)
            return
        win = ImportWindow(self, path)
        win.grab_set()
        self.wait_window(win)
        db_release_my_lock()
        self._update_online_bar()
        self._refresh()

    def _export(self):
        q  = self.var_search.get().strip()
        mf = self.var_mf.get()
        if mf == "All missions": mf = ""
        stores = db_get_stores(q, mf)
        if not stores:
            messagebox.showinfo("Export", "No data to export with current filters.")
            return
        path = filedialog.asksaveasfilename(
            title="Save export",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="store-mission-export.csv")
        if not path: return
        count = 0
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Store Details", "Mission Name", "Mission Link"])
            for store in stores:
                for mission, link in db_get_missions(store):
                    if not mf or mission == mf:
                        w.writerow([store, mission, link])
                        count += 1
        self._toast(f"Exported {count:,} rows")

    # ── Import DB file ────────────────────────────────────────────────────────
    def _import_db(self):
        """Import (replace or merge) from another .db file."""
        path = filedialog.askopenfilename(
            title="Select a StoreMissionLookup database file",
            filetypes=[("SQLite database", "*.db"), ("All files", "*.*")])
        if not path: return
        choice = messagebox.askyesnocancel(
            "Import Database",
            "How do you want to import this database?\n\n"
            "YES    = Merge with existing data (recommended)\n"
            "NO     = Replace all existing data\n"
            "CANCEL = Cancel")
        if choice is None: return
        try:
            src = sqlite3.connect(path)
            if choice:
                # Merge — insert rows, skip duplicates
                rows = src.execute("SELECT store,store_lc,mission,link,dedup FROM entries").fetchall()
                src.close()
                added, skipped = 0, 0
                conn = get_conn()
                cur  = conn.cursor()
                for row in rows:
                    try:
                        cur.execute(
                            "INSERT INTO entries (store,store_lc,mission,link,dedup) VALUES (?,?,?,?,?)", row)
                        added += 1
                    except sqlite3.IntegrityError:
                        skipped += 1
                conn.commit(); conn.close()
                db_log_import(f"[DB Import] {os.path.basename(path)}", added, skipped)
                messagebox.showinfo("Import Complete",
                                    f"Database merged.\n\n{added:,} rows added, {skipped:,} duplicates skipped.")
            else:
                # Replace
                src.close()
                target = get_db_path()
                shutil.copy2(path, target)
                messagebox.showinfo("Import Complete",
                                    "Database replaced successfully.\nThe app will now reload.")
            self._refresh()
        except Exception as e:
            messagebox.showerror("Import Failed", f"Could not import database:\n{e}")

    def _open_history(self):
        existing = self._child_windows.get("history")
        if existing and existing.winfo_exists():
            existing.lift(); existing.focus_force(); return
        win = HistoryWindow(self)
        self._register_window("history", win)

    def _open_shortcuts(self):
        existing = self._child_windows.get("shortcuts")
        if existing and existing.winfo_exists():
            existing.lift(); existing.focus_force(); return
        win = ShortcutsWindow(self)
        self._register_window("shortcuts", win)

    def _open_settings(self):
        existing = self._child_windows.get("settings")
        if existing and existing.winfo_exists():
            existing.lift(); existing.focus_force(); return
        win = SettingsWindow(self)
        self._register_window("settings", win)
        win.protocol("WM_DELETE_WINDOW",
                     lambda: (self._unregister_window("settings"),
                               win.destroy(),
                               self._schedule_auto_refresh(),
                               self._refresh()))

    def _copy(self, url):
        self.clipboard_clear()
        self.clipboard_append(url)
        self._toast("Link copied")

    def _toast(self, msg):
        self.lbl_result.configure(text=msg, text_color="#c8f060")
        self.after(2800, self._do_search)


# ── Import Window ─────────────────────────────────────────────────────────────
class ImportWindow(ctk.CTkToplevel):
    def __init__(self, parent, path):
        super().__init__(parent)
        self.title("Import File")
        self.geometry("600x420")
        self.resizable(True, True)
        self.minsize(540, 360)
        self.configure(fg_color="#1a1a1a")
        self.transient(parent)
        self.lift()
        self.path       = path
        self._cancelled = False
        self._sc = self._mc = self._lc = -1
        self._header    = []
        self._all_rows  = []
        self._read_file()
        self._build_ui()
        if self._sc != -1 and self._mc != -1 and self._lc != -1:
            self.after(400, self._start)

    def _read_file(self):
        headers, rows, err = read_file_rows(self.path)
        if err:
            self._header   = []
            self._all_rows = []
            self.after(100, lambda: messagebox.showerror("Read Error", err, parent=self))
            self.after(200, self.destroy)
            return
        self._header   = headers
        self._all_rows = rows
        self._sc, self._mc, self._lc = detect_columns(headers)

    def _build_ui(self):
        ext = os.path.splitext(self.path)[1].upper().lstrip(".")
        ctk.CTkLabel(self, text=f"Import {ext} File",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(18,2))
        ctk.CTkLabel(self, text=os.path.basename(self.path),
                     font=ctk.CTkFont(family="Courier New", size=11),
                     text_color="#888").pack()
        ctk.CTkLabel(self,
                     text=f"{len(self._all_rows):,} rows detected  •  {len(self._header)} columns",
                     font=ctk.CTkFont(size=11), text_color="#555").pack(pady=(2,0))

        # Validate columns — show warning if any missing
        val_err = validate_required_columns(self._sc, self._mc, self._lc, self._header)
        if val_err:
            warn_f = ctk.CTkFrame(self, fg_color="#2e1a00", corner_radius=6,
                                   border_color="#f0c060", border_width=1)
            warn_f.pack(fill="x", padx=20, pady=(8,4))
            ctk.CTkLabel(warn_f, text="⚠  " + val_err.split("\n")[0],
                         text_color="#f0c060", font=ctk.CTkFont(size=12),
                         anchor="w", wraplength=480
                         ).pack(padx=12, pady=6, anchor="w")

        need_map = self._sc == -1 or self._mc == -1 or self._lc == -1
        if need_map:
            ctk.CTkLabel(self, text="Map columns manually:",
                         font=ctk.CTkFont(size=12), text_color="#888").pack(pady=(6,2))
            opts = self._header or ["(col 1)", "(col 2)", "(col 3)"]
            mf   = ctk.CTkFrame(self, fg_color="transparent")
            mf.pack(pady=4)
            self._sc_m = _field_row(mf, "Store Details:", opts, self._sc, row=0)
            self._mc_m = _field_row(mf, "Mission Name:",  opts, self._mc, row=1)
            self._lc_m = _field_row(mf, "Mission Link:",  opts, self._lc, row=2)
            _btn(self, "Start Import", self._start_manual,
                 fg="#1a2208", border="#8aab30", txt="#c8f060", hov="#1f2d0a"
                 ).pack(pady=(8,4))

        pw = ctk.CTkFrame(self, fg_color="transparent")
        pw.pack(fill="x", padx=24, pady=(12,4))
        self.pbar = ctk.CTkProgressBar(pw)
        self.pbar.pack(fill="x")
        self.pbar.set(0)

        # Status label for normal progress messages
        self.lbl_prog = ctk.CTkLabel(self, text="Ready…",
                                      font=ctk.CTkFont(family="Courier New", size=11),
                                      text_color="#888")
        self.lbl_prog.pack(pady=(4,0))

        # Scrollable error box — hidden until an error occurs
        self.err_box = ctk.CTkTextbox(
            self, height=140,
            font=ctk.CTkFont(family="Courier New", size=11),
            text_color="#f06060",
            fg_color="#1a0808",
            border_color="#5a1010",
            border_width=1,
            wrap="word",
            state="disabled")
        # Not packed yet — only shown on error

        _btn(self, "Cancel", self._cancel, w=90).pack(pady=(6,14))


    def _start_manual(self):
        opts = self._header
        try:
            self._sc = opts.index(self._sc_m.get())
            self._mc = opts.index(self._mc_m.get())
            self._lc = opts.index(self._lc_m.get())
        except (ValueError, AttributeError):
            pass
        if len({self._sc, self._mc, self._lc}) < 3:
            messagebox.showerror("Column Error",
                                 "Please select three different columns.", parent=self)
            return
        self._start()

    def _start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        CHUNK  = 2000
        added  = skipped = processed = 0
        # ── Column presence check ──────────────────────────────────────────────
        col_missing = []
        if self._sc == -1: col_missing.append("Store Details")
        if self._mc == -1: col_missing.append("Mission Name")
        if col_missing:
            err_msg = ("Cannot import — required column(s) not assigned: "
                       + ", ".join(col_missing)
                       + "\nPlease use the column mappers to assign them.")
            self.after(0, self._error, err_msg)
            return

        # ── Pre-scan ALL rows for missing data — HARD BLOCK if any found ────────
        bad_rows = []  # (row_number, [missing_field_names])
        rows     = self._all_rows
        for i, row in enumerate(rows, start=2):  # row 1 is the header
            store   = row[self._sc].strip() if self._sc < len(row) else ""
            mission = row[self._mc].strip() if self._mc < len(row) else ""
            link    = row[self._lc].strip() if self._lc < len(row) else ""
            fields  = []
            if not store:   fields.append("Store Details")
            if not mission: fields.append("Mission Name")
            if not link:    fields.append("Mission Link")
            if fields:
                bad_rows.append((i, fields))

        if bad_rows:
            shown     = bad_rows[:15]
            lines     = ["  Row {:,}: {} is blank".format(r, " & ".join(f)) for r, f in shown]
            if len(bad_rows) > 15:
                lines.append("  ... and {:,} more rows with missing data".format(len(bad_rows) - 15))
            detail    = "\n".join(lines)
            total_bad = len(bad_rows)
            noun      = "rows" if total_bad != 1 else "row"
            verb      = "have" if total_bad != 1 else "has"
            err_msg   = (
                "Import blocked -- {:,} {} {} missing required data:\n\n{}\n\n"
                "All 3 columns (Store Details, Mission Name, Mission Link) "
                "must be filled in for every row.\n\n"
                "Please fix the file and try again."
            ).format(total_bad, noun, verb, detail)
            self.after(0, self._error, err_msg)
            return   # hard stop — nothing is imported

        try:
            total = len(rows)
            batch = []
            for row in rows:
                if self._cancelled: break
                processed += 1
                store   = row[self._sc].strip() if self._sc < len(row) else ""
                mission = row[self._mc].strip() if self._mc < len(row) else ""
                link    = row[self._lc].strip() if self._lc < len(row) else ""
                batch.append((store, mission, link))
                if len(batch) >= CHUNK:
                    a, s    = db_insert_batch(batch)
                    added  += a; skipped += s; batch = []
                    pct     = min(0.95, processed / total) if total else 0.5
                    self.after(0, self._update, processed, total, added, skipped, pct)
            if batch and not self._cancelled:
                a, s   = db_insert_batch(batch)
                added += a; skipped += s
            if not self._cancelled:
                db_log_import(os.path.basename(self.path), added, skipped)
            self.after(0, self._done, added, skipped)
        except Exception as e:
            self.after(0, self._error, str(e))

    def _update(self, processed, total, added, skipped, pct):
        self.pbar.set(pct)
        self.lbl_prog.configure(
            text=f"{processed:,} / {total:,} rows — {added:,} added, {skipped:,} skipped")

    def _done(self, added, skipped):
        self.pbar.set(1.0)
        msg = f"Done!  {added:,} rows imported"
        if skipped: msg += f",  {skipped:,} skipped (duplicates or missing data)"
        self.lbl_prog.configure(text=msg, text_color="#60f0a0")
        self.after(1600, self.destroy)

    def _error(self, msg):
        # Hide progress label, show scrollable error box
        self.lbl_prog.pack_forget()
        self.err_box.configure(state="normal")
        self.err_box.delete("0.0", "end")
        self.err_box.insert("0.0", msg)
        self.err_box.configure(state="disabled")
        self.err_box.pack(fill="both", expand=True, padx=20, pady=(4,4))
        # Grow window so the full error is visible
        self.geometry("620x520")
        self.resizable(True, True)

    def _cancel(self):
        self._cancelled = True
        self.destroy()


# ── History Window (with resizable columns) ───────────────────────────────────
class HistoryWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Activity History")
        self.geometry("960x540")
        self.configure(fg_color="#1a1a1a")
        self.transient(parent)
        self.lift()
        self.focus_force()
        self.bind("<Escape>", lambda e: self.destroy())
        self._sort_col = "performed_at"
        self._sort_dir = "DESC"
        self._build_ui()
        self._load()

    def _build_ui(self):
        ctk.CTkLabel(self, text="Activity History",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(16,4))
        ctk.CTkLabel(self,
                     text="Click column headers to sort  •  Drag column borders to resize  •  Esc to close",
                     font=ctk.CTkFont(size=12), text_color="#888").pack(pady=(0,10))

        frame = ctk.CTkFrame(self, fg_color="#111111", corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=(0,8))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("H.Treeview",
                        background="#111111", foreground="#c8c6c0",
                        fieldbackground="#111111", borderwidth=0,
                        rowheight=26, font=("Segoe UI", 11))
        style.configure("H.Treeview.Heading",
                        background="#1e1e1e", foreground="#888888",
                        relief="flat", font=("Segoe UI", 11, "bold"))
        style.map("H.Treeview",
                  background=[("selected","#1e2e0a")],
                  foreground=[("selected","#c8f060")])

        cols = ("date","action","detail","by","count","extra")
        # stretch=True on all columns enables user resizing by dragging headers
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                  style="H.Treeview")

        col_defs = [
            ("date",   "performed_at",  "Date & Time",  165, "w",      True),
            ("action", "action",        "Action",        90, "center", True),
            ("detail", "detail",        "Detail",       285, "w",      True),
            ("by",     "performed_by",  "Performed By", 130, "w",      True),
            ("count",  "count",         "Rows",          70, "center", True),
            ("extra",  "extra",         "Notes",        170, "w",      True),
        ]
        self._col_db = {c[0]: c[1] for c in col_defs}

        for col, db_col, heading, width, anchor, stretch in col_defs:
            self.tree.heading(col, text=heading,
                              command=lambda c=db_col: self._sort_by(c))
            self.tree.column(col, width=width, anchor=anchor,
                             stretch=stretch, minwidth=50)

        self.tree.tag_configure("import", foreground="#c8f060")
        self.tree.tag_configure("delete", foreground="#f06060")
        self.tree.tag_configure("undo",   foreground="#60b4f0")
        self.tree.bind("<Escape>", lambda e: self.destroy())

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=(4,14))
        _btn(bf, "↻ Refresh", self._load,   h=30, fs=12, w=100).pack(side="left", padx=8)
        _btn(bf, "Close",     self.destroy, h=30, fs=12, w=90 ).pack(side="left", padx=8)

    def _sort_by(self, db_col):
        if self._sort_col == db_col:
            self._sort_dir = "ASC" if self._sort_dir == "DESC" else "DESC"
        else:
            self._sort_col = db_col
            self._sort_dir = "DESC"
        self._load()
        heading_map = {
            "performed_at": "Date & Time", "action": "Action",
            "detail": "Detail", "performed_by": "Performed By",
            "count": "Rows", "extra": "Notes"
        }
        col_map_rev = {v: k for k, v in self._col_db.items()}
        for db_c, h_text in heading_map.items():
            col_id = col_map_rev.get(db_c, db_c)
            arr = (" ▼" if self._sort_dir == "DESC" else " ▲") if db_c == self._sort_col else ""
            self.tree.heading(col_id, text=h_text + arr)

    def _load(self):
        self.tree.delete(*self.tree.get_children())
        try:
            rows = db_get_history(self._sort_col, self._sort_dir)
        except Exception:
            rows = []
        if not rows:
            self.tree.insert("", "end",
                             values=("No activity recorded yet.","","","","",""))
            return
        icons = {"Import":"⬆","Delete":"🗑","Undo":"↩"}
        for action, detail, performed_at, performed_by, count, extra in rows:
            tag  = action.lower()
            icon = icons.get(action, "•")
            self.tree.insert("", "end", tags=(tag,),
                             values=(performed_at, f"{icon} {action}",
                                     detail, performed_by,
                                     f"{count:,}" if isinstance(count,int) else count,
                                     extra))


# ── Shortcuts Window ──────────────────────────────────────────────────────────
class ShortcutsWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Keyboard Shortcuts")
        self.geometry("500x560")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")
        self.transient(parent)
        self.lift()
        self.focus_force()
        self.bind("<Escape>", lambda e: self.destroy())
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(self, text="⌨  Keyboard Shortcuts",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(18,4))
        ctk.CTkLabel(self, text="All shortcuts work from anywhere in the app.  Press Esc to close.",
                     font=ctk.CTkFont(size=12), text_color="#888").pack(pady=(0,14))

        groups = [
            ("Navigation & Search", [
                ("Ctrl + F",      "Focus the search box and select all text"),
                ("Ctrl + R",      "Refresh — reloads data, clears search, filter & selection"),
                ("Escape",        "Clear the search box  /  close any popup window"),
            ]),
            ("Data Actions", [
                ("Ctrl + I",      "Open Import dialog  (CSV or Excel .xlsx/.xls)"),
                ("Ctrl + E",      "Export currently filtered data to CSV"),
                ("Delete",        "Delete all selected stores"),
                ("Ctrl + Z",      "Undo last delete  (30-second window)"),
            ]),
            ("Windows", [
                ("Ctrl + H",      "Open Activity History"),
                ("Ctrl + S",      "Open Settings"),
                ("Ctrl + `",      "Cycle focus between open popup windows"),
                ("F1",            "Open this Shortcuts window"),
                ("Ctrl + /",      "Open this Shortcuts window"),
                ("Escape",        "Close any open popup window"),
            ]),
            ("Mouse — Store List", [
                ("Click",         "Select a store and view its missions"),
                ("Ctrl + Click",  "Add / remove individual stores from selection"),
                ("Shift + Click", "Select all stores between two clicks"),
            ]),
            ("Search", [
                ("Fuzzy checkbox","Toggle fuzzy / approximate matching in toolbar"),
            ]),
        ]

        scroll = ctk.CTkScrollableFrame(self, fg_color="#111111", corner_radius=8)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0,8))

        for group_title, items in groups:
            ctk.CTkLabel(scroll, text=group_title,
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color="#c8f060", anchor="w"
                         ).pack(fill="x", padx=10, pady=(14,4))
            for shortcut, description in items:
                row = ctk.CTkFrame(scroll, fg_color="#1a1a1a",
                                   corner_radius=6, border_color="#2e2e2e", border_width=1)
                row.pack(fill="x", padx=6, pady=2)
                ctk.CTkLabel(row, text=shortcut,
                             font=ctk.CTkFont(family="Courier New", size=12, weight="bold"),
                             text_color="#e8e6e0", width=150, anchor="w"
                             ).pack(side="left", padx=(12,8), pady=7)
                ctk.CTkLabel(row, text=description,
                             font=ctk.CTkFont(size=12),
                             text_color="#888", anchor="w"
                             ).pack(side="left", padx=(0,12), pady=7)

        _btn(self, "Close", self.destroy, w=90).pack(pady=(4,14))


# ── Settings Window ───────────────────────────────────────────────────────────
class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("620x500")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")
        self.transient(parent)
        self.lift()
        self.focus_force()
        self.bind("<Escape>",    lambda e: self.destroy())
        self.bind("<Control-s>", lambda e: self._save())
        self.bind("<Control-S>", lambda e: self._save())
        self._build_ui()

    def _build_ui(self):
        s = load_settings()
        ctk.CTkLabel(self, text="Settings",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(18,4))

        # Your name
        f1 = ctk.CTkFrame(self, fg_color="transparent")
        f1.pack(fill="x", padx=24, pady=8)
        ctk.CTkLabel(f1, text="Your name:", font=ctk.CTkFont(size=13),
                     width=130, anchor="e").pack(side="left", padx=(0,10))
        self.var_name = ctk.StringVar(value=s.get("user_name",""))
        ctk.CTkEntry(f1, textvariable=self.var_name, width=340,
                     font=ctk.CTkFont(size=13)).pack(side="left")

        # DB path
        f2 = ctk.CTkFrame(self, fg_color="transparent")
        f2.pack(fill="x", padx=24, pady=8)
        ctk.CTkLabel(f2, text="Database path:", font=ctk.CTkFont(size=13),
                     width=130, anchor="e").pack(side="left", padx=(0,10))
        self.var_db = ctk.StringVar(value=s.get("db_path", DEFAULT_DB))
        ctk.CTkEntry(f2, textvariable=self.var_db, width=270,
                     font=ctk.CTkFont(size=12)).pack(side="left")
        _btn(f2, "Browse…", self._browse, h=32, fs=12, w=80
             ).pack(side="left", padx=(8,0))

        f3 = ctk.CTkFrame(self, fg_color="transparent")
        f3.pack(fill="x", padx=24, pady=(0,4))
        ctk.CTkLabel(f3, text="", width=130).pack(side="left", padx=(0,10))
        _btn(f3, "Reset to local default", self._reset_local, h=26, fs=11
             ).pack(side="left", padx=(0,8))
        _btn(f3, "📥 Import DB file", self._import_db_from_settings, h=26, fs=11
             ).pack(side="left")

        # Auto-refresh
        f4 = ctk.CTkFrame(self, fg_color="transparent")
        f4.pack(fill="x", padx=24, pady=8)
        ctk.CTkLabel(f4, text="Auto-refresh:", font=ctk.CTkFont(size=13),
                     width=130, anchor="e").pack(side="left", padx=(0,10))
        refresh_opts = ["Off", "10 minutes", "20 minutes", "30 minutes", "45 minutes"]
        refresh_vals = [0, 10, 20, 30, 45]
        cur_val   = s.get("auto_refresh", 0)
        cur_label = refresh_opts[refresh_vals.index(cur_val)] if cur_val in refresh_vals else "Off"
        self.var_refresh     = ctk.StringVar(value=cur_label)
        self._refresh_opts   = refresh_opts
        self._refresh_vals   = refresh_vals
        ctk.CTkOptionMenu(f4, variable=self.var_refresh, values=refresh_opts,
                          width=180, font=ctk.CTkFont(size=13)).pack(side="left")
        ctk.CTkLabel(f4, text="  (for shared database use)",
                     font=ctk.CTkFont(size=11), text_color="#555"
                     ).pack(side="left", padx=8)

        # Fuzzy search
        f6 = ctk.CTkFrame(self, fg_color="transparent")
        f6.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(f6, text="Fuzzy search:", font=ctk.CTkFont(size=13),
                     width=130, anchor="e").pack(side="left", padx=(0,10))
        self.var_fuzzy_s = ctk.BooleanVar(value=s.get("fuzzy", False))
        ctk.CTkCheckBox(f6, text="Enable fuzzy / approximate matching",
                        variable=self.var_fuzzy_s,
                        font=ctk.CTkFont(size=12),
                        fg_color="#2a4a10", hover_color="#3a5a18",
                        checkmark_color="#c8f060").pack(side="left")

        # Info box
        info = ctk.CTkFrame(self, fg_color="#1a2208", corner_radius=8,
                             border_color="#4a7a20", border_width=1)
        info.pack(fill="x", padx=24, pady=10)
        ctk.CTkLabel(info,
                     text="📁  Shared path examples:\n"
                          "  OneDrive:    C:\\Users\\YourName\\OneDrive\\StoreMissionLookup\\data.db\n"
                          "  SharePoint:  C:\\Users\\YourName\\CompanyName\\StoreMissionLookup\\data.db\n"
                          "  Google Drive: G:\\My Drive\\StoreMissionLookup\\data.db",
                     font=ctk.CTkFont(family="Courier New", size=11),
                     text_color="#8aab30", justify="left", anchor="w"
                     ).pack(padx=14, pady=10, anchor="w")

        self.lbl_status = ctk.CTkLabel(self, text="",
                                        font=ctk.CTkFont(size=12), text_color="#888")
        self.lbl_status.pack(pady=4)

        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=(4,16))
        _btn(bf, "Test Connection", self._test, h=32, fs=12, w=140
             ).pack(side="left", padx=8)
        _btn(bf, "Save", self._save,
             fg="#1a2208", border="#8aab30", txt="#c8f060", hov="#1f2d0a",
             h=32, fs=12, w=90).pack(side="left", padx=8)
        _btn(bf, "Cancel", self.destroy, h=32, fs=12, w=90).pack(side="left", padx=8)

    def _browse(self):
        path = filedialog.asksaveasfilename(
            title="Choose database file location",
            defaultextension=".db",
            filetypes=[("SQLite database","*.db"),("All files","*.*")],
            initialfile="data.db")
        if path: self.var_db.set(path)

    def _reset_local(self):
        self.var_db.set(DEFAULT_DB)

    def _import_db_from_settings(self):
        self.destroy()
        self.master._import_db()

    def _test(self):
        path = self.var_db.get().strip()
        try:
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            conn = sqlite3.connect(path, timeout=5)
            conn.execute("SELECT 1")
            conn.close()
            self.lbl_status.configure(text="✓  Connection successful", text_color="#60f0a0")
        except Exception as e:
            self.lbl_status.configure(text=f"✗  Failed: {e}", text_color="#f06060")

    def _save(self):
        db_path = self.var_db.get().strip() or DEFAULT_DB
        if not is_valid_db_path(db_path):
            messagebox.showerror("Invalid Path",
                "The database path must be a local file path ending in .db\n\n"
                "URLs (http://, https://) are not valid.", parent=self)
            return
        s = load_settings()
        s["user_name"]   = self.var_name.get().strip() or "User"
        s["db_path"]     = db_path
        s["fuzzy"]       = self.var_fuzzy_s.get()
        label = self.var_refresh.get()
        s["auto_refresh"] = self._refresh_vals[self._refresh_opts.index(label)] \
                            if label in self._refresh_opts else 0
        save_settings(s)
        try:
            init_db()
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not init database:\n{e}", parent=self)
            return
        self.destroy()


# ── Helpers ───────────────────────────────────────────────────────────────────
def _btn(parent, text, cmd, fg="#1e1e1e", border="#3a3a3a",
         txt="#e8e6e0", hov="#2a2a2a", w=None, h=32, fs=13):
    kw = dict(text=text, command=cmd, fg_color=fg, border_color=border,
              border_width=1, text_color=txt, hover_color=hov,
              font=ctk.CTkFont(size=fs), height=h, corner_radius=6)
    if w: kw["width"] = w
    return ctk.CTkButton(parent, **kw)

def _field_row(parent, label, opts, default_idx, row):
    ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=12),
                 width=120, anchor="e").grid(row=row, column=0, padx=8, pady=4)
    menu = ctk.CTkOptionMenu(parent, values=opts, width=220, font=ctk.CTkFont(size=12))
    menu.grid(row=row, column=1, padx=8, pady=4)
    if 0 <= default_idx < len(opts):
        menu.set(opts[default_idx])
    return menu


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("green")
    init_db()
    App().mainloop()
