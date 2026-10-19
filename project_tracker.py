"""Vektor Project Management - High-Fidelity Streamlit Application.
Install: pip install streamlit pandas plotly openpyxl pymongo
Run: streamlit run project_tracker_fixed_complete.py
Persists a shared application snapshot to MongoDB Vektor.Reports. Local SQLite is only a working cache.
"""
from __future__ import annotations

import base64
import calendar
import html
import io
import json
import os
import sqlite3
import re
import mimetypes
from pathlib import Path
from datetime import date, datetime, timedelta

try:
    from pymongo import MongoClient
    from pymongo.errors import PyMongoError
except ImportError:  # A clear setup message is shown if MongoDB is configured but pymongo is absent.
    MongoClient = None
    PyMongoError = Exception

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DB = os.environ.get("VEKTOR_DB_PATH", "vektor.db")
MONGO_DATABASE = "Vektor"
MONGO_COLLECTION = "Reports"
MONGO_SNAPSHOT_ID = "vektor_project_management_snapshot"

# Brand & Status Palette
GREEN = "#00875A"
TEAL_ACCENT = "#00A879"
DARK_TEAL_HEADER = "#053241"
BLUE = "#1473E6"
ORANGE = "#F28C00"
RED = "#EF4444"
GRAY = "#8495B0"
NAVY = "#10194F"
PURPLE = "#7C3AED"

st.set_page_config(
    page_title="Vektor | Project Management",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Database Layer & Schema Management (data is supplied through Excel)
# ---------------------------------------------------------------------------
def conn() -> sqlite3.Connection:
    """Local working cache; MongoDB is the durable source-of-truth snapshot."""
    c = sqlite3.connect(DB, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c


@st.cache_resource(show_spinner=False)
def _mongo_collection():
    """Resolve and reuse the MongoDB collection connection across reruns."""
    uri = ""
    try:
        uri = str(st.secrets.get("MONGODB_URI", "")).strip()
    except Exception:
        uri = ""
    uri = uri or os.environ.get("MONGODB_URI", "").strip()
    if not uri:
        return None
    if MongoClient is None:
        raise RuntimeError("MongoDB persistence requires pymongo. Install it with: pip install pymongo")
    client = MongoClient(uri, serverSelectionTimeoutMS=8000, connectTimeoutMS=8000)
    client.admin.command("ping")
    return client[MONGO_DATABASE][MONGO_COLLECTION]


def _snapshot_tables(c: sqlite3.Connection) -> dict:
    names = ("projects", "tasks", "subtasks", "task_attachments", "resources", "milestones", "settings", "saved_reports", "audit_log", "calendar_events", "project_assignments")
    snapshot = {}
    for name in names:
        rows = c.execute(f"SELECT * FROM {name}").fetchall()
        snapshot[name] = [dict(row) for row in rows]
    snapshot["saved_at"] = datetime.now().isoformat(timespec="seconds")
    snapshot["schema_version"] = 1
    return snapshot


def _persist_snapshot() -> None:
    """Persist every app table together in the requested Vektor.Reports collection."""
    collection = _mongo_collection()
    if collection is None:
        return
    try:
        with conn() as c:
            snapshot = _snapshot_tables(c)
        collection.replace_one({"_id": MONGO_SNAPSHOT_ID}, {"_id": MONGO_SNAPSHOT_ID, **snapshot}, upsert=True)
        st.session_state["_vektor_mongo_connected"] = True
    except Exception as exc:
        st.session_state["_vektor_mongo_connected"] = False
        st.error(f"MongoDB save failed. Your change is still in the local cache. Check MONGODB_URI and database permissions. Details: {exc}")


def _restore_snapshot() -> bool:
    """Restore the latest saved MongoDB snapshot into the local SQLite working cache."""
    collection = _mongo_collection()
    if collection is None:
        return False
    doc = collection.find_one({"_id": MONGO_SNAPSHOT_ID})
    if not doc:
        st.session_state["_vektor_mongo_connected"] = True
        return False
    tables = ("projects", "tasks", "subtasks", "task_attachments", "resources", "milestones", "settings", "saved_reports", "audit_log", "calendar_events", "project_assignments")
    with conn() as c:
        for table in tables:
            rows = doc.get(table, [])
            c.execute(f"DELETE FROM {table}")
            if rows:
                cols = [r[1] for r in c.execute(f"PRAGMA table_info({table})").fetchall()]
                valid_rows = [{k: v for k, v in row.items() if k in cols} for row in rows]
                for row in valid_rows:
                    if row:
                        fields = list(row)
                        placeholders = ",".join("?" for _ in fields)
                        c.execute(
                            f"INSERT OR REPLACE INTO {table} ({','.join(fields)}) VALUES ({placeholders})",
                            tuple(row[k] for k in fields),
                        )
        c.commit()
    st.session_state["_vektor_mongo_connected"] = True
    st.session_state["_vektor_has_real_data"] = True
    return True


def _recalculate_project_rollups() -> None:
    """Derive task progress from completion marks, then calculate project rollups and days left."""
    with conn() as c:
        subtask_metrics = {
            int(row["task_id"]): (int(row["total"] or 0), int(row["done"] or 0))
            for row in c.execute(
                "SELECT task_id, COUNT(*) AS total, COALESCE(SUM(CASE WHEN lower(CAST(completed AS TEXT)) IN ('1','true','yes','completed') THEN 1 ELSE 0 END),0) AS done FROM subtasks GROUP BY task_id"
            ).fetchall()
        }
        task_rows = c.execute("SELECT id, project_id, status, due_date FROM tasks").fetchall()
        task_metrics = {}
        for task in task_rows:
            sub_total, sub_done = subtask_metrics.get(int(task["id"]), (0, 0))
            status = str(task["status"] or "Not Started")
            if sub_total:
                progress = round(sub_done / sub_total * 100)
                if sub_done == sub_total:
                    status = "Completed"
                elif sub_done > 0:
                    status = "In Progress"
                elif status.strip().lower() == "completed":
                    status = "Not Started"
                completed = sub_done == sub_total
                total_units, done_units = sub_total, sub_done
            else:
                completed = status.strip().lower() == "completed"
                progress = 100 if completed else 0
                total_units, done_units = 1, int(completed)
            c.execute("UPDATE tasks SET status=?, progress=? WHERE id=?", (status, progress, task["id"]))
            task_metrics.setdefault(task["project_id"], []).append({
                "completed": completed, "total_units": total_units, "done_units": done_units,
                "due_date": task["due_date"], "status": status,
            })

        today = date.today()
        projects = c.execute("SELECT id, start_date, end_date, status FROM projects").fetchall()
        for project in projects:
            tasks = task_metrics.get(project["id"], [])
            total_tasks = len(tasks)
            done_tasks = sum(1 for task in tasks if task["completed"])
            total_units = sum(task["total_units"] for task in tasks)
            done_units = sum(task["done_units"] for task in tasks)
            progress = round(done_units / total_units * 100) if total_units else 0
            try:
                end_date = date.fromisoformat(str(project["end_date"])[:10]) if project["end_date"] else None
            except (TypeError, ValueError):
                end_date = None
            if total_tasks and done_tasks == total_tasks:
                status, health = "Completed", "Healthy"
            elif end_date and end_date < today and progress < 100:
                status, health = "Delayed", "Critical"
            elif end_date and (end_date - today).days <= 7 and progress < 75:
                status, health = "At Risk", "At Risk"
            elif total_tasks and done_units > 0:
                status, health = "In Progress", "Healthy"
            else:
                status, health = (project["status"] or "Not Started"), "Healthy"
            c.execute(
                "UPDATE projects SET tasks_done=?, tasks_total=?, progress=?, days_left=?, status=?, health=? WHERE id=?",
                (done_tasks, total_tasks, progress, (end_date - today).days if end_date else None, status, health, project["id"]),
            )
        c.commit()


def sql(q: str, p: tuple = ()) -> None:
    """Execute an app mutation, then persist it and recalculate dependent project metrics."""
    normalized = " ".join(q.strip().lower().split())
    match = re.search(r"\b(?:insert\s+into|update|delete\s+from)\s+([a-z_]+)", normalized)
    table = match.group(1) if match else ""
    is_entity_mutation = table in {"projects", "tasks", "subtasks", "task_attachments", "resources", "milestones", "saved_reports", "settings", "audit_log", "calendar_events", "project_assignments"}
    with conn() as c:
        c.execute(q, p)
        c.commit()
    if table in {"tasks", "subtasks", "projects", "milestones"}:
        _recalculate_project_rollups()
    if is_entity_mutation:
        _persist_snapshot()


def df(q: str, p: tuple = ()) -> pd.DataFrame:
    with conn() as c:
        return pd.read_sql_query(q, c, params=p)


def init_db(force_reset: bool = False) -> None:
    with conn() as c:
        # Auto-upgrade older local databases to the current schema
        cols = [r[1] for r in c.execute("PRAGMA table_info(projects)").fetchall()]
        if force_reset or ("category" not in cols and len(cols) > 0):
            c.executescript(
                """
                DROP TABLE IF EXISTS projects;
                DROP TABLE IF EXISTS tasks;
                DROP TABLE IF EXISTS subtasks;
                DROP TABLE IF EXISTS task_attachments;
                DROP TABLE IF EXISTS resources;
                DROP TABLE IF EXISTS milestones;
                DROP TABLE IF EXISTS settings;
                DROP TABLE IF EXISTS saved_reports;
                DROP TABLE IF EXISTS audit_log;
                DROP TABLE IF EXISTS calendar_events;
                DROP TABLE IF EXISTS project_assignments;
                """
            )

        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS projects(
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                owner TEXT,
                team TEXT,
                category TEXT,
                status TEXT,
                priority TEXT,
                start_date TEXT,
                end_date TEXT,
                progress INTEGER,
                days_left INTEGER,
                tasks_done INTEGER,
                tasks_total INTEGER,
                tags TEXT,
                bar_color TEXT,
                short_desc TEXT,
                description TEXT,
                health TEXT
            );
            CREATE TABLE IF NOT EXISTS tasks(
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                project_id INTEGER,
                team TEXT,
                status TEXT,
                priority TEXT,
                start_date TEXT,
                due_date TEXT,
                assignee TEXT,
                progress INTEGER,
                dependency TEXT,
                notes TEXT
            );
            CREATE TABLE IF NOT EXISTS subtasks(
                id INTEGER PRIMARY KEY,
                task_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                completed INTEGER DEFAULT 0,
                due_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE INDEX IF NOT EXISTS idx_subtasks_task_id ON subtasks(task_id);
            CREATE TABLE IF NOT EXISTS resources(
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT,
                team TEXT,
                role TEXT,
                status TEXT,
                capacity INTEGER,
                notes TEXT,
                skills TEXT
            );
            CREATE TABLE IF NOT EXISTS milestones(
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                project_id INTEGER,
                team TEXT,
                status TEXT,
                start_date TEXT,
                due_date TEXT,
                completion_date TEXT,
                dependency TEXT,
                notes TEXT
            );
            CREATE TABLE IF NOT EXISTS settings(
                key TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS saved_reports(
                id INTEGER PRIMARY KEY,
                name TEXT,
                report_type TEXT,
                config TEXT,
                last_generated TEXT,
                created_by TEXT
            );
            CREATE TABLE IF NOT EXISTS audit_log(
                id INTEGER PRIMARY KEY,
                event TEXT,
                item_type TEXT,
                item_id INTEGER,
                details TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS calendar_events(
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                event_type TEXT,
                project_id INTEGER,
                task_name TEXT,
                resources TEXT,
                start_date TEXT,
                start_time TEXT,
                end_date TEXT,
                end_time TEXT,
                allocation INTEGER,
                repeat_rule TEXT,
                description TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS project_assignments(
                id INTEGER PRIMARY KEY,
                resource_id INTEGER,
                project_id INTEGER,
                allocation INTEGER,
                start_date TEXT,
                end_date TEXT,
                status TEXT,
                role TEXT,
                notes TEXT
            );
            """
        )
        # Safe schema upgrades keep older local caches usable without resetting user data.
        _task_cols = {row[1] for row in c.execute("PRAGMA table_info(tasks)").fetchall()}
        for _col, _definition in {
            "parent_task_id": "INTEGER", "task_type": "TEXT", "tags": "TEXT", "description": "TEXT",
            "estimated_effort": "REAL", "actual_effort": "REAL", "created_at": "TEXT", "updated_at": "TEXT",
        }.items():
            if _col not in _task_cols:
                c.execute(f"ALTER TABLE tasks ADD COLUMN {_col} {_definition}")
        _resource_cols = {row[1] for row in c.execute("PRAGMA table_info(resources)").fetchall()}
        for _col, _definition in {
            "skills": "TEXT", "availability_start": "TEXT", "availability_end": "TEXT",
            "working_hours": "INTEGER", "employment_type": "TEXT", "photo_b64": "TEXT",
        }.items():
            if _col not in _resource_cols:
                c.execute(f"ALTER TABLE resources ADD COLUMN {_col} {_definition}")
        c.execute("CREATE TABLE IF NOT EXISTS task_attachments(id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, file_name TEXT NOT NULL, mime_type TEXT, file_size INTEGER, file_data BLOB, uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_task_attachments_task ON task_attachments(task_id)")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_project_assignments_resource_project ON project_assignments(resource_id, project_id)")

        c.commit()

project_columns = ['id', 'name', 'owner', 'team', 'category', 'status', 'priority', 'start_date', 'end_date', 'progress', 'days_left', 'tasks_done', 'tasks_total', 'tags', 'bar_color', 'short_desc', 'description', 'health']
task_columns = ['id', 'name', 'project_id', 'team', 'status', 'priority', 'start_date', 'due_date', 'assignee', 'progress', 'dependency', 'notes', 'parent_task_id', 'task_type', 'tags', 'description', 'estimated_effort', 'actual_effort', 'created_at', 'updated_at']
resource_columns = ['id', 'name', 'email', 'team', 'role', 'status', 'capacity', 'notes', 'skills', 'availability_start', 'availability_end', 'working_hours', 'employment_type', 'photo_b64']
milestone_columns = ['id', 'name', 'project_id', 'team', 'status', 'start_date', 'due_date', 'completion_date', 'dependency', 'notes']
report_columns = ['id', 'name', 'report_type', 'config', 'last_generated', 'created_by']
calendar_event_columns = ['id', 'title', 'event_type', 'project_id', 'task_name', 'resources', 'start_date', 'start_time', 'end_date', 'end_time', 'allocation', 'repeat_rule', 'description', 'created_at']
assignment_columns = ['id', 'resource_id', 'project_id', 'allocation', 'start_date', 'end_date', 'status', 'role', 'notes']


def _excel_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def import_excel_workbook(uploaded_file) -> dict:
    """Replace the tables represented by sheets in the uploaded Excel workbook."""
    sheet_map = {
        "Projects": ("projects", project_columns), "Tasks": ("tasks", task_columns),
        "Subtasks": ("subtasks", ["id", "task_id", "name", "completed", "due_date", "created_at"]),
        "Resources": ("resources", resource_columns), "Milestones": ("milestones", milestone_columns),
        "Calendar Events": ("calendar_events", calendar_event_columns),
        "Project Assignments": ("project_assignments", assignment_columns),
        "Saved Reports": ("saved_reports", report_columns), "Settings": ("settings", ["key", "value"]),
    }
    book = pd.ExcelFile(uploaded_file)
    actual_sheets = {s.strip().lower(): s for s in book.sheet_names}
    frames = {}
    for sheet, (table, columns) in sheet_map.items():
        actual = actual_sheets.get(sheet.lower())
        if not actual:
            continue
        frame = pd.read_excel(book, sheet_name=actual)
        frame.columns = [str(c).strip() for c in frame.columns]
        if table == "settings" and not {"key", "value"}.issubset(frame.columns):
            raise ValueError("Settings sheet requires 'key' and 'value' columns.")
        if table in {"projects", "tasks", "resources", "milestones", "saved_reports"} and not frame.empty and "name" not in frame:
            raise ValueError(f"{sheet} sheet requires a 'name' column.")
        frame = frame[[c for c in frame.columns if c in columns]].copy()
        for col in columns:
            if col not in frame:
                frame[col] = None
        frames[table] = frame[columns]
    if not frames:
        raise ValueError("No recognized Vektor sheets were found. Download and use the provided template.")

    if "projects" in frames:
        project_ids = set(pd.to_numeric(frames["projects"]["id"], errors="coerce").dropna().astype(int))
    else:
        project_ids = set(df("SELECT id FROM projects")["id"].tolist())
    for table in ("tasks", "milestones", "calendar_events", "project_assignments"):
        if table in frames and "project_id" in frames[table]:
            refs = set(pd.to_numeric(frames[table]["project_id"], errors="coerce").dropna().astype(int))
            missing = refs - project_ids
            if missing:
                raise ValueError(f"{table.title()} refers to unknown project IDs: {sorted(missing)[:10]}")
    if "project_assignments" in frames and "resources" in frames and "resource_id" in frames["project_assignments"]:
        resource_ids = set(pd.to_numeric(frames["resources"]["id"], errors="coerce").dropna().astype(int))
        assignment_resource_ids = set(pd.to_numeric(frames["project_assignments"]["resource_id"], errors="coerce").dropna().astype(int))
        missing_resources = assignment_resource_ids - resource_ids
        if missing_resources:
            raise ValueError(f"Project Assignments refers to unknown resource IDs: {sorted(missing_resources)[:10]}")

    if "subtasks" in frames and "tasks" in frames:
        task_ids = set(pd.to_numeric(frames["tasks"]["id"], errors="coerce").dropna().astype(int))
        subtask_task_ids = set(pd.to_numeric(frames["subtasks"]["task_id"], errors="coerce").dropna().astype(int))
        missing_tasks = subtask_task_ids - task_ids
        if missing_tasks:
            raise ValueError(f"Subtasks refers to unknown task IDs: {sorted(missing_tasks)[:10]}")

    order = ["projects", "resources", "tasks", "subtasks", "milestones", "calendar_events", "project_assignments", "saved_reports", "settings"]
    counts = {}
    with conn() as c:
        if "tasks" in frames:
            c.execute("DELETE FROM task_attachments")
        for table in order:
            if table not in frames:
                continue
            c.execute(f"DELETE FROM {table}")
            for _, row in frames[table].iterrows():
                record = {k: _excel_value(v) for k, v in row.to_dict().items()}
                record = {k: v for k, v in record.items() if v is not None}
                if table == "settings":
                    if not record.get("key"):
                        continue
                    record["value"] = str(record.get("value", ""))
                elif table in {"projects", "tasks", "resources", "milestones", "saved_reports", "subtasks"} and not record.get("name"):
                    continue
                if table == "subtasks" and not record.get("task_id"):
                    continue
                if not record:
                    continue
                cols = list(record)
                c.execute(f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) VALUES ({','.join('?' for _ in cols)})", [record[k] for k in cols])
            counts[table] = c.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('legacy_seed_cleanup_v1','true')")
        c.commit()
    st.session_state["_vektor_has_real_data"] = True
    _recalculate_project_rollups()
    _persist_snapshot()
    return counts



def get_setting(k: str, default: str = "") -> str:
    x = df("SELECT value FROM settings WHERE key=?", (k,))
    return str(x.iloc[0].value) if not x.empty else default


def put_setting(k: str, v: object) -> None:
    sql(
        "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (k, str(v)),
    )


def get_bool_setting(k: str, default: bool = False) -> bool:
    return get_setting(k, "true" if default else "false").strip().lower() in {"true", "1", "yes", "on"}


def log_event(event: str, typ: str = "", item_id: int | None = None, details: str = "") -> None:
    if get_setting("audit_logs", "true") == "true":
        sql(
            "INSERT INTO audit_log(event,item_type,item_id,details) VALUES(?,?,?,?)",
            (event, typ, item_id, details),
        )


def save_task_attachments(task_id: int, files) -> int:
    """Persist uploaded task files in SQLite and the shared Mongo snapshot."""
    saved = 0
    with conn() as c:
        for uploaded in files or []:
            data = uploaded.getvalue()
            if not data:
                continue
            if len(data) > 25 * 1024 * 1024:
                raise ValueError(f"{uploaded.name} exceeds the 25 MB per-file limit.")
            mime = getattr(uploaded, "type", None) or mimetypes.guess_type(uploaded.name)[0] or "application/octet-stream"
            c.execute("INSERT INTO task_attachments(task_id,file_name,mime_type,file_size,file_data) VALUES(?,?,?,?,?)", (task_id, uploaded.name, mime, len(data), sqlite3.Binary(data)))
            saved += 1
        c.commit()
    if saved:
        _persist_snapshot()
    return saved


def render_task_attachments(task_id: int, key_prefix: str) -> None:
    """Render real downloadable attachments with working delete actions."""
    files = df("SELECT id,file_name,mime_type,file_size,uploaded_at FROM task_attachments WHERE task_id=? ORDER BY id DESC", (int(task_id),))
    if files.empty:
        st.caption("No files attached to this task yet.")
        return
    for _, file_row in files.iterrows():
        c1, c2, c3 = st.columns([3.6, 1.2, 0.6])
        with c1:
            st.write(f"📎 {file_row['file_name']} · {round((int(file_row.get('file_size') or 0))/1024, 1)} KB")
        with c2:
            with conn() as _c:
                _blob = _c.execute("SELECT file_data FROM task_attachments WHERE id=?", (int(file_row['id']),)).fetchone()
            if _blob and _blob[0] is not None:
                st.download_button("Download", data=bytes(_blob[0]), file_name=str(file_row['file_name']), mime=str(file_row.get('mime_type') or "application/octet-stream"), key=f"{key_prefix}_download_{int(file_row['id'])}", use_container_width=True)
        with c3:
            if st.button("🗑", key=f"{key_prefix}_delete_{int(file_row['id'])}", help="Delete attachment"):
                sql("DELETE FROM task_attachments WHERE id=?", (int(file_row['id']),))
                st.rerun()


init_db()
try:
    _mongo_restored = _restore_snapshot()
    if not st.session_state.get("_vektor_mongo_connected", False):
        st.warning(
            "MongoDB is not configured. Add MONGODB_URI to Streamlit Secrets or the environment "
            "to enable durable shared storage. The app is currently using its local cache."
        )
except Exception as _mongo_startup_error:
    st.warning(
        "MongoDB is not connected. Check MONGODB_URI and database permissions. "
        "The app is currently using its local cache. "
        f"Details: {_mongo_startup_error}"
    )

# One-time migration removes legacy demonstration records from older local caches.
# No project/task/resource records are created by this script.
try:
    with conn() as _c:
        _already_clean = _c.execute("SELECT value FROM settings WHERE key IN ('legacy_seed_cleanup_v1','legacy_seed_cleanup_v2') LIMIT 1").fetchone()
        if not _already_clean:
            _demo_names = [r[0] for r in _c.execute("SELECT name FROM resources WHERE lower(coalesce(email,'')) LIKE '%@vektor.example'").fetchall()]
            _demo_project_ids = [r[0] for r in _c.execute("SELECT id FROM projects WHERE owner IN (" + (','.join('?' for _ in _demo_names) or "NULL") + ")", _demo_names).fetchall()] if _demo_names else []
            if _demo_project_ids:
                _ph = ','.join('?' for _ in _demo_project_ids)
                for _table in ("subtasks", "task_attachments"):
                    _c.execute(f"DELETE FROM {_table} WHERE task_id IN (SELECT id FROM tasks WHERE project_id IN ({_ph}))", _demo_project_ids)
                for _table in ("tasks", "milestones", "calendar_events", "project_assignments"):
                    _c.execute(f"DELETE FROM {_table} WHERE project_id IN ({_ph})", _demo_project_ids)
                _c.execute(f"DELETE FROM projects WHERE id IN ({_ph})", _demo_project_ids)
            _c.execute("DELETE FROM resources WHERE lower(coalesce(email,'')) LIKE '%@vektor.example'")
            _c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('legacy_seed_cleanup_v2','true')")
            _c.commit()
except Exception:
    pass

# Recalculate derived metrics on every run so Days Left is always relative to today's date.
_recalculate_project_rollups()

# Load Dataframes
DATE_FORMAT = get_setting("date_format", "MMM DD, YYYY")
P = df("SELECT * FROM projects ORDER BY id")
T = df("SELECT t.*, p.name AS project FROM tasks t LEFT JOIN projects p ON p.id=t.project_id ORDER BY t.due_date")
R = df("SELECT * FROM resources ORDER BY id")
M = df("SELECT m.*, p.name AS project FROM milestones m LEFT JOIN projects p ON p.id=m.project_id ORDER BY m.due_date")

projects_list = P.name.tolist()
pmap = dict(zip(P.name, P.id))
teams_list = sorted(set(R.team.dropna().tolist() + P.team.dropna().tolist()))

# Theme Configuration
active_theme = get_setting("theme", "HPE Green")
PRIMARY_ACCENT = GREEN if active_theme == "HPE Green" else (BLUE if active_theme == "Blue" else "#00A879")
SIDEBAR_BG = (
    "linear-gradient(180deg, #001918 0%, #002724 100%)"
    if active_theme == "HPE Green"
    else ("linear-gradient(180deg, #0B1938 0%, #132A5E 100%)" if active_theme == "Blue" else "linear-gradient(180deg, #0B0F19 0%, #151C2C 100%)")
)
APP_BG = "#F3F8FC" if active_theme != "Dark" else "#0F172A"
CARD_BG = "#FFFFFF" if active_theme != "Dark" else "#1E293B"
TEXT_COLOR = "#0F172A" if active_theme != "Dark" else "#F8FAFC"
BORDER_COLOR = "#E2ECF5" if active_theme != "Dark" else "#334155"


# ---------------------------------------------------------------------------
# Global CSS Styling (Hides Top Streamlit Icons & Matches Mockups Exactly)
# ---------------------------------------------------------------------------
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* Hide ALL Streamlit top icons, toolbar, menu, deploy button, and decoration */
header[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
.stDeployButton,
#MainMenu,
footer {{
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
}}

html, body, [class*="css"], [data-testid="stAppViewContainer"], button, input, textarea, select, label, p, div {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    color: {TEXT_COLOR};
    font-size: 14px;
}}
.stApp {{
    background: {APP_BG};
    color: {TEXT_COLOR};
}}
.block-container {{
    padding-top: 0.65rem !important;
    padding-bottom: 1rem !important;
    padding-left: 1.25rem !important;
    padding-right: 1.25rem !important;
    max-width: 1860px;
}}

/* Sidebar Styling */
[data-testid="stSidebar"] {{
    background: {SIDEBAR_BG} !important;
    border-right: 1px solid rgba(255,255,255,0.07);
    min-width: 230px !important;
    max-width: 230px !important;
}}
[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] {{
    display: none !important;
}}
[data-testid="stSidebar"] * {{
    color: #E2F2EF !important;
}}
[data-testid="stSidebar"] .stButton > button {{
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 0.62rem 0.95rem !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
    font-size: 14px !important;
    border: none !important;
    background: transparent !important;
    margin-bottom: 2px !important;
    transition: all 0.15s ease;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(255, 255, 255, 0.08) !important;
}}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {{
    background: linear-gradient(90deg, #00966C 0%, #007A58 100%) !important;
    color: #FFFFFF !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 12px rgba(0, 150, 108, 0.35);
}}

/* Primary & Secondary Buttons in Main Area */
.stButton > button, .stDownloadButton > button {{
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    border: 1px solid {BORDER_COLOR} !important;
    background: {CARD_BG};
    color: {TEXT_COLOR};
    transition: all 0.15s ease;
}}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
    background: {PRIMARY_ACCENT} !important;
    color: #FFFFFF !important;
    border-color: {PRIMARY_ACCENT} !important;
}}

/* Containers & Cards */
div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: {CARD_BG} !important;
    border: 1px solid {BORDER_COLOR} !important;
    border-radius: 12px !important;
    padding: 16px !important;
    box-shadow: 0 1px 4px rgba(15, 23, 42, 0.03);
}}

/* Tabs Styling */
.stTabs [data-baseweb="tab-list"] {{
    gap: 6px;
    background: {CARD_BG};
    padding: 6px 10px 0 10px;
    border-radius: 10px 10px 0 0;
    border: 1px solid {BORDER_COLOR};
    border-bottom: 2px solid #E2E8F0;
}}
.stTabs [data-baseweb="tab"] {{
    height: 40px;
    padding: 0 18px;
    font-weight: 600;
    font-size: 13.5px;
    color: #5A6A85;
    border-radius: 8px 8px 0 0;
}}
.stTabs [aria-selected="true"] {{
    color: {PRIMARY_ACCENT} !important;
    background: rgba(0, 135, 90, 0.08) !important;
    border-bottom: 3px solid {PRIMARY_ACCENT} !important;
}}

/* KPI Cards matching Overview & Projects Mockups */
.vk-kpi-card {{
    background: {CARD_BG};
    border: 1px solid {BORDER_COLOR};
    border-radius: 12px;
    padding: 14px 16px;
    display: flex;
    align-items: center;
    gap: 14px;
    min-height: 96px;
    box-shadow: 0 1px 4px rgba(15, 23, 42, 0.03);
}}
.vk-kpi-icon {{
    width: 46px;
    height: 46px;
    min-width: 46px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 20px;
    font-weight: 800;
}}
.vk-kpi-body {{
    flex: 1;
    min-width: 0;
}}
.vk-kpi-title {{
    font-size: 11.5px;
    font-weight: 700;
    color: #475569;
    margin-bottom: 2px;
}}
.vk-kpi-val {{
    font-size: 25px;
    font-weight: 800;
    color: #0F172A;
    line-height: 1.15;
}}
.vk-kpi-val span.pct {{
    font-size: 18px;
    font-weight: 700;
    color: #1E293B;
}}
.vk-kpi-sub {{
    font-size: 11.5px;
    color: #64748B;
    margin-top: 2px;
}}
.vk-kpi-bar {{
    height: 6px;
    background: #E2E8F0;
    border-radius: 999px;
    overflow: hidden;
    margin-top: 7px;
}}
.vk-kpi-bar-fill {{
    height: 100%;
    border-radius: 999px;
}}

/* Legacy metric card for other tabs */
.vk-metric-card {{
    display: flex;
    align-items: center;
    gap: 14px;
    background: {CARD_BG};
    border: 1px solid {BORDER_COLOR};
    border-radius: 12px;
    padding: 15px 18px;
    min-height: 94px;
}}
.vk-metric-icon {{
    width: 48px;
    height: 48px;
    min-width: 48px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    font-size: 21px;
}}
.vk-metric-label {{ font-size: 12px; font-weight: 600; color: #5A6A85; }}
.vk-metric-value {{ font-size: 26px; font-weight: 800; color: {TEXT_COLOR}; line-height: 1.15; }}
.vk-metric-note {{ font-size: 11.5px; color: #64748B; margin-top: 2px; }}

/* Section Headers */
.vk-section-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 15px;
    font-weight: 700;
    color: {TEXT_COLOR};
    margin-bottom: 10px;
}}

/* Custom HTML Table Styling */
.vk-table-wrap {{
    overflow-x: auto;
    border-radius: 8px;
    border: 1px solid #EDF2F7;
    background: {CARD_BG};
}}
table.vk-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
    text-align: left;
}}
table.vk-table thead tr {{
    background: #F8FAFC;
    border-bottom: 1px solid #E2E8F0;
}}
table.vk-table th {{
    padding: 6px 9px;
    font-weight: 600;
    font-size: 11.5px;
    color: #475569;
}}
table.vk-table td {{
    padding: 5px 9px;
    border-bottom: 1px solid #F1F5F9;
    color: #1E293B;
    vertical-align: middle;
}}
table.vk-table tbody tr:hover {{
    background: #F8FBFF;
}}

/* Status & Priority Badges */
.vk-badge {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 3px 10px;
    border-radius: 6px;
    font-size: 11.5px;
    font-weight: 600;
    white-space: nowrap;
}}
.vk-badge-green {{ background: #DCFCE7; color: #15803D; }}
.vk-badge-blue {{ background: #DBEAFE; color: #1D4ED8; }}
.vk-badge-orange {{ background: #FEF3C7; color: #B45309; }}
.vk-badge-red {{ background: #FEE2E2; color: #DC2626; }}
.vk-badge-gray {{ background: #E2E8F0; color: #475569; }}
.vk-badge-purple {{ background: #F3E8FF; color: #7C3AED; }}

/* Tag Chips */
.vk-tag {{
    display: inline-block;
    padding: 2px 9px;
    border-radius: 5px;
    font-size: 11px;
    font-weight: 600;
    margin-right: 5px;
}}

/* Progress Bar inside Tables */
.vk-prog-wrap {{
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 110px;
}}
.vk-prog-track {{
    flex: 1;
    height: 8px;
    background: #E2E8F0;
    border-radius: 999px;
    overflow: hidden;
}}
.vk-prog-fill {{
    height: 100%;
    border-radius: 999px;
}}
.vk-prog-text {{
    font-size: 12px;
    font-weight: 600;
    color: #334155;
    min-width: 32px;
    text-align: right;
}}

/* Avatar Chip */
.vk-avatar {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    font-size: 10px;
    font-weight: 700;
    color: white;
    margin-right: 7px;
}}

/* Compact density, readable at normal browser zoom */
[data-testid="stVerticalBlock"] {{ gap: 0.55rem !important; }}
[data-testid="stHorizontalBlock"] {{ gap: 0.65rem !important; }}
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataEditor"] [role="gridcell"] {{ padding-top: 2px !important; padding-bottom: 2px !important; }}
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataEditor"] [role="columnheader"] {{ padding-top: 4px !important; padding-bottom: 4px !important; }}
.stMarkdown p {{ margin-bottom: 0.35rem !important; }}
h1, h2, h3 {{ line-height: 1.2 !important; }}
.stButton > button, .stDownloadButton > button {{ min-height: 2.1rem !important; padding: 0.35rem 0.7rem !important; font-size: 14px !important; }}
/* Modal Dialog Width & Header Styling */
div[data-testid="stDialog"] > div > div {{
    max-width: 1020px !important;
    width: 95vw !important;
    border-radius: 14px !important;
    padding: 0 !important;
}}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# UI Helper Functions & Tile Toggle (No st.radio Anywhere)
# ---------------------------------------------------------------------------
def tile_toggle(label: str, options: list[str], key: str, default: str | None = None) -> str:
    """Button/tile toggle replacement for st.radio."""
    if key not in st.session_state:
        st.session_state[key] = default if default in options else options[0]
    if label:
        st.markdown(
            f"<div style='font-size:12.5px;font-weight:600;color:#475569;margin-bottom:4px'>{html.escape(label)}</div>",
            unsafe_allow_html=True,
        )
    cols = st.columns(len(options))
    for col, opt in zip(cols, options):
        is_sel = st.session_state[key] == opt
        with col:
            if st.button(
                opt,
                key=f"tile_{key}_{opt}",
                type="primary" if is_sel else "secondary",
                use_container_width=True,
            ):
                st.session_state[key] = opt
                st.rerun()
    return st.session_state[key]


def badge_html(val: object) -> str:
    s = str(val or "Unknown")
    mapping = {
        "On Track": "vk-badge-green",
        "Completed": "vk-badge-green",
        "Healthy": "vk-badge-green",
        "Available": "vk-badge-green",
        "Balanced": "vk-badge-green",
        "Connected": "vk-badge-green",
        "Active": "vk-badge-green",
        "In Progress": "vk-badge-blue",
        "Low": "vk-badge-blue",
        "Project Manager": "vk-badge-blue",
        "At Risk": "vk-badge-orange",
        "Partial": "vk-badge-orange",
        "Medium": "vk-badge-orange",
        "On Hold": "vk-badge-orange",
        "High": "vk-badge-red",
        "Delayed": "vk-badge-red",
        "Critical": "vk-badge-red",
        "Overdue": "vk-badge-red",
        "On Leave": "vk-badge-red",
        "Overloaded": "vk-badge-red",
        "Not Started": "vk-badge-gray",
        "Team Member": "vk-badge-gray",
        "Disconnected": "vk-badge-gray",
        "Admin": "vk-badge-purple",
    }
    cls = mapping.get(s, "vk-badge-gray")
    return f'<span class="vk-badge {cls}">{html.escape(s)}</span>'


def tag_pill_html(tag: str) -> str:
    t = tag.strip()
    styles = {
        "AI": ("#EDE9FE", "#6D28D9"),
        "Innovation": ("#FEE2E2", "#DC2626"),
        "Customer": ("#DBEAFE", "#1D4ED8"),
        "Web": ("#DBEAFE", "#1D4ED8"),
        "Data": ("#FEF3C7", "#B45309"),
        "Cloud": ("#DBEAFE", "#1D4ED8"),
        "IT": ("#DCFCE7", "#15803D"),
        "Development": ("#EDE9FE", "#6D28D9"),
        "Infrastructure": ("#DCFCE7", "#15803D"),
        "Mobile": ("#FEE2E2", "#DC2626"),
        "Automation": ("#DCFCE7", "#15803D"),
        "Operations": ("#DBEAFE", "#1D4ED8"),
        "Security": ("#F3E8FF", "#7C3AED"),
        "Compliance": ("#DBEAFE", "#1D4ED8"),
        "Sustainability": ("#DCFCE7", "#15803D"),
        "ESG": ("#F3E8FF", "#7C3AED"),
        "HR": ("#FEE2E2", "#DC2626"),
        "Training": ("#EDE9FE", "#6D28D9"),
        "Procurement": ("#FEF3C7", "#B45309"),
        "Marketing": ("#F3E8FF", "#7C3AED"),
        "Corporate": ("#DBEAFE", "#1D4ED8"),
    }
    bg, fg = styles.get(t, ("#F1F5F9", "#475569"))
    return f'<span class="vk-tag" style="background:{bg};color:{fg}">{html.escape(t)}</span>'


def progress_html(val: object, color_override: str | None = None) -> str:
    try:
        pct = max(0, min(100, int(float(val or 0))))
    except Exception:
        pct = 0
    color = color_override or (GREEN if pct >= 60 else (BLUE if pct >= 35 else ORANGE))
    return (
        f'<div class="vk-prog-wrap">'
        f'<div class="vk-prog-track"><div class="vk-prog-fill" style="width:{pct}%;background:{color}"></div></div>'
        f'<span class="vk-prog-text">{pct}%</span>'
        f"</div>"
    )


def avatar_html(name: object) -> str:
    s = str(name or "Unassigned").strip()
    if not s or s == "Unassigned":
        return '<span style="color:#94A3B8;font-style:italic">Unassigned</span>'
    parts = s.split()
    initials = (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()
    avatar_colors = {
    }
    bg = avatar_colors.get(s, "#475569")
    return f'<span class="vk-avatar" style="background:{bg}">{html.escape(initials)}</span>{html.escape(s)}'


def fmt_date(d_str: object) -> str:
    if d_str is None or (not isinstance(d_str, (date, datetime, pd.Timestamp)) and pd.isna(d_str)) or not str(d_str).strip():
        return "—"
    try:
        d = pd.to_datetime(d_str, errors="raise").date()
        fmt = DATE_FORMAT
        if fmt in {"DD/MM/YYYY", "24/04/2024", "dd/mm/yyyy"}:
            return d.strftime("%d/%m/%Y")
        if fmt in {"MM/DD/YYYY", "04/24/2024", "mm/dd/yyyy"}:
            return d.strftime("%m/%d/%Y")
        return d.strftime("%b %d, %Y").replace(" 0", " ")
    except Exception:
        return str(d_str)


def render_rich_table(
    data: pd.DataFrame,
    columns: list[tuple[str, str, str]],
    max_rows: int | None = None,
) -> None:
    if data.empty:
        st.info("No records match the current filters.")
        return
    view = data.head(max_rows) if max_rows else data
    ths = "".join(f"<th>{html.escape(label)}</th>" for _, label, _ in columns)
    rows_html = []
    for _, row in view.iterrows():
        tds = []
        for col, _, ctype in columns:
            val = row.get(col, "")
            if ctype == "badge":
                cell = badge_html(val)
            elif ctype == "progress":
                cell = progress_html(val)
            elif ctype == "avatar":
                cell = avatar_html(val)
            elif ctype == "date":
                cell = fmt_date(val)
            elif ctype == "bold":
                cell = f'<span style="font-weight:600;color:{TEXT_COLOR}">{html.escape(str(val if pd.notna(val) else "—"))}</span>'
            else:
                cell = html.escape(str(val if pd.notna(val) and str(val) != "" else "—"))
            tds.append(f"<td>{cell}</td>")
        rows_html.append("<tr>" + "".join(tds) + "</tr>")
    table_markup = (
        f'<div class="vk-table-wrap"><table class="vk-table">'
        f"<thead><tr>{ths}</tr></thead>"
        f'<tbody>{"".join(rows_html)}</tbody>'
        f"</table></div>"
    )
    st.markdown(table_markup, unsafe_allow_html=True)


def kpi_card_mockup(
    title: str,
    main_val: str,
    pct_text: str = "",
    sub_text: str = "",
    bar_pct: int | None = None,
    icon_html: str = "✓",
    icon_bg: str = "#DCFCE7",
    icon_fg: str = "#00875A",
    bar_color: str = "#00875A",
) -> None:
    pct_span = f' <span class="pct">{html.escape(pct_text)}</span>' if pct_text else ""
    bottom_part = (
        f'<div class="vk-kpi-bar"><div class="vk-kpi-bar-fill" style="width:{bar_pct}%;background:{bar_color}"></div></div>'
        if bar_pct is not None
        else f'<div class="vk-kpi-sub">{html.escape(sub_text)}</div>'
    )
    st.markdown(
        f"""
        <div class="vk-kpi-card">
            <div class="vk-kpi-icon" style="background:{icon_bg};color:{icon_fg}">{icon_html}</div>
            <div class="vk-kpi-body">
                <div class="vk-kpi-title">{html.escape(title)}</div>
                <div class="vk-kpi-val">{html.escape(main_val)}{pct_span}</div>
                {bottom_part}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label: str, value: object, note: str = "", icon: str = "📊", tone: str = "green") -> None:
    tones = {
        "green": ("#E6F6F0", GREEN),
        "blue": ("#E8F2FF", BLUE),
        "orange": ("#FFF4E5", ORANGE),
        "red": ("#FEECEB", RED),
        "purple": ("#F3E8FF", PURPLE),
    }
    bg, fg = tones.get(tone, tones["green"])
    st.markdown(
        f"""
        <div class="vk-metric-card">
            <div class="vk-metric-icon" style="background:{bg};color:{fg}">{icon}</div>
            <div>
                <div class="vk-metric-label">{html.escape(label)}</div>
                <div class="vk-metric-value">{html.escape(str(value))}</div>
                <div class="vk-metric-note">{html.escape(note)}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(title: str, right_text: str = "") -> None:
    right_span = f'<span style="font-size:12px;font-weight:500;color:#64748B">{html.escape(right_text)}</span>' if right_text else ""
    st.markdown(f'<div class="vk-section-header"><span>{html.escape(title)}</span>{right_span}</div>', unsafe_allow_html=True)


def chart_status_donut(data: pd.DataFrame, col: str, center_label: str = "Total", height: int = 240) -> None:
    if data.empty or col not in data:
        st.info("No data available.")
        return
    z = data[col].fillna("Unknown").value_counts().rename_axis("Status").reset_index(name="Count")
    colors = {
        "On Track": GREEN,
        "Healthy": GREEN,
        "Completed": GREEN,
        "Available": GREEN,
        "Balanced": GREEN,
        "Low": BLUE,
        "In Progress": BLUE,
        "At Risk": ORANGE,
        "Partial": ORANGE,
        "Medium": ORANGE,
        "High": RED,
        "Delayed": RED,
        "Critical": RED,
        "On Leave": RED,
        "Overdue": RED,
        "Overloaded": RED,
        "Not Started": GRAY,
    }
    total = int(z["Count"].sum())
    fig = px.pie(z, names="Status", values="Count", hole=0.62, color="Status", color_discrete_map=colors)
    fig.update_traces(textposition="inside", textinfo="percent", insidetextorientation="horizontal")
    fig.update_layout(
        height=height,
        margin=dict(l=0, r=0, t=6, b=6),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0, font=dict(size=11)),
        annotations=[
            dict(
                text=f"<b>{total}</b><br><span style='font-size:10px;color:#64748B'>{center_label}</span>",
                x=0.5,
                y=0.5,
                font_size=16,
                showarrow=False,
            )
        ],
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def progress_bar_chart(data: pd.DataFrame, name: str = "name", val: str = "progress", height: int = 250) -> None:
    if data.empty:
        st.info("No data available.")
        return
    z = data[[name, val]].copy()
    z[val] = pd.to_numeric(z[val], errors="coerce").fillna(0)
    bar_colors = [GREEN if v >= 75 else (BLUE if v >= 50 else (ORANGE if v >= 30 else RED)) for v in z[val]]
    fig = go.Figure(
        go.Bar(
            x=z[val],
            y=z[name],
            orientation="h",
            text=[f"{int(v)}%" for v in z[val]],
            textposition="outside",
            marker=dict(color=bar_colors, cornerradius=5),
            cliponaxis=False,
        )
    )
    fig.update_layout(
        height=max(height, len(z) * 38),
        margin=dict(l=0, r=38, t=6, b=6),
        xaxis=dict(range=[0, 112], showgrid=True, gridcolor="#EDF2F7", title=None),
        yaxis=dict(title=None),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Export a readable workbook with frozen headers, filters, widths, and progress formatting."""
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        for sheet_name, d in sheets.items():
            frame = d.copy() if isinstance(d, pd.DataFrame) else pd.DataFrame(d)
            frame.to_excel(w, sheet_name=sheet_name[:31], index=False)
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.formatting.rule import DataBarRule
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else f"A1:{ws.cell(1, max(1, ws.max_column)).coordinate}"
            ws.row_dimensions[1].height = 30
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="053241")
                cell.font = Font(name="Aptos", bold=True, color="FFFFFF", size=10)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(bottom=Side(style="medium", color="00A879"))
            for col_cells in ws.columns:
                col_letter = col_cells[0].column_letter
                sample_values = [str(c.value) for c in col_cells[:100] if c.value is not None]
                max_len = max((len(v) for v in sample_values), default=10)
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 38)
                header = str(ws.cell(1, col_cells[0].column).value or "").lower()
                if any(token in header for token in ("description", "notes", "details", "config")):
                    ws.column_dimensions[col_letter].width = 34
                if "date" in header or header.endswith("_at") or header in {"start", "end", "due_date", "start_date", "end_date", "completion_date"}:
                    for cell in col_cells[1:]:
                        if cell.value is not None:
                            cell.number_format = "mmm d, yyyy"
                if "progress" in header or header in {"allocation", "capacity", "task_progress"}:
                    if ws.max_row > 1:
                        ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="00A879", showValue=True))
            for row_idx in range(2, ws.max_row + 1):
                if row_idx % 2 == 0:
                    for cell in ws[row_idx]:
                        cell.fill = PatternFill("solid", fgColor="F4F8FA")
                for cell in ws[row_idx]:
                    cell.alignment = Alignment(vertical="top", wrap_text=cell.data_type == "s" and len(str(cell.value or "")) > 45)
            if ws.max_row > 1:
                ws.sheet_view.showGridLines = False
    return b.getvalue()


def tracker_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Export current tracker data with live Excel formulas for task/project/assignment progress."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.formatting.rule import DataBarRule
    from openpyxl.utils import get_column_letter
    raw = excel_bytes(sheets)
    wb = load_workbook(io.BytesIO(raw))
    if "Tasks" not in wb.sheetnames:
        return raw
    tasks_ws = wb["Tasks"]
    task_headers = {str(tasks_ws.cell(1, c).value or "").strip(): c for c in range(1, tasks_ws.max_column + 1)}
    # Separate manual input from calculated status so tasks with subtasks always show their true status.
    if "Manual Status" not in task_headers:
        manual_col = tasks_ws.max_column + 1
        tasks_ws.cell(1, manual_col).value = "Manual Status"
        task_headers["Manual Status"] = manual_col
        for r in range(2, max(tasks_ws.max_row, 1000) + 1):
            tasks_ws.cell(r, manual_col).value = tasks_ws.cell(r, task_headers.get("status", 5)).value if r <= tasks_ws.max_row else None
    for helper in ("Subtasks Total", "Subtasks Completed", "Calculated Status"):
        if helper not in task_headers:
            col = tasks_ws.max_column + 1
            tasks_ws.cell(1, col).value = helper
            task_headers[helper] = col
    def col_letter(ws, header, default):
        mapping = {str(ws.cell(1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
        return get_column_letter(mapping.get(header, default))
    task_id = col_letter(tasks_ws, "id", 1)
    task_project = col_letter(tasks_ws, "project_id", 3)
    task_status = col_letter(tasks_ws, "status", 5)
    task_manual = col_letter(tasks_ws, "Manual Status", task_headers["Manual Status"])
    task_progress = col_letter(tasks_ws, "progress", 10)
    total_col = col_letter(tasks_ws, "Subtasks Total", task_headers["Subtasks Total"])
    done_col = col_letter(tasks_ws, "Subtasks Completed", task_headers["Subtasks Completed"])
    calc_status_col = col_letter(tasks_ws, "Calculated Status", task_headers["Calculated Status"])
    subtasks_ws = wb["Subtasks"] if "Subtasks" in wb.sheetnames else None
    if subtasks_ws is not None:
        sub_task_id = col_letter(subtasks_ws, "task_id", 2)
        sub_completed = col_letter(subtasks_ws, "completed", 4)
        for r in range(2, 1001):
            tasks_ws[f"{total_col}{r}"] = f'=IF(${task_id}{r}="","",COUNTIF(Subtasks!${sub_task_id}:${sub_task_id},${task_id}{r}))'
            tasks_ws[f"{done_col}{r}"] = f'=IF(${task_id}{r}="","",COUNTIFS(Subtasks!${sub_task_id}:${sub_task_id},${task_id}{r},Subtasks!${sub_completed}:${sub_completed},TRUE()))'
    else:
        for r in range(2, 1001):
            tasks_ws[f"{total_col}{r}"] = f'=IF(${task_id}{r}="","",0)'
            tasks_ws[f"{done_col}{r}"] = f'=IF(${task_id}{r}="","",0)'
    for r in range(2, 1001):
        tasks_ws[f"{calc_status_col}{r}"] = f'=IF(${task_id}{r}="","",IF(${total_col}{r}>0,IF(${done_col}{r}=${total_col}{r},"Completed",IF(${done_col}{r}>0,"In Progress","Not Started")),IF(LOWER(${task_manual}{r})="completed","Completed",IF(LOWER(${task_manual}{r})="in progress","In Progress",IF(LOWER(${task_manual}{r})="at risk","At Risk",IF(LOWER(${task_manual}{r})="delayed","Delayed",IF(LOWER(${task_manual}{r})="on hold","On Hold","Not Started"))))))'
        tasks_ws[f"{task_status}{r}"] = f'=IF(${task_id}{r}="","",IF(${total_col}{r}>0,${calc_status_col}{r},IF(${task_manual}{r}="","Not Started",${task_manual}{r})))'
        tasks_ws[f"{task_progress}{r}"] = f'=IF(${task_id}{r}="","",IF(${total_col}{r}>0,IFERROR(${done_col}{r}/${total_col}{r}*100,0),IF(LOWER(${task_manual}{r})="completed",100,0)))'
    if "Projects" in wb.sheetnames:
        ws = wb["Projects"]
        headers = {str(ws.cell(1,c).value or "").strip():c for c in range(1,ws.max_column+1)}
        project_id = get_column_letter(headers.get("id",1))
        end_col = get_column_letter(headers.get("end_date",9))
        progress_col = get_column_letter(headers.get("progress",10))
        days_col = get_column_letter(headers.get("days_left",11))
        done_tasks_col = get_column_letter(headers.get("tasks_done",12))
        total_tasks_col = get_column_letter(headers.get("tasks_total",13))
        for r in range(2,1001):
            ws[f"{progress_col}{r}"] = f'=IF(${project_id}{r}="","",IFERROR((SUMIFS(Tasks!${done_col}:${done_col},Tasks!${task_project}:${task_project},${project_id}{r})+COUNTIFS(Tasks!${task_project}:${task_project},${project_id}{r},Tasks!${total_col}:${total_col},0,Tasks!${task_status}:${task_status},"Completed"))/(SUMIFS(Tasks!${total_col}:${total_col},Tasks!${task_project}:${task_project},${project_id}{r})+COUNTIFS(Tasks!${task_project}:${task_project},${project_id}{r},Tasks!${total_col}:${total_col},0))*100,0))'
            ws[f"{days_col}{r}"] = f'=IF(${project_id}{r}="","",IF(${end_col}{r}="","",${end_col}{r}-TODAY()))'
            ws[f"{done_tasks_col}{r}"] = f'=IF(${project_id}{r}="","",COUNTIFS(Tasks!${task_project}:${task_project},${project_id}{r},Tasks!${task_status}:${task_status},"Completed"))'
            ws[f"{total_tasks_col}{r}"] = f'=IF(${project_id}{r}="","",COUNTIF(Tasks!${task_project}:${task_project},${project_id}{r}))'
    if "Project Assignments" in wb.sheetnames and "Resources" in wb.sheetnames:
        ws = wb["Project Assignments"]
        headers = {str(ws.cell(1,c).value or "").strip():c for c in range(1,ws.max_column+1)}
        progress_c = headers.get("task_progress")
        if not progress_c:
            progress_c = ws.max_column + 1
            ws.cell(1,progress_c).value = "task_progress"
        resource_id_col = get_column_letter(headers.get("resource_id",2))
        project_col = get_column_letter(headers.get("project_id",3))
        resource_ws = wb["Resources"]
        for r in range(2,1001):
            ws.cell(r,progress_c).value = f'=IF(OR(${resource_id_col}{r}="",${project_col}{r}=""),"",IFERROR((SUMIFS(Tasks!${done_col}:${done_col},Tasks!${task_project}:${task_project},${project_col}{r},Tasks!$I:$I,VLOOKUP(${resource_id_col}{r},Resources!$A:$B,2,FALSE()))+COUNTIFS(Tasks!${task_project}:${task_project},${project_col}{r},Tasks!$I:$I,VLOOKUP(${resource_id_col}{r},Resources!$A:$B,2,FALSE()),Tasks!${total_col}:${total_col},0,Tasks!${task_status}:${task_status},"Completed"))/(SUMIFS(Tasks!${total_col}:${total_col},Tasks!${task_project}:${task_project},${project_col}{r},Tasks!$I:$I,VLOOKUP(${resource_id_col}{r},Resources!$A:$B,2,FALSE()))+COUNTIFS(Tasks!${task_project}:${task_project},${project_col}{r},Tasks!$I:$I,VLOOKUP(${resource_id_col}{r},Resources!$A:$B,2,FALSE()),Tasks!${total_col}:${total_col},0))*100,0))'
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1,ws.max_row)}"
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="053241")
            cell.font = Font(name="Aptos Display", bold=True, color="FFFFFF", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        for c in range(1, ws.max_column + 1):
            header = str(ws.cell(1,c).value or "")
            letter = get_column_letter(c)
            ws.column_dimensions[letter].width = min(max(len(header)+3, 12), 34)
            if header in {"Subtasks Total", "Subtasks Completed", "Calculated Status"}:
                ws.column_dimensions[letter].hidden = True
            if header.lower() in {"progress", "task_progress", "allocation", "capacity"}:
                for r in range(2, min(1000, max(ws.max_row, 2)) + 1):
                    ws[f"{letter}{r}"].number_format = '0"%"'
                ws.conditional_formatting.add(f"{letter}2:{letter}1000", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=200 if header.lower() in {"allocation", "capacity"} else 100, color="00A879", showValue=True))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def render_top_header(title: str, subtitle: str, show_filters: bool = True) -> None:
    """Clean page header without any profile bar at the top."""
    if show_filters:
        c1, c2, c3 = st.columns([2.2, 1, 1])
        with c1:
            st.markdown(
                f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>{html.escape(title)}</div>"
                f"<div style='font-size:13.5px;color:#64748B;margin-top:3px'>{html.escape(subtitle)}</div>",
                unsafe_allow_html=True,
            )
        with c2:
            st.selectbox("Project", ["All Projects"] + projects_list, key="filter_project", label_visibility="collapsed")
        with c3:
            st.selectbox("Team", ["All Teams"] + teams_list, key="filter_team", label_visibility="collapsed")
    else:
        st.markdown(
            f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15;margin-bottom:3px'>{html.escape(title)}</div>"
            f"<div style='font-size:13.5px;color:#64748B;margin-bottom:12px'>{html.escape(subtitle)}</div>",
            unsafe_allow_html=True,
        )


def apply_global_filters(data: pd.DataFrame, project_col: str = "project") -> pd.DataFrame:
    x = data.copy()
    p = st.session_state.get("filter_project", "All Projects")
    t = st.session_state.get("filter_team", "All Teams")
    if p != "All Projects" and project_col in x.columns:
        x = x[x[project_col] == p]
    if t != "All Teams" and "team" in x.columns:
        x = x[x["team"] == t]
    return x


# ---------------------------------------------------------------------------
# Sidebar Navigation (Exact Replica of Mockup Sidebar)
# ---------------------------------------------------------------------------
with st.sidebar:
    ws_title = get_setting("workspace_name", "Vektor")
    logo_b64 = get_setting("workspace_logo_b64", "")
    if logo_b64:
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" style="max-height:42px;margin-bottom:6px">',
            unsafe_allow_html=True,
        )
    st.markdown(
        f"""
        <div style="padding: 4px 6px 22px 6px;">
            <div style="font-size:31px;font-weight:800;letter-spacing:-1.2px;line-height:1;color:#FFFFFF !important;">
                {html.escape(ws_title)} <span style="color:#00C48C;font-weight:900">❯❯</span>
            </div>
            <div style="font-size:11.5px;color:#B2DFDB !important;margin-top:5px;font-weight:500;">
                Direction + Velocity
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    nav_items = {
        "⊞  Overview": "Overview",
        "📁  Projects": "Projects",
        "📊  Gantt Chart": "Gantt Chart",
        "☰  Task List": "Task List",
        "👥  Resources": "Resources",
        "📅  Calendar": "Calendar",
        "📄  Reports": "Reports",
        "⚙  Settings": "Settings",
    }
    default_page = get_setting("default_landing", "Overview")
    if "page" not in st.session_state:
        st.session_state.page = default_page if default_page in nav_items.values() else "Overview"

    for label, name in nav_items.items():
        is_active = st.session_state.page == name
        if st.button(
            label,
            key=f"nav_{name}",
            use_container_width=True,
            type="primary" if is_active else "secondary",
        ):
            st.session_state.page = name
            st.rerun()

page = st.session_state.page


# ===========================================================================
# PROJECT MODALS (View Project, Edit Project, Delete Project, New Project)
# ===========================================================================
def clean_html(raw: str) -> str:
    """Collapse indentation so Streamlit Markdown never renders HTML as a code block."""
    return "".join(line.strip() for line in raw.splitlines() if line.strip())


def render_timeline_stepper_html(project_id: int | None = None) -> str:
    # Render milestones for the selected project without hardcoded demo dates.
    if project_id is None:
        milestones = df("SELECT * FROM milestones ORDER BY due_date")
    else:
        milestones = df("SELECT * FROM milestones WHERE project_id=? ORDER BY due_date", (project_id,))
    if milestones.empty:
        return '<div style="padding:18px;color:#64748B;font-size:12px">No milestones are available for this project yet.</div>'
    steps = []
    for _, m in milestones.head(6).iterrows():
        status = str(m.get("status") or "Not Started")
        state = "done" if status.lower() == "completed" else ("current" if status.lower() == "in progress" else "todo")
        box_bg, box_fg = ("#DCFCE7", "#15803D") if state == "done" else (("#DBEAFE", "#1D4ED8") if state == "current" else ("#F1F5F9", "#475569"))
        steps.append((str(m.get("name") or "Milestone"), fmt_date(m.get("due_date")), state, str(m.get("name") or "Milestone"), status, box_bg, box_fg))
    nodes_html = ""
    for idx, (st_title, st_date, state, box_title, box_sub, box_bg, box_fg) in enumerate(steps):
        if state == "done":
            circle = '<div style="width:24px;height:24px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;z-index:2;position:relative">✓</div>'
        elif state == "current":
            circle = '<div style="width:24px;height:24px;border-radius:50%;background:#1473E6;border:3px solid #BFDBFE;color:white;display:flex;align-items:center;justify-content:center;font-size:10px;z-index:2;position:relative">●</div>'
        else:
            circle = '<div style="width:22px;height:22px;border-radius:50%;background:#94A3B8;color:white;display:flex;align-items:center;justify-content:center;font-size:10px;z-index:2;position:relative"></div>'
        line_color = "#00875A" if state == "done" else ("#1473E6" if state == "current" else "#CBD5E1")
        line_right = f'<div style="position:absolute;top:46px;left:50%;width:100%;height:3px;background:{line_color};z-index:1"></div>' if idx < len(steps)-1 else ""
        nodes_html += clean_html(f'''<div style="flex:1;position:relative;display:flex;flex-direction:column;align-items:center;text-align:center;padding:0 4px">{line_right}<div style="font-size:11.5px;font-weight:700;color:#1E293B">{html.escape(st_title)}</div><div style="font-size:10.5px;color:#64748B;margin-bottom:6px">{html.escape(st_date)}</div>{circle}<div style="margin-top:10px;background:{box_bg};border-radius:6px;padding:6px;width:94%;min-height:44px"><div style="font-size:10.5px;font-weight:600;color:{box_fg}">{html.escape(box_title)}</div><div style="font-size:10.5px;font-weight:700;color:{box_fg}">{html.escape(box_sub)}</div></div></div>''')
    return f'<div style="display:flex;align-items:flex-start;justify-content:space-between;padding:8px 0 4px 0">{nodes_html}</div>'


def render_project_progress_donut(pct: int, height: int = 195) -> go.Figure:
    fig = go.Figure(
        go.Pie(
            values=[max(1, pct), max(0, 100 - pct)],
            labels=["Complete", "Remaining"],
            hole=0.68,
            sort=False,
            direction="clockwise",
            marker=dict(colors=[GREEN, "#CBD5E1"]),
            textinfo="none",
            hoverinfo="label+percent",
        )
    )
    fig.update_layout(
        height=height,
        margin=dict(l=6, r=6, t=6, b=6),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        annotations=[
            dict(
                text=f"<b style='font-size:22px;color:#0F172A'>{pct}%</b><br><span style='font-size:11px;color:#475569'>Complete</span>",
                x=0.5,
                y=0.5,
                showarrow=False,
            )
        ],
    )
    return fig


def project_task_status_html(project_id: int) -> str:
    """Dynamic task-status summary for a project; derived from persisted task/subtask completion."""
    task_rows = df("SELECT status FROM tasks WHERE project_id=?", (project_id,))
    statuses = task_rows["status"].fillna("Not Started").astype(str).str.strip().str.lower() if not task_rows.empty else pd.Series(dtype=str)
    total = len(statuses)
    labels = [
        ("Completed", "completed", "#00875A"),
        ("In Progress", "in progress", "#1473E6"),
        ("At Risk", "at risk", "#F28C00"),
        ("Not Started", "not started", "#64748B"),
    ]
    rows = []
    for label, match, color in labels:
        count = int((statuses == match).sum())
        pct = round(count / total * 100) if total else 0
        rows.append(f"<div style='display:flex;justify-content:space-between'><span><span style='color:{color}'>●</span> {label}</span><b>{count} ({pct}%)</b></div>")
    return "<div style='font-size:12px;line-height:2.1;padding:4px 8px'>" + "".join(rows) + "</div>"


@st.dialog("View Project", width="large")
def view_project_modal(pid: int) -> None:
    row_df = df("SELECT * FROM projects WHERE id=?", (pid,))
    if row_df.empty:
        st.warning("Project not found.")
        return
    r = row_df.iloc[0]
    tags_html = "".join(tag_pill_html(t) for t in str(r.tags or "").split(",") if t.strip())
    full_desc = str(r.description or r.short_desc or "")
    pct = int(r.progress or 0)
    _project_start_date = pd.to_datetime(r.start_date, errors="coerce")
    _project_end_date = pd.to_datetime(r.end_date, errors="coerce")
    _project_duration = max(0, (_project_end_date.date() - _project_start_date.date()).days + 1) if pd.notna(_project_start_date) and pd.notna(_project_end_date) else "—"

    st.markdown(
        clean_html(
            f"""
            <div style="background:{DARK_TEAL_HEADER};color:white;padding:16px 22px;border-radius:10px;margin-bottom:14px;display:flex;align-items:center;gap:12px">
                <div style="font-size:26px">📄</div>
                <div>
                    <div style="font-size:19px;font-weight:800;color:white">View Project</div>
                    <div style="font-size:12px;color:#CBD5E1">Project details and overall progress at a glance.</div>
                </div>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;margin-bottom:4px">
                <div style="display:flex;align-items:center;gap:10px">
                    <span style="font-size:21px;font-weight:800;color:#0F172A">{html.escape(r['name'])}</span>
                    {badge_html(r.status)}
                </div>
                <div>{tags_html}</div>
            </div>
            <div style="font-size:13px;color:#475569;margin-bottom:14px">{html.escape(full_desc)}</div>
            <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px">
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:white">
                    <div style="width:36px;height:36px;border-radius:50%;background:#DCFCE7;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:16px">👤</div>
                    <div><div style="font-size:11px;color:#64748B">Project Manager</div><div style="font-size:13px;font-weight:700;color:#0F172A">{html.escape(str(r.owner))}</div></div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:white">
                    <div style="width:36px;height:36px;border-radius:50%;background:#DCFCE7;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:16px">🌐</div>
                    <div><div style="font-size:11px;color:#64748B">Project Category</div><div style="font-size:13px;font-weight:700;color:#0F172A">{html.escape(str(r.category or '—'))}</div></div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:white">
                    <div style="width:36px;height:36px;border-radius:50%;background:#DCFCE7;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:16px">📅</div>
                    <div><div style="font-size:11px;color:#64748B">Start Date</div><div style="font-size:13px;font-weight:700;color:#0F172A">{fmt_date(r.start_date)}</div></div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:white">
                    <div style="width:36px;height:36px;border-radius:50%;background:#DCFCE7;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:16px">📅</div>
                    <div><div style="font-size:11px;color:#64748B">Target End Date</div><div style="font-size:13px;font-weight:700;color:#0F172A">{fmt_date(r.end_date)}</div></div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;display:flex;align-items:center;gap:10px;background:white">
                    <div style="width:36px;height:36px;border-radius:50%;background:#DCFCE7;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:16px">🕒</div>
                    <div><div style="font-size:11px;color:#64748B">Days Left</div><div style="font-size:18px;font-weight:800;color:#00875A;line-height:1">{int(r.days_left) if pd.notna(r.days_left) else 0}</div><div style="font-size:10px;color:#64748B">Days remaining</div></div>
                </div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    c_left, c_right = st.columns([1.05, 1.95])
    with c_left, st.container(border=True):
        section_title("📄 Project Progress")
        d1, d2 = st.columns([1, 1.15])
        with d1:
            st.plotly_chart(render_project_progress_donut(pct, 165), use_container_width=True, config={"displayModeBar": False})
        with d2:
            st.markdown(project_task_status_html(pid), unsafe_allow_html=True)
    with c_right, st.container(border=True):
        section_title("⏳ Timeline", "View Full Timeline →")
        st.markdown(render_timeline_stepper_html(pid), unsafe_allow_html=True)

    b_left, b_right = st.columns([1.05, 1.95])
    with b_left, st.container(border=True):
        section_title("📄 Project Information")
        st.markdown(
            clean_html(
                f"""
                <div style="font-size:12.5px;line-height:1.95">
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Project Name</span><span style="font-weight:600">{html.escape(r['name'])}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px;line-height:1.4;margin:4px 0"><span style="color:#64748B">Description</span><span>{html.escape(full_desc)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Category</span><span>{html.escape(str(r.category or '—'))}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Status</span><span>{badge_html(r.status)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Priority</span><span>{badge_html(r.priority or '—')}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Project Manager</span><span>{html.escape(str(r.owner))}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Start Date</span><span>{fmt_date(r.start_date)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Target End Date</span><span>{fmt_date(r.end_date)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Duration</span><span>{_project_duration} days</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Days Left</span><span style="color:#00875A;font-weight:700">{int(r.days_left) if pd.notna(r.days_left) else 0} days</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px;margin-top:4px"><span style="color:#64748B">Tags</span><span>{tags_html}</span></div>
                </div>
                <hr style="border-color:#EEF2F6;margin:12px 0">
                <div style="font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:6px">📄 Project Summary</div>
                <div style="font-size:11.5px;color:#64748B;line-height:1.55">
                    {html.escape(full_desc or str(r['name']))}. Current progress is {pct}%. The project timeline and milestones shown here reflect the records currently loaded in the workspace.
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    with b_right:
        with st.container(border=True):
            section_title("🎯 Key Milestones", "View All Milestones →")
            ms_rows = [(int(ms_row.id), "#00875A" if str(ms_row.status).lower() == "completed" else ("#1473E6" if str(ms_row.status).lower() == "in progress" else "#64748B"), str(ms_row.name), fmt_date(ms_row.due_date), str(ms_row.status or "Not Started")) for _, ms_row in M[M.project_id == pid].iterrows()] if not M.empty else []
            tr_html = "".join(
                f"<tr><td><b>{idx}</b></td><td><span style='color:{dot};margin-right:6px'>●</span>{mname}</td><td>{mdue}</td><td>{badge_html(mstat)}</td></tr>"
                for idx, dot, mname, mdue, mstat in ms_rows
            )
            st.markdown(
                f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>#</th><th>Milestone</th><th>Due Date</th><th>Status</th></tr></thead><tbody>{tr_html}</tbody></table></div>',
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            section_title("🕒 Recent Activity", "View All Activity →")
            act_rows = []
            _project_activity = df("SELECT event,item_type,details,created_at FROM audit_log ORDER BY created_at DESC LIMIT 5")
            for _, _activity_row in _project_activity.iterrows():
                _event = str(_activity_row.get("event") or "Updated")
                _type = str(_activity_row.get("item_type") or "Record")
                _details = str(_activity_row.get("details") or "")
                _dot = "#00875A" if _event.lower() in {"created", "completed"} else "#1473E6"
                act_rows.append((str(_activity_row.get("created_at") or ""), _dot, f"{_type} {_event.lower()}", _details))
            act_tr = "".join(
                f"<tr><td>{dt}</td><td><span style='color:{dot};margin-right:6px'>●</span>{act}</td><td>{by}</td></tr>"
                for dt, dot, act, by in act_rows
            )
            st.markdown(
                f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Date & Time</th><th>Activity</th><th>By</th></tr></thead><tbody>{act_tr}</tbody></table></div>',
                unsafe_allow_html=True,
            )

    f1, f2, f3 = st.columns([4, 1, 1.2])
    with f2:
        if st.button("Close", use_container_width=True, key="close_view_modal"):
            st.session_state.pop("active_modal", None)
            st.rerun()
    with f3:
        if st.button("✎ Edit Project", type="primary", use_container_width=True, key="switch_to_edit_modal"):
            st.session_state.active_modal = ("edit", pid)
            st.rerun()


@st.dialog("Edit Project", width="large")
def edit_project_modal(pid: int | None = None) -> None:
    is_new = pid is None
    if not is_new:
        row_df = df("SELECT * FROM projects WHERE id=?", (pid,))
        r = row_df.iloc[0] if not row_df.empty else None
    else:
        r = None

    title_txt = "New Project" if is_new else "Edit Project"
    sub_txt = "Create a new project, schedule, and milestones." if is_new else "Update project details, schedule, and milestones."

    st.markdown(
        clean_html(
            f"""
            <div style="background:{DARK_TEAL_HEADER};color:white;padding:16px 22px;border-radius:10px;margin-bottom:14px;display:flex;align-items:center;gap:12px">
                <div style="font-size:26px">📝</div>
                <div>
                    <div style="font-size:19px;font-weight:800;color:white">{title_txt}</div>
                    <div style="font-size:12px;color:#CBD5E1">{sub_txt}</div>
                </div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    top_l, top_r = st.columns([2.1, 1])
    with top_l, st.container(border=True):
        section_title("📄 Project Details")
        f_a, f_b = st.columns(2)
        with f_a:
            p_name = st.text_input("Project Name *", value=str(r["name"]) if r is not None else "")
        with f_b:
            default_desc = str(r.description or "") if r is not None else ""
            p_desc = st.text_area("Description", value=default_desc, height=85)
            st.markdown(f"<div style='text-align:right;font-size:11px;color:#64748B;margin-top:-8px'>{len(p_desc)}/500</div>", unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        cats = ["Web Development", "IT & Development", "Mobile Development", "Cloud & Data", "Infrastructure", "Security & Compliance", "AI & Innovation", "Operations", "HR & Training", "Procurement", "Sustainability"]
        cur_cat = str(r.category) if r is not None and r.category in cats else "Web Development"
        with c1:
            p_cat = st.selectbox("Project Category *", cats, index=cats.index(cur_cat))
        stats = ["On Track", "In Progress", "At Risk", "Not Started", "Delayed"]
        cur_st = str(r.status) if r is not None and r.status in stats else "On Track"
        with c2:
            p_stat = st.selectbox("Status *", stats, index=stats.index(cur_st if r is not None else get_setting("default_project_status", "Not Started")) if (cur_st if r is not None else get_setting("default_project_status", "Not Started")) in stats else 0)
        prios = ["High", "Medium", "Low"]
        cur_pr = str(r.priority) if r is not None and r.priority in prios else "High"
        with c3:
            p_prio = st.selectbox("Priority *", prios, index=prios.index(cur_pr))

        c4, c5, c6 = st.columns(3)
        mgrs = R.name.tolist() if not R.empty else ["Unassigned"]
        cur_mgr = str(r.owner) if r is not None and r.owner in mgrs else mgrs[0]
        with c4:
            p_mgr = st.selectbox("Project Manager *", mgrs, index=mgrs.index(cur_mgr))
        with c5:
            s_dt = pd.to_datetime(r.start_date, errors="coerce").date() if r is not None and pd.notna(pd.to_datetime(r.start_date, errors="coerce")) else date.today()
            p_start = st.date_input("Start Date *", s_dt)
        with c6:
            e_dt = pd.to_datetime(r.end_date, errors="coerce").date() if r is not None and pd.notna(pd.to_datetime(r.end_date, errors="coerce")) else date.today() + timedelta(days=30)
            p_end = st.date_input("Target End Date *", e_dt)

        all_tags_pool = ["Web", "Marketing", "Corporate", "IT", "Development", "Mobile", "Customer", "Data", "Cloud", "Infrastructure", "Security", "Compliance", "AI", "Innovation", "Automation", "Operations", "HR", "Training", "Procurement", "Sustainability", "ESG"]
        cur_tags = [t.strip() for t in str(r.tags or "").split(",") if t.strip() in all_tags_pool]
        p_tags = st.multiselect("Tags", all_tags_pool, default=cur_tags)

    with top_r, st.container(border=True):
        section_title("📄 Project Progress (auto-calculated)")
        pct_val = int(r.progress or 0) if r is not None else 0
        st.plotly_chart(render_project_progress_donut(pct_val, 185), use_container_width=True, config={"displayModeBar": False})
        st.markdown(project_task_status_html(pid), unsafe_allow_html=True)

    with st.container(border=True):
        section_title("📋 Project Timeline & Milestones")
        st.markdown(render_timeline_stepper_html(pid), unsafe_allow_html=True)
        ms_edit_rows = [
            (int(ms_row.id), "#00875A" if str(ms_row.status).lower() == "completed" else ("#1473E6" if str(ms_row.status).lower() == "in progress" else "#64748B"), str(ms_row.name), fmt_date(ms_row.due_date), str(ms_row.status or "Not Started"))
            for _, ms_row in M[M.project_id == pid].iterrows()
        ] if not M.empty else []
        rows_mk = "".join(
            f"<tr><td>{i}</td><td><span style='color:{dot};margin-right:6px'>●</span>{html.escape(nm)}</td><td>{dt}</td><td>{badge_html(st_val)}</td><td>Manage below</td></tr>"
            for i, dot, nm, dt, st_val in ms_edit_rows
        )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>#</th><th>Milestone</th><th>Due Date</th><th>Status</th><th>Actions</th></tr></thead><tbody>{rows_mk}</tbody></table></div>',
            unsafe_allow_html=True,
        )
        if ms_edit_rows:
            st.caption("Milestone actions")
            for _ms_id, _ms_dot, _ms_name, _ms_date, _ms_status in ms_edit_rows:
                _mc1, _mc2, _mc3 = st.columns([3.2, 0.8, 0.8])
                with _mc1:
                    st.write(f"{_ms_name} · {_ms_date} · {_ms_status}")
                with _mc2:
                    if st.button("Edit", key=f"edit_project_milestone_{pid}_{_ms_id}", use_container_width=True):
                        st.session_state["project_milestone_edit"] = int(_ms_id)
                        st.rerun()
                with _mc3:
                    if st.button("Delete", key=f"delete_project_milestone_{pid}_{_ms_id}", use_container_width=True):
                        sql("DELETE FROM milestones WHERE id=? AND project_id=?", (int(_ms_id), int(pid)))
                        log_event("Deleted", "Milestone", int(_ms_id), _ms_name)
                        st.rerun()
        _milestone_edit_id = st.session_state.get("project_milestone_edit")
        if _milestone_edit_id is not None:
            _ms_form = df("SELECT * FROM milestones WHERE id=? AND project_id=?", (int(_milestone_edit_id), int(pid)))
            if not _ms_form.empty:
                _ms_record = _ms_form.iloc[0]
                _ms_name_in = st.text_input("Milestone name", value=str(_ms_record.get("name") or ""), key=f"project_ms_name_{_milestone_edit_id}")
                _ms_status_opts = ["Not Started", "In Progress", "Completed", "At Risk", "On Track", "Delayed"]
                _ms_status_current = str(_ms_record.get("status") or "Not Started")
                _ms_status_in = st.selectbox("Milestone status", _ms_status_opts, index=_ms_status_opts.index(_ms_status_current) if _ms_status_current in _ms_status_opts else 0, key=f"project_ms_status_{_milestone_edit_id}")
                _ms_due_dt = pd.to_datetime(_ms_record.get("due_date"), errors="coerce")
                _ms_due_in = st.date_input("Milestone due date", value=_ms_due_dt.date() if pd.notna(_ms_due_dt) else date.today(), key=f"project_ms_due_{_milestone_edit_id}")
                _ms_save_col, _ms_cancel_col = st.columns(2)
                with _ms_save_col:
                    if st.button("Save Milestone", type="primary", key=f"project_ms_save_{_milestone_edit_id}", use_container_width=True):
                        if not _ms_name_in.strip():
                            st.error("Enter a milestone name.")
                        else:
                            sql("UPDATE milestones SET name=?,status=?,due_date=? WHERE id=? AND project_id=?", (_ms_name_in.strip(), _ms_status_in, _ms_due_in.isoformat(), int(_milestone_edit_id), int(pid)))
                            log_event("Updated", "Milestone", int(_milestone_edit_id), _ms_name_in.strip())
                            st.session_state.pop("project_milestone_edit", None)
                            st.rerun()
                with _ms_cancel_col:
                    if st.button("Cancel milestone edit", key=f"project_ms_cancel_{_milestone_edit_id}", use_container_width=True):
                        st.session_state.pop("project_milestone_edit", None)
                        st.rerun()

    b1, b2, b3 = st.columns([4, 1, 1.3])
    with b2:
        if st.button("Cancel", use_container_width=True, key="cancel_edit_modal"):
            st.session_state.pop("active_modal", None)
            st.rerun()
    with b3:
        if st.button("💾 Save Changes", type="primary", use_container_width=True, key="save_edit_modal"):
            if p_end < p_start:
                st.error("Project target end date must be on or after its start date.")
                return
            tags_str = ",".join(p_tags)
            days_l = (p_end - date.today()).days
            health = "Critical" if p_stat == "Delayed" else ("At Risk" if p_stat == "At Risk" else "Healthy")
            if is_new:
                sql(
                    """
                    INSERT INTO projects(
                        name,owner,team,category,status,priority,start_date,end_date,
                        progress,days_left,tasks_done,tasks_total,tags,bar_color,short_desc,description,health
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        p_name.strip(), p_mgr, "IT", p_cat, p_stat, p_prio,
                        p_start.isoformat(), p_end.isoformat(), pct_val, days_l, 0, 0,
                        tags_str, GREEN, p_desc[:40], p_desc, health,
                    ),
                )
                log_event("Created", "Project", None, p_name.strip())
            else:
                sql(
                    """
                    UPDATE projects
                    SET name=?,owner=?,category=?,status=?,priority=?,start_date=?,end_date=?,tags=?,description=?,days_left=?,health=?
                    WHERE id=?
                    """,
                    (
                        p_name.strip(), p_mgr, p_cat, p_stat, p_prio,
                        p_start.isoformat(), p_end.isoformat(), tags_str, p_desc, days_l, health, pid,
                    ),
                )
                log_event("Updated", "Project", pid, p_name.strip())
            st.session_state.pop("active_modal", None)
            st.rerun()


@st.dialog("Delete Project")
def delete_project_modal(pid: int) -> None:
    row_df = df("SELECT * FROM projects WHERE id=?", (pid,))
    if row_df.empty:
        st.session_state.pop("active_modal", None)
        st.rerun()
    r = row_df.iloc[0]
    tags_html = "".join(tag_pill_html(t) for t in str(r.tags or "").split(",") if t.strip())

    st.markdown(
        clean_html(
            f"""
            <div style="text-align:center;padding:8px 12px 14px 12px">
                <div style="width:64px;height:64px;border-radius:50%;background:#FEE2E2;color:#DC2626;display:inline-flex;align-items:center;justify-content:center;font-size:30px;margin-bottom:10px">🗑</div>
                <div style="font-size:22px;font-weight:800;color:#0F172A;margin-bottom:6px">Delete Project</div>
                <div style="font-size:13.5px;font-weight:600;color:#334155;margin-bottom:6px">Are you sure you want to delete this project?</div>
                <div style="font-size:12.5px;color:#64748B;line-height:1.5;max-width:440px;margin:0 auto 16px auto">
                    This action cannot be undone. All project information, including tasks, milestones, and associated data, will be permanently removed.
                </div>
            </div>
            <div style="border:1px solid #E2E8F0;border-left:5px solid {r.bar_color or GREEN};border-radius:10px;padding:14px 16px;background:#F8FAFC;margin-bottom:18px">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
                    <span style="font-size:15.5px;font-weight:800;color:#0F172A">{html.escape(r['name'])}</span>
                    {badge_html(r.status)}
                </div>
                <div style="font-size:12.5px;color:#475569;margin-bottom:10px">{html.escape(str(r.description or ''))}</div>
                <div style="margin-bottom:10px">{tags_html}</div>
                <div style="font-size:12px;color:#334155;margin-bottom:6px">📅 &nbsp; {fmt_date(r.start_date)} – {fmt_date(r.end_date)}</div>
                <div style="font-size:12px;color:#334155">👤 &nbsp; {avatar_html(r.owner)}</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1.2, 1, 1.3])
    with c2:
        if st.button("Cancel", use_container_width=True, key="cancel_del_modal"):
            st.session_state.pop("active_modal", None)
            st.rerun()
    with c3:
        if st.button("🗑 Delete Project", type="primary", use_container_width=True, key="confirm_del_modal"):
            sql("DELETE FROM tasks WHERE project_id=?", (pid,))
            sql("DELETE FROM milestones WHERE project_id=?", (pid,))
            sql("DELETE FROM project_assignments WHERE project_id=?", (pid,))
            sql("DELETE FROM projects WHERE id=?", (pid,))
            log_event("Deleted", "Project", pid, r["name"])
            st.session_state.pop("active_modal", None)
            st.rerun()

# ===========================================================================
# PAGE 1: PROJECT OVERVIEW (Exact Replica of Image 2)
# ===========================================================================
if page == "Overview":
    # Header Row: Title on Left, All Projects + Date Range on Right
    h_left, h_r1, h_r2 = st.columns([2.3, 0.9, 1.1])
    with h_left:
        st.markdown(
            f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Project Overview</div>"
            f"<div style='font-size:13.5px;color:#64748B;margin-top:2px'>Track progress, monitor tasks, and keep your projects on schedule.</div>",
            unsafe_allow_html=True,
        )
    with h_r1:
        ov_proj_filter = st.selectbox("Project", ["All Projects"] + projects_list, key="ov_proj_sel", label_visibility="collapsed")
    with h_r2:
        st.date_input("Date Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="ov_date_sel", label_visibility="collapsed")

    st.write("")
    # KPI values are calculated from the current Excel-fed records.
    _ov_tasks = T.copy()
    _ov_status = _ov_tasks.status.fillna("Not Started").astype(str) if not _ov_tasks.empty else pd.Series(dtype=str)
    _ov_total_tasks = len(_ov_tasks)
    _ov_completed = int((_ov_status.str.lower() == "completed").sum())
    _ov_in_progress = int((_ov_status.str.lower() == "in progress").sum())
    _ov_on_hold = int((_ov_status.str.lower() == "on hold").sum())
    _ov_due = pd.to_datetime(_ov_tasks.due_date, errors="coerce") if not _ov_tasks.empty else pd.Series(dtype="datetime64[ns]")
    _ov_overdue = int(((_ov_due.dt.date < date.today()) & (_ov_status.str.lower() != "completed")).sum()) if not _ov_tasks.empty else 0
    def _ov_pct(n): return f"({round(n / _ov_total_tasks * 100)}%)" if _ov_total_tasks else "(0%)"
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: kpi_card_mockup("Total Projects", str(len(P)), sub_text="Current workspace", icon_html="❖", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k2: kpi_card_mockup("Total Tasks", str(_ov_total_tasks), sub_text="Across All Projects", icon_html="✓", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k3: kpi_card_mockup("Completed Tasks", str(_ov_completed), _ov_pct(_ov_completed), bar_pct=round(_ov_completed/_ov_total_tasks*100) if _ov_total_tasks else 0, icon_html="◕", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k4: kpi_card_mockup("In Progress", str(_ov_in_progress), _ov_pct(_ov_in_progress), bar_pct=round(_ov_in_progress/_ov_total_tasks*100) if _ov_total_tasks else 0, icon_html="▶", icon_bg="#DBEAFE", icon_fg="#1473E6", bar_color="#1473E6")
    with k5: kpi_card_mockup("On Hold", str(_ov_on_hold), _ov_pct(_ov_on_hold), bar_pct=round(_ov_on_hold/_ov_total_tasks*100) if _ov_total_tasks else 0, icon_html="❚❚", icon_bg="#FEF3C7", icon_fg="#F28C00", bar_color="#F28C00")
    with k6: kpi_card_mockup("Overdue", str(_ov_overdue), _ov_pct(_ov_overdue), bar_pct=round(_ov_overdue/_ov_total_tasks*100) if _ov_total_tasks else 0, icon_html="✕", icon_bg="#FEE2E2", icon_fg="#EF4444", bar_color="#EF4444")

    st.write("")
    # Row 2: Project Progress (Left) & Task Status Donut (Right)
    r2_l, r2_r = st.columns([1.55, 1])
    with r2_l, st.container(border=True):
        th1, th2 = st.columns([3, 1])
        with th1:
            st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>▤ &nbsp; Project Progress</div>", unsafe_allow_html=True)
        with th2:
            if st.button("View All Projects →", key="ov_to_proj", use_container_width=True):
                st.session_state.page = "Projects"
                st.rerun()

        ov_projs = P.copy()
        if ov_proj_filter != "All Projects":
            ov_projs = P[P.name == ov_proj_filter].copy()

        rows_html = ""
        for _, pr in ov_projs.iterrows():
            bcol = pr.bar_color or GREEN
            rows_html += (
                f"<tr>"
                f"<td><div style='display:flex;align-items:center;gap:10px'>"
                f"<span style='width:4px;height:22px;border-radius:3px;background:{bcol};display:inline-block'></span>"
                f"<span style='font-weight:600;color:#0F172A'>{html.escape(pr['name'])}</span></div></td>"
                f"<td>{fmt_date(pr.start_date)}</td>"
                f"<td>{fmt_date(pr.end_date)}</td>"
                f"<td>{progress_html(pr.progress, GREEN)}</td>"
                f"<td>{badge_html(pr.status)}</td>"
                f"<td style='text-align:center;font-weight:600'>{int(pr.days_left or 0)}</td>"
                f"</tr>"
            )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table">'
            f"<thead><tr><th>Project Name</th><th>Start Date</th><th>End Date</th><th>Progress</th><th>Status</th><th style='text-align:center'>Days Left</th></tr></thead>"
            f"<tbody>{rows_html}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    with r2_r, st.container(border=True):
        th1, th2 = st.columns([2.4, 1.1])
        with th1:
            st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>📅 &nbsp; Task Status</div>", unsafe_allow_html=True)
        with th2:
            if st.button("View Task List →", key="ov_to_tasks1", use_container_width=True):
                st.session_state.page = "Task List"
                st.rerun()

        d_col, l_col = st.columns([1.25, 1])
        with d_col:
            ts_df = pd.DataFrame({"Status": ["Completed", "In Progress", "On Hold", "Overdue"], "Count": [_ov_completed, _ov_in_progress, _ov_on_hold, _ov_overdue]})
            fig_ts = px.pie(
                ts_df,
                names="Status",
                values="Count",
                hole=0.60,
                color="Status",
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "On Hold": ORANGE, "Overdue": RED},
            )
            fig_ts.update_traces(textposition="inside", textinfo="percent", sort=False)
            fig_ts.update_layout(
                height=230,
                margin=dict(l=0, r=0, t=4, b=4),
                showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)",
                annotations=[
                    dict(
                        text=f"<b style='font-size:20px;color:#0F172A'>{_ov_total_tasks}</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_ts, use_container_width=True, config={"displayModeBar": False})
        with l_col:
            st.markdown(
                f"""
                <div style="font-size:12.5px;padding-top:34px;line-height:2.35">
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#00875A">●</span> &nbsp;Completed</span><b>{_ov_completed}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#1473E6">●</span> &nbsp;In Progress</span><b>{_ov_in_progress}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;On Hold</span><b>{_ov_on_hold}</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> &nbsp;Overdue</span><b>{_ov_overdue}</b></div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # Row 3: Upcoming Milestones (Left) & Task Priority Donut (Right)
    r3_l, r3_r = st.columns([1.55, 1])
    with r3_l, st.container(border=True):
        section_title("📅  Upcoming Milestones", "View All →")
        ms_ov = []
        if not M.empty:
            _upcoming = M.copy()
            _upcoming["_due"] = pd.to_datetime(_upcoming.due_date, errors="coerce")
            _upcoming = _upcoming[_upcoming["_due"].notna()].sort_values("_due").head(5)
            for _, _m in _upcoming.iterrows():
                _days_left = (_m["_due"].date() - date.today()).days
                _dot = "#00875A" if str(_m.status).lower() == "completed" else ("#F28C00" if _days_left < 0 else "#1473E6")
                ms_ov.append((_dot, str(_m.name), str(_m.project or ""), fmt_date(_m.due_date), _days_left, str(_m.status or "Not Started")))
        ms_html = "".join(
            f"<tr><td><span style='color:{dot};margin-right:8px'>●</span><span style='font-weight:500'>{mname}</span></td>"
            f"<td>{mproj}</td><td>{mdue}</td><td style='text-align:center'>{mdays}</td><td>{badge_html(mstat)}</td></tr>"
            for dot, mname, mproj, mdue, mdays, mstat in ms_ov
        )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table">'
            f"<thead><tr><th>Milestone</th><th>Project</th><th>Due Date</th><th style='text-align:center'>Days Left</th><th>Status</th></tr></thead>"
            f"<tbody>{ms_html}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    with r3_r, st.container(border=True):
        th1, th2 = st.columns([2.4, 1.1])
        with th1:
            st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>📅 &nbsp; Task Priority</div>", unsafe_allow_html=True)
        with th2:
            if st.button("View Task List →", key="ov_to_tasks2", use_container_width=True):
                st.session_state.page = "Task List"
                st.rerun()

        d_col, l_col = st.columns([1.25, 1])
        with d_col:
            _prio_series = T.priority.fillna("None").astype(str) if not T.empty else pd.Series(dtype=str)
            _prio_counts = {k: int((_prio_series.str.lower() == k.lower()).sum()) for k in ["High", "Medium", "Low", "None"]}
            tp_df = pd.DataFrame({"Priority": list(_prio_counts), "Count": list(_prio_counts.values())})
            fig_tp = px.pie(
                tp_df,
                names="Priority",
                values="Count",
                hole=0.60,
                color="Priority",
                color_discrete_map={"High": RED, "Medium": ORANGE, "Low": BLUE, "None": GREEN},
            )
            fig_tp.update_traces(textposition="inside", textinfo="percent", sort=False)
            fig_tp.update_layout(
                height=230,
                margin=dict(l=0, r=0, t=4, b=4),
                showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)",
                annotations=[
                    dict(
                        text=f"<b style='font-size:20px;color:#0F172A'>{_ov_total_tasks}</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_tp, use_container_width=True, config={"displayModeBar": False})
        with l_col:
            st.markdown(
                f"""
                <div style="font-size:12.5px;padding-top:34px;line-height:2.35">
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#EF4444">●</span> &nbsp;High</span><b>{_prio_counts['High']}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;Medium</span><b>{_prio_counts['Medium']}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#1473E6">●</span> &nbsp;Low</span><b>{_prio_counts['Low']}</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;None</span><b>{_prio_counts['None']}</b></div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # Row 4: Team Workload (Left) & Recent Activity (Right)
    r4_l, r4_r = st.columns([1.15, 1])
    with r4_l, st.container(border=True):
        th1, th2 = st.columns([2.6, 1.1])
        with th1:
            st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A'>👥 &nbsp; Team Workload</div>", unsafe_allow_html=True)
        with th2:
            if st.button("View Resources →", key="ov_to_res", use_container_width=True):
                st.session_state.page = "Resources"
                st.rerun()

        tw_rows = []
        for _, _resource in R.head(8).iterrows():
            _nm = str(_resource.get("name") or "")
            _mine = T[T.assignee.fillna("").astype(str) == _nm] if not T.empty else pd.DataFrame()
            _mine_status = _mine.status.fillna("").astype(str).str.lower() if not _mine.empty else pd.Series(dtype=str)
            _mine_due = pd.to_datetime(_mine.due_date, errors="coerce") if not _mine.empty else pd.Series(dtype="datetime64[ns]")
            _mine_over = int(((_mine_due.dt.date < date.today()) & (_mine_status != "completed")).sum()) if not _mine.empty else 0
            _util = max(0, min(100, 100 - int(_resource.get("capacity") or 0)))
            tw_rows.append((_nm, len(_mine), int((_mine_status == "in progress").sum()), int((_mine_status == "completed").sum()), _mine_over, _util))
        tw_html = "".join(
            f"<tr><td>{avatar_html(nm)}</td>"
            f"<td style='text-align:center'>{ass}</td>"
            f"<td style='text-align:center'>{inp}</td>"
            f"<td style='text-align:center'>{cmp}</td>"
            f"<td style='text-align:center;color:#EF4444;font-weight:700'>{ovd}</td>"
            f"<td>{progress_html(util, GREEN)}</td></tr>"
            for nm, ass, inp, cmp, ovd, util in tw_rows
        )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table">'
            f"<thead><tr><th>Team Member</th><th style='text-align:center'>Assigned Tasks</th><th style='text-align:center'>In Progress</th><th style='text-align:center'>Completed</th><th style='text-align:center'>Overdue</th><th>Utilization</th></tr></thead>"
            f"<tbody>{tw_html}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    with r4_r, st.container(border=True):
        section_title("🕒  Recent Activity", "View All →")
        ra_rows = []
        _activity = df("SELECT event,item_type,details,created_at FROM audit_log ORDER BY created_at DESC LIMIT 5")
        for _, _a in _activity.iterrows():
            _event_name = str(_a.get("event") or "Updated")
            _type_name = str(_a.get("item_type") or "Record")
            _details = str(_a.get("details") or "")
            _dot = "#00875A" if _event_name.lower() in {"created", "completed"} else "#1473E6"
            ra_rows.append((str(_a.get("created_at") or ""), _dot, f"{_type_name} {_event_name.lower()}", _details, ""))
        ra_html = "".join(
            f"<tr><td style='white-space:nowrap'>{dt}</td>"
            f"<td><span style='color:{dot};margin-right:6px'>●</span>{act}</td>"
            f"<td>{prj}</td><td>{tsk}</td></tr>"
            for dt, dot, act, prj, tsk in ra_rows
        )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table">'
            f"<thead><tr><th>Date & Time</th><th>Activity</th><th>Project</th><th>Task</th></tr></thead>"
            f"<tbody>{ra_html}</tbody></table></div>",
            unsafe_allow_html=True,
        )


# ===========================================================================
# PAGE 2: PROJECTS (Exact Replica of Images 1, 3, 4, 5, 6)
# ===========================================================================
elif page == "Projects":
    # Top Header Row: Title on Left, Search + Date Range + "+ New Project" on Right
    h_l, h_s, h_d, h_b = st.columns([1.9, 1.0, 1.0, 0.65])
    with h_l:
        st.markdown(
            f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Projects</div>"
            f"<div style='font-size:13.5px;color:#64748B;margin-top:2px'>Manage and track all your projects in one place.</div>",
            unsafe_allow_html=True,
        )
    with h_s:
        proj_q = st.text_input("Search", placeholder="🔍 Search projects...", key="p_search_box", label_visibility="collapsed")
    with h_d:
        st.date_input("Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="p_date_box", label_visibility="collapsed")
    with h_b:
        if st.button("＋ New Project", type="primary", use_container_width=True, key="btn_new_proj_top"):
            st.session_state.active_modal = ("edit", None)
            st.rerun()

    st.write("")
    # 5 KPI Cards Row matching Projects screenshots
    total_p = len(P)
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        kpi_card_mockup("TOTAL PROJECTS", str(total_p), sub_text="Current workspace", icon_html="❖", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    _p_status = P.status.fillna("Not Started").astype(str) if not P.empty else pd.Series(dtype=str)
    _p_counts = {name: int((_p_status.str.lower() == name.lower()).sum()) for name in ["On Track", "In Progress", "At Risk", "Delayed", "Completed", "Not Started"]}
    _overdue_projects = int(((pd.to_datetime(P.end_date, errors="coerce").dt.date < date.today()) & (_p_status.str.lower() != "completed")).sum()) if not P.empty else 0
    def _p_pct(n): return f"({round(n / total_p * 100)}%)" if total_p else "(0%)"
    with k2:
        kpi_card_mockup("ON TRACK", str(_p_counts["On Track"]), _p_pct(_p_counts["On Track"]), bar_pct=round(_p_counts["On Track"] / total_p * 100) if total_p else 0, icon_html="✓", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k3:
        kpi_card_mockup("IN PROGRESS", str(_p_counts["In Progress"]), _p_pct(_p_counts["In Progress"]), bar_pct=round(_p_counts["In Progress"] / total_p * 100) if total_p else 0, icon_html="▶", icon_bg="#DBEAFE", icon_fg="#1473E6", bar_color="#1473E6")
    with k4:
        kpi_card_mockup("AT RISK", str(_p_counts["At Risk"]), _p_pct(_p_counts["At Risk"]), bar_pct=round(_p_counts["At Risk"] / total_p * 100) if total_p else 0, icon_html="❚❚", icon_bg="#FEF3C7", icon_fg="#F28C00", bar_color="#F28C00")
    with k5:
        kpi_card_mockup("OVERDUE", str(_overdue_projects), _p_pct(_overdue_projects), bar_pct=round(_overdue_projects / total_p * 100) if total_p else 0, icon_html="✕", icon_bg="#FEE2E2", icon_fg="#EF4444", bar_color="#EF4444")

    st.write("")
    # Filter & View Mode Row (No radio buttons — uses tile buttons for List View / Card View)
    if "proj_view_mode" not in st.session_state:
        st.session_state.proj_view_mode = "Card View"

    all_managers = sorted(P.owner.dropna().unique().tolist())
    all_tags = sorted({t.strip() for raw in P.tags.dropna() for t in str(raw).split(",") if t.strip()})

    f1, f2, f3, f4, f_sp, f_clr, f_lst, f_crd = st.columns([0.95, 0.95, 0.95, 0.95, 0.7, 0.75, 0.68, 0.68])
    with f1:
        sel_stat = st.selectbox("Status", ["All Statuses", "On Track", "In Progress", "At Risk", "Not Started"], key="pf_stat", label_visibility="collapsed")
    with f2:
        sel_prio = st.selectbox("Priority", ["All Priorities", "High", "Medium", "Low"], key="pf_prio", label_visibility="collapsed")
    with f3:
        sel_mgr = st.selectbox("Manager", ["All Managers"] + all_managers, key="pf_mgr", label_visibility="collapsed")
    with f4:
        sel_tag = st.selectbox("Tag", ["All Tags"] + all_tags, key="pf_tag", label_visibility="collapsed")
    with f_clr:
        if st.button("▽ Clear Filters", use_container_width=True, key="btn_clr_pf"):
            for k, v in [("pf_stat", "All Statuses"), ("pf_prio", "All Priorities"), ("pf_mgr", "All Managers"), ("pf_tag", "All Tags"), ("p_search_box", "")]:
                st.session_state[k] = v
            st.rerun()
    with f_lst:
        if st.button(
            "☰ List View",
            type="primary" if st.session_state.proj_view_mode == "List View" else "secondary",
            use_container_width=True,
            key="toggle_list_view",
        ):
            st.session_state.proj_view_mode = "List View"
            st.rerun()
    with f_crd:
        if st.button(
            "⊞ Card View",
            type="primary" if st.session_state.proj_view_mode == "Card View" else "secondary",
            use_container_width=True,
            key="toggle_card_view",
        ):
            st.session_state.proj_view_mode = "Card View"
            st.rerun()

    # Apply Filters
    vp = P.copy()
    if proj_q:
        vp = vp[vp.astype(str).apply(lambda s: s.str.contains(proj_q, case=False, na=False)).any(axis=1)]
    if sel_stat != "All Statuses":
        vp = vp[vp.status == sel_stat]
    if sel_prio != "All Priorities":
        vp = vp[vp.priority == sel_prio]
    if sel_mgr != "All Managers":
        vp = vp[vp.owner == sel_mgr]
    if sel_tag != "All Tags":
        vp = vp[vp.tags.fillna("").str.contains(sel_tag, case=False)]

    # --- CARD VIEW (Image 1) ---
    if st.session_state.proj_view_mode == "Card View":
        sub_l, sub_r1, sub_r2 = st.columns([3.5, 0.35, 0.9])
        with sub_r1:
            st.markdown("<div style='font-size:12.5px;color:#475569;padding-top:8px;text-align:right'>Sort by:</div>", unsafe_allow_html=True)
        with sub_r2:
            sort_opt = st.selectbox(
                "Sort",
                ["Project Name (A - Z)", "Project Name (Z - A)", "Progress (High - Low)", "Days Left (Low - High)"],
                key="card_sort_opt",
                label_visibility="collapsed",
            )
        if sort_opt == "Project Name (A - Z)":
            vp = vp.sort_values("name", ascending=True)
        elif sort_opt == "Project Name (Z - A)":
            vp = vp.sort_values("name", ascending=False)
        elif sort_opt == "Progress (High - Low)":
            vp = vp.sort_values("progress", ascending=False)
        else:
            vp = vp.sort_values("days_left", ascending=True)

        with sub_l:
            st.markdown(
                f"<div style='font-size:13px;font-weight:500;color:#475569;padding-top:8px'>Showing {len(vp)} of {len(P)} projects</div>",
                unsafe_allow_html=True,
            )

        # Render 4-Column Card Grid
        records = vp.to_dict("records")
        for row_idx in range(0, len(records), 4):
            chunk = records[row_idx : row_idx + 4]
            cols = st.columns(4)
            for col, item in zip(cols, chunk):
                with col, st.container(border=True):
                    pid = int(item["id"])
                    pct = int(item.get("progress") or 0)
                    st_name = str(item.get("status") or "Not Started")
                    bar_col = GREEN if st_name == "On Track" else (ORANGE if st_name == "At Risk" else BLUE)
                    left_accent = str(item.get("bar_color") or GREEN)
                    tags_markup = "".join(tag_pill_html(t) for t in str(item.get("tags") or "").split(",") if t.strip())
                    tdone = int(item.get("tasks_done") or 0)
                    ttot = int(item.get("tasks_total") or 0)

                    st.markdown(
                        f"""
                        <div style="border-left:4px solid {left_accent};padding-left:10px;margin-bottom:8px">
                            <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:6px">
                                <span style="font-size:15px;font-weight:800;color:#0F172A;line-height:1.25">{html.escape(item['name'])}</span>
                                {badge_html(st_name)}
                            </div>
                            <div style="font-size:12px;color:#475569;margin-top:5px;min-height:34px;line-height:1.4">
                                {html.escape(str(item.get('description') or ''))}
                            </div>
                        </div>
                        <div style="margin-bottom:10px">{tags_markup}</div>
                        <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
                            <div style="flex:1;height:7px;background:#E2E8F0;border-radius:999px;overflow:hidden">
                                <div style="width:{pct}%;height:100%;background:{bar_col};border-radius:999px"></div>
                            </div>
                            <span style="font-size:12px;font-weight:700;color:#1E293B">{pct}%</span>
                        </div>
                        <div style="font-size:12px;color:#334155;margin-bottom:7px">
                            📅 &nbsp;{fmt_date(item.get('start_date'))} – {fmt_date(item.get('end_date'))}
                        </div>
                        <div style="font-size:12px;color:#334155;margin-bottom:8px">
                            👤 &nbsp;{avatar_html(item.get('owner'))}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    ft_l, ft_v, ft_d = st.columns([2.2, 0.65, 0.65])
                    with ft_l:
                        st.markdown(f"<div style='font-size:12px;color:#475569;padding-top:6px'>{tdone}/{ttot} tasks</div>", unsafe_allow_html=True)
                    with ft_v:
                        if st.button("👁", key=f"crd_view_{pid}", help="View Project", use_container_width=True):
                            st.session_state.active_modal = ("view", pid)
                            st.rerun()
                    with ft_d:
                        if st.button("🗑", key=f"crd_del_{pid}", help="Delete Project", use_container_width=True):
                            st.session_state.active_modal = ("delete", pid)
                            st.rerun()

    # --- LIST VIEW (Image 3) ---
    else:
        vp = vp.sort_values("id", ascending=True)
        with st.container(border=True):
            # Table Header
            h_cols = st.columns([0.45, 2.1, 1.35, 0.95, 0.95, 1.35, 0.9, 0.75, 0.65, 0.75, 1.4, 0.8])
            headers = ["#", "Project Name ⇅", "Project Manager", "Start Date", "Target End Date", "Progress", "Status", "Priority", "Days Left", "Tasks", "Tags", "Actions"]
            for c, h_txt in zip(h_cols, headers):
                c.markdown(f"<div style='font-size:11.5px;font-weight:700;color:#475569;padding-bottom:6px'>{h_txt}</div>", unsafe_allow_html=True)
            st.markdown("<hr style='margin:0 0 8px 0;border-color:#E2E8F0'>", unsafe_allow_html=True)

            for idx_num, (_, item) in enumerate(vp.iterrows(), start=1):
                pid = int(item["id"])
                pct = int(item.get("progress") or 0)
                st_name = str(item.get("status") or "Not Started")
                bar_col = GREEN if st_name == "On Track" else (ORANGE if st_name == "At Risk" else BLUE)
                left_accent = str(item.get("bar_color") or GREEN)
                tdone = int(item.get("tasks_done") or 0)
                ttot = int(item.get("tasks_total") or 0)
                task_pct = int(round((tdone / max(1, ttot)) * 100))
                # Show first 2 tags in List View to match screenshot
                tag_list = [t.strip() for t in str(item.get("tags") or "").split(",") if t.strip()][:2]
                tags_markup = "".join(tag_pill_html(t) for t in tag_list)

                r_cols = st.columns([0.45, 2.1, 1.35, 0.95, 0.95, 1.35, 0.9, 0.75, 0.65, 0.75, 1.4, 0.8])
                with r_cols[0]:
                    st.markdown(f"<div style='font-size:12.5px;font-weight:600;padding-top:6px'>☐ &nbsp;{idx_num}</div>", unsafe_allow_html=True)
                with r_cols[1]:
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:8px'>"
                        f"<span style='width:4px;height:30px;border-radius:3px;background:{left_accent};display:inline-block'></span>"
                        f"<div><div style='font-size:13px;font-weight:700;color:#0F172A'>{html.escape(item['name'])}</div>"
                        f"<div style='font-size:11px;color:#64748B'>{html.escape(str(item.get('short_desc') or ''))}</div></div></div>",
                        unsafe_allow_html=True,
                    )
                with r_cols[2]:
                    st.markdown(f"<div style='padding-top:4px;font-size:12.5px'>{avatar_html(item.get('owner'))}</div>", unsafe_allow_html=True)
                with r_cols[3]:
                    st.markdown(f"<div style='padding-top:6px;font-size:12.5px'>{fmt_date(item.get('start_date'))}</div>", unsafe_allow_html=True)
                with r_cols[4]:
                    st.markdown(f"<div style='padding-top:6px;font-size:12.5px'>{fmt_date(item.get('end_date'))}</div>", unsafe_allow_html=True)
                with r_cols[5]:
                    st.markdown(f"<div style='padding-top:6px'>{progress_html(pct, bar_col)}</div>", unsafe_allow_html=True)
                with r_cols[6]:
                    st.markdown(f"<div style='padding-top:4px'>{badge_html(st_name)}</div>", unsafe_allow_html=True)
                with r_cols[7]:
                    st.markdown(f"<div style='padding-top:4px'>{badge_html(item.get('priority') or 'Medium')}</div>", unsafe_allow_html=True)
                with r_cols[8]:
                    st.markdown(f"<div style='padding-top:6px;font-size:12.5px;font-weight:700;color:#00875A'>{int(item.get('days_left') or 0)}</div>", unsafe_allow_html=True)
                with r_cols[9]:
                    st.markdown(
                        f"<div style='padding-top:2px'><div style='font-size:11.5px;font-weight:600'>{tdone}/{ttot}</div>"
                        f"<div style='height:5px;background:#E2E8F0;border-radius:99px;overflow:hidden;margin-top:2px'>"
                        f"<div style='width:{task_pct}%;height:100%;background:#00875A'></div></div></div>",
                        unsafe_allow_html=True,
                    )
                with r_cols[10]:
                    st.markdown(f"<div style='padding-top:4px'>{tags_markup}</div>", unsafe_allow_html=True)
                with r_cols[11]:
                    ac1, ac2 = st.columns(2)
                    with ac1:
                        if st.button("👁", key=f"lst_view_{pid}", help="View Project"):
                            st.session_state.active_modal = ("view", pid)
                            st.rerun()
                    with ac2:
                        if st.button("🗑", key=f"lst_del_{pid}", help="Delete Project"):
                            st.session_state.active_modal = ("delete", pid)
                            st.rerun()
                st.markdown("<hr style='margin:6px 0;border-color:#F1F5F9'>", unsafe_allow_html=True)

            st.caption(f"Showing all {len(vp)} filtered projects")

    # Trigger active modal if set
    if "active_modal" in st.session_state:
        m_type, m_pid = st.session_state.active_modal
        if m_type == "view":
            view_project_modal(m_pid)
        elif m_type == "edit":
            edit_project_modal(m_pid)
        elif m_type == "delete":
            delete_project_modal(m_pid)


# ===========================================================================
# PAGE 3: GANTT CHART (Exact Replica of Month, Week, and Day Mockups)
# ===========================================================================
elif page == "Gantt Chart":
    # Top Header Row: Title on Left, Search + Date Range on Right
    gh_l, gh_s, gh_d = st.columns([2.2, 1.1, 1.0])
    with gh_l:
        st.markdown(
            clean_html(
                f"""
                <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Gantt Chart</div>
                <div style='font-size:13.5px;color:#64748B;margin-top:2px'>Visualize and manage project schedules, dependencies, and milestones.</div>
                """
            ),
            unsafe_allow_html=True,
        )
    with gh_s:
        gantt_q = st.text_input("Search", placeholder="🔍 Search projects or tasks...", key="gantt_search_top", label_visibility="collapsed")
    with gh_d:
        gantt_range = st.date_input("Date Range", (date.today().replace(day=1), date.today()+timedelta(days=90)), key="gantt_date_top", label_visibility="collapsed")

    st.write("")
    # Filter Bar Card
    with st.container(border=True):
        gf1, gf2, gf3, gf4, gf_sp, gf_more = st.columns([1.15, 1.15, 1.15, 1.15, 0.8, 0.7])
        gantt_projects = ["All Projects"] + projects_list
        with gf1:
            st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Project</div>", unsafe_allow_html=True)
            g_proj_sel = st.selectbox("Project", gantt_projects, key="g_flt_proj", label_visibility="collapsed")
        with gf2:
            st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>View</div>", unsafe_allow_html=True)
            g_view_sel = st.selectbox("View", ["All Tasks", "Tasks Only", "Milestones Only"], key="g_flt_view", label_visibility="collapsed")
        with gf3:
            st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Status</div>", unsafe_allow_html=True)
            g_stat_sel = st.selectbox("Status", ["All Statuses", "On Track", "Completed", "In Progress", "Not Started", "At Risk", "Overdue"], key="g_flt_stat", label_visibility="collapsed")
        with gf4:
            st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Assignee</div>", unsafe_allow_html=True)
            g_ass_sel = st.selectbox("Assignee", ["All Assignees"] + sorted(R.name.unique().tolist()), key="g_flt_ass", label_visibility="collapsed")
        with gf_more:
            st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
            if st.button("▽ More Filters", key="g_btn_more_flt", use_container_width=True):
                for k, v in [("g_flt_proj", "All Projects"), ("g_flt_view", "All Tasks"), ("g_flt_stat", "All Statuses"), ("g_flt_ass", "All Assignees")]:
                    st.session_state[k] = v
                st.rerun()

    # Legend + Tile Toggle Bar (Month | Week | Day)
    if "gantt_scale" not in st.session_state:
        st.session_state.gantt_scale = "Month"

    with st.container(border=True):
        lg_col, m_col, w_col, d_col, exp_col, dl_col = st.columns([4.5, 0.45, 0.45, 0.42, 0.25, 0.25])
        with lg_col:
            dep_marker = (
                "<span style='display:inline-block;width:11px;height:11px;border-radius:3px;background:#EF4444;margin-right:5px;vertical-align:middle'></span>Dependency"
                if st.session_state.gantt_scale == "Week"
                else "<span style='color:#64748B;font-weight:800;margin-right:5px'>⟶</span>Dependency"
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;padding-top:6px;font-size:12px;font-weight:500;color:#334155">
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#1473E6;margin-right:5px;vertical-align:middle"></span>Task</span>
                        <span><span style="color:#7C3AED;font-size:14px;margin-right:4px;vertical-align:middle">◆</span>Milestone</span>
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#00875A;margin-right:5px;vertical-align:middle"></span>Completed</span>
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#60A5FA;margin-right:5px;vertical-align:middle"></span>In Progress</span>
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#94A3B8;margin-right:5px;vertical-align:middle"></span>Not Started</span>
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#F59E0B;margin-right:5px;vertical-align:middle"></span>At Risk</span>
                        <span><span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:#EF4444;margin-right:5px;vertical-align:middle"></span>Overdue</span>
                        <span>{dep_marker}</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with m_col:
            if st.button("Month", key="g_scale_month", type="primary" if st.session_state.gantt_scale == "Month" else "secondary", use_container_width=True):
                st.session_state.gantt_scale = "Month"
                st.rerun()
        with w_col:
            if st.button("Week", key="g_scale_week", type="primary" if st.session_state.gantt_scale == "Week" else "secondary", use_container_width=True):
                st.session_state.gantt_scale = "Week"
                st.rerun()
        with d_col:
            if st.button("Day", key="g_scale_day", type="primary" if st.session_state.gantt_scale == "Day" else "secondary", use_container_width=True):
                st.session_state.gantt_scale = "Day"
                st.rerun()
        with exp_col:
            if st.button("⛶", key="g_fullscreen_btn", help="Toggle expanded Gantt view", use_container_width=True):
                st.session_state["gantt_expanded"] = not st.session_state.get("gantt_expanded", False)
                st.rerun()
        with dl_col:
            st.download_button(
                "⬇",
                T.to_csv(index=False).encode("utf-8-sig"),
                file_name="vektor_gantt_schedule.csv",
                mime="text/csv",
                key="g_dl_btn",
                use_container_width=True,
            )

        scale = st.session_state.gantt_scale
        _gantt_expanded = bool(st.session_state.get("gantt_expanded", False))
        if _gantt_expanded:
            st.caption("Expanded Gantt view · use the same button to return to compact view.")

        # Build Gantt rows from imported/current project, task, and milestone records.
        _range_start = gantt_range[0] if isinstance(gantt_range, (tuple, list)) and len(gantt_range) == 2 else date.today().replace(day=1)
        _range_end = gantt_range[1] if isinstance(gantt_range, (tuple, list)) and len(gantt_range) == 2 else (_range_start + timedelta(days=90))
        _range_days = max(1, (_range_end - _range_start).days + 1)
        def _bar_position(start_value, end_value):
            _s = pd.to_datetime(start_value, errors="coerce")
            _e = pd.to_datetime(end_value, errors="coerce")
            if pd.isna(_s) or pd.isna(_e): return 0.0, 0.0
            _left = max(0.0, min(100.0, ((_s.date() - _range_start).days / _range_days) * 100))
            _right = max(0.0, min(100.0, (((_e.date() - _range_start).days + 1) / _range_days) * 100))
            return _left, max(0.0, _right - _left)
        def _gantt_status(status_value, due_value):
            _status = str(status_value or "Not Started")
            _due = pd.to_datetime(due_value, errors="coerce")
            if _status.lower() != "completed" and pd.notna(_due) and _due.date() < date.today(): return "Overdue", "#EF4444"
            return _status, {"Completed":"#00875A", "In Progress":"#1473E6", "At Risk":"#F28C00", "Delayed":"#EF4444", "On Track":"#00875A", "Not Started":"#8495B0"}.get(_status, "#8495B0")
        month_day_rows=[]
        _gantt_project_ids={}
        for _, _p in P.iterrows():
            _pid=str(int(_p.id)); _gantt_project_ids[str(_p.name)]=_pid
            _ps=pd.to_datetime(_p.start_date,errors="coerce"); _pe=pd.to_datetime(_p.end_date,errors="coerce")
            _left,_width=_bar_position(_p.start_date,_p.end_date)
            _status,_color=_gantt_status(_p.status,_p.end_date)
            _duration=max(1,(_pe-_ps).days+1) if pd.notna(_ps) and pd.notna(_pe) else 0
            month_day_rows.append((_pid,0,"folder",str(_p.bar_color or GREEN),str(_p.name),fmt_date(_p.start_date),fmt_date(_p.end_date),f"{_duration} days" if _duration else "—",_status,f"{int(_p.progress or 0)}%",_left,_width,_color,100,"",False,"#F0FDF4",""))
            _pt=T[T.project_id==_p.id] if not T.empty else pd.DataFrame()
            for _ix,(_, _t) in enumerate(_pt.iterrows(),1):
                _left,_width=_bar_position(_t.get("start_date"),_t.get("due_date")); _status,_color=_gantt_status(_t.get("status"),_t.get("due_date"))
                _ts=pd.to_datetime(_t.get("start_date"),errors="coerce"); _te=pd.to_datetime(_t.get("due_date"),errors="coerce")
                _duration=max(1,(_te-_ts).days+1) if pd.notna(_ts) and pd.notna(_te) else 0
                _pct=100 if _status=="Completed" else int(_t.get("progress") or 0)
                month_day_rows.append((f"{_pid}.{_ix}",1,"play",_color,str(_t.get("name") or "Untitled task"),fmt_date(_t.get("start_date")),fmt_date(_t.get("due_date")),f"{_duration} days" if _duration else "—",_status,f"{_pct}%",_left,_width,_color,_pct,str(_t.get("dependency") or ""),bool(_t.get("dependency")),"#FFFFFF",""))
            _pm=M[M.project_id==_p.id] if not M.empty else pd.DataFrame()
            for _ix,(_, _m) in enumerate(_pm.iterrows(),1):
                _left,_width=_bar_position(_m.get("due_date"),_m.get("due_date")); _status,_color=_gantt_status(_m.get("status"),_m.get("due_date"))
                month_day_rows.append((f"{_pid}.M{_ix}",1,"doc",PURPLE,str(_m.get("name") or "Milestone"),fmt_date(_m.get("start_date")),fmt_date(_m.get("due_date")),"0 days",_status,"100%" if _status=="Completed" else "0%",_left,max(_width,0.8),PURPLE,100,str(_m.get("name") or "Milestone"),bool(_m.get("dependency")),"#FFFFFF",""))
        week_rows=month_day_rows
        day_rows=month_day_rows

        active_rows = week_rows if scale == "Week" else (day_rows if scale == "Day" else month_day_rows)

        # Apply Top Search & Dropdown Filters
        proj_prefix_map = _gantt_project_ids
        if g_proj_sel != "All Projects" and g_proj_sel in proj_prefix_map:
            pfx = proj_prefix_map[g_proj_sel]
            active_rows = [r for r in active_rows if r[0] == pfx or r[0].startswith(pfx + ".")]
        if g_stat_sel != "All Statuses":
            active_rows = [r for r in active_rows if r[8] == g_stat_sel or r[1] == 0]
        if g_view_sel == "Milestones Only":
            active_rows = [r for r in active_rows if r[14] or r[1] == 0]
        if gantt_q:
            q_low = gantt_q.lower()
            active_rows = [r for r in active_rows if q_low in r[4].lower() or q_low in r[0].lower()]

        # Build a date-aware timeline header from the selected range.
        _today_offset = max(0.0, min(100.0, ((date.today() - _range_start).days / _range_days) * 100))
        if scale == "Month":
            _month_ticks = []
            _cursor = _range_start.replace(day=1)
            while _cursor <= _range_end and len(_month_ticks) < 18:
                _month_ticks.append(_cursor)
                _cursor = (_cursor.replace(day=28) + timedelta(days=4)).replace(day=1)
            _month_cells = "".join(f'<div style="padding:6px 2px;border-right:1px solid #E2E8F0">{_d.strftime("%b %Y")}</div>' for _d in _month_ticks)
            _week_ticks = []
            _cursor = _range_start
            while _cursor <= _range_end and len(_week_ticks) < 60:
                _week_ticks.append(_cursor)
                _cursor += timedelta(days=7)
            _week_cells = "".join(f'<div style="padding:5px 0;border-right:1px solid #EDF2F7;{"color:#1D4ED8;font-weight:700" if _d <= date.today() < _d + timedelta(days=7) else ""}">{_d.strftime("%b %-d") if os.name != "nt" else _d.strftime("%b %d").replace(" 0", " ")}</div>' for _d in _week_ticks)
            tl_top_hdr = f'<div style="display:grid;grid-template-columns:repeat({max(1,len(_month_ticks))},1fr);text-align:center;background:#F8FAFC;font-size:11px;font-weight:700;color:#1E293B">{_month_cells}</div><div style="display:grid;grid-template-columns:repeat({max(1,len(_week_ticks))},1fr);text-align:center;background:#F8FAFC;font-size:10px;font-weight:600;color:#475569">{_week_cells}</div>'
        elif scale == "Week":
            _week_ticks = []
            _cursor = _range_start - timedelta(days=_range_start.weekday())
            while _cursor <= _range_end and len(_week_ticks) < 60:
                _week_ticks.append(_cursor)
                _cursor += timedelta(days=7)
            _week_cells = "".join(f'<div style="padding:5px 2px;border-right:1px solid #E2E8F0;background:{"#D1FAE5" if _d <= date.today() < _d + timedelta(days=7) else "#F8FAFC"}"><div style="font-size:10px;font-weight:700;color:#0F172A">{_d.strftime("%b %-d") if os.name != "nt" else _d.strftime("%b %d").replace(" 0", " ")}</div><div style="font-size:9px;color:#64748B">Week {_d.isocalendar().week}</div></div>' for _d in _week_ticks)
            tl_top_hdr = f'<div style="display:grid;grid-template-columns:repeat({max(1,len(_week_ticks))},1fr);text-align:center">{_week_cells}</div>'
        else:
            _day_ticks = []
            _cursor = _range_start
            while _cursor <= _range_end and len(_day_ticks) < 90:
                _day_ticks.append(_cursor)
                _cursor += timedelta(days=1)
            _day_cells = "".join(f'<div style="padding:5px 1px;border-right:1px solid #EDF2F7;background:{"#D1FAE5" if _d == date.today() else "#F8FAFC"}">{_d.strftime("%d")}</div>' for _d in _day_ticks)
            tl_top_hdr = f'<div style="display:flex;align-items:center;justify-content:space-between;padding:5px 12px;border-bottom:1px solid #E2E8F0;background:#F8FAFC"><div style="font-size:12px;font-weight:700;color:#0F172A">Daily timeline · {_range_start.strftime("%b %d, %Y")} – {_range_end.strftime("%b %d, %Y")}</div></div><div style="display:grid;grid-template-columns:repeat({max(1,len(_day_ticks))},minmax(18px,1fr));background:#F8FAFC;font-size:10px;font-weight:600;color:#475569;text-align:center;position:relative">{_day_cells}</div>'
        today_line_html = f'<div style="position:absolute;top:0;bottom:0;left:{_today_offset:.2f}%;width:1.5px;background:#10B981;z-index:4;pointer-events:none"></div>'

        # Render Rows HTML
        body_rows_html = ""
        for r_tuple in active_rows:
            rid, indent, ic_type, ic_col, rname, rstart, rend, rdur, rstat, rpct, bleft, bw, bcol, bfill, m_txt, has_dep, rbg = r_tuple[:17]
            right_bar_label = r_tuple[17] if len(r_tuple) > 17 else ""
            is_bold = indent in (0, 1) and (scale != "Week" or indent == 0)
            fw = "700" if is_bold else "500"
            pad_l = 8 + indent * 18
            caret = "⌄ &nbsp;" if (indent == 0 or (scale == "Month" and indent == 1) or (scale == "Day" and rid in ("1.2", "1.3"))) else "&nbsp;&nbsp;&nbsp;&nbsp;"
            if ic_type == "folder":
                ico_badge = f'<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:4px;background:{ic_col};color:white;font-size:10px;margin-right:6px">📁</span>'
            elif ic_type == "doc":
                ico_badge = f'<span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:4px;background:{ic_col};color:white;font-size:10px;margin-right:6px">▤</span>'
            else:
                ico_badge = f'<span style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:4px;border:1.5px solid {ic_col};color:{ic_col};font-size:8px;margin-right:6px">▶</span>'

            # Bar + Milestone + Dependency Connector
            inner_fill = (
                f'<div style="width:{bfill}%;height:100%;background:#1D4ED8;border-radius:4px 0 0 4px"></div>'
                if bfill < 100
                else ""
            )
            bar_el = (
                f'<div style="position:absolute;left:{bleft}%;width:{bw}%;height:14px;top:9px;background:{bcol};border-radius:4px;overflow:hidden;z-index:2">{inner_fill}</div>'
                if bw > 0
                else ""
            )
            dep_el = (
                f'<div style="position:absolute;left:{bleft + bw}%;top:15px;width:8px;height:16px;border-right:1.5px solid #64748B;border-bottom:1.5px solid #64748B;z-index:3"><span style="position:absolute;right:-4px;bottom:-6px;font-size:9px;color:#64748B">▾</span></div>'
                if has_dep and bw > 0
                else ""
            )
            ms_el = ""
            if m_txt:
                lbl = "" if m_txt == " end_diamond" else f'<span style="font-size:11px;font-weight:600;color:#1E293B;margin-left:4px">{html.escape(m_txt)}</span>'
                ms_el = f'<div style="position:absolute;left:{bleft + bw + 0.6}%;top:5px;display:flex;align-items:center;z-index:3;white-space:nowrap"><span style="color:#7C3AED;font-size:16px;line-height:1">◆</span>{lbl}</div>'

            day_txt_el = (
                f'<div style="position:absolute;left:{bleft + bw + 1.2}%;top:8px;font-size:11px;font-weight:500;color:#1E293B;white-space:nowrap;z-index:3">{html.escape(right_bar_label)}</div>'
                if right_bar_label and not m_txt
                else ""
            )

            grid_cols = 10 if scale == "Week" else (11 if scale == "Day" else 12)
            bg_cells = "".join(f'<div style="border-right:1px solid #F1F5F9;height:32px"></div>' for _ in range(grid_cols))

            body_rows_html += f"""
            <div style="display:grid;grid-template-columns:46% 54%;border-bottom:1px solid #EDF2F7;background:{rbg};align-items:center;font-size:12px">
                <div style="display:grid;grid-template-columns:52px 1.9fr 92px 92px 64px 92px 68px;align-items:center;border-right:1px solid #E2E8F0;height:32px">
                    <div style="padding-left:8px;font-weight:{fw};color:#334155;font-size:11.5px">{caret}{rid}</div>
                    <div style="padding-left:{pad_l}px;font-weight:{fw};color:#0F172A;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{ico_badge}{html.escape(rname)}</div>
                    <div style="font-weight:{fw};color:#334155;font-size:11.5px">{rstart}</div>
                    <div style="font-weight:{fw};color:#334155;font-size:11.5px">{rend}</div>
                    <div style="font-weight:{fw};color:#334155;font-size:11.5px">{rdur}</div>
                    <div>{badge_html(rstat)}</div>
                    <div style="font-weight:{fw};color:#0F172A;padding-left:6px">{rpct}</div>
                </div>
                <div style="position:relative;height:32px">
                    <div style="display:grid;grid-template-columns:repeat({grid_cols},1fr);position:absolute;inset:0">{bg_cells}</div>
                    {bar_el}{dep_el}{ms_el}{day_txt_el}
                </div>
            </div>
            """

        gantt_html = clean_html(
            f"""
            <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:{'auto' if _gantt_expanded else 'hidden'};max-height:{'78vh' if _gantt_expanded else 'none'};background:white;margin-top:8px">
                <div style="display:grid;grid-template-columns:46% 54%;border-bottom:1px solid #CBD5E1;background:#F8FAFC;align-items:end">
                    <div style="display:grid;grid-template-columns:52px 1.9fr 92px 92px 64px 92px 68px;padding:10px 0;font-size:11.5px;font-weight:700;color:#1E293B;border-right:1px solid #E2E8F0">
                        <div style="padding-left:14px">ID</div>
                        <div style="padding-left:8px">Task Name</div>
                        <div>Start Date</div>
                        <div>End Date</div>
                        <div>Duration</div>
                        <div>Status</div>
                        <div>% Complete</div>
                    </div>
                    <div>{tl_top_hdr}</div>
                </div>
                <div style="position:relative">
                    {today_line_html}
                    {body_rows_html}
                </div>
            </div>
            """
        )
        st.markdown(gantt_html, unsafe_allow_html=True)


# ===========================================================================
# PAGE 4: TASK LIST (Exact Replica of Main Table + Add/View/Edit Task Modals)
# ===========================================================================
elif page == "Task List":

    @st.dialog("Add Task", width="large")
    def add_task_modal() -> None:
        if not projects_list:
            st.warning("Create a project or import a Projects sheet from Settings before adding tasks.")
            return
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:14px">
                    <div style="width:38px;height:38px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:700">＋</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Add Task</div>
                        <div style="font-size:12.5px;color:#64748B">Enter the task details below.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        r1_a, r1_b = st.columns(2)
        with r1_a:
            proj = st.selectbox("Project *", projects_list, index=0, key="add_t_proj")
        with r1_b:
            parent_t = st.selectbox("Parent Task", ["None (Top Level Task)"] + (T.name.dropna().astype(str).tolist() if not T.empty else []), key="add_t_parent")

        r2_a, r2_b = st.columns(2)
        with r2_a:
            t_name = st.text_input("Task Name *", placeholder="Enter task name", key="add_t_name")
        with r2_b:
            st.text_input("Task ID (Auto)", value="Will be generated automatically", disabled=True, key="add_t_id_auto")

        r3_a, r3_b = st.columns([1.35, 1])
        with r3_a:
            t_desc = st.text_area("Description", placeholder="Enter task description, objectives, and details...", height=108, key="add_t_desc")
        with r3_b:
            _add_priority_labels = ["🔴 High", "🟠 Medium", "🔵 Low"]
            _add_priority_default = get_setting("default_task_priority", "Medium")
            _add_priority_idx = {"High": 0, "Medium": 1, "Low": 2}.get(_add_priority_default, 1)
            t_prio = st.selectbox("Priority *", _add_priority_labels, index=_add_priority_idx, key="add_t_prio")
            _add_status_labels = ["🟦 Not Started", "🔵 In Progress", "🟢 Completed", "🟠 At Risk"]
            _add_status_default = get_setting("default_task_status", "Not Started")
            _add_status_idx = {"Not Started": 0, "In Progress": 1, "Completed": 2, "At Risk": 3}.get(_add_status_default, 0)
            t_stat = st.selectbox("Status *", _add_status_labels, index=_add_status_idx, key="add_t_stat")

        r4_a, r4_b, r4_c = st.columns(3)
        with r4_a:
            t_start = st.date_input("Start Date *", date.today(), key="add_t_start")
        with r4_b:
            t_end = st.date_input("End Date *", date.today()+timedelta(days=7), key="add_t_end")
        with r4_c:
            _add_duration = max(1, (t_end - t_start).days + 1)
            st.number_input("Duration (Days)", min_value=1, max_value=365, value=min(365, _add_duration), key="add_t_dur", disabled=True)

        r5_a, r5_b, r5_c = st.columns([1.1, 1.1, 1])
        with r5_a:
            assignees_pool = R.name.dropna().astype(str).tolist() if not R.empty else ["Unassigned"]
            t_ass = st.selectbox("Assignee *", assignees_pool, index=0, key="add_t_ass")
        with r5_b:
            t_pct = st.slider("% Complete (auto)", 0, 100, 0, key="add_t_pct", disabled=True)
        with r5_c:
            st.selectbox("Task Type", ["Task", "Milestone", "Subtask"], key="add_t_type")

        r6_a, r6_b, r6_c = st.columns([1.1, 1.1, 1])
        with r6_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            st.toggle("◆ Mark as Milestone", value=False, key="add_t_ms_tog")
        with r6_b:
            t_dep = st.selectbox("Dependencies", ["None"] + (T.name.dropna().astype(str).tolist() if not T.empty else []), key="add_t_dep")
        with r6_c:
            st.multiselect("Tags", ["Planning", "Design", "Development", "Testing", "Launch", "Research", "Strategy"], placeholder="Select or enter tags", key="add_t_tags")

        r7_a, r7_b, r7_c = st.columns([1, 1, 1.2])
        with r7_a:
            st.number_input("Estimated Effort (Hours)", min_value=0.0, value=0.0, step=1.0, key="add_t_est")
        with r7_b:
            st.number_input("Actual Effort (Hours)", min_value=0.0, value=0.0, step=1.0, key="add_t_act")
        with r7_c:
            t_notes = st.text_area("Progress Notes", placeholder="Enter progress notes...", height=68, key="add_t_notes")

        st.markdown(
            clean_html(
                """
                <div style="font-size:12.5px;font-weight:700;color:#0F172A;margin:6px 0 6px 0">Attachments & Checklist / Subtasks</div>
                """
            ),
            unsafe_allow_html=True,
        )

        task_uploads = st.file_uploader("Attach files (PDF, DOCX, XLSX, PPTX, JPG, PNG; up to 25 MB each)", type=["pdf", "docx", "xlsx", "pptx", "jpg", "jpeg", "png"], accept_multiple_files=True, key="add_t_attachments")
        sc1, sc2 = st.columns([4.5, 0.8])
        with sc1:
            st.text_input("Subtask", placeholder="Enter subtask", key="add_t_sub_in", label_visibility="collapsed")
        with sc2:
            if st.button("＋ Add", type="primary", key="add_t_sub_btn", use_container_width=True):
                _subtask = st.session_state.get("add_t_sub_in", "").strip()
                if _subtask:
                    st.session_state.setdefault("add_t_sub_items", []).append(_subtask)
                    st.session_state["add_t_sub_in"] = ""
                    st.rerun()
        if st.session_state.get("add_t_sub_items"):
            for _idx, _subtask in enumerate(st.session_state["add_t_sub_items"], 1):
                st.caption(f"☐  {_subtask}")

        if st.session_state.get("add_t_sub_items"):
            st.caption(f"{len(st.session_state['add_t_sub_items'])} subtask(s) will be saved with this task. Completion progress starts at 0% and updates as items are checked.")


        f_chk, f_can, f_sav = st.columns([3.2, 0.9, 1.1])
        with f_chk:
            create_another = st.checkbox("Create another task after saving", key="add_t_another")
        with f_can:
            if st.button("Cancel", key="add_t_cancel", use_container_width=True):
                st.session_state.pop("task_modal", None)
                st.rerun()
        with f_sav:
            if st.button("Save Task", type="primary", key="add_t_save", use_container_width=True):
                if not t_name.strip():
                    st.error("Enter a task name before saving.")
                elif t_end < t_start:
                    st.error("Task end date must be on or after the start date.")
                else:
                    clean_prio = t_prio.split()[-1]
                    clean_stat = " ".join(t_stat.split()[1:])
                    sql(
                        "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id,task_type,tags,description,estimated_effort,actual_effort,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (
                            t_name.strip() or "New Project Task",
                            pmap.get(proj),
                            "IT",
                            clean_stat,
                            clean_prio,
                            t_start.isoformat(),
                            t_end.isoformat(),
                            t_ass,
                            int(t_pct),
                            "" if t_dep == "None" else t_dep,
                            t_notes or t_desc,
                            (int(T.loc[T.name.astype(str) == parent_t, "id"].iloc[0]) if parent_t != "None (Top Level Task)" and not T[T.name.astype(str) == parent_t].empty else None),
                            ("Milestone" if st.session_state.get("add_t_ms_tog", False) else st.session_state.get("add_t_type", "Task")),
                            ",".join(st.session_state.get("add_t_tags", [])),
                            t_desc.strip(),
                            (float(st.session_state.get("add_t_est") or 0) or None),
                            (float(st.session_state.get("add_t_act") or 0) or None),
                            datetime.now().isoformat(timespec="seconds"), datetime.now().isoformat(timespec="seconds"),
                        ),
                    )
                    with conn() as _c:
                        _new_task_id = _c.execute("SELECT id FROM tasks WHERE project_id=? AND name=? ORDER BY id DESC LIMIT 1", (pmap.get(proj), t_name.strip())).fetchone()[0]
                    for _subtask_name in st.session_state.get("add_t_sub_items", []):
                        sql("INSERT INTO subtasks(task_id,name,completed) VALUES(?,?,0)", (int(_new_task_id), str(_subtask_name)))
                    if st.session_state.get("add_t_ms_tog", False):
                        sql("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?)", (t_name.strip(), pmap.get(proj), "IT", "Not Started", t_start.isoformat(), t_end.isoformat(), "", t_notes or t_desc))
                    st.session_state["add_t_sub_items"] = []
                    if task_uploads:
                        try:
                            save_task_attachments(int(_new_task_id), task_uploads)
                        except Exception as _att_exc:
                            st.warning(f"Task was saved, but one or more attachments could not be saved: {_att_exc}")
                    log_event("Created", "Task", int(_new_task_id), t_name.strip() or "New Project Task")
                    if not create_another:
                        st.session_state.pop("task_modal", None)
                    st.rerun()
    @st.dialog("View Task", width="large")
    def view_task_modal(task_rec: dict) -> None:
        tname = task_rec.get("name", "")
        tproj = task_rec.get("project", projects_list[0] if projects_list else "")
        tstart = fmt_date(task_rec.get("raw_start") or task_rec.get("start_date"))
        tend = fmt_date(task_rec.get("raw_due") or task_rec.get("end_date"))
        tdur = task_rec.get("duration", "—")
        tpct = int(task_rec.get("pct", 0) or 0)
        tstat = task_rec.get("status", "Not Started")
        tprio = task_rec.get("priority", "Medium")
        tass = task_rec.get("assignee", "Unassigned")
        _task_id_for_view = int(task_rec.get("id") or 0)
        _task_detail_row = df("SELECT * FROM tasks WHERE id=?", (_task_id_for_view,))
        _task_details = _task_detail_row.iloc[0].to_dict() if not _task_detail_row.empty else {}
        _task_description = str(_task_details.get("description") or _task_details.get("notes") or "No description provided.")
        _task_parent_id = _task_details.get("parent_task_id")
        _task_parent_name = "None" if not _task_parent_id else str((T[T.id == int(_task_parent_id)].iloc[0].get("name")) if not T[T.id == int(_task_parent_id)].empty else "Parent task unavailable")
        _task_type = str(_task_details.get("task_type") or "Task").replace("💼 ", "").replace("◆ ", "")
        _task_assignee_row = R[R.name.astype(str) == str(tass)] if not R.empty else pd.DataFrame()
        _task_assignee_role = str(_task_assignee_row.iloc[0].get("role") or "Team Member") if not _task_assignee_row.empty else "Unassigned"
        _task_is_milestone = _task_type == "Milestone" or (not M.empty and bool(((M.project_id == _task_details.get("project_id")) & (M.name.astype(str) == str(tname))).any()))
        _view_subtasks = task_rec.get("subtasks", []) or []
        _view_sub_done = sum(1 for _sub in _view_subtasks if str(_sub.get("completed", 0)).strip().lower() in ("1", "true", "yes", "completed"))
        _view_sub_total = len(_view_subtasks)
        _view_subtasks_html = "".join(
            f"<div style='display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9;padding:4px 0'><span>{'✓' if str(_sub.get('completed', 0)).strip().lower() in ('1','true','yes','completed') else '○'} &nbsp;{html.escape(str(_sub.get('name') or 'Untitled subtask'))}</span><span style='color:#64748B'>{'Completed' if str(_sub.get('completed', 0)).strip().lower() in ('1','true','yes','completed') else 'Pending'}</span></div>"
            for _sub in _view_subtasks
        ) or "<div style='font-size:12px;color:#64748B'>No subtasks. Mark this task Completed in the task list to reach 100%.</div>"
        _view_activity = df("SELECT event,details,created_at FROM audit_log WHERE lower(item_type)='task' AND item_id=? ORDER BY created_at DESC LIMIT 5", (int(task_rec.get("id") or 0),))
        _view_activity_html = "".join(
            f"<div style='display:grid;grid-template-columns:130px 155px 1fr;gap:8px;border-bottom:1px solid #F1F5F9;align-items:center;padding:4px 0'><span>{html.escape(str(_a.get('event') or 'Updated'))}</span><span style='color:#64748B;font-size:11.5px'>{html.escape(fmt_date(_a.get('created_at')))}</span><span>{html.escape(str(_a.get('details') or ''))}</span></div>"
            for _, _a in _view_activity.iterrows()
        ) or "<div style='color:#64748B'>No activity recorded for this task yet.</div>"


        st.markdown(
            clean_html(
                f"""
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:14px">
                    <div style="width:38px;height:38px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:18px">👁</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">View Task</div>
                        <div style="font-size:12.5px;color:#64748B">See the complete details of this task.</div>
                    </div>
                </div>
                <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;margin-bottom:14px">
                    <div style="display:flex;align-items:center;gap:12px">
                        <div style="width:46px;height:46px;border-radius:10px;background:#1473E6;color:white;display:flex;align-items:center;justify-content:center;font-size:22px">📅</div>
                        <div>
                            <div style="font-size:20px;font-weight:800;color:#0F172A">{html.escape(tname)}</div>
                            <div style="font-size:13px;color:#475569">{html.escape(tproj)}</div>
                        </div>
                    </div>
                    <div style="display:flex;gap:8px;align-items:center">
                        <span class="vk-badge vk-badge-green" style="padding:5px 12px;font-size:12px">✓ &nbsp;{html.escape(tstat)}</span>
                        <span class="vk-badge vk-badge-red" style="padding:5px 12px;font-size:12px">⚑ &nbsp;{html.escape(tprio)} Priority</span>
                        <span class="vk-badge vk-badge-blue" style="padding:5px 14px;font-size:12px">Task</span>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;background:#F8FAFC">
                        <span style="font-size:20px;color:#1E293B">📅</span>
                        <div><div style="font-size:11px;color:#64748B">Start Date</div><div style="font-size:14px;font-weight:800;color:#0F172A">{tstart}</div></div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;background:#F8FAFC">
                        <span style="font-size:20px;color:#1E293B">📅</span>
                        <div><div style="font-size:11px;color:#64748B">End Date</div><div style="font-size:14px;font-weight:800;color:#0F172A">{tend}</div></div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;background:#F8FAFC">
                        <span style="font-size:20px;color:#1E293B">🕒</span>
                        <div><div style="font-size:11px;color:#64748B">Duration</div><div style="font-size:14px;font-weight:800;color:#0F172A">{tdur}</div></div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;background:#F8FAFC">
                        <span style="width:26px;height:26px;border-radius:50%;border:4px solid #00875A;display:inline-block"></span>
                        <div><div style="font-size:11px;color:#64748B">% Complete</div><div style="font-size:15px;font-weight:800;color:#0F172A">{tpct}%</div></div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:1.55fr 1fr;gap:12px;margin-bottom:12px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📄 &nbsp;Description</div>
                        <div style="background:#F8FAFC;border:1px solid #EDF2F7;border-radius:8px;padding:10px 12px;font-size:12.5px;color:#334155;line-height:1.55">
                            {html.escape(_task_description)}
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">👤 &nbsp;Assignee</div>
                        <div style="background:#F8FAFC;border:1px solid #EDF2F7;border-radius:8px;padding:14px;display:flex;align-items:center;gap:12px">
                            <div style="width:44px;height:44px;border-radius:50%;background:#0284C7;color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px">{html.escape("".join(part[0] for part in str(tass).split()[:2]).upper() or "?")}</div>
                            <div>
                                <div style="font-size:14px;font-weight:800;color:#0F172A">{html.escape(tass)}</div>
                                <div style="font-size:12px;color:#64748B">{html.escape(_task_assignee_role)}</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Project</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">📁 &nbsp;{html.escape(tproj)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Parent Task</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">☰ &nbsp;{html.escape(_task_parent_name)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Milestone</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600"><span style="color:#7C3AED">◆</span> &nbsp;{"Milestone" if _task_is_milestone else "No"}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Status</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600"><span style="color:#00875A">●</span> &nbsp;{html.escape(tstat)}</div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr 1.2fr;gap:10px;margin-bottom:12px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Priority</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600"><span style="color:#EF4444">⚑</span> &nbsp;{html.escape(tprio)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Task Type</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">💼 &nbsp;{html.escape(_task_type)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">🏷 &nbsp;Tags</div>
                        <div>{''.join(tag_pill_html(_tag.strip()) for _tag in str(task_rec.get('tags') or '').split(',') if _tag.strip()) or '—'}</div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:1.35fr 1fr;gap:12px;margin-bottom:12px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">☑ &nbsp;Checklist / Subtasks ({_view_sub_total})</div>
                        {f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px"><div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden"><div style="width:{round(_view_sub_done / _view_sub_total * 100) if _view_sub_total else 0}%;height:100%;background:#00875A"></div></div><span style="font-size:11.5px;color:#475569;font-weight:600">{_view_sub_done} of {_view_sub_total} completed</span></div>' if _view_sub_total else ''}
                        <div style="font-size:12.5px;line-height:2.1">{_view_subtasks_html}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📎 &nbsp;Attachments</div>
                        <div style="font-size:12px;color:#64748B">Task files are listed below.</div>
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:14px">
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">💬 &nbsp;Recent Activity / Notes</div>
                    {_view_activity_html}
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        st.markdown("**Task Attachments**")
        render_task_attachments(int(task_rec.get("id") or 0), "view_task_files")

        b_sp, b_cl, b_ed = st.columns([3.8, 0.9, 1.1])
        with b_cl:
            if st.button("Close", key="close_v_task", use_container_width=True):
                st.session_state.pop("task_modal", None)
                st.rerun()
        with b_ed:
            if st.button("✎ Edit Task", type="primary", key="to_edit_v_task", use_container_width=True):
                st.session_state.task_modal = ("edit", task_rec)
                st.rerun()

    @st.dialog("Edit Task", width="large")
    def edit_task_modal(task_rec: dict) -> None:
        if not projects_list:
            st.warning("Import or create a project before editing tasks.")
            return
        tname = task_rec.get("name", "")
        tproj = task_rec.get("project", projects_list[0] if projects_list else "")
        twbs = task_rec.get("wbs", "1.3")
        _edit_task_id = int(task_rec.get("id") or twbs)
        if st.session_state.get("_ed_sub_task_id") != _edit_task_id:
            _sub_df = df("SELECT id, task_id, name, completed, due_date FROM subtasks WHERE task_id=? ORDER BY id", (_edit_task_id,))
            st.session_state["ed_sub_items"] = _sub_df.to_dict("records") if not _sub_df.empty else []
            st.session_state["_ed_sub_task_id"] = _edit_task_id

        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:14px">
                    <div style="width:38px;height:38px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:18px">✎</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Edit Task</div>
                        <div style="font-size:12.5px;color:#64748B">Update the task details below.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        e1_ico, e1_nm, e1_pr, e1_id = st.columns([0.35, 1.8, 1.3, 0.7])
        with e1_ico:
            st.markdown("<div style='width:44px;height:44px;border-radius:10px;background:#1473E6;color:white;display:flex;align-items:center;justify-content:center;font-size:22px;margin-top:18px'>📅</div>", unsafe_allow_html=True)
        with e1_nm:
            ed_name = st.text_input("Task Name *", value=tname, key="ed_t_name")
        with e1_pr:
            p_opts = projects_list if tproj in projects_list else [tproj] + projects_list
            ed_proj = st.selectbox("Project *", p_opts, index=p_opts.index(tproj), key="ed_t_proj")
        with e1_id:
            st.text_input("Task ID (Auto)", value=twbs, disabled=True, key="ed_t_id")

        e2_l, e2_r = st.columns([1.65, 1])
        with e2_l:
            ed_desc = st.text_area(
                "Description *",
                value=str(task_rec.get("description") or task_rec.get("notes") or ""),
                height=92,
                key="ed_t_desc",
            )
        with e2_r:
            ass_list = sorted(set((R.name.dropna().astype(str).tolist() if not R.empty else []) + [str(task_rec.get("assignee") or "Unassigned")]))
            _ass_index = ass_list.index(str(task_rec.get("assignee") or "Unassigned")) if str(task_rec.get("assignee") or "Unassigned") in ass_list else 0
            ed_ass = st.selectbox("👤 Assignee *", ass_list or ["Unassigned"], index=_ass_index, key="ed_t_ass")

        e3_a, e3_b, e3_c, e3_d = st.columns(4)
        with e3_a:
            ed_start = st.date_input("Start Date *", pd.to_datetime(task_rec.get("raw_start"), errors="coerce").date() if pd.notna(pd.to_datetime(task_rec.get("raw_start"), errors="coerce")) else date.today(), key="ed_t_start")
        with e3_b:
            ed_end = st.date_input("End Date *", pd.to_datetime(task_rec.get("raw_due"), errors="coerce").date() if pd.notna(pd.to_datetime(task_rec.get("raw_due"), errors="coerce")) else date.today()+timedelta(days=7), key="ed_t_end")
        with e3_c:
            _ed_start_dt = pd.to_datetime(task_rec.get("raw_start"), errors="coerce")
            _ed_end_dt = pd.to_datetime(task_rec.get("raw_due"), errors="coerce")
            _ed_duration_default = max(1, (_ed_end_dt.date() - _ed_start_dt.date()).days + 1) if pd.notna(_ed_start_dt) and pd.notna(_ed_end_dt) else 7
            st.number_input("Duration (Days) *", min_value=1, max_value=365, value=min(365, _ed_duration_default), key="ed_t_dur", disabled=True)
        with e3_d:
            ed_pct = st.number_input("% Complete (auto)", min_value=0, max_value=100, value=int(task_rec.get("pct", 0)), key="ed_t_pct", disabled=True)

        e4_a, e4_b, e4_c, e4_d = st.columns(4)
        with e4_a:
            _priority_opts = ["⚑ High", "⚑ Medium", "⚑ Low"]
            _priority_current = str(task_rec.get("priority") or "Medium")
            _priority_index = {"High": 0, "Medium": 1, "Low": 2}.get(_priority_current, 1)
            ed_prio = st.selectbox("Priority *", _priority_opts, index=_priority_index, key="ed_t_prio")
        with e4_b:
            _status_opts = ["🟢 Completed", "🔵 In Progress", "⚪ Not Started", "🟠 At Risk"]
            _current_status = str(task_rec.get("status") or "Not Started")
            _status_index = {"Completed": 0, "In Progress": 1, "Not Started": 2, "At Risk": 3}.get(_current_status, 2)
            ed_stat = st.selectbox("Status *", _status_opts, index=_status_index, key="ed_t_stat")
        with e4_c:
            _ed_type_opts = ["Task", "Milestone", "Subtask"]
            _ed_type_current = str(task_rec.get("task_type") or "Task").replace("💼 ", "").replace("◆ ", "")
            st.selectbox("Task Type *", _ed_type_opts, index=_ed_type_opts.index(_ed_type_current) if _ed_type_current in _ed_type_opts else 0, key="ed_t_type")
        with e4_d:
            _parent_opts = ["None"] + (T[T.id != int(task_rec.get("id") or -1)].name.dropna().astype(str).tolist() if not T.empty else [])
            _parent_id = task_rec.get("parent_task_id")
            _parent_name = str(T[T.id == int(_parent_id)].iloc[0].get("name")) if _parent_id and not T[T.id == int(_parent_id)].empty else "None"
            st.selectbox("Parent Task", _parent_opts, index=_parent_opts.index(_parent_name) if _parent_name in _parent_opts else 0, key="ed_t_parent")

        ed_tags = st.multiselect("Tags", ["Planning", "Design", "Development", "Testing", "Launch", "Research", "Strategy"], default=[x.strip() for x in str(task_rec.get("tags") or "").split(",") if x.strip()], key="ed_t_tags")
        e5_a, e5_b, e5_c = st.columns([0.9, 0.9, 1.4])
        with e5_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            st.toggle("◆ Mark as Milestone", value=(str(task_rec.get("task_type") or "Task").replace("💼 ", "").replace("◆ ", "") == "Milestone" or (not M.empty and bool(((M.project_id == task_rec.get("project_id")) & (M.name.astype(str) == str(tname))).any()))), key="ed_t_ms")
        with e5_b:
            st.markdown("<div style='font-size:12.5px;margin-bottom:6px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("Milestone Type", ["Not a milestone", "Project Milestone"], index=1 if str(task_rec.get("task_type") or "") == "Milestone" else 0, key="ed_t_ms_sel", label_visibility="collapsed")
        with e5_c:
            st.caption("Tags are managed in the Tags field above.")

        e6_l, e6_r = st.columns([1.25, 1])
        with e6_l, st.container(border=True):
            st.markdown("<div style='font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px'>☑ &nbsp;Checklist / Subtasks</div>", unsafe_allow_html=True)
            cs1, cs2 = st.columns([3.5, 1])
            with cs1:
                st.text_input("Subtask", placeholder="Enter subtask...", key="ed_sub_in", label_visibility="collapsed")
            with cs2:
                if st.button("＋ Add", type="primary", key="ed_sub_add", use_container_width=True):
                    _subtask = st.session_state.get("ed_sub_in", "").strip()
                    if _subtask:
                        st.session_state.setdefault("ed_sub_items", []).append({"id": None, "task_id": _edit_task_id, "name": _subtask, "completed": 0, "due_date": None})
                        st.session_state["ed_sub_in"] = ""
                        st.rerun()
            _sub_items = st.session_state.get("ed_sub_items", [])
            _sub_done = 0
            for _idx, _subtask in enumerate(_sub_items):
                _sub_id = _subtask.get("id")
                _sub_key = f"ed_sub_done_{_edit_task_id}_{_sub_id if _sub_id is not None else 'new_'+str(_idx)}"
                _default_done = str(_subtask.get("completed", 0)).strip().lower() in ("1", "true", "yes", "completed")
                _is_done = st.checkbox(str(_subtask.get("name") or "Untitled subtask"), value=_default_done, key=_sub_key)
                _subtask["completed"] = int(_is_done)
                _sub_done += int(_is_done)
            if _sub_items:
                st.progress(_sub_done / len(_sub_items), text=f"{_sub_done} of {len(_sub_items)} subtasks completed")
            else:
                st.caption("No subtasks: mark the task Completed to reach 100%.")
        with e6_r, st.container(border=True):
            st.markdown("**📎 Attachments**")
            st.caption("Upload and manage real files using the attachment controls below.")

        with st.container(border=True):
            st.markdown("<div style='font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px'>📄 &nbsp;Additional Information <span style='font-weight:400;color:#64748B'>(Optional)</span></div>", unsafe_allow_html=True)
            ai1, ai2, ai3, ai4 = st.columns([0.8, 0.8, 1.2, 1.4])
            with ai1:
                st.number_input("Estimated Effort (Hours)", min_value=0.0, value=max(0.0, float(task_rec.get("estimated_effort") or 0)), step=1.0, key="ed_est_h")
            with ai2:
                st.number_input("Actual Effort (Hours)", min_value=0.0, value=max(0.0, float(task_rec.get("actual_effort") or 0)), step=1.0, key="ed_act_h")
            with ai3:
                _dep_opts = list(dict.fromkeys(["None"] + (T.name.dropna().astype(str).tolist() if not T.empty else []) + [str(task_rec.get("dependency") or "")]))
            _dep_opts = list(dict.fromkeys([x for x in _dep_opts if x is not None]))
            st.selectbox("Dependencies", _dep_opts, index=_dep_opts.index(str(task_rec.get("dependency") or "None")) if str(task_rec.get("dependency") or "None") in _dep_opts else 0, key="ed_dep_s")
            with ai4:
                ed_notes = st.text_area("Progress Notes", value=str(task_rec.get("notes") or ""), height=68, key="ed_p_notes")

        st.markdown("**Existing Attachments**")
        render_task_attachments(_edit_task_id, "edit_task_files")
        ed_task_uploads = st.file_uploader("Add attachments", type=["pdf", "docx", "xlsx", "pptx", "jpg", "jpeg", "png"], accept_multiple_files=True, key=f"ed_task_uploads_{_edit_task_id}")
        _edit_activity = df("SELECT event,created_at FROM audit_log WHERE lower(item_type)='task' AND item_id=? ORDER BY created_at DESC LIMIT 2", (_edit_task_id,))
        _last_activity = fmt_date(_edit_activity.iloc[0].get("created_at")) if not _edit_activity.empty else "No activity yet"
        f_meta, f_can, f_sav = st.columns([3.3, 0.85, 1.15])
        with f_meta:
            st.markdown(
                clean_html(
                    f"""
                    <div style="font-size:11.5px;color:#475569;padding-top:4px">Latest recorded task activity: {_last_activity}</div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with f_can:
            if st.button("Cancel", key="ed_t_cancel", use_container_width=True):
                st.session_state.pop("task_modal", None)
                st.rerun()
        with f_sav:
            if st.button("Save Changes", type="primary", key="ed_t_save", use_container_width=True):
                if not ed_name.strip():
                    st.error("Enter a task name before saving.")
                    return
                if ed_end < ed_start:
                    st.error("Task end date must be on or after the start date.")
                    return
                _clean_ed_status = " ".join(str(ed_stat).split()[1:])
                _clean_ed_priority = str(ed_prio).split()[-1]
                _new_project_id = int(pmap.get(ed_proj, task_rec.get("project_id") or 0))
                sql(
                    "UPDATE tasks SET name=?, project_id=?, assignee=?, start_date=?, due_date=?, priority=?, status=?, progress=?, dependency=?, notes=?, parent_task_id=?, task_type=?, tags=?, description=?, estimated_effort=?, actual_effort=?, updated_at=? WHERE id=?",
                    (ed_name.strip() or tname, _new_project_id, ed_ass, ed_start.isoformat(), ed_end.isoformat(), _clean_ed_priority, _clean_ed_status, int(ed_pct), "" if st.session_state.get("ed_dep_s") == "None" else st.session_state.get("ed_dep_s", ""), ed_notes, (int(T.loc[T.name.astype(str) == st.session_state.get("ed_t_parent"), "id"].iloc[0]) if st.session_state.get("ed_t_parent") not in (None, "None") and not T[T.name.astype(str) == st.session_state.get("ed_t_parent")].empty else None), ("Milestone" if st.session_state.get("ed_t_ms", False) else str(st.session_state.get("ed_t_type", task_rec.get("task_type") or "Task")).replace("💼 ", "").replace("◆ ", "")), ",".join(ed_tags), ed_desc, (float(st.session_state.get("ed_est_h") or 0) or None), (float(st.session_state.get("ed_act_h") or 0) or None), datetime.now().isoformat(timespec="seconds"), _edit_task_id),
                )
                if st.session_state.get("ed_t_ms", False):
                    _existing_ms = df("SELECT id FROM milestones WHERE project_id=? AND name=?", (_new_project_id, ed_name.strip() or tname))
                    if _existing_ms.empty:
                        sql("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?)", (ed_name.strip() or tname, _new_project_id, "", "Not Started", ed_start.isoformat(), ed_end.isoformat(), "", ed_notes or ed_desc))
                    else:
                        sql("UPDATE milestones SET start_date=?,due_date=?,notes=? WHERE id=?", (ed_start.isoformat(), ed_end.isoformat(), ed_notes or ed_desc, int(_existing_ms.iloc[0].id)))
                elif not st.session_state.get("ed_t_ms", False):
                    sql("DELETE FROM milestones WHERE project_id=? AND name=?", (_new_project_id, ed_name.strip() or tname))
                if ed_task_uploads:
                    try:
                        save_task_attachments(_edit_task_id, ed_task_uploads)
                    except Exception as _att_exc:
                        st.warning(f"Task changes were saved, but one or more attachments could not be saved: {_att_exc}")
                for _subtask in st.session_state.get("ed_sub_items", []):
                    if _subtask.get("id") is not None:
                        sql("UPDATE subtasks SET name=?, completed=?, due_date=? WHERE id=? AND task_id=?", (str(_subtask.get("name") or ""), int(_subtask.get("completed") or 0), _subtask.get("due_date"), int(_subtask["id"]), _edit_task_id))
                    elif str(_subtask.get("name") or "").strip():
                        sql("INSERT INTO subtasks(task_id,name,completed,due_date) VALUES(?,?,?,?)", (_edit_task_id, str(_subtask["name"]).strip(), int(_subtask.get("completed") or 0), _subtask.get("due_date")))
                log_event("Updated", "Task", _edit_task_id, ed_name)
                st.session_state.pop("task_modal", None)
                st.session_state.pop("_ed_sub_task_id", None)
                st.rerun()

    # --- TASK LIST MAIN HEADER ---
    th_l, th_s, th_d, th_b = st.columns([1.9, 1.05, 0.95, 0.65])
    with th_l:
        st.markdown(
            clean_html(
                f"""
                <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Task List</div>
                <div style='font-size:13.5px;color:#64748B;margin-top:2px'>Manage, track, and update all project tasks in one place.</div>
                """
            ),
            unsafe_allow_html=True,
        )
    with th_s:
        tl_top_q = st.text_input("Search", placeholder="🔍 Search tasks, projects or assignees...", key="tl_q_top", label_visibility="collapsed")
    with th_d:
        st.date_input("Date Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="tl_d_top", label_visibility="collapsed")
    with th_b:
        if st.button("＋ Add Task", type="primary", use_container_width=True, key="tl_btn_add_top"):
            st.session_state["add_t_sub_items"] = []
            st.session_state.task_modal = ("add", None)
            st.rerun()

    st.write("")
    # --- FILTER ROW ---
    tf1, tf2, tf3, tf4, tf5, tf6 = st.columns([1.05, 1.0, 1.0, 1.05, 1.35, 0.75])
    with tf1:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Project</div>", unsafe_allow_html=True)
        tl_proj_f = st.selectbox("Project", ["All Projects"] + projects_list, key="tl_f_proj", label_visibility="collapsed")
    with tf2:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Status</div>", unsafe_allow_html=True)
        tl_stat_f = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "Not Started", "At Risk", "Overdue"], key="tl_f_stat", label_visibility="collapsed")
    with tf3:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Priority</div>", unsafe_allow_html=True)
        tl_prio_f = st.selectbox("Priority", ["All Priorities", "High", "Medium", "Low"], key="tl_f_prio", label_visibility="collapsed")
    with tf4:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Assignee</div>", unsafe_allow_html=True)
        tl_ass_f = st.selectbox("Assignee", ["All Assignees"] + sorted(set(T.assignee.dropna().astype(str).tolist())), key="tl_f_ass", label_visibility="collapsed")
    with tf5:
        st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
        tl_sub_q = st.text_input("Search tasks", placeholder="🔍 Search tasks...", key="tl_q_sub", label_visibility="collapsed")
    with tf6:
        st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("▽ More Filters", key="tl_btn_more", use_container_width=True):
            for k, v in [("tl_f_proj", "All Projects"), ("tl_f_stat", "All Statuses"), ("tl_f_prio", "All Priorities"), ("tl_f_ass", "All Assignees"), ("tl_q_top", ""), ("tl_q_sub", "")]:
                st.session_state[k] = v
            st.rerun()

    st.write("")
    # --- 6 KPI CARDS ROW (Exact Match to Image 12) ---
    tk1, tk2, tk3, tk4, tk5, tk6 = st.columns(6)
    _task_status = T.status.fillna("Not Started").astype(str) if not T.empty else pd.Series(dtype=str)
    _task_total = len(T)
    _task_counts = {
        "Completed": int((_task_status.str.lower() == "completed").sum()),
        "In Progress": int((_task_status.str.lower() == "in progress").sum()),
        "Not Started": int((_task_status.str.lower() == "not started").sum()),
        "At Risk": int((_task_status.str.lower() == "at risk").sum()),
    }
    _due_dates = pd.to_datetime(T.due_date, errors="coerce") if not T.empty else pd.Series(dtype="datetime64[ns]")
    _overdue = int(((_due_dates.dt.date < date.today()) & (_task_status.str.lower() != "completed")).sum()) if not T.empty else 0
    def _pct_text(value):
        return f"({round(value / _task_total * 100)}%)" if _task_total else "(0%)"
    kpi_specs = [
        (tk1, "#1473E6", "#FFFFFF", "▤", "Total Tasks", "#475569", str(_task_total), ""),
        (tk2, "#00875A", "#FFFFFF", "✓", "Completed", "#475569", str(_task_counts["Completed"]), _pct_text(_task_counts["Completed"])),
        (tk3, "#1473E6", "#FFFFFF", "❂", "In Progress", "#475569", str(_task_counts["In Progress"]), _pct_text(_task_counts["In Progress"])),
        (tk4, "#64748B", "#FFFFFF", "●", "Not Started", "#475569", str(_task_counts["Not Started"]), _pct_text(_task_counts["Not Started"])),
        (tk5, "#F28C00", "#FFFFFF", "!", "At Risk", "#D97706", str(_task_counts["At Risk"]), _pct_text(_task_counts["At Risk"])),
        (tk6, "#EF4444", "#FFFFFF", "!", "Overdue", "#DC2626", str(_overdue), _pct_text(_overdue)),
    ]
    for col, ibg, ifg, ico, lbl, lcol, val, pct_s in kpi_specs:
        with col:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="min-height:84px;padding:12px 14px">
                        <div style="width:40px;height:40px;min-width:40px;border-radius:50%;background:{ibg};color:{ifg};display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:800">{ico}</div>
                        <div>
                            <div style="font-size:11.5px;font-weight:600;color:{lcol}">{lbl}</div>
                            <div style="font-size:21px;font-weight:800;color:#0F172A;line-height:1.2">{val} <span style="font-size:16px;font-weight:700">{pct_s}</span></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

    st.write("")
    # --- GROUPED WBS TASK TABLE (Exact Match to Image 12) ---
    task_groups = []
    _status_colors = {"Completed": "#00875A", "In Progress": "#1473E6", "At Risk": "#F28C00", "Delayed": "#EF4444", "Not Started": "#8495B0", "On Track": "#00875A"}
    for _, project_row in P.iterrows():
        project_tasks = T[T.project_id == project_row.id].copy() if not T.empty else pd.DataFrame()
        _ps = pd.to_datetime(project_row.start_date, errors="coerce")
        _pe = pd.to_datetime(project_row.end_date, errors="coerce")
        _dur = max(1, (int((_pe - _ps).days) + 1)) if pd.notna(_ps) and pd.notna(_pe) else 0
        _task_items = []
        for task_idx, (_, task_row) in enumerate(project_tasks.iterrows(), 1):
            _ts = pd.to_datetime(task_row.get("start_date"), errors="coerce")
            _te = pd.to_datetime(task_row.get("due_date"), errors="coerce")
            _days = max(1, (int((_te - _ts).days) + 1)) if pd.notna(_ts) and pd.notna(_te) else 0
            _stat = str(task_row.get("status") or "Not Started")
            _task_id = int(task_row.get("id", 0))
            _subtasks_df = df("SELECT id, task_id, name, completed, due_date FROM subtasks WHERE task_id=? ORDER BY id", (_task_id,))
            _subtasks = _subtasks_df.to_dict("records") if not _subtasks_df.empty else []
            _pct = int(task_row.get("progress") or 0)
            _task_items.append({"id": _task_id, "wbs": str(task_row.get("id", task_idx)), "ico": "doc", "ic_col": _status_colors.get(_stat, "#64748B"), "name": str(task_row.get("name") or "Untitled task"), "project": str(project_row.get("name") or ""), "start_date": fmt_date(task_row.get("start_date")), "end_date": fmt_date(task_row.get("due_date")), "duration": f"{_days} days" if _days else "—", "status": _stat, "priority": str(task_row.get("priority") or "Medium"), "assignee": str(task_row.get("assignee") or "Unassigned"), "pct": _pct, "raw_start": task_row.get("start_date"), "raw_due": task_row.get("due_date"), "project_id": int(task_row.get("project_id") or 0), "dependency": str(task_row.get("dependency") or ""), "notes": str(task_row.get("notes") or ""), "description": str(task_row.get("description") or ""), "task_type": str(task_row.get("task_type") or "Task"), "parent_task_id": task_row.get("parent_task_id"), "tags": str(task_row.get("tags") or ""), "estimated_effort": task_row.get("estimated_effort"), "actual_effort": task_row.get("actual_effort"), "subtasks": _subtasks, "bar_col": _status_colors.get(_stat, "#8495B0"), "tag": str(task_row.get("team") or ""), "tag_bg": "#DBEAFE", "tag_fg": "#1D4ED8"})
        _pct = int(project_row.get("progress") or 0)
        task_groups.append({"group_title": f"{project_row.get('name')} ({len(project_tasks)})", "project": str(project_row.get("name") or ""), "icon_bg": str(project_row.get("bar_color") or GREEN), "start": fmt_date(project_row.get("start_date")), "end": fmt_date(project_row.get("end_date")), "dur": f"{_dur} days" if _dur else "—", "pct": _pct, "bar_col": str(project_row.get("bar_color") or GREEN), "row_bg": "#F0FDF4", "tasks": _task_items})

    with st.container(border=True):
        col_weights = [0.55, 1.85, 1.05, 0.85, 0.85, 0.65, 0.85, 0.72, 1.25, 1.15, 0.9, 0.7]
        hdr_cols = st.columns(col_weights)
        hdr_titles = ["☐ &nbsp; ID", "Task Name &nbsp;▸", "Project", "Start Date", "End Date", "Duration", "Status", "Priority", "Assignee", "% Complete", "Tags", "Actions"]
        for c, h_t in zip(hdr_cols, hdr_titles):
            c.markdown(f"<div style='font-size:11.5px;font-weight:700;color:#1E293B;padding-bottom:6px'>{h_t}</div>", unsafe_allow_html=True)
        st.markdown("<hr style='margin:0 0 6px 0;border-color:#E2E8F0'>", unsafe_allow_html=True)

        q_combined = (tl_top_q + " " + tl_sub_q).strip().lower()

        for grp in task_groups:
            if tl_proj_f != "All Projects" and grp["project"] != tl_proj_f:
                continue

            filtered_tasks = []
            for t_item in grp["tasks"]:
                if tl_stat_f != "All Statuses" and t_item["status"] != tl_stat_f:
                    continue
                if tl_prio_f != "All Priorities" and t_item["priority"] != tl_prio_f:
                    continue
                if tl_ass_f != "All Assignees" and t_item["assignee"] != tl_ass_f:
                    continue
                if q_combined and not (q_combined in t_item["name"].lower() or q_combined in t_item["assignee"].lower() or q_combined in grp["project"].lower()):
                    continue
                filtered_tasks.append(t_item)

            if not filtered_tasks:
                continue

            # Project Group Banner Row
            st.markdown(
                clean_html(
                    f"""
                    <div style="display:grid;grid-template-columns:55px 2.8fr 1fr 1fr 0.8fr 0.7fr 3.6fr;align-items:center;background:{grp['row_bg']};padding:8px 10px;border-radius:6px;margin:4px 0 6px 0;font-size:12.5px;font-weight:700;color:#0F172A">
                        <div>⌄ &nbsp;<span style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:5px;background:{grp['icon_bg']};color:white;font-size:11px">▤</span></div>
                        <div>{html.escape(grp['group_title'])}</div>
                        <div>{grp['start']}</div>
                        <div>{grp['end']}</div>
                        <div>{grp['dur']}</div>
                        <div>{grp['pct']}%</div>
                        <div style="padding-right:16px">
                            <div style="height:11px;background:#CBD5E1;border-radius:99px;overflow:hidden;max-width:340px">
                                <div style="width:{grp['pct']}%;height:100%;background:{grp['bar_col']};border-radius:99px"></div>
                            </div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

            # Individual Task Rows
            for t_item in filtered_tasks:
                wbs_id = t_item["wbs"]
                rc = st.columns(col_weights)
                with rc[0]:
                    _task_checked = st.checkbox("Completed", value=(str(t_item["status"]).lower() == "completed"), key=f"tl_complete_{t_item['id']}", label_visibility="collapsed", disabled=bool(t_item.get("subtasks")), help="Completion is calculated from checked subtasks." if t_item.get("subtasks") else "Mark this task complete.")
                    if _task_checked != (str(t_item["status"]).lower() == "completed"):
                        sql("UPDATE tasks SET status=?, progress=? WHERE id=?", ("Completed" if _task_checked else "Not Started", 100 if _task_checked else 0, int(t_item["id"])))
                        log_event("Marked completed" if _task_checked else "Reopened", "Task", int(t_item["id"]), t_item["name"])
                        st.rerun()
                    st.caption(str(wbs_id))
                with rc[1]:
                    t_ico = "▤" if t_item["ico"] == "doc" else "▶"
                    st.markdown(
                        clean_html(
                            f"""
                            <div style="display:flex;align-items:center;gap:7px;padding-top:4px">
                                <span style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:4px;background:{t_item['ic_col']};color:white;font-size:9px">{t_ico}</span>
                                <span style="font-size:12.5px;font-weight:600;color:#0F172A">{html.escape(t_item['name'])}</span>
                            </div>
                            """
                        ),
                        unsafe_allow_html=True,
                    )
                with rc[2]:
                    st.markdown(f"<div style='font-size:12px;color:#64748B;padding-top:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{html.escape(t_item['project'])}</div>", unsafe_allow_html=True)
                with rc[3]:
                    st.markdown(f"<div style='font-size:12px;color:#1E293B;padding-top:5px'>{t_item['start_date']}</div>", unsafe_allow_html=True)
                with rc[4]:
                    st.markdown(f"<div style='font-size:12px;color:#1E293B;padding-top:5px'>{t_item['end_date']}</div>", unsafe_allow_html=True)
                with rc[5]:
                    st.markdown(f"<div style='font-size:12px;color:#1E293B;padding-top:5px'>{t_item['duration']}</div>", unsafe_allow_html=True)
                with rc[6]:
                    st.markdown(f"<div style='padding-top:3px'>{badge_html(t_item['status'])}</div>", unsafe_allow_html=True)
                with rc[7]:
                    st.markdown(f"<div style='padding-top:3px'>{badge_html(t_item['priority'])}</div>", unsafe_allow_html=True)
                with rc[8]:
                    ini = "".join(part[0] for part in str(t_item.get("assignee") or "").split()[:2]).upper() or "?"
                    abg = "#0284C7"
                    st.markdown(
                        clean_html(
                            f"""
                            <div style="display:flex;align-items:center;padding-top:3px;font-size:12px;color:#0F172A;white-space:nowrap">
                                <span class="vk-avatar" style="background:{abg}">{ini}</span>{html.escape(t_item['assignee'])}
                            </div>
                            """
                        ),
                        unsafe_allow_html=True,
                    )
                with rc[9]:
                    if t_item.get("subtasks"):
                        pct_v = int(t_item["pct"])
                        st.markdown(
                            clean_html(
                                f"""
                                <div style="display:flex;align-items:center;gap:8px;padding-top:6px">
                                    <span style="font-size:12px;font-weight:600;color:#1E293B;min-width:30px">{pct_v}%</span>
                                    <div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden">
                                        <div style="width:{pct_v}%;height:100%;background:{t_item['bar_col']};border-radius:99px"></div>
                                    </div>
                                </div>
                                """
                            ),
                            unsafe_allow_html=True,
                        )
                    else:
                        st.markdown("<div style='padding-top:6px;color:#94A3B8;font-size:12px'>—</div>", unsafe_allow_html=True)
                with rc[10]:
                    st.markdown(
                        f"<div style='padding-top:4px'><span class='vk-tag' style='background:{t_item['tag_bg']};color:{t_item['tag_fg']}'>{html.escape(t_item['tag'])}</span></div>",
                        unsafe_allow_html=True,
                    )
                with rc[11]:
                    b_v, b_e = st.columns(2)
                    with b_v:
                        if st.button("👁", key=f"tl_v_{wbs_id}", help="View Task"):
                            st.session_state.task_modal = ("view", t_item)
                            st.rerun()
                    with b_e:
                        if st.button("❐", key=f"tl_e_{wbs_id}", help="Edit Task"):
                            for _key in ("ed_t_name", "ed_t_proj", "ed_t_start", "ed_t_end", "ed_t_dur", "ed_t_pct", "ed_t_prio", "ed_t_stat", "ed_t_ass", "ed_dep_s", "ed_p_notes"):
                                st.session_state.pop(_key, None)
                            st.session_state.pop("_ed_sub_task_id", None)
                            st.session_state.task_modal = ("edit", t_item)
                            st.rerun()
                st.markdown("<hr style='margin:4px 0;border-color:#F1F5F9'>", unsafe_allow_html=True)

        st.caption(f"Task list contains {len(T)} task record(s) before filters")

    # Trigger Task Modal if active
    if "task_modal" in st.session_state:
        tm_type, tm_payload = st.session_state.task_modal
        if tm_type == "add":
            add_task_modal()
        elif tm_type == "view":
            view_task_modal(tm_payload)
        elif tm_type == "edit":
            edit_task_modal(tm_payload)


# ===========================================================================
# PAGE 5: RESOURCES (Exact Replica of Main View + Calendar, Assignments, Add & 7 Action Modals)
# ===========================================================================
elif page == "Resources":

    def _add_custom_resource_skill() -> None:
        _skill_value = str(st.session_state.get("ar_custom_skill", "")).strip()
        if _skill_value:
            _selected_skills = list(st.session_state.get("ar_skills", []))
            if _skill_value not in _selected_skills:
                _selected_skills.append(_skill_value)
            st.session_state["ar_skills"] = _selected_skills
            st.session_state["ar_custom_skill"] = ""

    @st.dialog("Add Resource", width="large")
    def add_resource_modal() -> None:
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:14px">
                    <div style="width:40px;height:40px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700">👥+</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Add Resource</div>
                        <div style="font-size:12.5px;color:#64748B">Add a new team member to the resource pool.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        top_l, top_r = st.columns([2.1, 1])
        with top_l, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>👤 &nbsp;Basic Information</div>", unsafe_allow_html=True)
            b1, b2 = st.columns(2)
            with b1:
                r_name = st.text_input("Full Name *", placeholder="Enter full name", key="ar_name")
            with b2:
                r_email = st.text_input("Email Address *", placeholder="name@hpe.com", key="ar_email")
            b3, b4, b5 = st.columns(3)
            with b3:
                r_role = st.selectbox(
                    "Role *",
                    ["Select role", "UX Designer", "Business Analyst", "Frontend Developer", "Project Manager", "Backend Developer", "QA Engineer", "Content Specialist", "DevOps Engineer", "IT Support", "Data Analyst"],
                    key="ar_role",
                )
            with b4:
                st.selectbox("Employment Type *", ["Employee", "Contractor", "Part-Time"], key="ar_emp_type")
            with b5:
                r_stat = st.selectbox("Status *", ["Available", "Allocated", "On Leave"], key="ar_stat")
            r_team = st.selectbox("Team *", sorted(set(teams_list)) or ["Others"], key="ar_team")

        with top_r, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:8px">👤 &nbsp;Profile Photo <span style="font-weight:400;color:#64748B;font-size:12px">(Optional)</span></div>
                    <div style="text-align:center;padding:6px 0">
                        <div style="width:74px;height:74px;border-radius:50%;background:#94A3B8;color:white;display:inline-flex;align-items:center;justify-content:center;font-size:34px;margin-bottom:10px">👤</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            r_photo = st.file_uploader("Upload Photo", type=["jpg", "jpeg", "png"], key="ar_photo", label_visibility="collapsed")
            if r_photo is not None:
                if len(r_photo.getvalue()) > 2 * 1024 * 1024:
                    st.error("Profile photos must be 2 MB or smaller.")
                else:
                    st.image(r_photo, width=96)
            st.markdown(
                "<div style='text-align:center;font-size:11px;color:#64748B;line-height:1.4'>JPG, PNG (Max 2MB)<br>Recommended size: 200 x 200</div>",
                unsafe_allow_html=True,
            )

        mid_l, mid_r = st.columns([1.15, 1])
        with mid_l, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>⚙ &nbsp;Skills</div>", unsafe_allow_html=True)
            skill_pool = ["UI/UX", "Figma", "React", "JavaScript", "Python", "API", "Testing", "Automation", "Content", "SEO", "Azure", "CI/CD", "Analysis", "Requirements", "Planning", "Stakeholder", "Infrastructure", "Support", "Power BI", "Excel"]
            st.multiselect("Select Skills *", skill_pool, default=[], placeholder="🔍 Search or select skills...", key="ar_skills")
            st.text_input("Add Custom Skill", placeholder="Enter a skill and press Add", key="ar_custom_skill")
            st.button("＋ Add Custom Skill", key="ar_custom_skill_btn", use_container_width=True, on_click=_add_custom_resource_skill)

        with mid_r, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>💼 &nbsp;Project Assignment</div>", unsafe_allow_html=True)
            st.multiselect("Assign to Project(s)", projects_list, placeholder="Select project(s)", key="ar_projs")
            al_c1, al_c2 = st.columns([3.5, 0.7])
            with al_c1:
                r_alloc = st.number_input("Allocation (%) ⓘ", min_value=0, max_value=200, value=100, step=10, key="ar_alloc")
            with al_c2:
                st.markdown("<div style='margin-top:28px;border:1px solid #CBD5E1;border-radius:6px;padding:7px 10px;text-align:center;font-weight:600;background:#F8FAFC'>%</div>", unsafe_allow_html=True)
            st.markdown(
                clean_html(
                    """
                    <div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;padding:10px 12px;margin-top:8px;display:flex;align-items:center;gap:10px;font-size:12px;color:#1D4ED8">
                        <span style="font-size:16px">ℹ</span>
                        <span>You can assign this resource to multiple projects and set the allocation for each project.</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        bot_l, bot_r = st.columns([1.15, 1])
        with bot_l, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>📅 &nbsp;Availability</div>", unsafe_allow_html=True)
            av1, av2 = st.columns(2)
            with av1:
                st.date_input("Availability Start Date *", date.today(), key="ar_av_start")
            with av2:
                st.date_input("Availability End Date", value=None, key="ar_av_end")
            av3, av4 = st.columns(2)
            with av3:
                st.number_input("Working Hours (per day)", min_value=1, max_value=24, value=8, key="ar_hrs")
            with av4:
                r_cap = st.number_input("Default Workload Capacity (%)", min_value=0, max_value=200, value=100, step=10, key="ar_cap")

        with bot_r, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>📄 &nbsp;Additional Information</div>", unsafe_allow_html=True)
            r_notes = st.text_area(
                "Notes (Optional)",
                placeholder="Enter additional notes about this resource (e.g. expertise, location, etc.)",
                height=105,
                key="ar_notes",
            )
            st.markdown(f"<div style='text-align:right;font-size:11px;color:#64748B;margin-top:-6px'>{len(r_notes)}/500</div>", unsafe_allow_html=True)

        f_sp, f_can, f_add = st.columns([3.6, 0.9, 1.2])
        with f_can:
            if st.button("Cancel", key="ar_cancel_btn", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()
        with f_add:
            if st.button("👥+ Add Resource", type="primary", key="ar_submit_btn", use_container_width=True):
                if not r_name.strip():
                    st.error("Enter a resource name before saving.")
                elif not r_email.strip() or "@" not in r_email:
                    st.error("Enter a valid email address.")
                elif st.session_state.get("ar_av_end") and st.session_state.get("ar_av_start") and st.session_state.get("ar_av_end") < st.session_state.get("ar_av_start"):
                    st.error("Availability end date must be on or after the start date.")
                else:
                    role_clean = "Team Member" if r_role == "Select role" else r_role
                    email_clean = r_email.strip()
                    if r_photo and len(r_photo.getvalue()) > 2 * 1024 * 1024:
                        st.error("Profile photos must be 2 MB or smaller.")
                        return
                    _photo_b64 = base64.b64encode(r_photo.getvalue()).decode("ascii") if r_photo else None
                    _availability_start = st.session_state.get("ar_av_start")
                    _availability_end = st.session_state.get("ar_av_end")
                    sql("INSERT INTO resources(name,email,team,role,status,capacity,notes,skills,availability_start,availability_end,working_hours,employment_type,photo_b64) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", (r_name.strip(), email_clean, r_team, role_clean, r_stat, int(r_cap), r_notes, ", ".join(st.session_state.get("ar_skills", [])), _availability_start.isoformat() if _availability_start else None, _availability_end.isoformat() if _availability_end else None, int(st.session_state.get("ar_hrs", 8)), st.session_state.get("ar_emp_type", "Employee"), _photo_b64))
                    with conn() as _c:
                        _rid = int(_c.execute("SELECT id FROM resources WHERE email=?", (email_clean,)).fetchone()[0])
                    for _project_name in st.session_state.get("ar_projs", []):
                        _pid = pmap.get(_project_name)
                        if _pid is not None:
                            _proj_row = P[P.id == _pid]
                            _end_date = str(_proj_row.iloc[0].end_date or "") if not _proj_row.empty else ""
                            sql("INSERT INTO project_assignments(resource_id,project_id,allocation,start_date,end_date,status,role,notes) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(resource_id,project_id) DO UPDATE SET allocation=excluded.allocation,start_date=excluded.start_date,end_date=excluded.end_date,status=excluded.status,role=excluded.role", (_rid, int(_pid), int(r_alloc), _availability_start.isoformat() if _availability_start else date.today().isoformat(), _end_date, "Assigned", role_clean, ""))
                    log_event("Created", "Resource", _rid, r_name.strip())
                    st.session_state.pop("res_modal", None)
                    st.rerun()

    @st.dialog("Resource Calendar", width="large")
    def resource_calendar_modal(res_rec: dict) -> None:
        rname = str(res_rec.get("name", "Unassigned"))
        month_key = f"rc_calendar_month_{res_rec.get('id', rname)}"
        date_key = f"rc_selected_date_{res_rec.get('id', rname)}"
        if month_key not in st.session_state: st.session_state[month_key] = date.today().replace(day=1)
        if date_key not in st.session_state: st.session_state[date_key] = date.today()
        month = st.session_state[month_key]; selected = st.session_state[date_key]
        st.markdown(clean_html(f"<div style='font-size:20px;font-weight:800;color:#0F172A'>Resource Calendar</div><div style='font-size:12.5px;color:#64748B;margin-bottom:10px'>{html.escape(rname)} · Scheduled tasks, milestones, and events</div>"),unsafe_allow_html=True)
        nav1, nav2, nav3, nav4 = st.columns([0.5,2.0,0.5,0.8])
        with nav1:
            if st.button("‹", key=f"rc_prev_{res_rec.get('id')}", use_container_width=True): st.session_state[month_key]=(month-timedelta(days=1)).replace(day=1); st.rerun()
        with nav2: st.markdown(f"<div style='text-align:center;font-size:15px;font-weight:800;padding-top:7px'>{month.strftime('%B %Y')}</div>",unsafe_allow_html=True)
        with nav3:
            if st.button("›", key=f"rc_next_{res_rec.get('id')}", use_container_width=True): st.session_state[month_key]=(month.replace(day=28)+timedelta(days=4)).replace(day=1); st.rerun()
        with nav4:
            if st.button("Today", key=f"rc_today_{res_rec.get('id')}", use_container_width=True): st.session_state[month_key]=date.today().replace(day=1); st.session_state[date_key]=date.today(); st.rerun()
        items=[]
        resource_tasks=df("SELECT t.*,p.name AS project FROM tasks t LEFT JOIN projects p ON p.id=t.project_id WHERE t.assignee=?",(rname,))
        for _, row in resource_tasks.iterrows():
            d=pd.to_datetime(row.get("due_date"),errors="coerce")
            if pd.notna(d): items.append({"date":d.date(),"title":str(row.get("name") or "Task"),"sub":f"Task · {row.get('project') or ''}","tone":"g"})
        resource_events=df("SELECT e.*,p.name AS project FROM calendar_events e LEFT JOIN projects p ON p.id=e.project_id WHERE instr(coalesce(e.resources,''),?)>0",(rname,))
        for _, row in resource_events.iterrows():
            d=pd.to_datetime(row.get("start_date"),errors="coerce")
            if pd.notna(d):
                typ=str(row.get("event_type") or "Meeting / Event"); items.append({"date":d.date(),"title":str(row.get("title") or "Event"),"sub":f"{typ} · {row.get('start_time') or ''}","tone":"r" if typ=="Leave" else "p"})
        resource_assignments=df("SELECT pa.*,p.name AS project FROM project_assignments pa LEFT JOIN projects p ON p.id=pa.project_id WHERE pa.resource_id=?",(int(res_rec.get("id") or 0),))
        for _, row in resource_assignments.iterrows():
            d=pd.to_datetime(row.get("start_date"),errors="coerce")
            if pd.notna(d): items.append({"date":d.date(),"title":str(row.get("project") or "Project Assignment"),"sub":f"Project Assignment · {int(row.get('allocation') or 0)}% allocation","tone":"g"})
        first=month.replace(day=1); start_day=first-timedelta(days=(first.weekday()+1)%7); n=((first.weekday()+1)%7+calendar.monthrange(month.year,month.month)[1]+6)//7*7
        days=[start_day+timedelta(days=i) for i in range(n)]; tones={"g":("#DCFCE7","#15803D"),"p":("#F3E8FF","#7C3AED"),"r":("#FEE2E2","#B91C1C")}
        rows_html=""
        for i in range(0,len(days),7):
            cells=""
            for d in days[i:i+7]:
                selected_cell=d==selected; bg="#EFF6FF" if selected_cell else ("#FFFFFF" if d.month==month.month else "#F8FAFC"); border="2px solid #3B82F6" if selected_cell else "1px solid #EDF2F7"
                pills=""
                for ev in [x for x in items if x["date"]==d][:2]:
                    tone=tones[ev["tone"]]; pills+=f'<div style="background:{tone[0]};color:{tone[1]};border-radius:5px;padding:3px 5px;font-size:9.5px;margin-top:3px">{html.escape(ev["title"])}</div>'
                cells+=f'<div style="min-height:76px;border:{border};background:{bg};padding:5px"><b>{d.day}</b>{pills}</div>'
            rows_html+=f'<div style="display:grid;grid-template-columns:repeat(7,1fr)">{cells}</div>'
        st.markdown(clean_html(f'<div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;background:white"><div style="display:grid;grid-template-columns:repeat(7,1fr);background:#F8FAFC;text-align:center;padding:7px 0;font-weight:700">Sun Mon Tue Wed Thu Fri Sat</div>{rows_html}</div>'),unsafe_allow_html=True)
        day_options=days; selected=selected if selected in day_options else date.today()
        chosen=st.selectbox("Selected date",day_options,index=day_options.index(selected),format_func=lambda d:d.strftime("%A, %B %d, %Y"),key=f"rc_inspect_{res_rec.get('id')}")
        if chosen!=st.session_state[date_key]: st.session_state[date_key]=chosen; st.rerun()
        chosen_items=[x for x in items if x["date"]==chosen]
        st.markdown("**Scheduled items**")
        if not chosen_items: st.caption("No scheduled items for this resource on the selected date.")
        for item in chosen_items: st.markdown(f"**{item['title']}**"); st.caption(item["sub"])
        close_col, add_col = st.columns([1,1])
        with close_col:
            if st.button("Close",key=f"rc_close_{res_rec.get('id')}",use_container_width=True): st.session_state.pop("res_modal",None); st.rerun()
        with add_col:
            if st.button("＋ Add Event",key=f"rc_add_{res_rec.get('id')}",type="primary",use_container_width=True):
                st.session_state["cal_selected_date"]=chosen; st.session_state["cal_calendar_month"]=chosen.replace(day=1); st.session_state["page"]="Calendar"; st.session_state["cal_modal"]="add_event"; st.session_state.pop("cal_edit_event_id",None)
                for _key in ("ae_type", "ae_title", "ae_proj", "ae_task", "ae_res", "ae_sd", "ae_st", "ae_ed", "ae_et", "ae_alloc", "ae_rep", "ae_desc"): st.session_state.pop(_key,None)
                st.rerun()

    @st.dialog("Project Assignments", width="large")
    def project_assignments_modal(res_rec: dict) -> None:
        rname = res_rec.get("name", "Unassigned")
        rrole = res_rec.get("role", "Team Member")
        remail = res_rec.get("email", "")
        rini = res_rec.get("ini", "U")
        rcolor = res_rec.get("av_col", "#0284C7")
        ralloc = int(res_rec.get("alloc", 100))
        rstat = res_rec.get("status", "Allocated")

        st.markdown(
            clean_html(
                f"""
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:12px">
                    <div style="width:40px;height:40px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700">👥</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Project Assignments</div>
                        <div style="font-size:12.5px;color:#64748B">View and manage project assignments for this resource.</div>
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 18px;background:#F8FAFC;display:grid;grid-template-columns:1.8fr 1fr 1fr 1fr 1.2fr;align-items:center;gap:12px;margin-bottom:14px">
                    <div style="display:flex;align-items:center;gap:12px">
                        <div style="width:46px;height:46px;border-radius:50%;background:{rcolor};color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px">{rini}</div>
                        <div>
                            <div style="font-size:15px;font-weight:800;color:#0F172A">{html.escape(rname)}</div>
                            <div style="font-size:12px;color:#475569">{html.escape(rrole)}</div>
                            <div style="font-size:11.5px;color:#64748B">{html.escape(remail)}</div>
                        </div>
                    </div>
                    <div style="border-left:1px solid #E2E8F0;padding-left:14px">
                        <div style="font-size:11px;color:#64748B;margin-bottom:3px">Availability Status</div>
                        {badge_html(rstat)}
                    </div>
                    <div style="border-left:1px solid #E2E8F0;padding-left:14px">
                        <div style="font-size:11px;color:#64748B">Working Hours</div>
                        <div style="font-size:15px;font-weight:800;color:#0F172A">8 hrs/day</div>
                    </div>
                    <div style="border-left:1px solid #E2E8F0;padding-left:14px">
                        <div style="font-size:11px;color:#64748B">Default Capacity</div>
                        <div style="font-size:15px;font-weight:800;color:#0F172A">100%</div>
                    </div>
                    <div style="border-left:1px solid #E2E8F0;padding-left:14px">
                        <div style="font-size:11px;color:#64748B">Total Allocation</div>
                        <div style="font-size:15px;font-weight:800;color:#0F172A">{ralloc}%</div>
                        <div style="height:7px;background:#E2E8F0;border-radius:99px;overflow:hidden;margin-top:3px"><div style="width:{min(100, ralloc)}%;height:100%;background:#00875A"></div></div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        pa_h1, pa_h2 = st.columns([3.5, 1.1])
        with pa_h1:
            st.markdown(f"<div style='font-size:14.5px;font-weight:700;color:#0F172A;padding-top:6px'>💼 &nbsp;Current Project Assignments ({len(_assigned_project_records)})</div>", unsafe_allow_html=True)
        with pa_h2:
            if st.button("＋ Assign to Project", type="primary", key="pa_assign_btn", use_container_width=True):
                st.session_state.res_modal = ("assign", res_rec)
                st.rerun()

        _assigned_tasks = T[T.assignee.fillna("").astype(str) == rname].copy() if not T.empty else pd.DataFrame()
        _resource_row = R[R.name.astype(str) == rname] if not R.empty else pd.DataFrame()
        _resource_id = int(_resource_row.iloc[0].id) if not _resource_row.empty else -1
        _assigned_project_records = df("SELECT pa.*,p.name AS project FROM project_assignments pa LEFT JOIN projects p ON p.id=pa.project_id WHERE pa.resource_id=? ORDER BY pa.start_date", (_resource_id,))
        _assigned_projects = []
        for _, _assignment in _assigned_project_records.iterrows():
            _project_tasks = _assigned_tasks[_assigned_tasks.project_id == _assignment.project_id] if not _assigned_tasks.empty else pd.DataFrame()
            _task_names = ", ".join(_project_tasks.name.dropna().astype(str).head(3).tolist()) if not _project_tasks.empty else "No tasks assigned yet"
            _units_total = _units_done = 0
            for _, _task in _project_tasks.iterrows():
                _sub_counts = df("SELECT COUNT(*) AS total, COALESCE(SUM(CASE WHEN lower(CAST(completed AS TEXT)) IN ('1','true','yes','completed') THEN 1 ELSE 0 END),0) AS done FROM subtasks WHERE task_id=?", (int(_task.id),))
                _sub_total = int(_sub_counts.iloc[0]["total"] or 0) if not _sub_counts.empty else 0
                _sub_done = int(_sub_counts.iloc[0]["done"] or 0) if not _sub_counts.empty else 0
                if _sub_total:
                    _units_total += _sub_total
                    _units_done += _sub_done
                else:
                    _units_total += 1
                    _units_done += int(str(_task.status or "").lower() == "completed")
            _completion = round(_units_done / _units_total * 100) if _units_total else 0
            _assigned_projects.append((str(_assignment.get("project") or "Unassigned"), str(_assignment.get("role") or rrole), int(_assignment.get("allocation") or 0), _completion, fmt_date(_assignment.get("start_date")), fmt_date(_assignment.get("end_date")), str(_assignment.get("status") or "Assigned"), _task_names))
        _assignment_html = "".join(f"<tr><td><b>{html.escape(prj)}</b></td><td>{html.escape(role)}</td><td>{progress_html(alloc, GREEN)}</td><td>{progress_html(completion, BLUE) if task_name != 'No tasks assigned yet' else '—'}</td><td>{sd}</td><td>{ed}</td><td>{badge_html(status)}</td><td>{html.escape(task_name)}</td><td style='text-align:center'>—</td></tr>" for prj,role,alloc,completion,sd,ed,status,task_name in _assigned_projects)
        st.markdown(clean_html(f"""<div class="vk-table-wrap" style="margin-top:6px;margin-bottom:14px"><table class="vk-table"><thead><tr><th>Project</th><th>Role</th><th>Allocation</th><th>Task Progress</th><th>Start Date</th><th>End Date</th><th>Status</th><th>Primary Task(s)</th><th style="text-align:center">Actions</th></tr></thead><tbody>{_assignment_html}</tbody></table></div>"""),unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("<div style='font-size:14.5px;font-weight:700;color:#0F172A;margin-bottom:6px'>🕒 &nbsp;Workload Overview</div>",unsafe_allow_html=True)
            _assigned_count=len(_assigned_tasks); _assigned_projects_count=len(set(_assigned_tasks.project.dropna().astype(str))) if not _assigned_tasks.empty else 0
            wa,wb,wc=st.columns(3)
            with wa: metric_card("Assigned Tasks",_assigned_count,"Tasks assigned to this resource","📋","blue")
            with wb: metric_card("Projects",_assigned_projects_count,"Distinct active projects","📁","green")
            with wc: metric_card("Availability",f"{max(0,100-ralloc)}%","Available capacity","👤","orange")

        pa_tabs=st.tabs([f"📅 Assigned Tasks ({len(_assigned_tasks)})", "☆ Milestones", "🕒 Time Allocation"])
        with pa_tabs[0]:
            if _assigned_tasks.empty:
                st.info("No tasks are assigned to this resource in the current data.")
            else:
                _view=_assigned_tasks.copy()
                _view["Start Date"]=_view.start_date.map(fmt_date)
                _view["Due Date"]=_view.due_date.map(fmt_date)
                _view["Progress"]=_view.progress.fillna(0).astype(int).astype(str)+"%"
                st.dataframe(_view[["id","name","project","Start Date","Due Date","status","priority","Progress"]].rename(columns={"id":"ID","name":"Task Name","project":"Project","status":"Status","priority":"Priority"}),use_container_width=True,hide_index=True)
        with pa_tabs[1]:
            _assigned_project_ids=set(_assigned_tasks.project_id.dropna().astype(int).tolist()) if not _assigned_tasks.empty else set()
            _linked_milestones=M[M.project_id.isin(_assigned_project_ids)] if not M.empty else pd.DataFrame()
            if _linked_milestones.empty: st.info("No milestones are linked to this resource's assigned projects.")
            else: st.dataframe(_linked_milestones[["name","project","status","due_date"]].rename(columns={"name":"Milestone","project":"Project","status":"Status","due_date":"Due Date"}),use_container_width=True,hide_index=True)
        with pa_tabs[2]:
            st.caption(f"Assigned task count: {_assigned_count}. Allocation shown here is based on the resource's current capacity field.")

        f_sp, f_cl = st.columns([4.5, 0.8])
        with f_cl:
            if st.button("Close", key="pa_close_btn", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()

    @st.dialog("Resource Action")
    def resource_action_submodal(act_mode: str, res_rec: dict) -> None:
        rname = res_rec.get("name", "Unassigned")
        rrole = res_rec.get("role", "Team Member")
        remail = res_rec.get("email", "")
        rini = res_rec.get("ini", "U")
        rcolor = res_rec.get("av_col", "#0284C7")
        rskills = res_rec.get("skills", ["UI/UX", "Figma"])
        rproj = res_rec.get("project", "Unassigned")
        ralloc = int(res_rec.get("alloc", 100))

        header_titles = {
            "view": ("👁", "View Resource"),
            "edit": ("👥", "Edit Resource"),
            "assign": ("👥", "Assign to Project"),
            "skills": ("⚙", "Manage Skills"),
            "availability": ("📅", "Change Availability"),
            "leave": ("👤", "Set on Leave"),
            "delete": ("🗑", "Delete Resource"),
        }
        ico, h_title = header_titles.get(act_mode, ("👤", "Resource Details"))

        if act_mode == "delete":
            st.markdown(
                clean_html(
                    f"""
                    <div style="padding:10px 6px">
                        <div style="width:52px;height:52px;border-radius:50%;background:#EF4444;color:white;display:flex;align-items:center;justify-content:center;font-size:24px;margin-bottom:12px">🗑</div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;margin-bottom:6px">Delete Resource</div>
                        <div style="font-size:13.5px;color:#334155;line-height:1.5;margin-bottom:18px">
                            Are you sure you want to delete <b>{html.escape(rname)}</b>?<br>
                            <span style="color:#DC2626;font-weight:600">This action cannot be undone.</span>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_del_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Delete", type="primary", key="sub_del_ok", use_container_width=True):
                    sql("DELETE FROM project_assignments WHERE resource_id=(SELECT id FROM resources WHERE name=?)", (rname,))
                    sql("DELETE FROM resources WHERE name=?", (rname,))
                    log_event("Deleted", "Resource", None, rname)
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            return

        # Common Mini Header + Avatar Banner for the other 6 Sub-Modals
        st.markdown(
            clean_html(
                f"""
                <div style="display:flex;align-items:center;gap:8px;padding-bottom:10px;border-bottom:1px solid #E2E8F0;margin-bottom:12px">
                    <span style="width:26px;height:26px;border-radius:50%;background:#0284C7;color:white;display:inline-flex;align-items:center;justify-content:center;font-size:13px">{ico}</span>
                    <span style="font-size:16px;font-weight:800;color:#0F172A">{h_title}</span>
                </div>
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">
                    <div style="width:38px;height:38px;border-radius:50%;background:{rcolor};color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:14px">{rini}</div>
                    <div>
                        <div style="font-size:14px;font-weight:800;color:#0F172A">{html.escape(rname)}</div>
                        <div style="font-size:11.5px;color:#64748B">{html.escape(rrole)}</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        if act_mode == "view":
            sk_badges = "".join(f'<span class="vk-tag" style="background:#DBEAFE;color:#1D4ED8">{html.escape(s)}</span>' for s in rskills)
            st.markdown(
                clean_html(
                    f"""
                    <div style="font-size:12.5px;line-height:2.0;margin-bottom:12px">
                        <div style="font-weight:700;color:#0F172A;margin-bottom:2px">Basic Information</div>
                        <div style="display:grid;grid-template-columns:110px 1fr"><span style="color:#64748B">Email</span><span>{html.escape(remail)}</span></div>
                        <div style="display:grid;grid-template-columns:110px 1fr"><span style="color:#64748B">Role</span><span>{html.escape(rrole)}</span></div>
                        <div style="display:grid;grid-template-columns:110px 1fr"><span style="color:#64748B">Status</span><span style="color:#1473E6;font-weight:700">{html.escape(res_rec.get('status', 'Allocated'))}</span></div>
                        <div style="display:grid;grid-template-columns:110px 1fr"><span style="color:#64748B">Employment Type</span><span>Employee</span></div>
                        <div style="font-weight:700;color:#0F172A;margin-top:8px;margin-bottom:4px">Skills</div>
                        <div>{sk_badges}</div>
                        <div style="font-weight:700;color:#0F172A;margin-top:8px;margin-bottom:2px">Current Projects</div>
                        <div>{html.escape(rproj)} ({ralloc}%)</div>
                        <div style="font-weight:700;color:#0F172A;margin-top:8px;margin-bottom:2px">Working Hours</div>
                        <div>8 hrs/day</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            if st.button("Close", key="sub_v_close", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()

        elif act_mode == "edit":
            new_nm = st.text_input("Full Name *", value=rname, key="sub_ed_nm")
            new_em = st.text_input("Email Address *", value=remail, key="sub_ed_em")
            roles_opt = sorted(set(R.role.dropna().astype(str).tolist() + [rrole])) or ["Team Member"]
            new_rl = st.selectbox("Role *", roles_opt, index=roles_opt.index(rrole) if rrole in roles_opt else 0, key="sub_ed_rl")
            new_team = st.selectbox("Team", sorted(set(teams_list + [str(res_rec.get("team") or "Others")])) or ["Others"], index=(sorted(set(teams_list + [str(res_rec.get("team") or "Others")])).index(str(res_rec.get("team") or "Others")) if str(res_rec.get("team") or "Others") in sorted(set(teams_list + [str(res_rec.get("team") or "Others")])) else 0), key="sub_ed_team")
            new_status = st.selectbox("Status", ["Available", "Partial", "Allocated", "On Leave", "Inactive"], index=(["Available", "Partial", "Allocated", "On Leave", "Inactive"].index(str(res_rec.get("status") or "Available")) if str(res_rec.get("status") or "Available") in ["Available", "Partial", "Allocated", "On Leave", "Inactive"] else 0), key="sub_ed_status")
            new_capacity = st.number_input("Capacity (%)", min_value=0, max_value=200, value=max(0, min(200, int(res_rec.get("capacity") or 0))), key="sub_ed_capacity")
            new_skills = st.text_input("Skills (comma-separated)", value=str(res_rec.get("skills") or ""), key="sub_ed_skills")
            new_notes = st.text_area("Notes", value=str(res_rec.get("notes") or ""), key="sub_ed_notes")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_ed_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Save Changes", type="primary", key="sub_ed_sav", use_container_width=True):
                    sql("UPDATE resources SET name=?, email=?, role=?, team=?, status=?, capacity=?, skills=?, notes=? WHERE name=?", (new_nm.strip(), new_em.strip(), new_rl, st.session_state.get("sub_ed_team", res_rec.get("team") or "Others"), st.session_state.get("sub_ed_status", res_rec.get("status") or "Available"), int(st.session_state.get("sub_ed_capacity", res_rec.get("capacity") or 100)), st.session_state.get("sub_ed_skills", res_rec.get("skills") or ""), st.session_state.get("sub_ed_notes", res_rec.get("notes") or ""), rname))
                    log_event("Updated", "Resource", None, new_nm.strip())
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "assign":
            st.markdown("<div style='font-size:12.5px;font-weight:700;color:#0F172A;margin-bottom:6px'>Select Project(s) *</div>", unsafe_allow_html=True)
            _assignment_choices = projects_list
            _selected_assignment_projects = st.multiselect("Projects", _assignment_choices, default=[rproj] if rproj in _assignment_choices else [], key="sub_as_projects")
            st.number_input("Allocation (%)", min_value=0, max_value=200, value=ralloc, step=10, key="sub_as_pct")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_as_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Assign", type="primary", key="sub_as_ok", use_container_width=True):
                    _chosen_projects = st.session_state.get("sub_as_projects", [])
                    _allocation_value = int(st.session_state.get("sub_as_pct", ralloc))
                    _resource_rows = R[R.name.astype(str) == rname] if not R.empty else pd.DataFrame()
                    if _resource_rows.empty:
                        st.error("Resource record was not found. Refresh the page and try again.")
                    else:
                        _rid = int(_resource_rows.iloc[0].id)
                        for _project_name in _chosen_projects:
                            _project_id = pmap.get(_project_name)
                            if _project_id is not None:
                                _project_row = P[P.id == _project_id]
                                _end = str(_project_row.iloc[0].end_date or "") if not _project_row.empty else ""
                                sql("INSERT INTO project_assignments(resource_id,project_id,allocation,start_date,end_date,status,role,notes) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(resource_id,project_id) DO UPDATE SET allocation=excluded.allocation,start_date=excluded.start_date,end_date=excluded.end_date,status=excluded.status,role=excluded.role", (_rid,int(_project_id),_allocation_value,date.today().isoformat(),_end,"Assigned",rrole,""))
                        sql("UPDATE resources SET capacity=? WHERE id=?", (max(0,100-_allocation_value),_rid))
                        log_event("Assigned", "Resource", _rid, rname + " -> " + ", ".join(_chosen_projects))
                        st.session_state.pop("res_modal", None)
                        st.rerun()

        elif act_mode == "skills":
            st.text_input("Select Skills", placeholder="🔍 Search skills...", key="sub_sk_q")
            sk_c1, sk_c2 = st.columns(2)
            with sk_c1:
                st.checkbox("UI/UX", value=True, key="sk_1")
                st.checkbox("React", value=False, key="sk_2")
                st.checkbox("API", value=False, key="sk_3")
                st.checkbox("Testing", value=False, key="sk_4")
            with sk_c2:
                st.checkbox("Figma", value=True, key="sk_5")
                st.checkbox("JavaScript", value=False, key="sk_6")
                st.checkbox("Testing ", value=False, key="sk_7")
                st.checkbox("Automation", value=False, key="sk_8")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_sk_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Save Skills", type="primary", key="sub_sk_ok", use_container_width=True):
                    selected_skills = [label for key,label in [("sk_1","UI/UX"),("sk_2","React"),("sk_3","API"),("sk_4","Testing"),("sk_5","Figma"),("sk_6","JavaScript"),("sk_8","Automation")] if st.session_state.get(key, False)]
                    sql("UPDATE resources SET skills=? WHERE name=?", (", ".join(selected_skills), rname))
                    log_event("Updated Skills", "Resource", None, rname + ": " + ", ".join(selected_skills))
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "availability":
            st.date_input("Availability Start Date *", date.today(), key="sub_av_st")
            st.date_input("Availability End Date", value=None, key="sub_av_en")
            av_a, av_b = st.columns(2)
            with av_a:
                st.number_input("Working Hours (per day)", min_value=1, max_value=24, value=8, key="sub_av_hr")
            with av_b:
                new_cap = st.number_input("Default Capacity (%)", min_value=0, max_value=200, value=100, step=10, key="sub_av_cp")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_av_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Update", type="primary", key="sub_av_ok", use_container_width=True):
                    sql("UPDATE resources SET capacity=?, availability_start=?, availability_end=?, working_hours=? WHERE name=?", (int(new_cap), st.session_state.get("sub_av_st").isoformat() if st.session_state.get("sub_av_st") else None, st.session_state.get("sub_av_en").isoformat() if st.session_state.get("sub_av_en") else None, int(st.session_state.get("sub_av_hr", 8)), rname))
                    log_event("Updated Availability", "Resource", None, rname)
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "leave":
            lv1, lv2 = st.columns(2)
            with lv1:
                st.date_input("Leave Start Date *", date.today(), key="sub_lv_st")
            with lv2:
                st.date_input("Leave End Date *", date.today()+timedelta(days=7), key="sub_lv_en")
            st.selectbox("Reason (Optional)", ["Vacation", "Sick Leave", "Personal Leave", "Training"], key="sub_lv_rs")
            st.text_input("Notes (Optional)", placeholder="Enter notes...", key="sub_lv_nt")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_lv_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Set on Leave", type="primary", key="sub_lv_ok", use_container_width=True):
                    _leave_start = st.session_state.get("sub_lv_st")
                    _leave_end = st.session_state.get("sub_lv_en")
                    if _leave_end and _leave_start and _leave_end < _leave_start:
                        st.error("Leave end date must be on or after the start date.")
                    else:
                        _resource_rows = R[R.name.astype(str) == rname] if not R.empty else pd.DataFrame()
                        _resource_project = int(_resource_rows.iloc[0].id) if not _resource_rows.empty else None
                        sql("UPDATE resources SET status='On Leave', capacity=0 WHERE name=?", (rname,))
                        sql("INSERT INTO calendar_events(title,event_type,resources,start_date,start_time,end_date,end_time,allocation,repeat_rule,description) VALUES(?,?,?,?,?,?,?,?,?,?)", (f"{rname} - Leave", "Leave", rname, _leave_start.isoformat() if _leave_start else date.today().isoformat(), "09:00 AM", _leave_end.isoformat() if _leave_end else (_leave_start or date.today()).isoformat(), "05:00 PM", 0, "Does not repeat", f"{st.session_state.get('sub_lv_rs','')} {st.session_state.get('sub_lv_nt','')}".strip()))
                        log_event("Set on Leave", "Resource", _resource_project, rname)
                        st.session_state.pop("res_modal", None)
                        st.rerun()

    # --- RESOURCES TOP HEADER ROW ---
    rh_l, rh_s, rh_d, rh_b = st.columns([1.9, 1.05, 0.95, 0.7])
    with rh_l:
        st.markdown(
            clean_html(
                f"""
                <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Resources</div>
                <div style='font-size:13.5px;color:#64748B;margin-top:2px'>Manage team members, roles, skills and workload across all projects.</div>
                """
            ),
            unsafe_allow_html=True,
        )
    with rh_s:
        res_top_q = st.text_input("Search", placeholder="🔍 Search team members, skills, or projects...", key="res_q_top", label_visibility="collapsed")
    with rh_d:
        st.date_input("Date Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="res_d_top", label_visibility="collapsed")
    with rh_b:
        if st.button("＋ Add Resource", type="primary", use_container_width=True, key="res_btn_add_top"):
            st.session_state.res_modal = ("add", None)
            st.rerun()

    st.write("")
    # --- 5 KPI CARDS ROW (Exact Match to Image 20) ---
    rk1, rk2, rk3, rk4, rk5 = st.columns(5)
    _resource_status_top = R.status.fillna("Available").astype(str).str.lower() if not R.empty else pd.Series(dtype=str)
    _total_resources_top = len(R)
    _available_top = int((_resource_status_top == "available").sum())
    _allocated_top = int(_resource_status_top.isin(["allocated", "partial"]).sum())
    _leave_top = int((_resource_status_top == "on leave").sum())
    _assignments_top = df("SELECT resource_id,SUM(allocation) AS allocation FROM project_assignments GROUP BY resource_id")
    _overallocated_top = int((pd.to_numeric(_assignments_top.get("allocation", pd.Series(dtype=float)), errors="coerce") > 100).sum()) if not _assignments_top.empty else 0
    _top_pct = lambda _n: f"({_n / _total_resources_top:.0%})" if _total_resources_top else "(0%)"
    res_kpis = [
        (rk1, "#1473E6", "👥", "Total Resources", "#334155", str(_total_resources_top), ""),
        (rk2, "#00875A", "✓", "Available", "#334155", str(_available_top), _top_pct(_available_top)),
        (rk3, "#1473E6", "🕒", "Allocated", "#334155", str(_allocated_top), _top_pct(_allocated_top)),
        (rk4, "#F28C00", "!", "Overallocated", "#DC2626", str(_overallocated_top), _top_pct(_overallocated_top)),
        (rk5, "#64748B", "👤", "On Leave", "#334155", str(_leave_top), _top_pct(_leave_top)),
    ]
    for col, ibg, ico, lbl, lcol, val, pct_s in res_kpis:
        with col:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="min-height:86px;padding:14px 16px">
                        <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:{ibg};color:white;display:flex;align-items:center;justify-content:center;font-size:19px;font-weight:800">{ico}</div>
                        <div>
                            <div style="font-size:12px;font-weight:600;color:{lcol}">{lbl}</div>
                            <div style="font-size:23px;font-weight:800;color:#0F172A;line-height:1.2">{val} <span style="font-size:18px;font-weight:800">{pct_s}</span></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

    st.write("")
    # --- 3 ANALYTICS CHARTS ROW (Exact Match to Image 20) ---
    rc_1, rc_2, rc_3 = st.columns([1.15, 1.05, 1.05])
    with rc_1, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
                    <span style="font-size:15px;font-weight:700;color:#0F172A">Resource Allocation</span>
                    <div style="font-size:11.5px;color:#475569;display:flex;gap:12px">
                        <span><span style="color:#00875A">●</span> Allocated</span>
                        <span><span style="color:#1473E6">●</span> Available</span>
                        <span><span style="color:#94A3B8">●</span> On Leave</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        _assignment_history = df("SELECT start_date, allocation FROM project_assignments")
        _leave_history = df("SELECT start_date FROM calendar_events WHERE lower(coalesce(event_type,''))='leave'")
        _month_starts = [((date.today().replace(day=1) - timedelta(days=1)).replace(day=1) - timedelta(days=30*i)).replace(day=1) for i in (2,1,0)]
        _ra_rows = []
        for _month_start in _month_starts:
            _next_month = (_month_start.replace(day=28) + timedelta(days=4)).replace(day=1)
            _a_dates = pd.to_datetime(_assignment_history.get("start_date", pd.Series(dtype=str)), errors="coerce") if not _assignment_history.empty else pd.Series(dtype="datetime64[ns]")
            _a_mask = (_a_dates.dt.date >= _month_start) & (_a_dates.dt.date < _next_month) if not _assignment_history.empty else pd.Series(dtype=bool)
            _allocated = round(pd.to_numeric(_assignment_history.loc[_a_mask, "allocation"], errors="coerce").fillna(0).sum()/100, 1) if not _assignment_history.empty else 0
            _l_dates = pd.to_datetime(_leave_history.get("start_date", pd.Series(dtype=str)), errors="coerce") if not _leave_history.empty else pd.Series(dtype="datetime64[ns]")
            _leave_count = int(((_l_dates.dt.date >= _month_start) & (_l_dates.dt.date < _next_month)).sum()) if not _leave_history.empty else 0
            _available = round(max(0, (pd.to_numeric(R.get("capacity", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()/100) - _allocated), 1) if not R.empty else 0
            _ra_rows.append({"Month": _month_start.strftime("%b %Y"), "Allocated": _allocated, "Available": _available, "On Leave": _leave_count})
        ra_df = pd.DataFrame(_ra_rows)
        fig_ra = px.bar(
            ra_df,
            x="Month",
            y=["Allocated", "Available", "On Leave"],
            color_discrete_map={"Allocated": GREEN, "Available": BLUE, "On Leave": "#94A3B8"},
        )
        fig_ra.update_layout(
            height=195,
            margin=dict(l=0, r=0, t=6, b=0),
            showlegend=False,
            xaxis_title=None,
            yaxis=dict(title=None, rangemode="tozero"),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            bargap=0.45,
        )
        st.plotly_chart(fig_ra, use_container_width=True, config={"displayModeBar": False})

    with rc_2, st.container(border=True):
        st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;margin-bottom:6px'>Utilization</div>", unsafe_allow_html=True)
        u_l, u_r = st.columns([1.15, 1])
        with u_l:
            _assignments_live = df("SELECT resource_id, allocation FROM project_assignments")
        _total_capacity = max(1, len(R) * 100)
        _allocated_pct = min(100, round(pd.to_numeric(_assignments_live.get("allocation", pd.Series(dtype=float)), errors="coerce").fillna(0).sum() / _total_capacity * 100)) if not _assignments_live.empty else 0
        _leave_pct = min(100 - _allocated_pct, round(_leave_top / max(1, len(R)) * 100)) if len(R) else 0
        _available_pct = max(0, 100 - _allocated_pct - _leave_pct)
        _util_values = [_allocated_pct, _available_pct, _leave_pct]
        fig_ut = go.Figure(go.Pie(values=_util_values, labels=["Allocated", "Available", "On Leave"], hole=0.66, sort=False, direction="clockwise", marker=dict(colors=[GREEN, BLUE, "#94A3B8"]), textinfo="none"))
        fig_ut.update_layout(height=195, margin=dict(l=4, r=4, t=4, b=4), showlegend=False, paper_bgcolor="rgba(0,0,0,0)", annotations=[dict(text=f"<b style='font-size:20px;color:#0F172A'>{_allocated_pct}%</b><br><span style='font-size:10.5px;color:#475569'>Allocated</span>", x=0.5, y=0.5, showarrow=False)])
        st.plotly_chart(fig_ut, use_container_width=True, config={"displayModeBar": False})
        with u_r:
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:12.5px;line-height:2.4;padding-top:32px">
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;Allocated</span><b>{_allocated_pct}%</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> &nbsp;Available</span><b>{_available_pct}%</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#94A3B8">●</span> &nbsp;On Leave</span><b>{_leave_pct}%</b></div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

    with rc_3, st.container(border=True):
        st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;margin-bottom:12px'>Resource Status</div>", unsafe_allow_html=True)
        _resource_status_norm = R.status.fillna("Available").astype(str).str.lower() if not R.empty else pd.Series(dtype=str)
        _alloc_by_resource = df("SELECT resource_id, SUM(allocation) AS allocation FROM project_assignments GROUP BY resource_id")
        _overallocated_ids = set(pd.to_numeric(_alloc_by_resource.loc[pd.to_numeric(_alloc_by_resource.get("allocation", pd.Series(dtype=float)), errors="coerce") > 100, "resource_id"], errors="coerce").dropna().astype(int)) if not _alloc_by_resource.empty else set()
        rs_bars = [
            ("Available", int((_resource_status_norm == "available").sum()), 0, "#6EE7B7"),
            ("Allocated", int(_resource_status_norm.isin(["allocated", "partial"]).sum()), 0, "#1473E6"),
            ("Overallocated", len(_overallocated_ids), 0, "#F87171"),
            ("On Leave", int((_resource_status_norm == "on leave").sum()), 0, "#E2E8F0"),
        ]
        _rs_max = max(1, max((x[1] for x in rs_bars), default=1))
        rs_bars = [(label, count, round(count/_rs_max*100), color) for label,count,_,color in rs_bars]
        rs_html = ""
        for r_lbl, r_cnt, r_w, r_c in rs_bars:
            rs_html += f"""
            <div style="display:grid;grid-template-columns:95px 1fr 28px;align-items:center;gap:10px;margin-bottom:14px;font-size:12.5px">
                <span style="color:#334155;font-weight:500">{r_lbl}</span>
                <div style="height:14px;background:#F1F5F9;border-radius:4px;overflow:hidden">
                    <div style="width:{r_w}%;height:100%;background:{r_c};border-radius:4px"></div>
                </div>
                <span style="text-align:right;font-weight:700;color:#0F172A">{r_cnt}</span>
            </div>
            """
        st.markdown(clean_html(rs_html), unsafe_allow_html=True)

    # --- FILTER ROW ---
    rf_q, rf_rl, rf_sk, rf_st, rf_pr, rf_mr = st.columns([1.55, 0.9, 0.9, 0.9, 0.9, 0.72])
    with rf_q:
        st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
        res_sub_q = st.text_input("Search", placeholder="🔍 Search by name, role, skill or project...", key="res_q_sub", label_visibility="collapsed")
    with rf_rl:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Role</div>", unsafe_allow_html=True)
        res_role_f = st.selectbox("Role", ["All Roles"] + sorted(R.role.dropna().astype(str).unique().tolist()), key="res_f_role", label_visibility="collapsed")
    with rf_sk:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Skill</div>", unsafe_allow_html=True)
        res_skill_f = st.selectbox("Skill", ["All Skills"] + sorted({skill.strip() for raw in R.get("skills", pd.Series(dtype=str)).dropna().astype(str) for skill in raw.split(",") if skill.strip()}), key="res_f_skill", label_visibility="collapsed")
    with rf_st:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Status</div>", unsafe_allow_html=True)
        res_stat_f = st.selectbox("Status", ["All Statuses", "Allocated", "Available", "Overallocated", "On Leave"], key="res_f_stat", label_visibility="collapsed")
    with rf_pr:
        st.markdown("<div style='font-size:11.5px;font-weight:600;color:#475569;margin-bottom:3px'>Project</div>", unsafe_allow_html=True)
        res_proj_f = st.selectbox("Project", ["All Projects"] + projects_list, key="res_f_proj", label_visibility="collapsed")
    with rf_mr:
        st.markdown("<div style='font-size:11.5px;margin-bottom:3px'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("▽ More Filters", key="res_btn_more", use_container_width=True):
            for k, v in [("res_f_role", "All Roles"), ("res_f_skill", "All Skills"), ("res_f_stat", "All Statuses"), ("res_f_proj", "All Projects"), ("res_q_top", ""), ("res_q_sub", "")]:
                st.session_state[k] = v
            st.rerun()

    st.write("")
    # --- RESOURCES DIRECTORY TABLE (Exact 10 Rows from Image 20 & 25) ---
    res_directory = []
    _res_task_map = {}
    if not T.empty:
        for _, _task in T.iterrows():
            _assignee=str(_task.get("assignee") or "").strip()
            if _assignee:
                _res_task_map.setdefault(_assignee,[]).append(str(_task.get("project") or ""))
    _pa_live = df("SELECT pa.*,p.name AS project,r.name AS resource FROM project_assignments pa LEFT JOIN projects p ON p.id=pa.project_id LEFT JOIN resources r ON r.id=pa.resource_id")
    if not _pa_live.empty:
        for _, _assignment in _pa_live.iterrows():
            _res_task_map.setdefault(str(_assignment.get("resource") or ""), []).append(str(_assignment.get("project") or ""))
    for _, _resource in R.iterrows():
        _name=str(_resource.get("name") or "")
        _capacity=max(0,min(100,int(_resource.get("capacity") or 0)))
        _assigned=sorted({x for x in _res_task_map.get(_name,[]) if x})
        _project_name=_assigned[0] if _assigned else "Unassigned"
        _status="On Leave" if str(_resource.get("status") or "").lower()=="on leave" else ("Allocated" if _capacity==0 else "Available")
        _initials="".join(x[0] for x in _name.split()[:2]).upper() or "U"
        _skills=[x.strip() for x in str(_resource.get("skills") or "").split(",") if x.strip()]
        res_directory.append({"id":int(_resource.get("id") or 0),"ini":_initials,"av_col":"#0284C7","name":_name,"email":str(_resource.get("email") or ""),"role":str(_resource.get("role") or "Team Member"),"skills":_skills,"sk_cols":[],"project":_project_name,"alloc":100-_capacity,"avail":f"{_capacity}%","status":_status})

    q_res = (res_top_q + " " + res_sub_q).strip().lower()
    filtered_res = []
    for item in res_directory:
        if res_role_f != "All Roles" and item["role"] != res_role_f:
            continue
        if res_skill_f != "All Skills" and res_skill_f not in item["skills"]:
            continue
        if res_stat_f != "All Statuses" and item["status"] != res_stat_f:
            continue
        if res_proj_f != "All Projects" and item["project"] != res_proj_f:
            continue
        if q_res and not (q_res in item["name"].lower() or q_res in item["role"].lower() or q_res in item["project"].lower() or any(q_res in s.lower() for s in item["skills"])):
            continue
        filtered_res.append(item)

    with st.container(border=True):
        r_weights = [0.3, 1.85, 1.15, 1.35, 1.35, 1.25, 0.75, 0.9, 1.05]
        h_cols = st.columns(r_weights)
        for c, h_lbl in zip(h_cols, ["☐", "Name", "Role", "Skills", "Current Project(s)", "Allocation", "Availability", "Status", "Actions"]):
            c.markdown(f"<div style='font-size:11.5px;font-weight:700;color:#1E293B;padding-bottom:6px'>{h_lbl}</div>", unsafe_allow_html=True)
        st.markdown("<hr style='margin:0 0 8px 0;border-color:#E2E8F0'>", unsafe_allow_html=True)

        for r_item in filtered_res:
            rid = r_item["id"]
            rc = st.columns(r_weights)
            with rc[0]:
                st.markdown("<div style='padding-top:10px;color:#64748B;font-size:13px'>☐</div>", unsafe_allow_html=True)
            with rc[1]:
                st.markdown(
                    clean_html(
                        f"""
                        <div style="display:flex;align-items:center;gap:10px;padding-top:2px">
                            <div style="width:34px;height:34px;min-width:34px;border-radius:50%;background:{r_item['av_col']};color:white;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px">{r_item['ini']}</div>
                            <div>
                                <div style="font-size:13px;font-weight:700;color:#0F172A;line-height:1.2">{html.escape(r_item['name'])}</div>
                                <div style="font-size:11px;color:#64748B">{html.escape(r_item['email'])}</div>
                            </div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )
            with rc[2]:
                st.markdown(f"<div style='padding-top:9px;font-size:12.5px;color:#334155'>{html.escape(r_item['role'])}</div>", unsafe_allow_html=True)
            with rc[3]:
                sk_html = "".join(
                    f'<span class="vk-tag" style="background:{bg};color:{fg}">{html.escape(sk)}</span>'
                    for sk, (bg, fg) in zip(r_item["skills"], r_item["sk_cols"])
                )
                st.markdown(f"<div style='padding-top:7px'>{sk_html}</div>", unsafe_allow_html=True)
            with rc[4]:
                st.markdown(f"<div style='padding-top:9px;font-size:12.5px;color:#334155'>{html.escape(r_item['project'])}</div>", unsafe_allow_html=True)
            with rc[5]:
                al_val = r_item["alloc"]
                al_col = RED if al_val > 100 else GREEN
                st.markdown(
                    clean_html(
                        f"""
                        <div style="display:flex;align-items:center;gap:8px;padding-top:9px">
                            <span style="font-size:12px;font-weight:600;color:#1E293B;min-width:34px">{al_val}%</span>
                            <div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden">
                                <div style="width:{min(100, al_val)}%;height:100%;background:{al_col};border-radius:99px"></div>
                            </div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )
            with rc[6]:
                av_str = r_item["avail"]
                av_color = "#DC2626" if av_str.startswith("-") else "#1E293B"
                st.markdown(f"<div style='padding-top:9px;font-size:12.5px;font-weight:600;color:{av_color}'>{av_str}</div>", unsafe_allow_html=True)
            with rc[7]:
                st_str = r_item["status"]
                st_cls = "vk-badge-blue" if st_str == "Allocated" else ("vk-badge-green" if st_str == "Available" else "vk-badge-red")
                st.markdown(f"<div style='padding-top:7px'><span class='vk-badge {st_cls}' style='min-width:86px'>{st_str}</span></div>", unsafe_allow_html=True)
            with rc[8]:
                ac1, ac2, ac3 = st.columns(3)
                with ac1:
                    if st.button("📅", key=f"res_cal_{rid}", help="Resource Calendar"):
                        st.session_state.res_modal = ("calendar", r_item)
                        st.rerun()
                with ac2:
                    if st.button("👥", key=f"res_pa_{rid}", help="Project Assignments"):
                        st.session_state.res_modal = ("assignments", r_item)
                        st.rerun()
                with ac3:
                    with st.popover("⋮"):
                        if st.button("👁 View Details", key=f"rm_v_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("view", r_item)
                            st.rerun()
                        if st.button("✎ Edit Resource", key=f"rm_e_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("edit", r_item)
                            st.rerun()
                        if st.button("💼 Assign to Project", key=f"rm_a_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("assign", r_item)
                            st.rerun()
                        if st.button("⚙ Manage Skills", key=f"rm_s_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("skills", r_item)
                            st.rerun()
                        if st.button("📅 Change Availability", key=f"rm_c_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("availability", r_item)
                            st.rerun()
                        if st.button("👤 Set on Leave", key=f"rm_l_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("leave", r_item)
                            st.rerun()
                        if st.button("🗑 Delete Resource", key=f"rm_d_{rid}", use_container_width=True):
                            st.session_state.res_modal = ("delete", r_item)
                            st.rerun()
            st.markdown("<hr style='margin:5px 0;border-color:#F1F5F9'>", unsafe_allow_html=True)

        st.caption(f"Showing all {len(filtered_res)} visible resource(s)")

    if "res_modal" in st.session_state:
        rm_type, rm_payload = st.session_state.res_modal
        if rm_type == "add":
            add_resource_modal()
        elif rm_type == "calendar":
            resource_calendar_modal(rm_payload)
        elif rm_type == "assignments":
            project_assignments_modal(rm_payload)
        else:
            resource_action_submodal(rm_type, rm_payload)


# ===========================================================================
# PAGE 6: CALENDAR (Exact Replica of Main Calendar + Inspector + Add Event Modal)
# ===========================================================================
elif page == "Calendar":

    @st.dialog("Add Event")
    def add_event_modal() -> None:
        _editing_event_id = int(st.session_state.get("cal_edit_event_id") or 0) if st.session_state.get("cal_modal") == "edit_event" else 0
        _editing_event_df = df("SELECT * FROM calendar_events WHERE id=?", (_editing_event_id,)) if _editing_event_id else pd.DataFrame()
        _editing_event = _editing_event_df.iloc[0].to_dict() if not _editing_event_df.empty else {}
        _event_types = ["Project Assignment", "Meeting / Event", "Leave", "Milestone"]
        _event_type_value = str(_editing_event.get("event_type") or "Meeting / Event")
        _event_type_index = _event_types.index(_event_type_value) if _event_type_value in _event_types else 1
        _event_projects = ["Select project"] + projects_list
        _event_project_name = str(_editing_event.get("project") or "Select project")
        _event_project_id = _editing_event.get("project_id")
        if _event_project_id is not None and not P.empty and int(_event_project_id) in P.id.astype(int).tolist():
            _event_project_name = str(P[P.id.astype(int) == int(_event_project_id)].iloc[0].get("name"))
        _event_tasks = ["Select task (optional)"] + T.name.dropna().astype(str).tolist()
        _event_task_name = str(_editing_event.get("task_name") or "Select task (optional)")
        _event_resource_defaults = [x.strip() for x in str(_editing_event.get("resources") or "").split(",") if x.strip()]
        _event_start_default = pd.to_datetime(_editing_event.get("start_date"), errors="coerce")
        _event_end_default = pd.to_datetime(_editing_event.get("end_date"), errors="coerce")
        _event_start_default = _event_start_default.date() if pd.notna(_event_start_default) else st.session_state.get("cal_selected_date", date.today())
        _event_end_default = _event_end_default.date() if pd.notna(_event_end_default) else _event_start_default
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:14px">
                    <div style="width:40px;height:40px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:700">＋</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Add Event</div>
                        <div style="font-size:12.5px;color:#64748B">Create a new event, task, or availability entry.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        e1, e2 = st.columns(2)
        with e1:
            _event_type_labels = ["🟢 Project Assignment", "🔵 Meeting / Event", "🔴 Leave", "🟣 Milestone"]
            ev_type = st.selectbox("Event Type *", _event_type_labels, index=_event_type_index, key="ae_type")
        with e2:
            ev_title = st.text_input("Title *", value=str(_editing_event.get("title") or ""), placeholder="Enter event title", key="ae_title")

        e3, e4 = st.columns(2)
        with e3:
            ev_proj = st.selectbox("Project (Optional)", _event_projects, index=_event_projects.index(_event_project_name) if _event_project_name in _event_projects else 0, key="ae_proj")
        with e4:
            ev_task = st.selectbox("Task (Optional)", _event_tasks, index=_event_tasks.index(_event_task_name) if _event_task_name in _event_tasks else 0, key="ae_task")

        _event_resource_options = R.name.dropna().astype(str).tolist()
        st.multiselect(
            "Resource(s) *",
            _event_resource_options,
            default=[x for x in _event_resource_defaults if x in _event_resource_options],
            key="ae_res",
        )

        _event_default_date = _event_start_default
        dt1, dt2, dt3, dt4 = st.columns([1.2, 0.9, 1.2, 0.9])
        with dt1:
            ev_start_d = st.date_input("Start Date & Time *", _event_default_date, key="ae_sd")
        with dt2:
            st.markdown("<div style='font-size:12.5px;margin-bottom:4px'>&nbsp;</div>", unsafe_allow_html=True)
            _start_times = ["09:00 AM", "10:00 AM", "01:00 PM", "02:00 PM"]
            st.selectbox("Start Time", _start_times, index=_start_times.index(str(_editing_event.get("start_time") or "09:00 AM")) if str(_editing_event.get("start_time") or "09:00 AM") in _start_times else 0, key="ae_st", label_visibility="collapsed")
        with dt3:
            ev_end_d = st.date_input("End Date & Time *", _event_end_default, key="ae_ed")
        with dt4:
            st.markdown("<div style='font-size:12.5px;margin-bottom:4px'>&nbsp;</div>", unsafe_allow_html=True)
            _end_times = ["05:00 PM", "11:00 AM", "03:00 PM", "04:00 PM"]
            st.selectbox("End Time", _end_times, index=_end_times.index(str(_editing_event.get("end_time") or "05:00 PM")) if str(_editing_event.get("end_time") or "05:00 PM") in _end_times else 0, key="ae_et", label_visibility="collapsed")

        al1, al2, rp_col = st.columns([1.5, 0.45, 1.95])
        with al1:
            st.text_input("Allocation (Optional)", value=str(_editing_event.get("allocation") if _editing_event.get("allocation") is not None else 50), key="ae_alloc")
            st.markdown("<div style='font-size:11px;color:#64748B;margin-top:-6px'>Enter allocation percentage for this assignment.</div>", unsafe_allow_html=True)
        with al2:
            st.markdown("<div style='margin-top:28px;border:1px solid #CBD5E1;border-radius:6px;padding:7px 8px;text-align:center;font-weight:600;background:#F8FAFC'>%</div>", unsafe_allow_html=True)
        with rp_col:
            _repeat_options = ["Does not repeat", "Daily", "Weekly", "Monthly"]
            st.selectbox("Repeat (Optional)", _repeat_options, index=_repeat_options.index(str(_editing_event.get("repeat_rule") or "Does not repeat")) if str(_editing_event.get("repeat_rule") or "Does not repeat") in _repeat_options else 0, key="ae_rep")

        ev_desc = st.text_area(
            "Description (Optional)",
            placeholder="Enter event description, meeting agenda, or additional details...",
            height=95,
            value=str(_editing_event.get("description") or ""),
            key="ae_desc",
        )
        st.markdown(f"<div style='text-align:right;font-size:11px;color:#64748B;margin-top:-6px;margin-bottom:10px'>{len(ev_desc)}/500</div>", unsafe_allow_html=True)

        f_sp, f_can, f_cr = st.columns([2.2, 1, 1.2])
        with f_can:
            if st.button("Cancel", key="ae_can_btn", use_container_width=True):
                st.session_state.pop("cal_modal", None)
                st.rerun()
        with f_cr:
            if st.button("Save Changes" if _editing_event_id else "Create Event", type="primary", key="ae_save_btn", use_container_width=True):
                if not ev_title.strip():
                    st.error("Enter an event title before saving.")
                elif ev_end_d < ev_start_d:
                    st.error("End date must be on or after the start date.")
                elif not str(st.session_state.get("ae_alloc", "50")).strip().replace(".", "", 1).isdigit() or not 0 <= float(st.session_state.get("ae_alloc", "50") or 0) <= 200:
                    st.error("Allocation must be a number between 0 and 200.")
                else:
                    _event_values = (ev_title.strip(), ev_type.replace("🟢 ", "").replace("🔵 ", "").replace("🔴 ", "").replace("🟣 ", ""), pmap.get(ev_proj), "" if ev_task == "Select task (optional)" else ev_task, ", ".join(st.session_state.get("ae_res", [])), ev_start_d.isoformat(), st.session_state.get("ae_st", "09:00 AM"), ev_end_d.isoformat(), st.session_state.get("ae_et", "05:00 PM"), int(float(st.session_state.get("ae_alloc", "50") or 0)), st.session_state.get("ae_rep", "Does not repeat"), ev_desc.strip())
                    if _editing_event_id:
                        sql("UPDATE calendar_events SET title=?,event_type=?,project_id=?,task_name=?,resources=?,start_date=?,start_time=?,end_date=?,end_time=?,allocation=?,repeat_rule=?,description=? WHERE id=?", _event_values + (_editing_event_id,))
                        log_event("Updated", "Calendar Event", _editing_event_id, ev_title)
                    else:
                        sql("INSERT INTO calendar_events(title,event_type,project_id,task_name,resources,start_date,start_time,end_date,end_time,allocation,repeat_rule,description) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)", _event_values)
                        log_event("Created", "Calendar Event", None, ev_title)
                    st.session_state.pop("cal_modal", None)
                    st.session_state.pop("cal_edit_event_id", None)
                    st.rerun()

    # --- CALENDAR TOP HEADER & FILTER BAR ---
    ch_l, ch_f1, ch_f2, ch_f3, ch_f4, ch_btn = st.columns([2.15, 0.72, 0.72, 0.72, 0.78, 0.68])
    with ch_l:
        st.markdown(clean_html(f"<div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};line-height:1.15'>Calendar</div><div style='font-size:13.5px;color:#64748B;margin-top:2px'>Live project timelines, task schedules, milestones, and saved events.</div>"),unsafe_allow_html=True)
    with ch_f1:
        cal_proj_f=st.selectbox("Project",["All Projects"]+projects_list,key="cal_f_proj",label_visibility="collapsed")
    with ch_f2:
        cal_team_f=st.selectbox("Team",["All Teams"]+teams_list,key="cal_f_team",label_visibility="collapsed")
    with ch_f3:
        cal_res_f=st.selectbox("Resource",["All Resources"]+R.name.dropna().astype(str).tolist(),key="cal_f_res",label_visibility="collapsed")
    with ch_f4:
        cal_type_f=st.selectbox("Event Type",["All Event Types","Task","Project Assignment","Meeting / Event","Leave","Milestone"],key="cal_f_type",label_visibility="collapsed")
    with ch_btn:
        if st.button("＋ Add Event",type="primary",key="cal_top_add_btn",use_container_width=True):
            st.session_state.cal_modal="add_event"
            st.session_state.pop("cal_edit_event_id", None)
            for _key in ("ae_type", "ae_title", "ae_proj", "ae_task", "ae_res", "ae_sd", "ae_st", "ae_ed", "ae_et", "ae_alloc", "ae_rep", "ae_desc"):
                st.session_state.pop(_key, None)
            st.rerun()

    if "cal_calendar_month" not in st.session_state: st.session_state.cal_calendar_month=date.today().replace(day=1)
    if "cal_selected_date" not in st.session_state: st.session_state.cal_selected_date=date.today()
    bar_l,bar_m,bar_r,bar_today,bar_search=st.columns([0.45,1.5,0.45,0.55,1.4])
    with bar_l:
        if st.button("‹",key="cal_prev_month",use_container_width=True):
            cur=st.session_state.cal_calendar_month; st.session_state.cal_calendar_month=(cur-timedelta(days=1)).replace(day=1); st.rerun()
    with bar_m:
        st.markdown(f"<div style='text-align:center;font-size:16px;font-weight:800;padding-top:7px'>{st.session_state.cal_calendar_month.strftime('%B %Y')}</div>",unsafe_allow_html=True)
    with bar_r:
        if st.button("›",key="cal_next_month",use_container_width=True):
            cur=st.session_state.cal_calendar_month; st.session_state.cal_calendar_month=(cur.replace(day=28)+timedelta(days=4)).replace(day=1); st.rerun()
    with bar_today:
        if st.button("Today",key="cal_tod_btn",use_container_width=True):
            st.session_state.cal_selected_date=date.today(); st.session_state.cal_calendar_month=date.today().replace(day=1); st.rerun()
    with bar_search:
        cal_search_q=st.text_input("Search calendar",key="cal_q_box",placeholder="Search tasks, events, projects...",label_visibility="collapsed")
    cal_view=tile_toggle("",["Month","Week","Day","Agenda"],key="cal_view_mode",default="Month")

    event_rows=[]
    task_data=df("SELECT t.*,p.name AS project,p.team AS project_team FROM tasks t LEFT JOIN projects p ON p.id=t.project_id")
    for _,r0 in task_data.iterrows():
        d0=pd.to_datetime(r0.get("due_date"),errors="coerce")
        if pd.notna(d0): event_rows.append({"date":d0.date(),"title":str(r0.get("name") or "Task"),"sub":f"Task · {r0.get('project') or 'No project'} · {r0.get('status') or 'Not Started'}","type":"Task","project":str(r0.get('project') or ''),"team":str(r0.get('project_team') or r0.get('team') or ''),"resource":str(r0.get('assignee') or ''),"tone":"b"})
    milestone_data=df("SELECT m.*,p.name AS project,p.team AS project_team FROM milestones m LEFT JOIN projects p ON p.id=m.project_id")
    for _,r0 in milestone_data.iterrows():
        d0=pd.to_datetime(r0.get("due_date"),errors="coerce")
        if pd.notna(d0): event_rows.append({"date":d0.date(),"title":str(r0.get("name") or "Milestone"),"sub":f"Milestone · {r0.get('project') or 'No project'}","type":"Milestone","project":str(r0.get('project') or ''),"team":str(r0.get('project_team') or r0.get('team') or ''),"resource":"","tone":"p"})
    event_data=df("SELECT e.*,p.name AS project,p.team AS project_team FROM calendar_events e LEFT JOIN projects p ON p.id=e.project_id")
    for _,r0 in event_data.iterrows():
        d0=pd.to_datetime(r0.get("start_date"),errors="coerce")
        if pd.notna(d0):
            typ=str(r0.get("event_type") or "Meeting / Event"); tone="r" if typ=="Leave" else ("g" if typ=="Project Assignment" else ("p" if typ=="Milestone" else "b"))
            event_rows.append({"date":d0.date(),"title":str(r0.get("title") or "Event"),"sub":f"{typ} · {r0.get('start_time') or ''}","type":typ,"project":str(r0.get('project') or ''),"team":str(r0.get('project_team') or ''),"resource":str(r0.get('resources') or ''),"tone":tone,"event_id":int(r0.get("id") or 0)})
    _assignment_events = df("SELECT pa.*,p.name AS project,p.team AS project_team,r.name AS resource FROM project_assignments pa LEFT JOIN projects p ON p.id=pa.project_id LEFT JOIN resources r ON r.id=pa.resource_id")
    for _, _assignment in _assignment_events.iterrows():
        _adate=pd.to_datetime(_assignment.get("start_date"),errors="coerce")
        if pd.notna(_adate): event_rows.append({"date":_adate.date(),"title":f"{_assignment.get('resource') or 'Resource'} · {_assignment.get('project') or 'Project Assignment'}","sub":f"Project Assignment · {int(_assignment.get('allocation') or 0)}% allocation","type":"Project Assignment","project":str(_assignment.get("project") or ""),"team":str(_assignment.get("project_team") or ""),"resource":str(_assignment.get("resource") or ""),"tone":"g"})
    # Expand saved recurring events into occurrences for the currently viewed month.
    _calendar_month = st.session_state.get("cal_calendar_month", date.today().replace(day=1))
    _visible_start = _calendar_month
    _visible_end = (_calendar_month.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    _recurring_rows = []
    for _, _ev in event_data.iterrows():
        _rule = str(_ev.get("repeat_rule") or "Does not repeat")
        if _rule == "Does not repeat":
            continue
        _start = pd.to_datetime(_ev.get("start_date"), errors="coerce")
        _end = pd.to_datetime(_ev.get("end_date"), errors="coerce")
        if pd.isna(_start):
            continue
        _occ = _start.date()
        _guard = 0
        while _occ < _visible_start and _guard < 500:
            if _rule == "Daily": _occ += timedelta(days=1)
            elif _rule == "Weekly": _occ += timedelta(days=7)
            elif _rule == "Monthly": _occ = (_occ.replace(day=28) + timedelta(days=4)).replace(day=min(_occ.day, 28))
            else: break
            _guard += 1
        while _occ <= _visible_end and _guard < 700:
            if _occ >= _visible_start and _occ != _start.date():
                _typ = str(_ev.get("event_type") or "Meeting / Event")
                _tone = "r" if _typ == "Leave" else ("g" if _typ == "Project Assignment" else ("p" if _typ == "Milestone" else "b"))
                _recurring_rows.append({"date": _occ, "title": str(_ev.get("title") or "Event"), "sub": f"{_typ} · { _ev.get('start_time') or '' } · Recurring", "type": _typ, "project": str(_ev.get("project") or ""), "team": str(_ev.get("project_team") or ""), "resource": str(_ev.get("resources") or ""), "tone": _tone, "event_id": int(_ev.get("id") or 0)})
            if _rule == "Daily": _occ += timedelta(days=1)
            elif _rule == "Weekly": _occ += timedelta(days=7)
            elif _rule == "Monthly":
                _next_month = (_occ.replace(day=28) + timedelta(days=4)).replace(day=1)
                _occ = _next_month + timedelta(days=min(_occ.day, calendar.monthrange(_next_month.year, _next_month.month)[1]) - 1)
            else: break
            _guard += 1
    event_rows.extend(_recurring_rows)
    if cal_proj_f!="All Projects": event_rows=[e for e in event_rows if e["project"]==cal_proj_f]
    if cal_team_f!="All Teams": event_rows=[e for e in event_rows if e["team"]==cal_team_f]
    if cal_res_f!="All Resources": event_rows=[e for e in event_rows if cal_res_f.lower() in e["resource"].lower()]
    if cal_type_f!="All Event Types": event_rows=[e for e in event_rows if e["type"]==cal_type_f or (cal_type_f=="Project Assignment" and e["type"]=="Task")]
    if cal_search_q: event_rows=[e for e in event_rows if cal_search_q.lower() in (e["title"]+" "+e["sub"]).lower()]

    main_l,main_r=st.columns([2.75,1.1])
    with main_l:
        selected=st.session_state.cal_selected_date; month=st.session_state.cal_calendar_month
        if cal_view=="Agenda":
            visible=sorted([e for e in event_rows if e["date"]>=selected],key=lambda e:e["date"])
            if not visible: st.info("No matching calendar items.")
            for ev in visible[:40]:
                with st.container(border=True):
                    a,b=st.columns([0.75,3])
                    with a: st.markdown(f"**{ev['date'].strftime('%b %d')}**"); st.caption(ev['date'].strftime('%A'))
                    with b: st.markdown(f"**{ev['title']}**"); st.caption(ev['sub'])
        else:
            first=month.replace(day=1)
            start_day=first-timedelta(days=(first.weekday()+1)%7)
            if cal_view=="Week": start_day=selected-timedelta(days=(selected.weekday()+1)%7); days=[start_day+timedelta(days=i) for i in range(7)]
            elif cal_view=="Day": days=[selected]
            else:
                n=((first.weekday()+1)%7+calendar.monthrange(month.year,month.month)[1]+6)//7*7
                days=[start_day+timedelta(days=i) for i in range(n)]
            rows_html=''
            for i in range(0,len(days),7):
                cells=''
                for d0 in days[i:i+7]:
                    sel=d0==selected; bg='#EFF6FF' if sel else ('#FFFFFF' if d0.month==month.month else '#F8FAFC'); border='2px solid #2563EB' if sel else '1px solid #EDF2F7'
                    pills=''
                    for ev in [x for x in event_rows if x['date']==d0][:3]:
                        tone={'g':('#DCFCE7','#065F46'),'b':('#DBEAFE','#1D4ED8'),'r':('#FEE2E2','#B91C1C'),'p':('#F3E8FF','#7C3AED')}[ev['tone']]
                        pills+=f'<div style="background:{tone[0]};color:{tone[1]};border-radius:6px;padding:3px 5px;font-size:10px;margin-top:3px"><b>{html.escape(ev["title"])}</b><div style="font-size:9px">{html.escape(ev["sub"])}</div></div>'
                    cells+=f'<div style="min-height:92px;border:{border};background:{bg};padding:5px 6px"><b>{d0.day}</b>{pills}</div>'
                rows_html+=f'<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr))">{cells}</div>'
            header=''.join(f'<div>{x}</div>' for x in ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'])
            st.markdown(clean_html(f'<div style="border:1px solid #E2E8F0;border-radius:12px;overflow:hidden;background:white"><div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));background:#F8FAFC;text-align:center;padding:9px 0;font-weight:700">{header}</div>{rows_html}</div>'),unsafe_allow_html=True)
            day_options=days
            current=selected if selected in day_options else day_options[0]
            pick=st.selectbox("Inspect date",day_options,index=day_options.index(current),format_func=lambda d:d.strftime('%A, %B %d, %Y'),key="cal_inspect_date")
            if pick!=st.session_state.cal_selected_date: st.session_state.cal_selected_date=pick; st.rerun()
    with main_r,st.container(border=True):
        selected=st.session_state.cal_selected_date; todays=[e for e in event_rows if e['date']==selected]
        st.markdown(f"<div style='font-size:18px;font-weight:800;color:#0F172A'>{selected.strftime('%a, %b %d, %Y')}</div>",unsafe_allow_html=True)
        st.caption(f"{len(todays)} scheduled item(s)")
        tabs=st.tabs(["Day View",f"Team ({len({e['resource'] for e in todays if e['resource']})})",f"Projects ({len({e['project'] for e in todays if e['project']})})"])
        with tabs[0]:
            if not todays: st.info("No tasks, milestones, or saved events on this date.")
            for ev in todays:
                st.markdown(f"**{ev['title']}**")
                st.caption(ev['sub'])
                if ev.get("event_id"):
                    _ec1, _ec2 = st.columns(2)
                    with _ec1:
                        if st.button("Edit event", key=f"cal_edit_event_{ev['event_id']}_{ev['date'].isoformat()}", use_container_width=True):
                            st.session_state["cal_edit_event_id"] = int(ev["event_id"])
                            st.session_state["cal_modal"] = "edit_event"
                            for _key in ("ae_type", "ae_title", "ae_proj", "ae_task", "ae_res", "ae_sd", "ae_st", "ae_ed", "ae_et", "ae_alloc", "ae_rep", "ae_desc"):
                                st.session_state.pop(_key, None)
                            st.rerun()
                    with _ec2:
                        if st.button("Delete event", key=f"cal_delete_event_{ev['event_id']}_{ev['date'].isoformat()}", use_container_width=True):
                            sql("DELETE FROM calendar_events WHERE id=?", (int(ev["event_id"]),))
                            log_event("Deleted", "Calendar Event", int(ev["event_id"]), str(ev["title"]))
                            st.rerun()
                st.divider()
        with tabs[1]:
            members=sorted({x.strip() for e in todays for x in e['resource'].split(',') if x.strip()}); st.write('\n'.join(members) if members else 'No resources scheduled.')
        with tabs[2]:
            names=sorted({e['project'] for e in todays if e['project']}); st.write('\n'.join(names) if names else 'No projects scheduled.')
        st.markdown("**Quick Actions**")
        q1,q2=st.columns(2)
        with q1:
            if st.button("＋ Add Event",key="qa_add_ev",use_container_width=True):
                st.session_state.cal_modal="add_event"; st.session_state.pop("cal_edit_event_id",None)
                for _key in ("ae_type", "ae_title", "ae_proj", "ae_task", "ae_res", "ae_sd", "ae_st", "ae_ed", "ae_et", "ae_alloc", "ae_rep", "ae_desc"): st.session_state.pop(_key,None)
                st.rerun()
        with q2:
            if st.button("View Team",key="qa_v_team",use_container_width=True): st.session_state.page="Resources"; st.rerun()
        q3,q4=st.columns(2)
        with q3:
            if st.button("Set Leave",key="qa_set_lv",use_container_width=True):
                st.session_state.cal_modal="add_event"; st.session_state.pop("cal_edit_event_id",None)
                for _key in ("ae_title", "ae_proj", "ae_task", "ae_res", "ae_sd", "ae_st", "ae_ed", "ae_et", "ae_alloc", "ae_rep", "ae_desc"): st.session_state.pop(_key,None)
                st.session_state["ae_type"]="🔴 Leave"; st.rerun()
        with q4:
            if st.button("Availability",key="qa_ed_av",use_container_width=True): st.session_state.page="Resources"; st.rerun()

    if st.session_state.get("cal_modal") in {"add_event", "edit_event"}: add_event_modal()

# ===========================================================================
# PAGE 7: REPORTS (Exact Replica of All 7 Report Mockup Screens)
# ===========================================================================
elif page == "Reports":
    _report_p_status = P.status.fillna("Not Started").astype(str) if not P.empty else pd.Series(dtype=str)
    _report_p_total = len(P)
    _report_p_counts = {name: int((_report_p_status.str.lower() == name.lower()).sum()) for name in ["On Track", "In Progress", "At Risk", "Delayed", "Completed", "Not Started"]}
    def _report_p_pct(n): return f"({round(n / _report_p_total * 100)}%)" if _report_p_total else "(0%)"
    _report_t_status = T.status.fillna("Not Started").astype(str) if not T.empty else pd.Series(dtype=str)
    _t_total = len(T); _t_done = int((_report_t_status.str.lower() == "completed").sum()); _t_progress = int((_report_t_status.str.lower() == "in progress").sum()); _t_notstarted = int((_report_t_status.str.lower() == "not started").sum())
    def _t_pct(n): return f"{round(n / _t_total * 100)}%" if _t_total else "0%"
    _report_r_status = R.status.fillna("Available").astype(str) if not R.empty else pd.Series(dtype=str)
    _r_total = len(R); _r_avail = int((_report_r_status.str.lower() == "available").sum()); _r_partial = int((_report_r_status.str.lower() == "partial").sum()); _r_leave = int((_report_r_status.str.lower() == "on leave").sum())
    _report_m_status = M.status.fillna("Not Started").astype(str) if not M.empty else pd.Series(dtype=str)
    _m_total = len(M); _m_done = int((_report_m_status.str.lower() == "completed").sum()); _m_prog = int((_report_m_status.str.lower() == "in progress").sum()); _m_not = int((_report_m_status.str.lower() == "not started").sum())
    if "report_tab" not in st.session_state:
        st.session_state.report_tab = "Overview"

    active_rtab = st.session_state.report_tab

    sub_headers = {
        "Overview": ("", "Gain insights into project progress, resource utilization, and team performance."),
        "Projects": ("Project Report", "Track project progress, milestones, task status, and overall project health."),
        "Resources": ("Resource Report", "View team availability, assignments, and workload across projects."),
        "Tasks": ("Tasks Report", "Track task status, progress, and workload across projects and teams."),
        "Milestones": ("Milestones Report", "Track milestone status, progress, and upcoming key dates across projects."),
        "Workload": ("Workload Report", "View team workload, task distribution, and capacity across projects."),
        "Custom": ("Custom Report", "Create and customize reports to analyze project data based on your needs."),
    }
    sub_h, sub_desc = sub_headers.get(active_rtab, sub_headers["Overview"])

    # --- TOP HEADER & FILTER BAR ---
    rh_l, rh_f1, rh_f2, rh_f3, rh_exp = st.columns([2.2, 1.05, 0.85, 0.85, 0.78])
    with rh_l:
        sub_h_html = f"<div style='font-size:19px;font-weight:800;color:{TEXT_COLOR};margin-top:2px'>{html.escape(sub_h)}</div>" if sub_h else ""
        st.markdown(
            clean_html(
                f"""
                <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.1'>Reports</div>
                {sub_h_html}
                <div style='font-size:13px;color:#64748B;margin-top:2px'>{html.escape(sub_desc)}</div>
                """
            ),
            unsafe_allow_html=True,
        )
    with rh_f1:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Date Range</div>", unsafe_allow_html=True)
        st.date_input("Date Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="rep_top_dates", label_visibility="collapsed")
    with rh_f2:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Project</div>", unsafe_allow_html=True)
        rep_proj_f = st.selectbox("Project", ["All Projects"] + projects_list, key="rep_top_proj", label_visibility="collapsed")
    with rh_f3:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Team</div>", unsafe_allow_html=True)
        rep_team_f = st.selectbox("Team", ["All Teams"] + teams_list, key="rep_top_team", label_visibility="collapsed")

    # Apply the report header filters to every report tab and its export.
    _report_dates = st.session_state.get("rep_top_dates", (date.today().replace(day=1), date.today()+timedelta(days=30)))
    _report_start, _report_end = (_report_dates if isinstance(_report_dates, (tuple, list)) and len(_report_dates) == 2 else (date.today().replace(day=1), date.today()+timedelta(days=30)))
    if _report_start and _report_end and _report_start > _report_end:
        _report_start, _report_end = _report_end, _report_start
    if rep_proj_f != "All Projects":
        P = P[P.name.astype(str) == rep_proj_f].copy()
        T = T[T.project.astype(str) == rep_proj_f].copy() if not T.empty else T.copy()
        M = M[M.project.astype(str) == rep_proj_f].copy() if not M.empty else M.copy()
    if rep_team_f != "All Teams":
        P = P[P.team.fillna("").astype(str) == rep_team_f].copy() if not P.empty else P.copy()
        T = T[T.team.fillna("").astype(str) == rep_team_f].copy() if not T.empty else T.copy()
        R = R[R.team.fillna("").astype(str) == rep_team_f].copy() if not R.empty else R.copy()
        M = M[M.team.fillna("").astype(str) == rep_team_f].copy() if not M.empty else M.copy()
    if _report_start and _report_end:
        if not P.empty:
            _ps = pd.to_datetime(P.start_date, errors="coerce").dt.date
            _pe = pd.to_datetime(P.end_date, errors="coerce").dt.date
            P = P[((_pe.isna()) | (_pe >= _report_start)) & ((_ps.isna()) | (_ps <= _report_end))].copy()
        if not T.empty:
            _td = pd.to_datetime(T.due_date, errors="coerce").dt.date
            T = T[_td.isna() | ((_td >= _report_start) & (_td <= _report_end))].copy()
        if not M.empty:
            _md = pd.to_datetime(M.due_date, errors="coerce").dt.date
            M = M[_md.isna() | ((_md >= _report_start) & (_md <= _report_end))].copy()

    _report_p_status = P.status.fillna("Not Started").astype(str) if not P.empty else pd.Series(dtype=str)
    _report_p_total = len(P)
    _report_p_counts = {name: int((_report_p_status.str.lower() == name.lower()).sum()) for name in ["On Track", "In Progress", "At Risk", "Delayed", "Completed", "Not Started"]}
    _report_t_status = T.status.fillna("Not Started").astype(str) if not T.empty else pd.Series(dtype=str)
    _t_total = len(T); _t_done = int((_report_t_status.str.lower() == "completed").sum()); _t_progress = int((_report_t_status.str.lower() == "in progress").sum()); _t_notstarted = int((_report_t_status.str.lower() == "not started").sum())
    _report_r_status = R.status.fillna("Available").astype(str) if not R.empty else pd.Series(dtype=str)
    _r_total = len(R); _r_avail = int((_report_r_status.str.lower() == "available").sum()); _r_partial = int((_report_r_status.str.lower() == "partial").sum()); _r_leave = int((_report_r_status.str.lower() == "on leave").sum())
    _report_m_status = M.status.fillna("Not Started").astype(str) if not M.empty else pd.Series(dtype=str)
    _m_total = len(M); _m_done = int((_report_m_status.str.lower() == "completed").sum()); _m_prog = int((_report_m_status.str.lower() == "in progress").sum()); _m_not = int((_report_m_status.str.lower() == "not started").sum())
    with rh_exp:
        st.markdown("<div style='font-size:11px;margin-bottom:2px'>&nbsp;</div>", unsafe_allow_html=True)
        st.download_button(
            "⬇ Export Report ▾",
            excel_bytes({"Projects": P, "Tasks": T, "Resources": R, "Milestones": M}),
            file_name=f"vektor_{active_rtab.lower()}_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="rep_top_export_btn",
        )

    st.write("")
    # --- 7-TAB TILE NAVIGATION BAR (No Radio Buttons) ---
    rtab_names = ["Overview", "Projects", "Resources", "Tasks", "Milestones", "Workload", "Custom"]
    nav_cols = st.columns([1, 1, 1, 1, 1, 1, 1, 2.2])
    for idx, r_name in enumerate(rtab_names):
        with nav_cols[idx]:
            if st.button(
                r_name,
                key=f"rep_nav_btn_{r_name}",
                type="primary" if active_rtab == r_name else "secondary",
                use_container_width=True,
            ):
                st.session_state.report_tab = r_name
                st.rerun()

    st.write("")

    # Helper for Donut + Right Legend inside Report Cards
    def render_report_donut_card(
        center_big: str,
        center_sub: str,
        labels: list[str],
        values: list[int],
        colors: list[str],
        right_rows: list[tuple[str, str, str]],
        height: int = 210,
    ) -> None:
        d_c, l_c = st.columns([1.05, 1.15])
        with d_c:
            fig = go.Figure(
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.66,
                    sort=False,
                    direction="clockwise",
                    marker=dict(colors=colors, line=dict(color="#FFFFFF", width=2)),
                    textinfo="none",
                )
            )
            fig.update_layout(
                height=height,
                margin=dict(l=4, r=4, t=4, b=4),
                showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)",
                annotations=[
                    dict(
                        text=f"<b style='font-size:20px;color:#0F172A'>{center_big}</b><br><span style='font-size:11px;color:#64748B'>{center_sub}</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with l_c:
            rows_mk = "".join(
                f"<div style='display:flex;justify-content:space-between;align-items:center;padding:5px 0;font-size:12.5px'>"
                f"<span><span style='color:{dot};font-size:14px'>●</span> &nbsp;{html.escape(lbl)}</span>"
                f"<b style='color:#0F172A'>{html.escape(val)}</b></div>"
                for dot, lbl, val in right_rows
            )
            st.markdown(clean_html(f"<div style='padding-top:34px'>{rows_mk}</div>"), unsafe_allow_html=True)

    # Helper for Pagination Footer
    def render_pagination_footer(left_text: str, pages: list[int], active_p: int = 1) -> None:
        # The associated HTML tables currently render all filtered rows; avoid fake page buttons.
        st.caption(left_text)

    # =========================================================================
    # REPORT TAB 1: OVERVIEW (Exact Match to Images 30 & 31)
    # =========================================================================
    if active_rtab == "Overview":
        ov_k1, ov_k2, ov_k3, ov_k4 = st.columns(4)
        with ov_k1:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:22px">💼</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Projects</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">{len(P)}</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> On Track</span><b>{_report_p_counts['On Track']}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> At Risk</span><b>{_report_p_counts['At Risk']}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> Delayed</span><b>{_report_p_counts['Delayed']}</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k2:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:22px">📋</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Tasks</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">{_t_total}</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:115px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>{_t_done}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>{_t_progress}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>{_t_notstarted}</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k3:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#EDE9FE;color:#7C3AED;display:flex;align-items:center;justify-content:center;font-size:22px">👥</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Team Members</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">{_r_total}</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Available</span><b>{_r_avail}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> Partial</span><b>{_r_partial}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> On Leave</span><b>{_r_leave}</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k4:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:58px;height:58px;border-radius:50%;border:6px solid #00875A;border-left-color:#3B82F6;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:800;color:#0F172A">{round(_t_done / _t_total * 100) if _t_total else 0}%</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B;margin-bottom:4px">Overall Progress</div>
                                <div style="font-size:11.5px;line-height:1.7">
                                    <div><span style="color:#00875A">●</span> <b>{round(_t_done / _t_total * 100) if _t_total else 0}%</b> Completed</div>
                                    <div><span style="color:#1473E6">●</span> <b>{round(_t_progress / _t_total * 100) if _t_total else 0}%</b> In Progress</div>
                                    <div><span style="color:#64748B">●</span> <b>{round(_t_notstarted / _t_total * 100) if _t_total else 0}%</b> Not Started</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        st.write("")
        # Row 1: Project Status | Task Status | Team Availability (Monthly Stacked)
        r1_c1, r1_c2, r1_c3 = st.columns(3)
        with r1_c1, st.container(border=True):
            section_title("Project Status")
            render_report_donut_card(
                str(_report_p_total), "Projects",
                ["On Track", "At Risk", "Delayed"], [_report_p_counts["On Track"], _report_p_counts["At Risk"], _report_p_counts["Delayed"]], [GREEN, ORANGE, RED],
                [(GREEN, "On Track", f"{_report_p_counts['On Track']} {_report_p_pct(_report_p_counts['On Track'])}"), (ORANGE, "At Risk", f"{_report_p_counts['At Risk']} {_report_p_pct(_report_p_counts['At Risk'])}"), (RED, "Delayed", f"{_report_p_counts['Delayed']} {_report_p_pct(_report_p_counts['Delayed'])}")],
            )
        with r1_c2, st.container(border=True):
            section_title("Task Status")
            render_report_donut_card(
                str(_t_total), "Tasks",
                ["Completed", "In Progress", "Not Started"], [_t_done, _t_progress, _t_notstarted], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", f"{_t_done} ({_t_pct(_t_done)})"), (BLUE, "In Progress", f"{_t_progress} ({_t_pct(_t_progress)})"), ("#94A3B8", "Not Started", f"{_t_notstarted} ({_t_pct(_t_notstarted)})")],
            )
        with r1_c3, st.container(border=True):
            section_title("Team Availability")
            ta_df = pd.DataFrame({"Current Status": ["Current"], "Available": [_r_avail], "Partial": [_r_partial], "On Leave": [_r_leave]})
            fig_ta = px.bar(
                ta_df,
                x="Current Status",
                y=["Available", "Partial", "On Leave"],
                color_discrete_map={"Available": GREEN, "Partial": ORANGE, "On Leave": RED},
            )
            fig_ta.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, range=[0, 26]),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.42,
            )
            st.plotly_chart(fig_ta, use_container_width=True, config={"displayModeBar": False})

        # Row 2: Project Progress | Task Completion Trend | Team Workload
        r2_c1, r2_c2, r2_c3 = st.columns(3)
        with r2_c1, st.container(border=True):
            section_title("Project Progress")
            pp_items = [(str(r0.name), int(r0.progress or 0)) for _, r0 in P.sort_values("progress", ascending=False).head(8).iterrows()]
            pp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 38px;align-items:center;gap:10px;margin-bottom:13px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='height:13px;background:#E2E8F0;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{p_v}%;height:100%;background:#00875A;border-radius:4px'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{p_v}%</b></div>"
                for p_nm, p_v in pp_items
            )
            st.markdown(clean_html(f"<div style='padding-top:8px'>{pp_html}</div>"), unsafe_allow_html=True)

        with r2_c2, st.container(border=True):
            section_title("Task Status by Due Month")
            _trend = T.copy()
            _trend["_due_month"] = pd.to_datetime(_trend.due_date, errors="coerce").dt.strftime("%b %Y") if not _trend.empty else pd.Series(dtype=str)
            _trend = _trend[_trend["_due_month"].notna()] if not _trend.empty else _trend
            if not _trend.empty:
                tct_df = _trend.assign(_status=_trend.status.fillna("Not Started").astype(str)).pivot_table(index="_due_month", columns="_status", values="id", aggfunc="count", fill_value=0).reset_index().rename(columns={"_due_month":"Month"})
                for _status_col in ["Completed", "In Progress", "Not Started"]:
                    if _status_col not in tct_df: tct_df[_status_col] = 0
                tct_df = tct_df[["Month", "Completed", "In Progress", "Not Started"]].tail(6)
            else:
                tct_df = pd.DataFrame({"Month": ["No dates"], "Completed": [0], "In Progress": [0], "Not Started": [0]})
            fig_tct = px.line(
                tct_df,
                x="Month",
                y=["Completed", "In Progress", "Not Started"],
                markers=True,
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": "#94A3B8"},
            )
            fig_tct.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=4, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, range=[0, 210]),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_tct, use_container_width=True, config={"displayModeBar": False})

        with r2_c3, st.container(border=True):
            section_title("Team Workload")
            tw_bars = []
            for _, _person in R.head(5).iterrows():
                _person_name = str(_person.get("name") or "")
                _assigned_count = int((T.assignee.fillna("").astype(str) == _person_name).sum()) if not T.empty else 0
                _work_pct = min(100, _assigned_count * 10)
                _work_color = RED if _work_pct >= 80 else (ORANGE if _work_pct >= 60 else (BLUE if _work_pct >= 35 else GREEN))
                tw_bars.append(("".join(part[0] for part in _person_name.split()[:2]).upper(), "#0284C7", _person_name, _work_pct, _work_color))
            tw_mk = "".join(
                f"<div style='display:grid;grid-template-columns:135px 1fr 38px;align-items:center;gap:10px;margin-bottom:11px;font-size:12.5px'>"
                f"<div style='display:flex;align-items:center'><span class='vk-avatar' style='background:{abg}'>{ini}</span><span style='font-weight:500'>{nm}</span></div>"
                f"<div style='height:12px;background:#E2E8F0;border-radius:4px;overflow:hidden'><div style='width:{pct}%;height:100%;background:{bcol};border-radius:4px'></div></div>"
                f"<b style='text-align:right;color:#0F172A'>{pct}%</b></div>"
                for ini, abg, nm, pct, bcol in tw_bars
            )
            st.markdown(clean_html(f"<div style='padding-top:4px'>{tw_mk}</div>"), unsafe_allow_html=True)

        # Row 3: Upcoming Milestones | Top Projects by Progress | Project Health
        r3_c1, r3_c2, r3_c3 = st.columns([1.2, 1, 1])
        with r3_c1, st.container(border=True):
            section_title("Upcoming Milestones", "View All")
            _um = M.copy()
            _um["_due"] = pd.to_datetime(_um.due_date, errors="coerce") if not _um.empty else pd.Series(dtype="datetime64[ns]")
            _um = _um[_um["_due"].notna()].sort_values("_due").head(5) if not _um.empty else _um
            um_rows = [(str(r0.name), str(r0.project or ""), fmt_date(r0.due_date), str(r0.status or "Not Started")) for _, r0 in _um.iterrows()]
            um_html = "".join(
                f"<tr><td><b>{m_n}</b></td><td>{p_n}</td><td>{d_s}</td><td>{badge_html(st_s)}</td></tr>"
                for m_n, p_n, d_s, st_s in um_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Milestone</th><th>Project</th><th>Due Date</th><th>Status</th></tr></thead><tbody>{um_html}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

        with r3_c2, st.container(border=True):
            section_title("Top Projects by Progress", "View All")
            tp_rows = [(str(r0.name), int(r0.progress or 0)) for _, r0 in P.sort_values("progress", ascending=False).head(5).iterrows()]
            tp_html = "".join(
                f"<tr><td><b>{p_n}</b></td><td>{progress_html(pct, GREEN)}</td></tr>"
                for p_n, pct in tp_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Project</th><th>Progress</th></tr></thead><tbody>{tp_html}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

        with r3_c3, st.container(border=True):
            section_title("Project Health", "View All")
            ph_rows = [(str(r0.name), int(r0.progress or 0), str(r0.bar_color or GREEN), str(r0.health or r0.status or "Unknown")) for _, r0 in P.head(8).iterrows()]
            ph_html = "".join(
                f"<tr><td><b>{p_n}</b></td><td>{progress_html(pct, bcol)}</td><td>{badge_html(h_s)}</td></tr>"
                for p_n, pct, bcol, h_s in ph_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Project</th><th>Progress</th><th>Health</th></tr></thead><tbody>{ph_html}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

    # =========================================================================
    # REPORT TAB 2: PROJECTS (Exact Match to Image 28)
    # =========================================================================
    elif active_rtab == "Projects":
        pk1, pk2, pk3, pk4 = st.columns(4)
        with pk1:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:98px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:22px">📁</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Projects</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">{_report_p_total}</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> On Track</span><b>{_report_p_counts["On Track"]}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> At Risk</span><b>{_report_p_counts["At Risk"]}</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> Delayed</span><b>{_report_p_counts["Delayed"]}</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with pk2:
            kpi_card_mockup("Completed Projects", str(_report_p_counts["Completed"]), sub_text=_report_p_pct(_report_p_counts["Completed"]), icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with pk3:
            kpi_card_mockup("In Progress Projects", str(_report_p_counts["In Progress"]), sub_text=_report_p_pct(_report_p_counts["In Progress"]), icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with pk4:
            kpi_card_mockup("Not Started Projects", str(_report_p_counts["Not Started"]), sub_text=_report_p_pct(_report_p_counts["Not Started"]), icon_html="❚❚", icon_bg="#E2E8F0", icon_fg="#475569")

        st.write("")
        pr1_1, pr1_2, pr1_3 = st.columns(3)
        with pr1_1, st.container(border=True):
            section_title("Project Status")
            render_report_donut_card(
                str(_report_p_total), "Projects",
                ["On Track", "At Risk", "Delayed"], [_report_p_counts["On Track"], _report_p_counts["At Risk"], _report_p_counts["Delayed"]], [GREEN, ORANGE, RED],
                [(GREEN, "On Track", f"{_report_p_counts['On Track']} {_report_p_pct(_report_p_counts['On Track'])}"), (ORANGE, "At Risk", f"{_report_p_counts['At Risk']} {_report_p_pct(_report_p_counts['At Risk'])}"), (RED, "Delayed", f"{_report_p_counts['Delayed']} {_report_p_pct(_report_p_counts['Delayed'])}")],
            )
        with pr1_2, st.container(border=True):
            section_title("Project Progress")
            pp_items = [(str(r0.name), int(r0.progress or 0)) for _, r0 in P.sort_values("progress", ascending=False).head(8).iterrows()]
            pp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 38px;align-items:center;gap:10px;margin-bottom:13px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='height:13px;background:#E2E8F0;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{p_v}%;height:100%;background:#00875A;border-radius:4px'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{p_v}%</b></div>"
                for p_nm, p_v in pp_items
            )
            st.markdown(clean_html(f"<div style='padding-top:8px'>{pp_html}</div>"), unsafe_allow_html=True)
        with pr1_3, st.container(border=True):
            section_title("Projects by Status (Trend)")
            pst_df = pd.DataFrame({"Month": [date.today().strftime("%b %Y")], "On Track": [_report_p_counts["On Track"]], "At Risk": [_report_p_counts["At Risk"]], "Delayed": [_report_p_counts["Delayed"]]})
            fig_pst = px.bar(
                pst_df,
                x="Month",
                y=["On Track", "At Risk", "Delayed"],
                color_discrete_map={"On Track": GREEN, "At Risk": ORANGE, "Delayed": RED},
            )
            fig_pst.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, range=[0, 13]),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.45,
            )
            st.plotly_chart(fig_pst, use_container_width=True, config={"displayModeBar": False})

        pr2_1, pr2_2, pr2_3 = st.columns(3)
        with pr2_1, st.container(border=True):
            section_title("Project Health")
            render_report_donut_card(
                str(_report_p_total), "Projects",
                ["Healthy", "At Risk", "Critical"], [int((_report_p_status.str.lower() == "healthy").sum()), int((_report_p_status.str.lower() == "at risk").sum()), int((_report_p_status.str.lower() == "critical").sum())], [GREEN, ORANGE, RED],
                [(GREEN, "Healthy", str(int((_report_p_status.str.lower() == "healthy").sum()))), (ORANGE, "At Risk", str(int((_report_p_status.str.lower() == "at risk").sum()))), (RED, "Critical", str(int((_report_p_status.str.lower() == "critical").sum())))],
            )
        with pr2_2, st.container(border=True):
            section_title("Milestone Status")
            render_report_donut_card(
                str(_m_total), "Milestones",
                ["Completed", "In Progress", "Not Started"], [_m_done, _m_prog, _m_not], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", f"{_m_done} ({_m_pct(_m_done)})"), (BLUE, "In Progress", f"{_m_prog} ({_m_pct(_m_prog)})"), ("#94A3B8", "Not Started", f"{_m_not} ({_m_pct(_m_not)})")],
            )
        with pr2_3, st.container(border=True):
            section_title("Top Projects by Progress")
            tpp_items = [(str(r0.name), int(r0.progress or 0)) for _, r0 in P.sort_values("progress", ascending=False).head(8).iterrows()]
            tpp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 38px;align-items:center;gap:10px;margin-bottom:13px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500'>{p_nm}</span>"
                f"<div style='height:13px;background:#E2E8F0;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{p_v}%;height:100%;background:#00875A;border-radius:4px'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{p_v}%</b></div>"
                for p_nm, p_v in tpp_items
            )
            st.markdown(clean_html(f"<div style='padding-top:8px'>{tpp_html}</div>"), unsafe_allow_html=True)

        pr3_l, pr3_r = st.columns([1.85, 1])
        with pr3_l, st.container(border=True):
            ph_c1, ph_c2 = st.columns([2.2, 1.2])
            with ph_c1:
                st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;padding-top:4px'>Project List</div>", unsafe_allow_html=True)
            with ph_c2:
                pr_search = st.text_input("Search", placeholder="🔍 Search projects...", key="rep_proj_list_q", label_visibility="collapsed")
            plist_rows = []
            for _, _pr in P.iterrows():
                _pr_tasks = T[T.project_id == _pr.id] if not T.empty else pd.DataFrame()
                _pr_ms = M[M.project_id == _pr.id] if not M.empty else pd.DataFrame()
                plist_rows.append((str(_pr.name), str(_pr.status or "Not Started"), fmt_date(_pr.start_date), fmt_date(_pr.end_date), int(_pr.progress or 0), f"{int((_pr_ms.status.astype(str).str.lower() == 'completed').sum())} / {len(_pr_ms)}", f"{int((_pr_tasks.status.astype(str).str.lower() == 'completed').sum())} / {len(_pr_tasks)}", str(_pr.health or "Healthy")))
            if pr_search:
                plist_rows = [r for r in plist_rows if pr_search.lower() in r[0].lower()]
            pl_tr = "".join(
                f"<tr><td><b>{p_n}</b></td><td>{badge_html(st_s)}</td><td>{sd}</td><td>{ed}</td><td>{progress_html(pct, GREEN)}</td><td style='text-align:center'>{ms_c}</td><td style='text-align:center'>{tk_c}</td><td>{badge_html(hl_s)}</td></tr>"
                for p_n, st_s, sd, ed, pct, ms_c, tk_c, hl_s in plist_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Project Name</th><th>Status</th><th>Start Date</th><th>End Date</th><th>Progress</th><th style="text-align:center">Milestones</th><th style="text-align:center">Tasks</th><th>Health</th></tr></thead><tbody>{pl_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer(f"Showing all {len(plist_rows)} filtered project(s)", [1], 1)

        with pr3_r, st.container(border=True):
            section_title("Upcoming Project Milestones", "View All")
            _upm = M.copy()
            _upm["_due"] = pd.to_datetime(_upm.due_date, errors="coerce") if not _upm.empty else pd.Series(dtype="datetime64[ns]")
            _upm = _upm[_upm["_due"].notna()].sort_values("_due").head(5) if not _upm.empty else _upm
            upm_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), fmt_date(r0.get("due_date")), str(r0.get("status") or "Not Started")) for _, r0 in _upm.iterrows()]
            upm_tr = "".join(
                f"<tr><td><b>{m_n}</b></td><td>{p_n}</td><td>{dd}</td><td>{badge_html(st_s)}</td></tr>"
                for m_n, p_n, dd, st_s in upm_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Milestone</th><th>Project</th><th>Due Date</th><th>Status</th></tr></thead><tbody>{upm_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

    # =========================================================================
    # REPORT TAB 3: RESOURCES (Exact Match to Image 29)
    # =========================================================================
    elif active_rtab == "Resources":
        rk1, rk2, rk3, rk4 = st.columns(4)
        _r_status=R.status.fillna("Available").astype(str) if not R.empty else pd.Series(dtype=str)
        _r_total=len(R); _r_avail=int((_r_status.str.lower()=="available").sum()); _r_partial=int((_r_status.str.lower()=="partial").sum()); _r_leave=int((_r_status.str.lower()=="on leave").sum())
        def _r_pct(n): return f"{round(n/_r_total*100)}%" if _r_total else "0%"
        with rk1: kpi_card_mockup("Total Resources", str(_r_total), icon_html="👥", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with rk2: kpi_card_mockup("Available", str(_r_avail), sub_text=_r_pct(_r_avail), icon_html="👤", icon_bg="#D1FAE5", icon_fg="#00875A")
        with rk3: kpi_card_mockup("Partial Allocation", str(_r_partial), sub_text=_r_pct(_r_partial), icon_html="◔", icon_bg="#FEF3C7", icon_fg="#F28C00")
        with rk4: kpi_card_mockup("On Leave", str(_r_leave), sub_text=_r_pct(_r_leave), icon_html="👤", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        rr1_1, rr1_2, rr1_3 = st.columns(3)
        with rr1_1, st.container(border=True):
            section_title("Resource Status")
            render_report_donut_card(
                str(_r_total), "Resources",
                ["Available", "Partial", "On Leave"], [_r_avail, _r_partial, _r_leave], [GREEN, ORANGE, RED],
                [(GREEN, "Available", f"{_r_avail} ({_r_pct(_r_avail)})"), (ORANGE, "Partial", f"{_r_partial} ({_r_pct(_r_partial)})"), (RED, "On Leave", f"{_r_leave} ({_r_pct(_r_leave)})")],
            )
        with rr1_2, st.container(border=True):
            section_title("Team Distribution")
            td_df = R.assign(team=R.team.fillna("Others").astype(str)).groupby("team", as_index=False).size().rename(columns={"team": "Team", "size": "Members"}) if not R.empty else pd.DataFrame(columns=["Team", "Members"])
            fig_td = px.bar(
                td_df,
                x="Team",
                y="Members",
                text="Members",
                color="Team",
                color_discrete_sequence=[GREEN, BLUE, PURPLE, ORANGE, "#64748B"],
            )
            fig_td.update_traces(textposition="outside", cliponaxis=False)
            fig_td.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=12, b=0),
                showlegend=False,
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_td, use_container_width=True, config={"displayModeBar": False})

        with rr1_3, st.container(border=True):
            section_title("Assignments by Project")
            abp_items = []
            _max_assign = max(1, max((len(T[T.project == str(pr.name)]) for _, pr in P.iterrows()), default=1)) if not T.empty else 1
            for _, _pr in P.iterrows():
                _pt = T[T.project == str(_pr.name)] if not T.empty else pd.DataFrame()
                _n = int(_pt.assignee.dropna().astype(str).nunique()) if not _pt.empty else 0
                abp_items.append((str(_pr.name), _n, round(_n/_max_assign*100), str(_pr.bar_color or GREEN)))
            abp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 24px;align-items:center;gap:10px;margin-bottom:13px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='height:13px;background:#F1F5F9;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{w_p}%;height:100%;background:{bcol};border-radius:4px'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{cnt}</b></div>"
                for p_nm, cnt, w_p, bcol in abp_items
            )
            st.markdown(clean_html(f"<div style='padding-top:8px'>{abp_html}</div>"), unsafe_allow_html=True)

        rr2_1, rr2_2 = st.columns(2)
        with rr2_1, st.container(border=True):
            section_title("Resource Availability (by Team)")
            _r_team_status = R.assign(team=R.team.fillna("Others").astype(str), status=R.status.fillna("Available").astype(str)) if not R.empty else pd.DataFrame(columns=["team", "status"])
            rat_df = _r_team_status.groupby(["team", "status"]).size().unstack(fill_value=0).rename_axis(None, axis=1).reset_index().rename(columns={"team": "Team"}) if not _r_team_status.empty else pd.DataFrame(columns=["Team", "Available", "Partial", "On Leave"])
            for _col in ["Available", "Partial", "On Leave"]:
                if _col not in rat_df: rat_df[_col] = 0
            rat_df = rat_df[["Team", "Available", "Partial", "On Leave"]]
            fig_rat = px.bar(
                rat_df,
                x="Team",
                y=["Available", "Partial", "On Leave"],
                color_discrete_map={"Available": GREEN, "Partial": ORANGE, "On Leave": RED},
            )
            fig_rat.update_layout(
                height=195,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.48,
            )
            st.plotly_chart(fig_rat, use_container_width=True, config={"displayModeBar": False})

        with rr2_2, st.container(border=True):
            section_title("Resource Assignments (Count)")
            _assignment_detail = df("SELECT pa.resource_id, r.team, COUNT(*) AS assignment_count FROM project_assignments pa LEFT JOIN resources r ON r.id=pa.resource_id GROUP BY pa.resource_id, r.team")
            if _assignment_detail.empty:
                rac_df = pd.DataFrame(columns=["Team", "Primary Assignment", "Additional Assignment"])
            else:
                _assignment_detail["team"] = _assignment_detail["team"].fillna("Others")
                _assignment_detail["Primary Assignment"] = (_assignment_detail["assignment_count"] == 1).astype(int)
                _assignment_detail["Additional Assignment"] = (_assignment_detail["assignment_count"] - 1).clip(lower=0)
                rac_df = _assignment_detail.groupby("team", as_index=False)[["Primary Assignment", "Additional Assignment"]].sum().rename(columns={"team": "Team"})
            fig_rac = px.bar(
                rac_df,
                x="Team",
                y=["Primary Assignment", "Additional Assignment"],
                color_discrete_map={"Primary Assignment": GREEN, "Additional Assignment": "#1D9BF0"},
            )
            fig_rac.update_layout(
                height=195,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.48,
            )
            st.plotly_chart(fig_rac, use_container_width=True, config={"displayModeBar": False})

        rr3_l, rr3_r = st.columns([1.85, 1])
        with rr3_l, st.container(border=True):
            rl_c1, rl_c2, rl_c3 = st.columns([1.8, 1.1, 0.8])
            with rl_c1:
                st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;padding-top:4px'>Resource List</div>", unsafe_allow_html=True)
            with rl_c2:
                rl_q = st.text_input("Search", placeholder="🔍 Search resources...", key="rep_res_q", label_visibility="collapsed")
            with rl_c3:
                rl_st = st.selectbox("Status", ["All Statuses", "Available", "Partial", "On Leave"], key="rep_res_st", label_visibility="collapsed")

            rlist_rows = []
            for _, _res in R.iterrows():
                _rn = str(_res.get("name") or ""); _ri = "".join(x[0] for x in _rn.split()[:2]).upper()
                _rt = T[T.assignee.fillna("").astype(str) == _rn] if not T.empty else pd.DataFrame()
                _next_m = M[(M.team == _res.get("team")) & (pd.to_datetime(M.due_date, errors="coerce").dt.date >= date.today())] if not M.empty else pd.DataFrame()
                _next_due = fmt_date(_next_m.sort_values("due_date").iloc[0].due_date) if not _next_m.empty else "—"
                _rstat = "On Leave" if str(_res.get("status") or "").lower() == "on leave" else ("Partial" if str(_res.get("status") or "").lower() == "partial" else "Available")
                rlist_rows.append((_ri, "#0284C7", _rn, str(_res.get("team") or ""), _rstat, len(set(_rt.project.dropna().astype(str))) if not _rt.empty else 0, len(_rt), _next_due, str(_res.get("notes") or "—")))
            if rl_q:
                rlist_rows = [r for r in rlist_rows if rl_q.lower() in r[2].lower() or rl_q.lower() in r[3].lower()]
            if rl_st != "All Statuses":
                rlist_rows = [r for r in rlist_rows if r[4] == rl_st]

            rl_tr = "".join(
                f"<tr><td><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td><td>{tm}</td><td>{badge_html(st_s)}</td><td style='text-align:center'>{pa}</td><td style='text-align:center'>{tk}</td><td>{um}</td><td style='color:#64748B'>{nt}</td></tr>"
                for ini, abg, nm, tm, st_s, pa, tk, um, nt in rlist_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Name</th><th>Team</th><th>Status</th><th style="text-align:center">Project Assignments</th><th style="text-align:center">Tasks</th><th>Upcoming Milestone</th><th>Notes</th></tr></thead><tbody>{rl_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer(f"Showing all {len(rlist_rows)} filtered resource(s)", [1], 1)

        with rr3_r:
            with st.container(border=True):
                section_title("Upcoming Leave", "View All")
                _leave_events = df("SELECT * FROM calendar_events WHERE lower(event_type)='leave' ORDER BY start_date")
                ul_rows = []
                for _, _leave in _leave_events.iterrows():
                    _leave_names = [x.strip() for x in str(_leave.get("resources") or "").split(",") if x.strip()]
                    for _leave_name in _leave_names:
                        ul_rows.append(("".join(x[0] for x in _leave_name.split()[:2]).upper(), "#0284C7", _leave_name, str(_leave.get("title") or "Leave"), fmt_date(_leave.get("start_date")), fmt_date(_leave.get("end_date"))))
                ul_tr = "".join(
                    f"<tr><td><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td><td>{lt}</td><td>{sd}</td><td>{ed}</td></tr>"
                    for ini, abg, nm, lt, sd, ed in ul_rows
                )
                st.markdown(
                    clean_html(
                        f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Name</th><th>Leave Type</th><th>Start Date</th><th>End Date</th></tr></thead><tbody>{ul_tr}</tbody></table></div>'
                    ),
                    unsafe_allow_html=True,
                )
            with st.container(border=True):
                section_title("Upcoming Assignments", "View All")
                _ua = T.copy()
                _ua["_start"] = pd.to_datetime(_ua.start_date, errors="coerce") if not _ua.empty else pd.Series(dtype="datetime64[ns]")
                _ua = _ua[_ua["_start"].notna()].sort_values("_start").head(5) if not _ua.empty else _ua
                ua_rows = [("".join(x[0] for x in str(r0.get("assignee") or "Unassigned").split()[:2]).upper(), "#0284C7", str(r0.get("assignee") or "Unassigned"), str(r0.get("project") or ""), fmt_date(r0.get("start_date"))) for _, r0 in _ua.iterrows()]
                ua_tr = "".join(
                    f"<tr><td><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td><td>{pr}</td><td>{sd}</td></tr>"
                    for ini, abg, nm, pr, sd in ua_rows
                )
                st.markdown(
                    clean_html(
                        f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Name</th><th>Project</th><th>Start Date</th></tr></thead><tbody>{ua_tr}</tbody></table></div>'
                    ),
                    unsafe_allow_html=True,
                )

    # =========================================================================
    # REPORT TAB 4: TASKS (Exact Match to Image 32)
    # =========================================================================
    elif active_rtab == "Tasks":
        tk1, tk2, tk3, tk4 = st.columns(4)
        _t_status=T.status.fillna("Not Started").astype(str) if not T.empty else pd.Series(dtype=str)
        _t_total=len(T); _t_done=int((_t_status.str.lower()=="completed").sum()); _t_progress=int((_t_status.str.lower()=="in progress").sum()); _t_notstarted=int((_t_status.str.lower()=="not started").sum())
        def _t_pct(n): return f"{round(n/_t_total*100)}%" if _t_total else "0%"
        with tk1: kpi_card_mockup("Total Tasks", str(_t_total), icon_html="📋", icon_bg="#D1FAE5", icon_fg="#00875A")
        with tk2: kpi_card_mockup("Completed Tasks", str(_t_done), sub_text=_t_pct(_t_done), icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with tk3: kpi_card_mockup("In Progress Tasks", str(_t_progress), sub_text=_t_pct(_t_progress), icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with tk4: kpi_card_mockup("Not Started Tasks", str(_t_notstarted), sub_text=_t_pct(_t_notstarted), icon_html="•••", icon_bg="#E2E8F0", icon_fg="#475569")

        st.write("")
        tr1_1, tr1_2, tr1_3 = st.columns(3)
        with tr1_1, st.container(border=True):
            section_title("Task Status")
            render_report_donut_card(
                str(_t_total), "Tasks",
                ["Completed", "In Progress", "Not Started"], [_t_done, _t_progress, _t_notstarted], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", f"{_t_done} ({_t_pct(_t_done)})"), (BLUE, "In Progress", f"{_t_progress} ({_t_pct(_t_progress)})"), ("#94A3B8", "Not Started", f"{_t_notstarted} ({_t_pct(_t_notstarted)})")],
            )
        with tr1_2, st.container(border=True):
            section_title("Tasks by Project")
            tbp_rows = []
            for _, _pr in P.head(5).iterrows():
                _pt = T[T.project_id == _pr.id] if not T.empty else pd.DataFrame()
                _st = _pt.status.fillna("Not Started").astype(str).str.lower() if not _pt.empty else pd.Series(dtype=str)
                _n = len(_pt)
                _done = int((_st == "completed").sum())
                _inprog = int((_st == "in progress").sum())
                _not = max(0, _n - _done - _inprog)
                tbp_rows.append((str(_pr.name), round(_done / _n * 100) if _n else 0, round(_inprog / _n * 100) if _n else 0, round(_not / _n * 100) if _n else 0, _n))
            tbp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 28px;align-items:center;gap:10px;margin-bottom:11px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='display:flex;height:13px;background:#F1F5F9;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{c_w}%;background:#00875A'></div>"
                f"<div style='width:{i_w}%;background:#1473E6'></div>"
                f"<div style='width:{n_w}%;background:#CBD5E1'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{tot}</b></div>"
                for p_nm, c_w, i_w, n_w, tot in tbp_rows
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="padding-top:4px">{tbp_html}</div>
                    <div style="display:flex;justify-content:center;gap:16px;font-size:11.5px;color:#475569;margin-top:8px">
                        <span><span style="color:#00875A">●</span> Completed</span>
                        <span><span style="color:#1473E6">●</span> In Progress</span>
                        <span><span style="color:#94A3B8">●</span> Not Started</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with tr1_3, st.container(border=True):
            section_title("Tasks by Priority")
            _priority = T.priority.fillna("Medium").astype(str).str.title() if not T.empty else pd.Series(dtype=str)
            _high = int((_priority == "High").sum()); _medium = int((_priority == "Medium").sum()); _low = int((_priority == "Low").sum())
            _p_pct = lambda n: f"{round(n / len(T) * 100)}%" if len(T) else "0%"
            render_report_donut_card(
                str(len(T)), "Tasks",
                ["High", "Medium", "Low"], [_high, _medium, _low], [RED, ORANGE, GREEN],
                [(RED, "High", f"{_high} ({_p_pct(_high)})"), (ORANGE, "Medium", f"{_medium} ({_p_pct(_medium)})"), (GREEN, "Low", f"{_low} ({_p_pct(_low)})")],
            )

        tr2_1, tr2_2, tr2_3 = st.columns([1.2, 1.15, 0.85])
        with tr2_1, st.container(border=True):
            section_title("Task Trend")
            if not T.empty:
                _task_trend = T.copy()
                _task_trend["_month"] = pd.to_datetime(_task_trend.get("start_date"), errors="coerce").dt.to_period("M").astype(str)
                _task_trend["_status"] = _task_trend.status.fillna("Not Started").astype(str).str.title()
                _months = sorted([m for m in _task_trend["_month"].unique().tolist() if m and m != "NaT"])[-6:]
                tt_df = _task_trend[_task_trend["_month"].isin(_months)].groupby(["_month", "_status"]).size().unstack(fill_value=0).reindex(columns=["Completed", "In Progress", "Not Started"], fill_value=0).reset_index().rename(columns={"_month": "Month"})
                tt_df["Month"] = pd.to_datetime(tt_df["Month"] + "-01", errors="coerce").dt.strftime("%b %Y")
            else:
                tt_df = pd.DataFrame(columns=["Month", "Completed", "In Progress", "Not Started"])
            fig_tt = px.line(
                tt_df,
                x="Month",
                y=["Completed", "In Progress", "Not Started"],
                markers=True,
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": "#94A3B8"},
            )
            fig_tt.update_layout(
                height=195,
                margin=dict(l=0, r=0, t=4, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_tt, use_container_width=True, config={"displayModeBar": False})

        with tr2_2, st.container(border=True):
            section_title("Tasks by Team")
            if not T.empty:
                _team_tasks = T.assign(team=T.team.fillna("Others").astype(str), status=T.status.fillna("Not Started").astype(str).str.title())
                tbt_df = _team_tasks.groupby(["team", "status"]).size().unstack(fill_value=0).reindex(columns=["Completed", "In Progress", "Not Started"], fill_value=0).reset_index().rename(columns={"team": "Team"})
            else:
                tbt_df = pd.DataFrame(columns=["Team", "Completed", "In Progress", "Not Started"])
            fig_tbt = px.bar(
                tbt_df,
                x="Team",
                y=["Completed", "In Progress", "Not Started"],
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": "#CBD5E1"},
            )
            fig_tbt.update_layout(
                height=195,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.42,
            )
            st.plotly_chart(fig_tbt, use_container_width=True, config={"displayModeBar": False})

        with tr2_3, st.container(border=True):
            section_title("Overdue Tasks")
            _overdue_due = pd.to_datetime(T.due_date, errors="coerce").dt.date if not T.empty else pd.Series(dtype=object)
            _overdue_mask = (_overdue_due.notna() & (_overdue_due < date.today()) & (T.status.fillna("").astype(str).str.lower() != "completed")) if not T.empty else pd.Series(dtype=bool)
            _overdue_report_tasks = T[_overdue_mask].copy() if not T.empty else T.copy()
            _overdue_priority = _overdue_report_tasks.priority.fillna("Medium").astype(str).str.title().value_counts() if not _overdue_report_tasks.empty else pd.Series(dtype=int)
            _overdue_pct = round(len(_overdue_report_tasks) / len(T) * 100) if len(T) else 0
            st.markdown(
                clean_html(
                    f"""
                    <div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:14px;margin-bottom:12px">
                        <div style="width:40px;height:40px;border-radius:50%;background:#FEE2E2;color:#DC2626;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:800">⚠</div>
                        <div style="display:flex;align-items:baseline;gap:10px">
                            <span style="font-size:28px;font-weight:800;color:#DC2626">{len(_overdue_report_tasks)}</span>
                            <div style="font-size:11.5px;color:#B91C1C;line-height:1.3"><b>Overdue Tasks</b><br>{_overdue_pct}% of total</div>
                        </div>
                    </div>
                    <div style="font-size:12px;font-weight:700;color:#0F172A;margin-bottom:6px">By Priority</div>
                    <div style="font-size:12.5px;line-height:2.0">
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#EF4444">●</span> &nbsp;High</span><b>{int(_overdue_priority.get("High", 0))}</b></div>
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;Medium</span><b>{int(_overdue_priority.get("Medium", 0))}</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;Low</span><b>{int(_overdue_priority.get("Low", 0))}</b></div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        tr3_l, tr3_r = st.columns([1.85, 1])
        with tr3_l, st.container(border=True):
            tl_c1, tl_c2, tl_c3 = st.columns([1.8, 1.1, 0.8])
            with tl_c1:
                st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;padding-top:4px'>Task List</div>", unsafe_allow_html=True)
            with tl_c2:
                tk_q = st.text_input("Search", placeholder="🔍 Search tasks...", key="rep_tk_q", label_visibility="collapsed")
            with tl_c3:
                tk_st = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "Not Started"], key="rep_tk_st", label_visibility="collapsed")

            tlist_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), str(r0.get("team") or ""), str(r0.get("status") or "Not Started"), str(r0.get("priority") or "Medium"), fmt_date(r0.get("start_date")), fmt_date(r0.get("due_date")), str(r0.get("assignee") or "Unassigned")) for _, r0 in T.iterrows()]
            if tk_q:
                tlist_rows = [r for r in tlist_rows if tk_q.lower() in r[0].lower() or tk_q.lower() in r[1].lower()]
            if tk_st != "All Statuses":
                tlist_rows = [r for r in tlist_rows if r[3] == tk_st]

            tl_tr = "".join(
                f"<tr><td><b>{tn}</b></td><td>{pn}</td><td>{tm}</td><td>{badge_html(st_s)}</td><td>{badge_html(pr_s)}</td><td>{sd}</td><td>{dd}</td><td>{ass}</td></tr>"
                for tn, pn, tm, st_s, pr_s, sd, dd, ass in tlist_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Task Name</th><th>Project</th><th>Team</th><th>Status</th><th>Priority</th><th>Start Date</th><th>Due Date</th><th>Assignee</th></tr></thead><tbody>{tl_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer(f"Showing all {len(tlist_rows)} filtered task(s)", [1], 1)

        with tr3_r:
            with st.container(border=True):
                section_title("Upcoming Tasks (Next 14 Days)", "View All")
                _ut = T.copy()
                _ut["_due"] = pd.to_datetime(_ut.due_date, errors="coerce") if not _ut.empty else pd.Series(dtype="datetime64[ns]")
                _ut = _ut[(_ut["_due"].dt.date >= date.today()) & (_ut["_due"].dt.date <= date.today()+timedelta(days=14))].sort_values("_due").head(5) if not _ut.empty else _ut
                ut_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), fmt_date(r0.get("due_date")), str(r0.get("priority") or "Medium")) for _, r0 in _ut.iterrows()]
                ut_tr = "".join(
                    f"<tr><td><b>{tn}</b></td><td>{pn}</td><td>{dd}</td><td>{badge_html(pr_s)}</td></tr>"
                    for tn, pn, dd, pr_s in ut_rows
                )
                st.markdown(
                    clean_html(
                        f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Task Name</th><th>Project</th><th>Due Date</th><th>Priority</th></tr></thead><tbody>{ut_tr}</tbody></table></div>'
                    ),
                    unsafe_allow_html=True,
                )
            with st.container(border=True):
                section_title("Recently Completed Tasks", "View All")
                _rc = T[T.status.fillna("").astype(str).str.lower() == "completed"].copy() if not T.empty else pd.DataFrame()
                rc_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), fmt_date(r0.get("due_date"))) for _, r0 in _rc.sort_values("due_date", ascending=False).head(5).iterrows()]
                rc_tr = "".join(
                    f"<tr><td><b>{tn}</b></td><td>{pn}</td><td>{cd}</td></tr>"
                    for tn, pn, cd in rc_rows
                )
                st.markdown(
                    clean_html(
                        f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Task Name</th><th>Project</th><th>Completed Date</th></tr></thead><tbody>{rc_tr}</tbody></table></div>'
                    ),
                    unsafe_allow_html=True,
                )

    # =========================================================================
    # REPORT TAB 5: MILESTONES (Exact Match to Image 35)
    # =========================================================================
    elif active_rtab == "Milestones":
        mk1, mk2, mk3, mk4, mk5 = st.columns(5)
        _m_status=M.status.fillna("Not Started").astype(str) if not M.empty else pd.Series(dtype=str)
        _m_total=len(M); _m_done=int((_m_status.str.lower()=="completed").sum()); _m_prog=int((_m_status.str.lower()=="in progress").sum()); _m_not=int((_m_status.str.lower()=="not started").sum())
        _m_due=pd.to_datetime(M.due_date,errors="coerce") if not M.empty else pd.Series(dtype="datetime64[ns]")
        _m_over=int(((_m_due.dt.date<date.today())&(_m_status.str.lower()!="completed")).sum()) if not M.empty else 0
        def _m_pct(n): return f"{round(n/_m_total*100)}%" if _m_total else "0%"
        with mk1: kpi_card_mockup("Total Milestones", str(_m_total), icon_html="⚑", icon_bg="#D1FAE5", icon_fg="#00875A")
        with mk2: kpi_card_mockup("Completed", str(_m_done), sub_text=_m_pct(_m_done), icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with mk3: kpi_card_mockup("In Progress", str(_m_prog), sub_text=_m_pct(_m_prog), icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with mk4: kpi_card_mockup("Not Started", str(_m_not), sub_text=_m_pct(_m_not), icon_html="🕒", icon_bg="#E2E8F0", icon_fg="#475569")
        with mk5: kpi_card_mockup("Overdue", str(_m_over), sub_text=_m_pct(_m_over), icon_html="⚠", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        mr1_1, mr1_2, mr1_3 = st.columns(3)
        with mr1_1, st.container(border=True):
            section_title("Milestone Status")
            render_report_donut_card(
                str(_m_total), "Milestones",
                ["Completed", "In Progress", "Not Started"], [_m_done, _m_prog, _m_not], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", f"{_m_done} ({_m_pct(_m_done)})"), (BLUE, "In Progress", f"{_m_prog} ({_m_pct(_m_prog)})"), ("#94A3B8", "Not Started", f"{_m_not} ({_m_pct(_m_not)})")],
            )
        with mr1_2, st.container(border=True):
            section_title("Milestones by Project")
            mbp_rows = []
            for _, _pr in P.iterrows():
                _pm = M[M.project_id == _pr.id] if not M.empty else pd.DataFrame()
                _done = int((_pm.status.fillna("").astype(str).str.lower() == "completed").sum()) if not _pm.empty else 0
                _inprog = int((_pm.status.fillna("").astype(str).str.lower() == "in progress").sum()) if not _pm.empty else 0
                _notstarted = int((_pm.status.fillna("").astype(str).str.lower() == "not started").sum()) if not _pm.empty else 0
                mbp_rows.append((str(_pr.name), round(_done/len(_pm)*100) if len(_pm) else 0, round(_inprog/len(_pm)*100) if len(_pm) else 0, round(_notstarted/len(_pm)*100) if len(_pm) else 0, len(_pm)))
            mbp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 24px;align-items:center;gap:10px;margin-bottom:11px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='display:flex;height:13px;background:#F1F5F9;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{c_w}%;background:#00875A'></div>"
                f"<div style='width:{i_w}%;background:#1473E6'></div>"
                f"<div style='width:{n_w}%;background:#CBD5E1'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{tot}</b></div>"
                for p_nm, c_w, i_w, n_w, tot in mbp_rows
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="padding-top:4px">{mbp_html}</div>
                    <div style="display:flex;justify-content:center;gap:16px;font-size:11.5px;color:#475569;margin-top:8px">
                        <span><span style="color:#00875A">●</span> Completed</span>
                        <span><span style="color:#1473E6">●</span> In Progress</span>
                        <span><span style="color:#94A3B8">●</span> Not Started</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with mr1_3, st.container(border=True):
            section_title("Milestones by Due Date (Trend)")
            if not M.empty:
                _milestone_trend = M.copy()
                _milestone_trend["_month"] = pd.to_datetime(_milestone_trend.get("due_date"), errors="coerce").dt.to_period("M").astype(str)
                _milestone_trend["_status"] = _milestone_trend.status.fillna("Not Started").astype(str).str.title()
                _milestone_trend["_overdue"] = (pd.to_datetime(_milestone_trend.get("due_date"), errors="coerce").dt.date < date.today()) & (_milestone_trend["_status"] != "Completed")
                _months = sorted([m for m in _milestone_trend["_month"].unique().tolist() if m and m != "NaT"])[-6:]
                mdd_df = _milestone_trend[_milestone_trend["_month"].isin(_months)].groupby(["_month", "_status"]).size().unstack(fill_value=0).reindex(columns=["Completed", "In Progress", "Not Started"], fill_value=0).reset_index().rename(columns={"_month": "Month"})
                mdd_df["Month"] = pd.to_datetime(mdd_df["Month"] + "-01", errors="coerce").dt.strftime("%b %Y")
                _overdue_counts = _milestone_trend[_milestone_trend["_month"].isin(_months)].groupby("_month")["_overdue"].sum().reset_index(name="Overdue")
                _overdue_counts["Month"] = pd.to_datetime(_overdue_counts["_month"] + "-01", errors="coerce").dt.strftime("%b %Y")
                mdd_df = mdd_df.merge(_overdue_counts[["Month", "Overdue"]], on="Month", how="left").fillna(0)
            else:
                mdd_df = pd.DataFrame(columns=["Month", "Completed", "In Progress", "Not Started", "Overdue"])
            fig_mdd = px.bar(
                mdd_df,
                x="Month",
                y=["Completed", "In Progress", "Not Started", "Overdue"],
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": "#CBD5E1", "Overdue": RED},
            )
            fig_mdd.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.45,
            )
            st.plotly_chart(fig_mdd, use_container_width=True, config={"displayModeBar": False})

        mr2_l, mr2_r = st.columns(2)
        with mr2_l, st.container(border=True):
            section_title("Upcoming Milestones (Next 30 Days)", "View All")
            _um30 = M.copy()
            _um30["_due"] = pd.to_datetime(_um30.due_date, errors="coerce") if not _um30.empty else pd.Series(dtype="datetime64[ns]")
            _um30 = _um30[(_um30["_due"].dt.date >= date.today()) & (_um30["_due"].dt.date <= date.today()+timedelta(days=30))].sort_values("_due").head(5) if not _um30.empty else _um30
            um30_rows = []
            for _, _m0 in _um30.iterrows():
                _left = (_m0["_due"].date() - date.today()).days
                _tone = ("#FEF3C7", "#D97706") if _left <= 7 else ("#DCFCE7", "#15803D")
                um30_rows.append((str(_m0.get("name") or ""), str(_m0.get("project") or ""), fmt_date(_m0.get("due_date")), str(_m0.get("status") or "Not Started"), str(_left), _tone[0], _tone[1]))
            um30_tr = "".join(
                f"<tr><td><b>{mn}</b></td><td>{pn}</td><td>{dd}</td><td>{badge_html(st_s)}</td>"
                f"<td style='text-align:center'><span class='vk-badge' style='background:{dbg};color:{dfg};min-width:44px'>{dl}</span></td></tr>"
                for mn, pn, dd, st_s, dl, dbg, dfg in um30_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Milestone Name</th><th>Project</th><th>Due Date</th><th>Status</th><th style="text-align:center">Days Left</th></tr></thead><tbody>{um30_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

        with mr2_r, st.container(border=True):
            section_title("Overdue Milestones", "View All")
            _om = M.copy()
            _om["_due"] = pd.to_datetime(_om.due_date, errors="coerce") if not _om.empty else pd.Series(dtype="datetime64[ns]")
            _om = _om[(_om["_due"].dt.date < date.today()) & (_om.status.fillna("").astype(str).str.lower() != "completed")].sort_values("_due").head(5) if not _om.empty else _om
            om_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), fmt_date(r0.get("due_date")), (date.today()-pd.to_datetime(r0.get("due_date")).date()).days if pd.notna(pd.to_datetime(r0.get("due_date"),errors="coerce")) else 0, str(r0.get("status") or "Not Started")) for _, r0 in _om.iterrows()]
            om_tr = "".join(
                f"<tr><td><b>{mn}</b></td><td>{pn}</td><td>{dd}</td><td style='text-align:center;color:#DC2626;font-weight:800'>{do}</td><td>{badge_html('Overdue')}</td></tr>"
                for mn, pn, dd, do in om_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Milestone Name</th><th>Project</th><th>Due Date</th><th style="text-align:center">Days Overdue</th><th>Status</th></tr></thead><tbody>{om_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            ml_c1, ml_c2, ml_c3, ml_c4, ml_c5 = st.columns([2.2, 1.1, 0.85, 0.85, 0.85])
            with ml_c1:
                st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;padding-top:4px'>Milestone List</div>", unsafe_allow_html=True)
            with ml_c2:
                ml_q = st.text_input("Search", placeholder="🔍 Search milestones...", key="rep_ms_q", label_visibility="collapsed")
            with ml_c3:
                st.selectbox("Project", ["All Projects"] + projects_list, key="rep_ms_pr", label_visibility="collapsed")
            with ml_c4:
                ml_st = st.selectbox("Status", ["All Statuses", "Completed", "In Progress", "Not Started", "Overdue"], key="rep_ms_st", label_visibility="collapsed")
            with ml_c5:
                st.selectbox("Team", ["All Teams"] + teams_list, key="rep_ms_tm", label_visibility="collapsed")

            mlist_rows = [(str(r0.get("name") or ""), str(r0.get("project") or ""), str(r0.get("team") or ""), str(r0.get("status") or "Not Started"), fmt_date(r0.get("start_date")), fmt_date(r0.get("due_date")), fmt_date(r0.get("completion_date")), str(r0.get("dependency") or "—"), str(r0.get("notes") or "")) for _, r0 in M.iterrows()]
            if ml_q:
                mlist_rows = [r for r in mlist_rows if ml_q.lower() in r[0].lower() or ml_q.lower() in r[1].lower()]
            if ml_st != "All Statuses":
                mlist_rows = [r for r in mlist_rows if r[3] == ml_st]

            ml_tr = "".join(
                f"<tr><td><b>{mn}</b></td><td>{pn}</td><td>{tm}</td><td>{badge_html(st_s)}</td><td>{sd}</td><td>{dd}</td><td>{cd}</td><td>{dp}</td><td style='color:#475569'>{nt}</td></tr>"
                for mn, pn, tm, st_s, sd, dd, cd, dp, nt in mlist_rows
            )
            st.markdown(
                clean_html(
                    f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Milestone Name</th><th>Project</th><th>Team</th><th>Status</th><th>Start Date</th><th>Due Date</th><th>Completion Date</th><th>Dependency</th><th>Notes</th></tr></thead><tbody>{ml_tr}</tbody></table></div>'
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer(f"Showing all {len(mlist_rows)} filtered milestone(s)", [1], 1)

    # =========================================================================
    # REPORT TAB 6: WORKLOAD (Exact Match to Image 27)
    # =========================================================================
    elif active_rtab == "Workload":
        wk1, wk2, wk3, wk4, wk5 = st.columns(5)
        _w_count=len(R); _w_tasks=len(T); _w_assigned=T.assignee.dropna().astype(str).nunique() if not T.empty else 0
        with wk1: kpi_card_mockup("Total Team Members", str(_w_count), icon_html="👥", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with wk2: kpi_card_mockup("Total Active Tasks", str(_w_tasks), icon_html="❖", icon_bg="#D1FAE5", icon_fg="#00875A")
        with wk3: kpi_card_mockup("Assigned Members", str(_w_assigned), sub_text=f"{round(_w_assigned/_w_count*100) if _w_count else 0}%", icon_html="👤", icon_bg="#D1FAE5", icon_fg="#00875A")
        with wk4: kpi_card_mockup("Unassigned Tasks", str(int(T.assignee.isna().sum()) if not T.empty else 0), sub_text=f"{round((int(T.assignee.isna().sum()) if not T.empty else 0)/_w_tasks*100) if _w_tasks else 0}%", icon_html="△", icon_bg="#FEF3C7", icon_fg="#F28C00")
        with wk5: kpi_card_mockup("Overdue Tasks", str(int(((pd.to_datetime(T.due_date, errors="coerce").dt.date < date.today()) & (T.status.fillna("").astype(str).str.lower() != "completed")).sum()) if not T.empty else 0), sub_text="", icon_html="⚠", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        wr1_1, wr1_2, wr1_3 = st.columns(3)
        with wr1_1, st.container(border=True):
            section_title("Workload Distribution")
            _loads = []
            for _, _wr in R.iterrows():
                _wn = str(_wr.get("name") or "")
                _count = int((T.assignee.fillna("").astype(str) == _wn).sum()) if not T.empty else 0
                _loads.append("Overloaded" if _count >= 10 else ("High Workload" if _count >= 5 else "Balanced"))
            _balanced = _loads.count("Balanced"); _high_load = _loads.count("High Workload"); _over_load = _loads.count("Overloaded")
            def _load_pct(n): return f"{round(n/len(R)*100)}%" if len(R) else "0%"
            render_report_donut_card(
                str(len(R)), "Team Members",
                ["Balanced", "High Workload", "Overloaded"], [_balanced, _high_load, _over_load], [GREEN, ORANGE, RED],
                [(GREEN, "Balanced", f"{_balanced} ({_load_pct(_balanced)})"), (ORANGE, "High Workload", f"{_high_load} ({_load_pct(_high_load)})"), (RED, "Overloaded", f"{_over_load} ({_load_pct(_over_load)})")],
            )
        with wr1_2, st.container(border=True):
            section_title("Task Distribution by Team")
            _team_names = sorted(set(T.team.dropna().astype(str).tolist())) if not T.empty else []
            _td_rows = []
            for _team_name in _team_names:
                _team_tasks = T[T.team.astype(str) == _team_name]
                _team_status = _team_tasks.status.fillna("Not Started").astype(str).str.lower()
                _team_prio = _team_tasks.priority.fillna("").astype(str).str.lower()
                _td_rows.append({"Team":_team_name,"Completed":int((_team_status=="completed").sum()),"In Progress":int((_team_status=="in progress").sum()),"Not Started":int((_team_status=="not started").sum()),"High Priority":int((_team_prio=="high").sum())})
            tdt_df = pd.DataFrame(_td_rows,columns=["Team","Completed","In Progress","Not Started","High Priority"])
            fig_tdt = px.bar(
                tdt_df,
                x="Team",
                y=["Completed", "In Progress", "Not Started", "High Priority"],
                color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": "#CBD5E1", "High Priority": ORANGE},
            )
            fig_tdt.update_layout(
                height=210,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, range=[0, 52]),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=10.5)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.45,
            )
            st.plotly_chart(fig_tdt, use_container_width=True, config={"displayModeBar": False})

        with wr1_3, st.container(border=True):
            section_title("Workload by Project")
            wbp_rows = []
            for _, _pr in P.iterrows():
                _pt = T[T.project_id == _pr.id] if not T.empty else pd.DataFrame()
                _status = _pt.status.fillna("Not Started").astype(str).str.lower() if not _pt.empty else pd.Series(dtype=str)
                _total = len(_pt)
                _done = int((_status == "completed").sum()); _prog = int((_status == "in progress").sum()); _not = int((_status == "not started").sum())
                _high = int((_pt.priority.fillna("").astype(str).str.lower() == "high").sum()) if not _pt.empty else 0
                wbp_rows.append((str(_pr.name), _total, _done, _prog, _not, _high))
            wbp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr 28px;align-items:center;gap:10px;margin-bottom:11px;font-size:12.5px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='display:flex;height:13px;background:#F1F5F9;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{c_w}%;background:#00875A'></div>"
                f"<div style='width:{i_w}%;background:#1473E6'></div>"
                f"<div style='width:{n_w}%;background:#CBD5E1'></div>"
                f"<div style='width:{h_w}%;background:#F28C00'></div></div>"
                f"<b style='color:#0F172A;text-align:right'>{tot}</b></div>"
                for p_nm, c_w, i_w, n_w, h_w, tot in wbp_rows
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="padding-top:4px">{wbp_html}</div>
                    <div style="display:flex;justify-content:center;gap:12px;font-size:11px;color:#475569;margin-top:8px">
                        <span><span style="color:#00875A">●</span> Completed</span>
                        <span><span style="color:#1473E6">●</span> In Progress</span>
                        <span><span style="color:#94A3B8">●</span> Not Started</span>
                        <span><span style="color:#F28C00">●</span> High Priority</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        wr2_1, wr2_2, wr2_3 = st.columns(3)
        with wr2_1, st.container(border=True):
            section_title("Team Workload")
            _teams_for_workload = sorted(set(R.team.fillna("Others").astype(str).tolist() if not R.empty else []) | set(T.team.fillna("Others").astype(str).tolist() if not T.empty else []))
            _team_workload_rows = []
            _assignment_team = df("SELECT r.team, SUM(pa.allocation) AS allocation FROM project_assignments pa LEFT JOIN resources r ON r.id=pa.resource_id GROUP BY r.team")
            for _team_name in _teams_for_workload:
                _team_resources = R[R.team.fillna("Others").astype(str) == _team_name] if not R.empty else pd.DataFrame()
                _available_cap = round(pd.to_numeric(_team_resources.get("capacity", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()/100, 1) if not _team_resources.empty else 0
                _team_tasks_count = int((T.team.fillna("Others").astype(str) == _team_name).sum()) if not T.empty else 0
                _team_assigned = round(float(pd.to_numeric(_assignment_team.loc[_assignment_team.team.fillna("Others").astype(str) == _team_name, "allocation"], errors="coerce").fillna(0).sum())/100, 1) if not _assignment_team.empty else 0
                _team_workload_rows.append({"Team": _team_name, "Available Capacity": _available_cap, "Assigned Tasks": _team_tasks_count, "Remaining Capacity": max(0, round(_available_cap-_team_assigned, 1))})
            twc_df = pd.DataFrame(_team_workload_rows, columns=["Team", "Available Capacity", "Assigned Tasks", "Remaining Capacity"])
            fig_twc = px.bar(
                twc_df,
                x="Team",
                y=["Available Capacity", "Assigned Tasks", "Remaining Capacity"],
                color_discrete_map={"Available Capacity": GREEN, "Assigned Tasks": BLUE, "Remaining Capacity": "#CBD5E1"},
            )
            fig_twc.update_layout(
                height=215,
                margin=dict(l=0, r=0, t=6, b=0),
                xaxis_title=None,
                yaxis=dict(title=None, rangemode="tozero"),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=10.5)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.45,
            )
            st.plotly_chart(fig_twc, use_container_width=True, config={"displayModeBar": False})

        with wr2_2, st.container(border=True):
            section_title("Individual Workload (Top 10)")
            iw_items = []
            _resource_loads = []
            for _, _r0 in R.iterrows():
                _rn = str(_r0.get("name") or "")
                _count = int((T.assignee.fillna("").astype(str) == _rn).sum()) if not T.empty else 0
                _resource_loads.append((_rn, _count))
            _max_resource_load = max([x[1] for x in _resource_loads], default=1) or 1
            for _rn, _count in sorted(_resource_loads, key=lambda x:x[1], reverse=True)[:8]:
                iw_items.append((_rn, _count, round(_count/_max_resource_load*100), RED if _count >= 10 else (ORANGE if _count >= 5 else GREEN)))
            iw_html = "".join(
                f"<div style='display:grid;grid-template-columns:100px 1fr 24px;align-items:center;gap:8px;margin-bottom:5px;font-size:11.5px'>"
                f"<span style='color:#1E293B;font-weight:500'>{nm}</span>"
                f"<div style='height:9px;background:#F1F5F9;border-radius:4px;overflow:hidden'><div style='width:{w_p}%;height:100%;background:{bcol};border-radius:4px'></div></div>"
                f"<b style='text-align:right;color:#0F172A'>{cnt}</b></div>"
                for nm, cnt, w_p, bcol in iw_items
            )
            st.markdown(clean_html(iw_html), unsafe_allow_html=True)

        with wr2_3, st.container(border=True):
            section_title("Workload Status by Project")
            wsp_rows = []
            for _, _pr in P.iterrows():
                _pt = T[T.project_id == _pr.id] if not T.empty else pd.DataFrame()
                _st = _pt.status.fillna("Not Started").astype(str).str.lower() if not _pt.empty else pd.Series(dtype=str)
                _n = len(_pt)
                _c = round(int((_st == "completed").sum())/_n*100) if _n else 0
                _i = round(int((_st == "in progress").sum())/_n*100) if _n else 0
                _u = max(0,100-_c-_i)
                wsp_rows.append((str(_pr.name),_c,_i,_u))
            wsp_html = "".join(
                f"<div style='display:grid;grid-template-columns:145px 1fr;align-items:center;gap:10px;margin-bottom:11px;font-size:12px'>"
                f"<span style='color:#1E293B;font-weight:500;text-align:right'>{p_nm}</span>"
                f"<div style='display:flex;height:13px;border-radius:3px;overflow:hidden'>"
                f"<div style='width:{b_w}%;background:#00875A'></div>"
                f"<div style='width:{h_w}%;background:#F28C00'></div>"
                f"<div style='width:{o_w}%;background:#EF4444'></div></div></div>"
                for p_nm, b_w, h_w, o_w in wsp_rows
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="padding-top:4px">{wsp_html}</div>
                    <div style="display:grid;grid-template-columns:145px 1fr;gap:10px;font-size:10.5px;color:#64748B;margin-top:4px">
                        <div></div>
                        <div style="display:flex;justify-content:space-between"><span>0%</span><span>20%</span><span>40%</span><span>60%</span><span>80%</span><span>100%</span></div>
                    </div>
                    <div style="display:flex;justify-content:center;gap:14px;font-size:11.5px;color:#475569;margin-top:8px">
                        <span><span style="color:#00875A">●</span> Balanced</span>
                        <span><span style="color:#F28C00">●</span> High Workload</span>
                        <span><span style="color:#EF4444">●</span> Overloaded</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        wr3_l, wr3_r = st.columns([1.85, 1])
        with wr3_l, st.container(border=True):
            wl_c1, wl_c2, wl_c3, wl_c4 = st.columns([1.6, 1.1, 0.85, 1.0])
            with wl_c1:
                st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;padding-top:4px'>Team Member Workload List</div>", unsafe_allow_html=True)
            with wl_c2:
                wl_q = st.text_input("Search", placeholder="🔍 Search team members...", key="rep_wl_q", label_visibility="collapsed")
            with wl_c3:
                st.selectbox("Team", ["All Teams"] + teams_list, key="rep_wl_tm", label_visibility="collapsed")
            with wl_c4:
                wl_st = st.selectbox("Workload", ["All Workload Status", "Overloaded", "High", "Balanced"], key="rep_wl_st", label_visibility="collapsed")

            wlist_rows = []
            for _, _wr in R.iterrows():
                _wn = str(_wr.get("name") or "")
                _wt = T[T.assignee.fillna("").astype(str) == _wn] if not T.empty else pd.DataFrame()
                _ws = _wt.status.fillna("").astype(str).str.lower() if not _wt.empty else pd.Series(dtype=str)
                _active = len(_wt); _complete = int((_ws == "completed").sum()); _inprog = int((_ws == "in progress").sum()); _notstarted = int((_ws == "not started").sum()); _high = int((_wt.priority.fillna("").astype(str).str.lower() == "high").sum()) if not _wt.empty else 0
                _load = "Overloaded" if _active >= 10 else ("High" if _active >= 5 else "Balanced")
                wlist_rows.append(("".join(x[0] for x in _wn.split()[:2]).upper(), "#0284C7", _wn, str(_wr.get("team") or ""), _active, _complete, _inprog, _notstarted, _high, _load))
            if wl_q:
                wlist_rows = [r for r in wlist_rows if wl_q.lower() in r[2].lower() or wl_q.lower() in r[3].lower()]
            if wl_st != "All Workload Status":
                wlist_rows = [r for r in wlist_rows if r[9] == wl_st]

            wl_badge_map = {
                "Overloaded": "vk-badge-red",
                "High": "vk-badge-orange",
                "Balanced": "vk-badge-green",
            }
            wl_tr = "".join(
                f"<tr><td><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td><td>{tm}</td><td style='text-align:center'><b>{at}</b></td><td style='text-align:center'>{cm}</td><td style='text-align:center'>{ip}</td><td style='text-align:center'>{ns}</td><td style='text-align:center'>{hp}</td>"
                f"<td><span class='vk-badge {wl_badge_map[ws]}' style='min-width:84px'>{ws}</span></td></tr>"
                for ini, abg, nm, tm, at, cm, ip, ns, hp, ws in wlist_rows
            )
            st.markdown(
                f"""<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>Name</th><th>Team</th><th style="text-align:center">Active Tasks</th><th style="text-align:center">Completed</th><th style="text-align:center">In Progress</th><th style="text-align:center">Not Started</th><th style="text-align:center">High Priority</th><th>Workload Status</th></tr></thead><tbody>{wl_tr}</tbody></table></div>""",
                unsafe_allow_html=True,
            )
# =========================================================================
    # REPORT TAB 7: CUSTOM REPORT (Exact Match to Image 33)
    # =========================================================================
    elif active_rtab == "Custom":
        if "custom_viz_type" not in st.session_state:
            st.session_state.custom_viz_type = "Bar Chart"

        # --- SECTION 1: BUILD CUSTOM REPORT ---
        with st.container(border=True):
            b_left, b_right = st.columns([2.15, 1.05])

            with b_left:
                st.markdown(
                    clean_html(
                        """
                        <div style="font-size:16px;font-weight:800;color:#0F172A">Build Custom Report</div>
                        <div style="font-size:12.5px;color:#64748B;margin-bottom:12px">Select the data and options you want to include in your report.</div>
                        """
                    ),
                    unsafe_allow_html=True,
                )

                cfg1, cfg2, cfg3 = st.columns([1, 1, 1.45])
                with cfg1:
                    st.markdown("<div style='font-size:12px;font-weight:700;color:#0F172A;margin-bottom:4px'>Report Type</div>", unsafe_allow_html=True)
                    st.selectbox(
                        "Report Type",
                        ["📁 Project Report", "📋 Task Report", "👥 Resource Report", "⚑ Milestone Report", "📊 Workload Report"],
                        key="cr_rep_type",
                        label_visibility="collapsed",
                    )
                with cfg2:
                    st.markdown("<div style='font-size:12px;font-weight:700;color:#0F172A;margin-bottom:4px'>Group By</div>", unsafe_allow_html=True)
                    st.selectbox(
                        "Group By",
                        ["▤ Project", "👥 Team", "● Status", "⚑ Priority", "👤 Assignee"],
                        key="cr_group_by",
                        label_visibility="collapsed",
                    )
                with cfg3:
                    st.markdown("<div style='font-size:12px;font-weight:700;color:#0F172A;margin-bottom:4px'>Metrics / Data to Include</div>", unsafe_allow_html=True)
                    st.multiselect(
                        "Metrics / Data to Include",
                        ["Project Status", "Task Status", "Milestone Status", "Team Members", "Progress", "Priority", "Overdue"],
                        default=["Project Status", "Task Status", "Milestone Status", "Team Members"],
                        key="cr_metrics",
                        label_visibility="collapsed",
                    )

                st.markdown("<div style='font-size:12px;font-weight:700;color:#0F172A;margin:8px 0 4px 0'>Filters</div>", unsafe_allow_html=True)
                fl1, fl2, fl3, fl4 = st.columns([1, 1, 1, 1.35])
                with fl1:
                    cr_f_proj = st.selectbox("Filter Project", ["All Projects"] + projects_list, key="cr_flt_proj", label_visibility="collapsed")
                with fl2:
                    cr_f_team = st.selectbox("Filter Team", ["All Teams"] + teams_list, key="cr_flt_team", label_visibility="collapsed")
                with fl3:
                    cr_f_stat = st.selectbox("Filter Status", ["All Statuses", "On Track", "At Risk", "Delayed"], key="cr_flt_stat", label_visibility="collapsed")
                with fl4:
                    st.date_input("Filter Date Range", (date.today().replace(day=1), date.today()+timedelta(days=30)), key="cr_flt_dates", label_visibility="collapsed")

                st.write("")
                act_b1, act_b2, _ = st.columns([0.95, 0.7, 2.6])
                with act_b1:
                    if st.button("▷ Generate Report", type="primary", key="cr_gen_btn", use_container_width=True):
                        st.session_state.custom_generated = True
                        st.rerun()
                with act_b2:
                    if st.button("↻ Reset", key="cr_reset_btn", use_container_width=True):
                        st.session_state.custom_viz_type = "Bar Chart"
                        st.session_state["cr_flt_proj"] = "All Projects"
                        st.session_state["cr_flt_team"] = "All Teams"
                        st.session_state["cr_flt_stat"] = "All Statuses"
                        st.rerun()

            with b_right, st.container(border=True):
                st.markdown("<div style='font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:10px'>Visualization Type</div>", unsafe_allow_html=True)
                viz_row1 = [("📊", "Bar Chart"), ("📊", "Column Chart"), ("◔", "Pie Chart")]
                viz_row2 = [("📈", "Line Chart"), ("⊞", "Table"), ("▤", "Summary Cards")]

                vc1 = st.columns(3)
                for col, (v_ico, v_lbl) in zip(vc1, viz_row1):
                    with col:
                        is_sel = st.session_state.custom_viz_type == v_lbl
                        if st.button(
                            f"{v_ico}\n{v_lbl}",
                            key=f"cr_viz_{v_lbl}",
                            type="primary" if is_sel else "secondary",
                            use_container_width=True,
                        ):
                            st.session_state.custom_viz_type = v_lbl
                            st.rerun()

                vc2 = st.columns(3)
                for col, (v_ico, v_lbl) in zip(vc2, viz_row2):
                    with col:
                        is_sel = st.session_state.custom_viz_type == v_lbl
                        if st.button(
                            f"{v_ico}\n{v_lbl}",
                            key=f"cr_viz_{v_lbl}",
                            type="primary" if is_sel else "secondary",
                            use_container_width=True,
                        ):
                            st.session_state.custom_viz_type = v_lbl
                            st.rerun()

        st.write("")
        # --- DATASET FOR CUSTOM REPORT PREVIEW & TABLE VIEW ---
        _custom_rows = []
        for _, _pr in P.iterrows():
            _pt = T[T.project_id == _pr.id] if not T.empty else pd.DataFrame()
            _pm = M[M.project_id == _pr.id] if not M.empty else pd.DataFrame()
            _ts = _pt.status.fillna("Not Started").astype(str).str.lower() if not _pt.empty else pd.Series(dtype=str)
            _due = pd.to_datetime(_pt.due_date, errors="coerce") if not _pt.empty else pd.Series(dtype="datetime64[ns]")
            _over = int(((_due.dt.date < date.today()) & (_ts != "completed")).sum()) if not _pt.empty else 0
            _team_members = _pt.assignee.dropna().astype(str).nunique() if not _pt.empty else 0
            _custom_rows.append({"Project Name":str(_pr.name),"Project Status":str(_pr.status or "Not Started"),"Total Tasks":len(_pt),"Completed":int((_ts=="completed").sum()),"In Progress":int((_ts=="in progress").sum()),"Not Started":int((_ts=="not started").sum()),"Overdue":_over,"Total Milestones":len(_pm),"Team Members":_team_members})
        custom_proj_df = pd.DataFrame(_custom_rows,columns=["Project Name","Project Status","Total Tasks","Completed","In Progress","Not Started","Overdue","Total Milestones","Team Members"])

        if cr_f_proj != "All Projects" and cr_f_proj in custom_proj_df["Project Name"].values:
            custom_proj_df = custom_proj_df[custom_proj_df["Project Name"] == cr_f_proj].copy()
        if cr_f_stat != "All Statuses":
            custom_proj_df = custom_proj_df[custom_proj_df["Project Status"] == cr_f_stat].copy()

        # --- SECTION 2: REPORT PREVIEW ---
        with st.container(border=True):
            rp_h1, rp_h2, rp_h3 = st.columns([2.8, 0.25, 0.85])
            with rp_h1:
                st.markdown(
                    clean_html(
                        """
                        <div style="font-size:16px;font-weight:800;color:#0F172A">Report Preview</div>
                        <div style="font-size:12.5px;color:#64748B">Project status and task distribution by project.</div>
                        """
                    ),
                    unsafe_allow_html=True,
                )
            with rp_h2:
                st.markdown("<div style='font-size:12.5px;color:#475569;padding-top:8px;text-align:right'>Show:</div>", unsafe_allow_html=True)
            with rp_h3:
                st.selectbox(
                    "Show",
                    ["Project Status", "Task Status", "Milestone Status", "Team Members"],
                    key="cr_preview_show",
                    label_visibility="collapsed",
                )

            pv_l, pv_r = st.columns([2.15, 1.05])
            with pv_l:
                sel_viz = st.session_state.custom_viz_type
                if sel_viz in ("Bar Chart", "Column Chart"):
                    fig_cr = go.Figure()
                    series_cfg = [
                        ("Completed", GREEN),
                        ("In Progress", BLUE),
                        ("Not Started", ORANGE),
                        ("Overdue", RED),
                    ]
                    for s_name, s_col in series_cfg:
                        vals = custom_proj_df[s_name].tolist()
                        fig_cr.add_trace(
                            go.Bar(
                                name=s_name,
                                x=custom_proj_df["Project Name"],
                                y=vals,
                                text=[str(v) if v > 0 else "" for v in vals],
                                textposition="inside",
                                insidetextanchor="middle",
                                textfont=dict(color="white", size=11),
                                marker=dict(color=s_col, line=dict(color="white", width=1)),
                            )
                        )
                    fig_cr.update_layout(
                        barmode="stack" if sel_viz == "Bar Chart" else "group",
                        height=265,
                        margin=dict(l=0, r=10, t=10, b=30),
                        xaxis=dict(title=None, tickfont=dict(size=12, color="#1E293B")),
                        yaxis=dict(title=None, range=[0, 52], dtick=10, gridcolor="#EDF2F7"),
                        legend=dict(orientation="h", yanchor="top", y=-0.16, xanchor="center", x=0.5, font=dict(size=12)),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        bargap=0.38,
                    )
                    st.plotly_chart(fig_cr, use_container_width=True, config={"displayModeBar": False})

                elif sel_viz == "Pie Chart":
                    totals = [
                        int(custom_proj_df["Completed"].sum()),
                        int(custom_proj_df["In Progress"].sum()),
                        int(custom_proj_df["Not Started"].sum()),
                        int(custom_proj_df["Overdue"].sum()),
                    ]
                    fig_cr = px.pie(
                        names=["Completed", "In Progress", "Not Started", "Overdue"],
                        values=totals,
                        hole=0.55,
                        color=["Completed", "In Progress", "Not Started", "Overdue"],
                        color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": ORANGE, "Overdue": RED},
                    )
                    fig_cr.update_layout(height=265, margin=dict(l=0, r=0, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_cr, use_container_width=True, config={"displayModeBar": False})

                elif sel_viz == "Line Chart":
                    fig_cr = px.line(
                        custom_proj_df,
                        x="Project Name",
                        y=["Completed", "In Progress", "Not Started", "Overdue"],
                        markers=True,
                        color_discrete_map={"Completed": GREEN, "In Progress": BLUE, "Not Started": ORANGE, "Overdue": RED},
                    )
                    fig_cr.update_layout(
                        height=265,
                        margin=dict(l=0, r=10, t=10, b=30),
                        xaxis_title=None,
                        yaxis_title=None,
                        legend_title_text="",
                        legend=dict(orientation="h", yanchor="top", y=-0.16, xanchor="center", x=0.5),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(fig_cr, use_container_width=True, config={"displayModeBar": False})

                else:
                    st.dataframe(custom_proj_df, hide_index=True, use_container_width=True)

            with pv_r:
                st.markdown(
                    clean_html(
                        f"""
                        <div style="font-size:14.5px;font-weight:800;color:#0F172A;margin-bottom:10px;padding-top:4px">Summary</div>
                        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:20px">📁</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Projects</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">{len(P)}</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:20px">☰</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Tasks</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">{len(T)}</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:20px">⚑</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Milestones</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">{len(M)}</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:20px">👥</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Team Members</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">{len(R)}</div>
                                </div>
                            </div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )

        st.write("")
        # --- SECTION 3: REPORT DATA (TABLE VIEW) & SAVED REPORTS ---
        bot_l, bot_r = st.columns([1.85, 1.05])

        with bot_l, st.container(border=True):
            th_1, th_2, th_3 = st.columns([1.65, 1.0, 0.85])
            with th_1:
                st.markdown("<div style='font-size:15.5px;font-weight:800;color:#0F172A;padding-top:6px'>Report Data (Table View)</div>", unsafe_allow_html=True)
            with th_2:
                cr_tbl_q = st.text_input("Search", placeholder="🔍 Search...", key="cr_tbl_search", label_visibility="collapsed")
            with th_3:
                st.download_button(
                    "📗 Export to Excel ▾",
                    excel_bytes({"Custom Report": custom_proj_df}),
                    file_name="vektor_custom_report_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="cr_tbl_excel_btn",
                    use_container_width=True,
                )

            tbl_df = custom_proj_df.copy()
            if cr_tbl_q:
                tbl_df = tbl_df[tbl_df["Project Name"].str.contains(cr_tbl_q, case=False, na=False)]

            cr_tr = "".join(
                f"<tr>"
                f"<td><b>{html.escape(r['Project Name'])}</b></td>"
                f"<td>{badge_html(r['Project Status'])}</td>"
                f"<td style='text-align:center'><b>{r['Total Tasks']}</b></td>"
                f"<td style='text-align:center'>{r['Completed']}</td>"
                f"<td style='text-align:center'>{r['In Progress']}</td>"
                f"<td style='text-align:center'>{r['Not Started']}</td>"
                f"<td style='text-align:center;color:#DC2626;font-weight:700'>{r['Overdue']}</td>"
                f"<td style='text-align:center'>{r['Total Milestones']}</td>"
                f"<td style='text-align:center'>{r['Team Members']}</td>"
                f"</tr>"
                for _, r in tbl_df.iterrows()
            )
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-table-wrap">
                        <table class="vk-table">
                            <thead>
                                <tr>
                                    <th>Project Name</th>
                                    <th>Project Status</th>
                                    <th style="text-align:center">Total Tasks</th>
                                    <th style="text-align:center">Completed</th>
                                    <th style="text-align:center">In Progress</th>
                                    <th style="text-align:center">Not Started</th>
                                    <th style="text-align:center">Overdue</th>
                                    <th style="text-align:center">Total Milestones</th>
                                    <th style="text-align:center">Team Members</th>
                                </tr>
                            </thead>
                            <tbody>{cr_tr}</tbody>
                        </table>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer(f"Showing all {len(tbl_df)} filtered project(s)", [1], 1)

        with bot_r, st.container(border=True):
            sr_h1, sr_h2 = st.columns([1.4, 1.1])
            with sr_h1:
                st.markdown("<div style='font-size:15.5px;font-weight:800;color:#0F172A;padding-top:6px'>Saved Reports</div>", unsafe_allow_html=True)
            with sr_h2:
                if st.button("＋ Save Current Report", type="primary", key="cr_save_cur_btn", use_container_width=True):
                    saved_name = f"{active_rtab} Summary — {date.today().isoformat()}"
                    report_payload = {
                        "tab": active_rtab,
                        "project_filter": st.session_state.get("cr_flt_proj", "All Projects"),
                        "team_filter": st.session_state.get("cr_flt_team", "All Teams"),
                        "status_filter": st.session_state.get("cr_flt_stat", "All Statuses"),
                        "metrics": st.session_state.get("cr_metrics", []),
                        "visualization": st.session_state.get("custom_viz_type", "Bar Chart"),
                        "data": custom_proj_df.where(pd.notna(custom_proj_df), None).to_dict(orient="records"),
                    }
                    sql(
                        "INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)",
                        (saved_name, active_rtab, json.dumps(report_payload, default=str), datetime.now().isoformat(timespec="seconds"), str(st.session_state.get("current_user_name", "Workspace User"))),
                    )
                    log_event("Saved", "Custom Report", None, saved_name)
                    st.toast("Current report saved.")
                    st.rerun()

            saved_reports_df = df("SELECT * FROM saved_reports ORDER BY id DESC")
            if saved_reports_df.empty:
                st.info("No saved reports yet. Use ‘Save Current Report’ to save this report configuration and its current data.")
            else:
                st.caption(f"{len(saved_reports_df)} saved report(s)")
                for _, saved_row in saved_reports_df.iterrows():
                    report_id = int(saved_row.get("id") or 0)
                    report_name = str(saved_row.get("name") or f"Report {report_id}")
                    report_type = str(saved_row.get("report_type") or "Custom")
                    report_date = fmt_date(str(saved_row.get("last_generated") or "")[:10])
                    report_owner = str(saved_row.get("created_by") or "—")
                    try:
                        payload = json.loads(str(saved_row.get("config") or "{}"))
                    except Exception:
                        payload = {}
                    saved_data = payload.get("data", []) if isinstance(payload, dict) else []
                    saved_frame = pd.DataFrame(saved_data)
                    with st.container(border=True):
                        info_col, view_col, export_col, delete_col = st.columns([2.6, 0.75, 0.85, 0.75])
                        with info_col:
                            st.markdown(f"**{html.escape(report_name)}**")
                            st.caption(f"{html.escape(report_type)} · {report_date} · {html.escape(report_owner)}")
                        with view_col:
                            if st.button("View", key=f"saved_report_view_{report_id}", use_container_width=True):
                                st.session_state["saved_report_view_id"] = report_id
                                st.rerun()
                        with export_col:
                            export_frame = saved_frame if not saved_frame.empty else pd.DataFrame({"Message": ["This saved report has no stored table data."]})
                            st.download_button(
                                "Export", excel_bytes({"Saved Report": export_frame}),
                                file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', report_name).strip('_') or 'saved_report'}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"saved_report_export_{report_id}", use_container_width=True,
                            )
                        with delete_col:
                            if st.button("Delete", key=f"saved_report_delete_{report_id}", use_container_width=True):
                                st.session_state["saved_report_delete_id"] = report_id
                                st.rerun()

                view_report_id = st.session_state.get("saved_report_view_id")
                if view_report_id is not None:
                    chosen = saved_reports_df[saved_reports_df["id"] == view_report_id]
                    if not chosen.empty:
                        st.markdown("**Saved report preview**")
                        try:
                            saved_payload = json.loads(str(chosen.iloc[0].get("config") or "{}"))
                        except Exception:
                            saved_payload = {}
                        saved_view_df = pd.DataFrame(saved_payload.get("data", [])) if isinstance(saved_payload, dict) else pd.DataFrame()
                        if saved_view_df.empty:
                            st.info("This saved report has no stored table data.")
                        else:
                            st.dataframe(saved_view_df, hide_index=True, use_container_width=True)
                    if st.button("Close preview", key="saved_report_close_preview"):
                        st.session_state.pop("saved_report_view_id", None)
                        st.rerun()

                delete_report_id = st.session_state.get("saved_report_delete_id")
                if delete_report_id is not None:
                    confirm_col, cancel_col = st.columns([1, 1])
                    with confirm_col:
                        if st.button("Confirm delete saved report", type="primary", key="saved_report_confirm_delete"):
                            sql("DELETE FROM saved_reports WHERE id=?", (int(delete_report_id),))
                            log_event("Deleted", "Saved Report", int(delete_report_id), "Saved report deleted")
                            st.session_state.pop("saved_report_delete_id", None)
                            st.session_state.pop("saved_report_view_id", None)
                            st.rerun()
                    with cancel_col:
                        if st.button("Cancel delete", key="saved_report_cancel_delete"):
                            st.session_state.pop("saved_report_delete_id", None)
                            st.rerun()


# ===========================================================================
# PAGE 8: SETTINGS (Exact Single-Screen Dashboard Replica of Image 34)
# ===========================================================================
elif page == "Settings":

    @st.dialog("Add User")
    def add_user_settings_modal() -> None:
        st.markdown("<div style='font-size:18px;font-weight:800;color:#0F172A;margin-bottom:8px'>＋ Add User</div>", unsafe_allow_html=True)
        u_nm = st.text_input("Full Name *", key="set_add_u_nm")
        u_em = st.text_input("Email *", placeholder="name@hpe.com", key="set_add_u_em")
        c1, c2 = st.columns(2)
        with c1:
            u_rl = st.selectbox("Role", ["Admin", "Project Manager", "Team Member"], index=2, key="set_add_u_rl")
        with c2:
            u_tm = st.selectbox("Team", teams_list or ["Others"], key="set_add_u_tm")
        b1, b2 = st.columns(2)
        with b1:
            if st.button("Cancel", key="set_add_u_can", use_container_width=True):
                st.session_state.pop("settings_modal", None)
                st.rerun()
        with b2:
            if st.button("Save User", type="primary", key="set_add_u_sav", use_container_width=True):
                if u_nm.strip():
                    em = u_em.strip() or f"{u_nm.strip().lower().replace(' ', '.')}@hpe.com"
                    sql(
                        "INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                        (u_nm.strip(), em, u_tm, u_rl, "Available", 100, ""),
                    )
                    log_event("Created", "User", None, u_nm.strip())
                st.session_state.pop("settings_modal", None)
                st.rerun()

    @st.dialog("Edit User")
    def edit_user_settings_modal(user_id: int) -> None:
        user_df = df("SELECT * FROM resources WHERE id=?", (user_id,))
        if user_df.empty:
            st.warning("This user no longer exists.")
            if st.button("Close", key="set_edit_user_missing_close"):
                st.session_state.pop("settings_edit_user_id", None)
                st.rerun()
            return
        user = user_df.iloc[0]
        name = st.text_input("Full Name *", value=str(user.get("name") or ""), key=f"set_edit_user_name_{user_id}")
        email = st.text_input("Email *", value=str(user.get("email") or ""), key=f"set_edit_user_email_{user_id}")
        roles = ["Admin", "Project Manager", "Team Member"]
        role_default = str(user.get("role") or "Team Member")
        role = st.selectbox("Role", roles, index=roles.index(role_default) if role_default in roles else 2, key=f"set_edit_user_role_{user_id}")
        teams = sorted(set(teams_list + [str(user.get("team") or "Others")])) or ["Others"]
        team_default = str(user.get("team") or "Others")
        team = st.selectbox("Team", teams, index=teams.index(team_default) if team_default in teams else 0, key=f"set_edit_user_team_{user_id}")
        statuses = ["Available", "Partial", "On Leave", "Inactive"]
        status_default = str(user.get("status") or "Available")
        status = st.selectbox("Status", statuses, index=statuses.index(status_default) if status_default in statuses else 0, key=f"set_edit_user_status_{user_id}")
        capacity = st.number_input("Available Capacity (%)", min_value=0, max_value=100, value=max(0,min(100,int(user.get("capacity") or 0))), key=f"set_edit_user_capacity_{user_id}")
        skills = st.text_input("Skills (comma-separated)", value=str(user.get("skills") or ""), key=f"set_edit_user_skills_{user_id}")
        notes = st.text_area("Notes", value=str(user.get("notes") or ""), key=f"set_edit_user_notes_{user_id}")
        save_col, cancel_col = st.columns(2)
        with save_col:
            if st.button("Save Changes", type="primary", key=f"set_edit_user_save_{user_id}", use_container_width=True):
                if not name.strip() or not email.strip():
                    st.error("Name and email are required.")
                else:
                    sql("UPDATE resources SET name=?,email=?,role=?,team=?,status=?,capacity=?,skills=?,notes=? WHERE id=?", (name.strip(),email.strip(),role,team,status,int(capacity),skills.strip(),notes.strip(),user_id))
                    log_event("Updated", "User", user_id, name.strip())
                    st.session_state.pop("settings_edit_user_id", None)
                    st.rerun()
        with cancel_col:
            if st.button("Cancel", key=f"set_edit_user_cancel_{user_id}", use_container_width=True):
                st.session_state.pop("settings_edit_user_id", None)
                st.rerun()

    st.markdown(
        clean_html(
            f"""
            <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Settings</div>
            <div style='font-size:13.5px;color:#64748B;margin-top:2px;margin-bottom:14px'>Manage workspace, users, defaults, notifications, integrations, and system preferences.</div>
            """
        ),
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown("**Excel data source**")
        st.caption("Upload a workbook to replace data in the matching sheets. Use the template to keep sheet names and columns consistent.")
        ex1,ex2=st.columns([2,1])
        with ex1: uploaded_data_book=st.file_uploader("Upload Vektor data workbook (.xlsx)",type=["xlsx"],key="settings_data_workbook")
        with ex2: st.download_button("Download Excel Template",data=open(Path(__file__).resolve().with_name("vektor_data_template_automated_final.xlsx"),"rb").read() if Path(__file__).resolve().with_name("vektor_data_template_automated_final.xlsx").exists() else excel_bytes({"Projects":pd.DataFrame(columns=project_columns),"Tasks":pd.DataFrame(columns=task_columns),"Subtasks":pd.DataFrame(columns=["id","task_id","name","completed","due_date","created_at"]),"Resources":pd.DataFrame(columns=resource_columns),"Milestones":pd.DataFrame(columns=milestone_columns),"Calendar Events":pd.DataFrame(columns=calendar_event_columns),"Project Assignments":pd.DataFrame(columns=assignment_columns),"Saved Reports":pd.DataFrame(columns=report_columns),"Settings":pd.DataFrame(columns=["key","value"])}),file_name="vektor_data_template.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.download_button("Export Current Data", data=tracker_excel_bytes({"Projects": df("SELECT * FROM projects"), "Tasks": df("SELECT * FROM tasks"), "Subtasks": df("SELECT * FROM subtasks"), "Resources": df("SELECT * FROM resources"), "Milestones": df("SELECT * FROM milestones"), "Calendar Events": df("SELECT * FROM calendar_events"), "Project Assignments": df("SELECT * FROM project_assignments"), "Saved Reports": df("SELECT * FROM saved_reports"), "Settings": df("SELECT * FROM settings")}), file_name="vektor_current_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if uploaded_data_book is not None and st.button("Import Excel and Replace Current Data",type="primary",key="settings_import_excel"):
            try:
                uploaded_data_book.seek(0)
                imported=import_excel_workbook(uploaded_data_book)
                st.success("Imported: "+", ".join(f"{k.replace('_',' ').title()} {v}" for k,v in imported.items()))
                st.rerun()
            except Exception as exc: st.error(f"Import failed; existing data was not replaced: {exc}")

    # --- ROW 1: WORKSPACE SETTINGS (Left) | USER MANAGEMENT (Right) ---
    s_r1_l, s_r1_r = st.columns([1.05, 1.35])

    with s_r1_l, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:22px;color:#0F172A">🏢</span>
                    <div>
                        <div style="font-size:16px;font-weight:800;color:#0F172A;line-height:1.15">Workspace Settings</div>
                        <div style="font-size:12px;color:#64748B">Set up your workspace details and preferences.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        ws_c1, ws_c2 = st.columns(2)
        with ws_c1:
            ws_name_in = st.text_input("Workspace Name", value=get_setting("workspace_name", "Vektor"), key="set_ws_name")
            ws_desc_in = st.text_area(
                "Description",
                value=get_setting("workspace_description", "Project management workspace for planning, tracking, and delivery."),
                height=86,
                key="set_ws_desc",
            )
            st.markdown("<div style='font-size:12px;font-weight:700;color:#0F172A;margin-bottom:6px'>Workspace Logo</div>", unsafe_allow_html=True)
            st.markdown(
                clean_html(
                    """
                    <div style="border:1.5px dashed #CBD5E1;border-radius:10px;padding:14px 12px;background:#F8FAFC;display:flex;align-items:center;justify-content:space-between;gap:10px">
                        <div style="font-size:23px;font-weight:800;color:#0F172A;letter-spacing:-0.8px;white-space:nowrap">Vektor <span style="color:#00875A">❯❯</span></div>
                        <div style="font-size:11px;color:#64748B;line-height:1.35">Click to upload logo<br>(Recommended 200 x 60)</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            _logo_upload = st.file_uploader("Upload workspace logo", type=["png", "jpg", "jpeg"], key="set_workspace_logo_upload", label_visibility="collapsed")
            if _logo_upload is not None:
                _new_logo_b64 = base64.b64encode(_logo_upload.getvalue()).decode("ascii")
                if _new_logo_b64 != get_setting("workspace_logo_b64", ""):
                    put_setting("workspace_logo_b64", _new_logo_b64)
                    st.success("Workspace logo saved.")

        with ws_c2:
            theme_opts = ["🟩 HPE (Green)", "🟦 Blue", "⬛ Dark"]
            cur_theme_val = get_setting("theme", "HPE Green")
            t_idx = 1 if cur_theme_val == "Blue" else (2 if cur_theme_val == "Dark" else 0)
            sel_theme = st.selectbox("Theme", theme_opts, index=t_idx, key="set_ws_theme")
            mapped_theme = "Blue" if "Blue" in sel_theme else ("Dark" if "Dark" in sel_theme else "HPE Green")
            if mapped_theme != cur_theme_val:
                put_setting("theme", mapped_theme)
                st.rerun()

            st.selectbox("Language", ["English", "Filipino"], index=0 if get_setting("language", "English") == "English" else 1, key="set_ws_lang")
            st.selectbox("Date Format", ["MMM DD, YYYY", "DD/MM/YYYY", "MM/DD/YYYY"], index=["MMM DD, YYYY", "DD/MM/YYYY", "MM/DD/YYYY"].index(get_setting("date_format", "MMM DD, YYYY")) if get_setting("date_format", "MMM DD, YYYY") in ["MMM DD, YYYY", "DD/MM/YYYY", "MM/DD/YYYY"] else 0, key="set_ws_dfmt")
            land_opts = ["Overview", "Projects", "Gantt Chart", "Task List", "Resources", "Calendar", "Reports"]
            cur_land = get_setting("default_landing", "Overview")
            sel_land = st.selectbox("Default Landing Page", land_opts, index=land_opts.index(cur_land) if cur_land in land_opts else 0, key="set_ws_land")
            if sel_land != cur_land:
                put_setting("default_landing", sel_land)
            if st.session_state.get("set_ws_lang") != get_setting("language", "English"):
                put_setting("language", st.session_state.get("set_ws_lang", "English"))
            if st.session_state.get("set_ws_dfmt") != get_setting("date_format", "MMM DD, YYYY"):
                put_setting("date_format", st.session_state.get("set_ws_dfmt", "MMM DD, YYYY"))
            if ws_name_in != get_setting("workspace_name", "Vektor"):
                put_setting("workspace_name", ws_name_in)
            if ws_desc_in != get_setting("workspace_description", ""):
                put_setting("workspace_description", ws_desc_in)

    with s_r1_r, st.container(border=True):
        um_h1, um_h2, um_h3 = st.columns([1.65, 1.05, 0.65])
        with um_h1:
            st.markdown(
                clean_html(
                    """
                    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
                        <span style="font-size:22px;color:#0F172A">👥</span>
                        <div>
                            <div style="font-size:16px;font-weight:800;color:#0F172A;line-height:1.15">User Management</div>
                            <div style="font-size:12px;color:#64748B">Manage team members, roles, and access permissions.</div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with um_h2:
            um_q = st.text_input("Search users", placeholder="🔍 Search users...", key="set_um_q", label_visibility="collapsed")
        with um_h3:
            if st.button("＋ Add User", type="primary", key="set_um_add_btn", use_container_width=True):
                st.session_state.settings_modal = "add_user"
                st.rerun()

        um_users = []
        live_users = df("SELECT * FROM resources ORDER BY name")
        for _, u in live_users.iterrows():
            initials = "".join(part[0] for part in str(u.get("name", "")).split()[:2]).upper() or "U"
            active = "Inactive" if str(u.get("status", "")).lower() in {"inactive", "on leave"} else "Active"
            um_users.append((initials, "#0284C7", str(u.get("name", "")), str(u.get("email", "")), str(u.get("role", "Team Member")), str(u.get("team", "")), active))
        if um_q:
            um_users = [u for u in um_users if um_q.lower() in u[2].lower() or um_q.lower() in u[3].lower() or um_q.lower() in u[4].lower()]

        um_tr = "".join(
            f"<tr>"
            f"<td style='padding:7px 10px'><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td>"
            f"<td style='padding:7px 10px;color:#334155'>{em}</td>"
            f"<td style='padding:7px 10px'>{rl}</td>"
            f"<td style='padding:7px 10px'>{tm}</td>"
            f"<td style='padding:7px 10px'><span class='vk-badge {'vk-badge-green' if st_s == 'Active' else 'vk-badge-red'}' style='min-width:64px'>{st_s}</span></td>"
            f"<td style='padding:7px 10px;white-space:nowrap;color:#64748B'>Actions below</td></tr>"
            for ini, abg, nm, em, rl, tm, st_s in um_users
        )
        st.markdown(
            clean_html(
                f"""
                <div class="vk-table-wrap">
                    <table class="vk-table">
                        <thead>
                            <tr><th>Name</th><th>Email</th><th>Role</th><th>Team</th><th>Status</th><th>Actions</th></tr>
                        </thead>
                        <tbody>{um_tr}</tbody>
                    </table>
                </div>
                <div style="padding-top:10px;font-size:12px;color:#475569">Showing all {len(um_users)} visible users</div>
                """
            ),
            unsafe_allow_html=True,
        )

    if not live_users.empty:
        st.caption("Edit or delete a resource record using these row actions.")
        for _, user_row in live_users.iterrows():
            u1, u2, u3 = st.columns([4.2, 0.55, 0.55])
            with u1:
                st.markdown(f"<div style='padding:6px 0;font-size:12px;color:#334155'>{html.escape(str(user_row.get('name') or ''))} · {html.escape(str(user_row.get('email') or ''))}</div>", unsafe_allow_html=True)
            with u2:
                if st.button("✎", key=f"settings_edit_user_{int(user_row.id)}", help="Edit user"):
                    st.session_state["settings_edit_user_id"] = int(user_row.id)
                    st.rerun()
            with u3:
                if st.button("🗑", key=f"settings_delete_user_{int(user_row.id)}", help="Delete user"):
                    sql("DELETE FROM project_assignments WHERE resource_id=?", (int(user_row.id),))
                    sql("DELETE FROM resources WHERE id=?", (int(user_row.id),))
                    log_event("Deleted", "User", int(user_row.id), str(user_row.get("name") or ""))
                    st.rerun()

    st.write("")
    # --- ROW 2: PROJECT SETTINGS | TASK SETTINGS | MILESTONE SETTINGS ---
    s_r2_1, s_r2_2, s_r2_3 = st.columns(3)

    with s_r2_1, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:20px;color:#0F172A">📁</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">Project Settings</div>
                        <div style="font-size:12px;color:#64748B">Configure default settings for new projects.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        _project_status_opts = ["Not Started", "On Track", "At Risk", "Delayed"]
        _project_status_default = get_setting("default_project_status", "On Track")
        st.selectbox("Default Project Status", _project_status_opts, index=_project_status_opts.index(_project_status_default) if _project_status_default in _project_status_opts else 1, key="set_ps_stat")
        _project_template_opts = ["Standard Project", "Software Delivery", "Marketing Campaign", "Infrastructure"]
        _project_template_default = get_setting("project_template", "Standard Project")
        st.selectbox("Default Project Template", _project_template_opts, index=_project_template_opts.index(_project_template_default) if _project_template_default in _project_template_opts else 0, key="set_ps_tmpl")
        st.toggle("Enable project key/ID auto-generation", value=get_bool_setting("auto_project_id", True), key="set_ps_t1")
        st.toggle("Require project description", value=get_bool_setting("require_project_desc", True), key="set_ps_t2")
        st.toggle("Allow project attachments", value=get_bool_setting("allow_project_attachments", True), key="set_ps_t3")
        st.toggle("Enable project archiving", value=get_bool_setting("project_archiving", False), key="set_ps_t4")

    with s_r2_2, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:20px;color:#0F172A">☑</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">Task Settings</div>
                        <div style="font-size:12px;color:#64748B">Configure task management options.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        _task_status_opts = ["Not Started", "In Progress", "Completed"]
        _task_status_default = get_setting("default_task_status", "Not Started")
        st.selectbox("Default Task Status", _task_status_opts, index=_task_status_opts.index(_task_status_default) if _task_status_default in _task_status_opts else 0, key="set_ts_stat")
        _task_priority_opts = ["Medium", "High", "Low"]
        _task_priority_default = get_setting("default_task_priority", "Medium")
        st.selectbox("Default Task Priority", _task_priority_opts, index=_task_priority_opts.index(_task_priority_default) if _task_priority_default in _task_priority_opts else 0, key="set_ts_prio")
        st.toggle("Enable task dependencies", value=get_bool_setting("task_dependencies", True), key="set_ts_t1")
        st.toggle("Allow subtasks", value=get_bool_setting("allow_subtasks", True), key="set_ts_t2")
        st.toggle("Enable recurring tasks", value=get_bool_setting("recurring_tasks", False), key="set_ts_t3")
        st.toggle("Require assignee", value=get_bool_setting("require_assignee", True), key="set_ts_t4")
        st.toggle("Allow task attachments", value=get_bool_setting("task_attachments", True), key="set_ts_t5")

    with s_r2_3, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:20px;color:#0F172A">⚑</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">Milestone Settings</div>
                        <div style="font-size:12px;color:#64748B">Configure milestone options.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        _milestone_status_opts = ["Not Started", "In Progress", "Completed"]
        _milestone_status_default = get_setting("default_milestone_status", "Not Started")
        st.selectbox("Default Milestone Status", _milestone_status_opts, index=_milestone_status_opts.index(_milestone_status_default) if _milestone_status_default in _milestone_status_opts else 0, key="set_ms_stat")
        st.write("")
        st.toggle("Enable milestones", value=get_bool_setting("milestones_enabled", True), key="set_ms_t1")
        st.toggle("Require due date", value=get_bool_setting("milestones_require_due", True), key="set_ms_t2")
        st.toggle("Allow milestone dependencies", value=get_bool_setting("milestone_dependencies", True), key="set_ms_t3")
        st.toggle("Auto-notify on milestone updates", value=get_bool_setting("milestone_notifications", True), key="set_ms_t4")

    st.write("")
    # --- ROW 3: NOTIFICATIONS | INTEGRATIONS | SYSTEM SETTINGS ---
    s_r3_1, s_r3_2, s_r3_3 = st.columns([1.15, 1.0, 0.85])

    with s_r3_1, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
                    <span style="font-size:20px;color:#0F172A">🔔</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">Notifications</div>
                        <div style="font-size:12px;color:#64748B">Configure email and in-app notifications.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        nf_l, nf_r = st.columns([1.35, 1])
        with nf_l:
            notif_items = [
                ("Task assigned to me", True),
                ("Task due soon", True),
                ("Task overdue", True),
                ("Milestone updates", True),
                ("Project status changes", True),
                ("Project assignment", True),
                ("Mention in comments", False),
                ("Weekly project digest", True),
            ]
            for idx, (n_lbl, n_def) in enumerate(notif_items):
                _notif_setting_key = "notification_" + re.sub(r"[^a-z0-9]+", "_", n_lbl.lower()).strip("_")
                st.toggle(n_lbl, value=get_bool_setting(_notif_setting_key, n_def), key=f"set_nf_{idx}")
        with nf_r:
            st.markdown("<div style='padding-top:26px'></div>", unsafe_allow_html=True)
            _notification_frequency_opts = ["Real-time", "Hourly", "Daily digest", "Weekly digest"]
            _notification_frequency_default = get_setting("notification_frequency", "Daily digest")
            st.selectbox("Notification Frequency", _notification_frequency_opts, index=_notification_frequency_opts.index(_notification_frequency_default) if _notification_frequency_default in _notification_frequency_opts else 2, key="set_nf_freq")

    with s_r3_2, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:20px;color:#0F172A">🔗</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">Integrations</div>
                        <div style="font-size:12px;color:#64748B">Connect with external tools and services.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        int_services = [
            ("Microsoft 365", "integration_Microsoft 365", "Integration preference; API credentials are required for live sync.", "⊞", "#EA580C"),
            ("Slack", "integration_Slack", "Integration preference; Slack API setup is required for delivery.", "❖", "#E11D48"),
            ("Teams", "integration_Microsoft Teams", "Integration preference; Teams credentials are required for live sync.", "👥", "#4F46E5"),
            ("OneDrive", "integration_OneDrive", "Integration preference; Microsoft Graph setup is required for file access.", "☁", "#0284C7"),
        ]
        for s_nm, s_setting_key, s_ds, s_ic, s_cl in int_services:
            ic1, ic2 = st.columns([2.6, 1.0])
            with ic1:
                st.markdown(
                    clean_html(
                        f"""
                        <div style="display:flex;align-items:center;gap:10px;padding:4px 0">
                            <span style="width:32px;height:32px;min-width:32px;border-radius:8px;background:#F8FAFC;border:1px solid #E2E8F0;color:{s_cl};display:inline-flex;align-items:center;justify-content:center;font-size:18px;font-weight:800">{s_ic}</span>
                            <div>
                                <div style="font-size:13px;font-weight:700;color:#0F172A;line-height:1.2">{s_nm}</div>
                                <div style="font-size:11px;color:#64748B">{s_ds}</div>
                            </div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )
            with ic2:
                is_c = get_bool_setting(s_setting_key, s_nm in {"Microsoft 365", "Slack"})
                if st.button("Enabled" if is_c else "Enable", key=f"set_int_{s_nm}", type="primary" if is_c else "secondary", use_container_width=True):
                    put_setting(s_setting_key, "false" if is_c else "true")
                    st.rerun()
            st.markdown("<hr style='margin:4px 0;border-color:#F1F5F9'>", unsafe_allow_html=True)

    with s_r3_3, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                    <span style="font-size:20px;color:#0F172A">⚙</span>
                    <div>
                        <div style="font-size:15.5px;font-weight:800;color:#0F172A;line-height:1.15">System Settings</div>
                        <div style="font-size:12px;color:#64748B">Manage system-wide settings.</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        st.toggle("Enable audit logs", value=get_bool_setting("audit_logs", True), key="set_sys_t1")
        st.toggle("Allow data export", value=get_bool_setting("data_export", True), key="set_sys_t2")
        st.toggle("Enable automatic backups", value=get_bool_setting("automatic_backups", True), key="set_sys_t3")
        st.toggle("Enable maintenance mode", value=get_bool_setting("maintenance_mode", False), key="set_sys_t4")

        st_l, st_r = st.columns([1.1, 1.3])
        with st_l:
            st.markdown("<div style='font-size:12.5px;font-weight:700;color:#0F172A;padding-top:10px'>Session Timeout</div>", unsafe_allow_html=True)
        with st_r:
            _session_timeout_opts = ["8 hours", "4 hours", "2 hours", "1 hour", "30 minutes"]
            _session_timeout_default = get_setting("session_timeout", "8 hours")
            st.selectbox("Session Timeout", _session_timeout_opts, index=_session_timeout_opts.index(_session_timeout_default) if _session_timeout_default in _session_timeout_opts else 0, key="set_sys_to", label_visibility="collapsed")

        st.write("")
        if st.button("🗑 Clear Cache", key="set_clear_cache_btn"):
            st.cache_data.clear()
            st.toast("Application cache cleared!")

    # Persist the settings controls so they survive navigation and app restarts.
    _settings_widget_map = {
        "default_project_status": "set_ps_stat", "project_template": "set_ps_tmpl",
        "auto_project_id": "set_ps_t1", "require_project_desc": "set_ps_t2",
        "allow_project_attachments": "set_ps_t3", "project_archiving": "set_ps_t4",
        "default_task_status": "set_ts_stat", "default_task_priority": "set_ts_prio",
        "task_dependencies": "set_ts_t1", "allow_subtasks": "set_ts_t2",
        "recurring_tasks": "set_ts_t3", "require_assignee": "set_ts_t4", "task_attachments": "set_ts_t5",
        "default_milestone_status": "set_ms_stat", "milestones_enabled": "set_ms_t1",
        "milestones_require_due": "set_ms_t2", "milestone_dependencies": "set_ms_t3", "milestone_notifications": "set_ms_t4",
        "notification_frequency": "set_nf_freq", "audit_logs": "set_sys_t1", "data_export": "set_sys_t2",
        "automatic_backups": "set_sys_t3", "maintenance_mode": "set_sys_t4", "session_timeout": "set_sys_to",
    }
    _settings_updates = {}
    for _setting_name, _widget_key in _settings_widget_map.items():
        if _widget_key in st.session_state:
            _setting_value = st.session_state[_widget_key]
            if isinstance(_setting_value, bool):
                _setting_value = "true" if _setting_value else "false"
            if get_setting(_setting_name, "") != str(_setting_value):
                _settings_updates[_setting_name] = str(_setting_value)
    for _idx, (_notif_label, _notif_default) in enumerate([
        ("Task assigned to me", True), ("Task due soon", True), ("Task overdue", True),
        ("Milestone updates", True), ("Project status changes", True), ("Project assignment", True),
        ("Mention in comments", False), ("Weekly project digest", True),
    ]):
        _notif_key = f"set_nf_{_idx}"
        if _notif_key in st.session_state:
            _setting_name = "notification_" + re.sub(r"[^a-z0-9]+", "_", _notif_label.lower()).strip("_")
            _setting_value = "true" if st.session_state[_notif_key] else "false"
            if get_setting(_setting_name, "") != _setting_value:
                _settings_updates[_setting_name] = _setting_value
    if _settings_updates:
        with conn() as _c:
            _c.executemany("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", list(_settings_updates.items()))
            _c.commit()
        _persist_snapshot()


    if st.session_state.get("settings_modal") == "add_user":
        add_user_settings_modal()
    if st.session_state.get("settings_edit_user_id") is not None:
        edit_user_settings_modal(int(st.session_state["settings_edit_user_id"]))


# ---------------------------------------------------------------------------
# Deployment notes and lightweight runtime integrity checks
# ---------------------------------------------------------------------------
# Configure Streamlit Secrets (or environment variable) without committing credentials:
# .streamlit/secrets.toml
# MONGODB_URI = "mongodb+srv://<db-user>:<password>@<cluster-host>/?retryWrites=true&w=majority"
# The URI is intentionally never embedded in this source file.
