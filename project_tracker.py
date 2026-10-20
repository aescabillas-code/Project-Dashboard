"""Vektor Project Management - High-Fidelity Streamlit Application.
Install: pip install streamlit pandas plotly openpyxl pymongo
Run: streamlit run vektor_project_management_app.py
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
# Database Layer & Automatic Schema / 12-Project Seeding
# ---------------------------------------------------------------------------
def conn() -> sqlite3.Connection:
    """Local working cache; MongoDB is the durable source-of-truth snapshot."""
    c = sqlite3.connect(DB, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c


def _mongo_collection():
    """Resolve credentials from Streamlit Secrets or MONGODB_URI environment variable."""
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
    names = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log", "task_subtasks", "task_attachments", "task_activity", "calendar_events")
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
    tables = ("projects", "tasks", "resources", "milestones", "settings", "saved_reports", "audit_log", "task_subtasks", "task_attachments", "task_activity", "calendar_events")
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


def _clear_mock_records(except_table: str = "") -> None:
    """Legacy compatibility shim: no mock records are seeded or auto-deleted."""
    return


def _recalculate_project_rollups() -> None:
    """Recompute task status, parent rollups, project progress, health and days left from current dates."""
    today = date.today()
    with conn() as c:
        projects = c.execute("SELECT id, start_date, end_date, status FROM projects").fetchall()
        for project in projects:
            rows = c.execute("SELECT id,parent_task_id,status,progress,due_date,start_date FROM tasks WHERE project_id=?", (project["id"],)).fetchall()
            row_ids = {int(r["id"]) for r in rows}
            parents = {int(r["parent_task_id"]) for r in rows if r["parent_task_id"] is not None and int(r["parent_task_id"]) in row_ids}
            # Due-date driven status applies to non-completed leaf tasks; parent status is rolled up below.
            progress_by_id = {}
            status_by_id = {}
            due_by_id = {}
            for task in rows:
                tid = int(task["id"])
                try: due = date.fromisoformat(str(task["due_date"])[:10]) if task["due_date"] else None
                except Exception: due = None
                due_by_id[tid] = due
                status = str(task["status"] or "Not Started")
                progress = max(0, min(100, int(task["progress"] or 0)))
                if status.lower() == "completed": progress = 100
                elif due and due < today:
                    status = "Overdue"
                elif due and (due - today).days <= 7:
                    status = "At Risk"
                elif status.lower() in {"overdue", "at risk", "delayed"}:
                    status = "In Progress" if progress > 0 else "Not Started"
                progress_by_id[tid] = progress
                status_by_id[tid] = status
            # Parent progress/status is derived from its descendant leaf items, not entered separately.
            for task in reversed(rows):
                tid = int(task["id"])
                children = [r for r in rows if r["parent_task_id"] is not None and int(r["parent_task_id"]) == tid]
                if children:
                    descendants = []
                    stack = [int(ch["id"]) for ch in children]
                    while stack:
                        cid = stack.pop()
                        crow = next((x for x in rows if int(x["id"]) == cid), None)
                        if crow is None: continue
                        cchildren = [x for x in rows if x["parent_task_id"] is not None and int(x["parent_task_id"]) == cid]
                        if cchildren: stack.extend(int(x["id"]) for x in cchildren)
                        else: descendants.append(cid)
                    if descendants:
                        progress_by_id[tid] = round(sum(progress_by_id[x] for x in descendants) / len(descendants))
                        leaf_statuses = [status_by_id[x] for x in descendants]
                        if all(x == "Completed" for x in leaf_statuses): status_by_id[tid] = "Completed"
                        elif "Overdue" in leaf_statuses: status_by_id[tid] = "Overdue"
                        elif "At Risk" in leaf_statuses: status_by_id[tid] = "At Risk"
                        elif any(x == "In Progress" for x in leaf_statuses): status_by_id[tid] = "In Progress"
                        else: status_by_id[tid] = "Not Started"
            for task in rows:
                tid = int(task["id"])
                c.execute("UPDATE tasks SET status=?,progress=? WHERE id=?", (status_by_id[tid], progress_by_id[tid], tid))
            leaf_ids = [int(r["id"]) for r in rows if int(r["id"]) not in parents]
            total = len(leaf_ids)
            done = sum(1 for tid in leaf_ids if status_by_id.get(tid, "") == "Completed")
            progress = round(done / total * 100) if total else 0
            try: end_date = date.fromisoformat(str(project["end_date"])[:10]) if project["end_date"] else None
            except Exception: end_date = None
            if total and done == total: status, health = "Completed", "Healthy"
            elif end_date and end_date < today: status, health = "Overdue", "Critical"
            elif end_date and (end_date - today).days <= 7 and progress < 75: status, health = "At Risk", "At Risk"
            elif progress > 0: status, health = "In Progress", "Healthy"
            else: status, health = "Not Started", "Healthy"
            c.execute("UPDATE projects SET tasks_done=?,tasks_total=?,progress=?,days_left=?,status=?,health=? WHERE id=?",
                      (done, total, progress, (end_date - today).days if end_date else None, status, health, project["id"]))
        c.commit()


def sql(q: str, p: tuple = ()) -> int | None:
    """Execute an app mutation, then persist it and recalculate dependent project metrics."""
    normalized = " ".join(q.strip().lower().split())
    match = re.search(r"\b(?:insert\s+into|update|delete\s+from)\s+([a-z_]+)", normalized)
    table = match.group(1) if match else ""
    is_entity_mutation = table in {"projects", "tasks", "resources", "milestones", "saved_reports", "settings", "audit_log", "task_subtasks", "task_attachments", "task_activity", "calendar_events"}
    is_insert = normalized.startswith("insert into")
    with conn() as c:
        cursor = c.execute(q, p)
        lastrowid = cursor.lastrowid if normalized.startswith("insert into") else None
        c.commit()
    if table in {"tasks", "projects", "milestones"}:
        _recalculate_project_rollups()
    if is_entity_mutation:
        _persist_snapshot()
    return lastrowid


def df(q: str, p: tuple = ()) -> pd.DataFrame:
    with conn() as c:
        return pd.read_sql_query(q, c, params=p)


def init_db(force_reset: bool = False) -> None:
    with conn() as c:
        # Auto-upgrade if an older schema without the 12-project columns exists
        cols = [r[1] for r in c.execute("PRAGMA table_info(projects)").fetchall()]
        if force_reset or ("category" not in cols and len(cols) > 0):
            c.executescript(
                """
                DROP TABLE IF EXISTS projects;
                DROP TABLE IF EXISTS tasks;
                DROP TABLE IF EXISTS resources;
                DROP TABLE IF EXISTS milestones;
                DROP TABLE IF EXISTS settings;
                DROP TABLE IF EXISTS saved_reports;
                DROP TABLE IF EXISTS audit_log;
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
                health TEXT,
                budget REAL DEFAULT 0
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
            CREATE TABLE IF NOT EXISTS resources(
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT,
                team TEXT,
                role TEXT,
                status TEXT,
                capacity INTEGER,
                notes TEXT
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
                id INTEGER PRIMARY KEY, event TEXT, item_type TEXT, item_id INTEGER, details TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS task_subtasks(
                id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, name TEXT NOT NULL, completed INTEGER DEFAULT 0, due_date TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS task_attachments(
                id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, filename TEXT NOT NULL, mime_type TEXT, content BLOB, size_bytes INTEGER, uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS task_activity(
                id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, actor TEXT DEFAULT 'User', action TEXT, details TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS calendar_events(
                id INTEGER PRIMARY KEY, title TEXT NOT NULL, project_id INTEGER, task_id INTEGER, team TEXT, event_type TEXT, start_date TEXT, end_date TEXT, start_time TEXT, end_time TEXT, resources TEXT, allocation REAL DEFAULT 0, repeat_rule TEXT, description TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            """
        )

        # Operational tables are intentionally empty on a new install.
        # The portfolio workbook is the only supported source for project/task data.
        default_settings = [
            ("workspace_name", "Vektor"), ("workspace_description", "Project management workspace for planning, tracking, and delivery."),
            ("theme", "HPE Green"), ("language", "English"), ("date_format", "MMM DD, YYYY"),
            ("default_landing", "Overview"), ("default_project_status", "On Track"), ("project_template", "Standard Project"),
            ("auto_project_id", "true"), ("require_project_desc", "true"), ("allow_project_attachments", "true"),
            ("project_archiving", "false"), ("default_task_status", "Not Started"), ("default_task_priority", "Medium"),
            ("task_dependencies", "true"), ("allow_subtasks", "true"), ("recurring_tasks", "false"),
            ("require_assignee", "true"), ("task_attachments", "true"), ("default_milestone_status", "Not Started"),
            ("milestones_enabled", "true"), ("milestones_require_due", "true"), ("milestone_dependencies", "true"),
            ("milestone_notifications", "true"), ("notification_frequency", "Daily digest"),
            ("integration_Microsoft 365", "true"), ("integration_Slack", "true"),
            ("integration_Microsoft Teams", "false"), ("integration_OneDrive", "false"),
            ("audit_logs", "true"), ("data_export", "true"), ("automatic_backups", "true"),
            ("maintenance_mode", "false"), ("session_timeout", "8 hours"),
        ]
        c.executemany("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", default_settings)
        # Add project metadata and hierarchy fields to existing installations without resetting data.
        project_cols = {r[1] for r in c.execute("PRAGMA table_info(projects)").fetchall()}
        if "budget" not in project_cols:
            c.execute("ALTER TABLE projects ADD COLUMN budget REAL DEFAULT 0")
        _event_cols = {r[1] for r in c.execute("PRAGMA table_info(calendar_events)").fetchall()}
        if "task_id" not in _event_cols:
            c.execute("ALTER TABLE calendar_events ADD COLUMN task_id INTEGER")
        task_cols = {r[1] for r in c.execute("PRAGMA table_info(tasks)").fetchall()}
        for col_name, definition in (("parent_task_id", "INTEGER"), ("source_id", "TEXT"), ("description", "TEXT"),
                                     ("task_type", "TEXT DEFAULT 'Task'"), ("tags", "TEXT"), ("completion_date", "TEXT"),
                                     ("estimated_effort", "REAL"), ("actual_effort", "REAL"), ("duration", "INTEGER")):
            if col_name not in task_cols:
                c.execute(f"ALTER TABLE tasks ADD COLUMN {col_name} {definition}")
        c.commit()


def get_setting(k: str, default: str = "") -> str:
    x = df("SELECT value FROM settings WHERE key=?", (k,))
    return str(x.iloc[0].value) if not x.empty else default


def put_setting(k: str, v: object) -> None:
    sql(
        "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (k, str(v)),
    )


def log_event(event: str, typ: str = "", item_id: int | None = None, details: str = "") -> None:
    if get_setting("audit_logs", "true") == "true":
        sql(
            "INSERT INTO audit_log(event,item_type,item_id,details) VALUES(?,?,?,?)",
            (event, typ, item_id, details),
        )


def _as_iso_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == "":
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    return parsed.date().isoformat() if pd.notna(parsed) else ""

def import_portfolio_excel(uploaded_file) -> dict:
    """Import the single-sheet portfolio workbook as the app's operational source of truth."""
    uploaded_file.seek(0)
    book = pd.ExcelFile(uploaded_file, engine="openpyxl")
    if "Portfolio" not in book.sheet_names:
        raise ValueError("The workbook must contain a Portfolio sheet.")
    raw = pd.read_excel(book, sheet_name="Portfolio", header=None)
    header_row = 2 if raw.shape[0] > 2 and "Initiative / Milestone / Task" in [str(x).strip() for x in raw.iloc[2].tolist()] else 0
    frame = pd.read_excel(book, sheet_name="Portfolio", header=header_row)
    frame.columns = [str(x).strip() for x in frame.columns]
    required = {"Program", "Initiative / Milestone / Task"}
    if not required.issubset(frame.columns):
        raise ValueError("Portfolio must include Program and Initiative / Milestone / Task columns.")
    frame = frame[frame["Program"].notna()].copy()
    frame["Program"] = frame["Program"].astype(str).str.strip()
    frame = frame[frame["Program"].ne("")]
    if frame.empty:
        raise ValueError("No project rows were found in the Portfolio sheet.")
    valid_tasks = frame[frame["Initiative / Milestone / Task"].notna()].copy()
    valid_tasks = valid_tasks[valid_tasks["Initiative / Milestone / Task"].astype(str).str.strip().ne("")]
    if valid_tasks.empty:
        raise ValueError("No task/subtask rows were found in the Portfolio sheet.")
    def val(row, *names, default=""):
        for name in names:
            if name in row.index and pd.notna(row[name]) and str(row[name]).strip().lower() not in {"nan", "nat", "none"}:
                return str(row[name]).strip()
        return default
    program_names = list(dict.fromkeys(frame["Program"].astype(str).str.strip().tolist()))
    # Import is a deliberate replacement operation; reset only operational tables, not settings.
    with conn() as c:
        c.execute("PRAGMA foreign_keys=OFF")
        for table in ("task_activity", "task_attachments", "task_subtasks", "milestones", "calendar_events", "tasks", "resources", "projects", "saved_reports", "audit_log"):
            if c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone():
                c.execute(f"DELETE FROM {table}")
        project_ids = {}
        for name in program_names:
            rows = frame[frame["Program"].astype(str).str.strip() == name]
            starts = [_as_iso_date(v) for v in rows.get("Start Date", []) if _as_iso_date(v)]
            ends = [_as_iso_date(v) for v in rows.get("End Date", []) if _as_iso_date(v)]
            project_owner = next((val(r, "Project Owner", "PC", "Owner") for _, r in rows.iterrows() if val(r, "Project Owner", "PC", "Owner")), "")
            project_team = next((val(r, "Project Team", "Team") for _, r in rows.iterrows() if val(r, "Project Team", "Team")), "")
            project_category = next((val(r, "Project Category") for _, r in rows.iterrows() if val(r, "Project Category")), "Portfolio")
            project_priority = next((val(r, "Project Priority") for _, r in rows.iterrows() if val(r, "Project Priority")), "Medium")
            project_tags = next((val(r, "Project Tags") for _, r in rows.iterrows() if val(r, "Project Tags")), "")
            project_desc = next((val(r, "Project Description") for _, r in rows.iterrows() if val(r, "Project Description")), "")
            project_budget = pd.to_numeric(pd.Series([next((r.get("Project Budget") for _, r in rows.iterrows() if pd.notna(r.get("Project Budget"))), 0)]), errors="coerce").fillna(0).iloc[0]
            cur = c.execute("""INSERT INTO projects(name,owner,team,category,status,priority,start_date,end_date,progress,days_left,tasks_done,tasks_total,tags,bar_color,short_desc,description,health,budget)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (name, project_owner, project_team, project_category, "Not Started", project_priority, min(starts) if starts else "", max(ends) if ends else "", 0, None, 0, 0, project_tags, GREEN, name, project_desc, "Healthy", float(project_budget or 0)))
            project_ids[name] = cur.lastrowid
        source_to_db = {}
        pending = []
        owners = set()
        for _, row in valid_tasks.iterrows():
            program = val(row, "Program")
            source_id = val(row, "App Key", "ID")
            parent_source = val(row, "Parent Key", "Parent ID")
            if not parent_source and source_id and "." in source_id:
                parent_source = source_id.rsplit(".", 1)[0]
            owner = val(row, "Owner")
            if owner: owners.add(owner)
            status = val(row, "Status", "Task Status", "Automated Status", default="Not Started")
            status_l = status.lower()
            status = "Completed" if status_l in {"complete", "completed", "closed", "done"} else ("In Progress" if status_l in {"in progress", "ongoing", "started"} else ("At Risk" if status_l in {"at risk"} else ("Overdue" if status_l in {"overdue", "delayed"} else ("On Hold" if status_l in {"on hold", "paused"} else "Not Started"))))
            _due_auto = _as_iso_date(row.get("End Date"))
            if status != "Completed" and _due_auto:
                try:
                    _due_date_auto = date.fromisoformat(_due_auto)
                    if _due_date_auto < date.today(): status = "Overdue"
                    elif _due_date_auto <= date.today() + timedelta(days=7): status = "At Risk"
                except Exception: pass
            pct = pd.to_numeric(pd.Series([row.get("Completion %", row.get("Progress", 100 if status == "Completed" else 0))]), errors="coerce").fillna(100 if status == "Completed" else 0).iloc[0]
            typ = val(row, "Task Type", default=("Milestone" if "milestone" in val(row, "Initiative / Milestone / Task").lower() else "Task"))
            cur = c.execute("""INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id,source_id,description,task_type,tags,completion_date,estimated_effort,actual_effort,duration)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (val(row, "Initiative / Milestone / Task"), project_ids[program], val(row, "Team", "PC"), status, val(row, "Priority", default="Medium"),
                 _as_iso_date(row.get("Start Date")), _as_iso_date(row.get("End Date")), owner, max(0, min(100, int(pct))), val(row, "Dependency"),
                 val(row, "Remarks", "Update"), None, source_id, val(row, "Description", "Remarks", "Update"), typ, val(row, "Tags"),
                 _as_iso_date(row.get("Completion Date")), pd.to_numeric(pd.Series([row.get("Estimated Effort", 0)]), errors="coerce").fillna(0).iloc[0],
                 pd.to_numeric(pd.Series([row.get("Actual Effort", 0)]), errors="coerce").fillna(0).iloc[0],
                 pd.to_numeric(pd.Series([row.get("Duration", 0)]), errors="coerce").fillna(0).iloc[0]))
            if source_id: source_to_db[source_id] = cur.lastrowid
            pending.append((cur.lastrowid, parent_source))
            if typ.strip().lower() == "milestone":
                c.execute("INSERT INTO milestones(name,project_id,team,status,start_date,due_date,completion_date,dependency,notes) VALUES(?,?,?,?,?,?,?,?,?)",
                    (val(row, "Initiative / Milestone / Task"), project_ids[program], val(row, "Team", "PC"), status, _as_iso_date(row.get("Start Date")), _as_iso_date(row.get("End Date")), _as_iso_date(row.get("Completion Date")), val(row, "Dependency"), val(row, "Remarks", "Update")))
        for task_id, parent_source in pending:
            if parent_source in source_to_db:
                c.execute("UPDATE tasks SET parent_task_id=? WHERE id=?", (source_to_db[parent_source], task_id))
        for owner in sorted(owners):
            c.execute("INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                      (owner, "", "", "Team Member", "Available", 100, ""))
        c.commit()
    _recalculate_project_rollups()
    _persist_snapshot()
    st.session_state["_vektor_has_real_data"] = True
    return {"projects": len(program_names), "tasks": len(valid_tasks), "owners": len(owners)}

def clear_operational_data(selected_tables: list[str]) -> None:
    allowed = {"projects", "tasks", "resources", "milestones", "calendar_events", "saved_reports", "audit_log"}
    chosen = allowed.intersection(selected_tables)
    # Clear dependent children when parent entities are selected.
    if "projects" in chosen:
        chosen.update({"tasks", "milestones", "calendar_events"})
    if "tasks" in chosen:
        chosen.update({"task_activity", "task_attachments", "task_subtasks"})
    with conn() as c:
        c.execute("PRAGMA foreign_keys=OFF")
        for table in ("task_activity", "task_attachments", "task_subtasks", "milestones", "calendar_events", "tasks", "resources", "projects", "saved_reports", "audit_log"):
            if table in chosen and c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone():
                c.execute(f"DELETE FROM {table}")
        c.commit()
    st.session_state["_vektor_has_real_data"] = False
    _recalculate_project_rollups()
    _persist_snapshot()


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

# Load Dataframes
_recalculate_project_rollups()
P = df("SELECT * FROM projects ORDER BY id")
T = df("SELECT t.*, p.name AS project FROM tasks t LEFT JOIN projects p ON p.id=t.project_id ORDER BY t.due_date")
R = df("SELECT * FROM resources ORDER BY id")
M = df("SELECT m.*, p.name AS project FROM milestones m LEFT JOIN projects p ON p.id=m.project_id ORDER BY m.due_date")

projects_list = P.name.tolist()
pmap = dict(zip(P.name, P.id))
teams_list = sorted(set(R.team.dropna().tolist() + P.team.dropna().tolist()))
if P.empty:
    st.markdown("<div class='project-title'>PROJECT TRACKER</div>", unsafe_allow_html=True)
    st.info("No portfolio data is loaded. Upload the approved Excel source below to populate the app. No sample projects or tasks are created automatically.")
    _initial_portfolio = st.file_uploader("Upload portfolio Excel source", type=["xlsx", "xlsm"], key="initial_portfolio_source")
    if _initial_portfolio is not None and st.button("Import Portfolio", type="primary", key="initial_import_portfolio"):
        try:
            _import_counts = import_portfolio_excel(_initial_portfolio)
            st.success(f"Imported {_import_counts['projects']} projects, {_import_counts['tasks']} task/subtask rows, and {_import_counts['owners']} owners.")
            st.rerun()
        except Exception as _import_error:
            st.error(f"Import failed: {_import_error}")
    st.stop()

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
/* Remove duplicate native dialog heading; retain close control and the icon/title inside each modal. */
div[data-testid="stDialog"] [data-testid="stDialogHeader"] > div:first-child {{ display:none !important; }}
div[data-testid="stDialog"] [data-testid="stDialogHeader"] {{ min-height: 20px !important; padding-top: 6px !important; padding-bottom: 0 !important; }}
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
        "Alice Morgan": "#00875A",
        "Mark Wilson": "#00875A",
        "Laura Torres": "#00875A",
        "Dana Scott": "#1473E6",
        "Sarah Kim": "#1473E6",
        "Kevin Lee": "#7C3AED",
    }
    bg = avatar_colors.get(s, "#475569")
    return f'<span class="vk-avatar" style="background:{bg}">{html.escape(initials)}</span>{html.escape(s)}'


def fmt_date(d_str: object) -> str:
    if not d_str or pd.isna(d_str) or not str(d_str).strip():
        return "—"
    try:
        d = date.fromisoformat(str(d_str)[:10])
        return d.strftime("%b %-d, %Y") if os.name != "nt" else d.strftime("%b %d, %Y").replace(" 0", " ")
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
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        for sheet_name, d in sheets.items():
            d.to_excel(w, sheet_name=sheet_name[:31], index=False)
    return b.getvalue()


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


def render_timeline_stepper_html() -> str:
    steps = [
        ("Project Start", "Apr 24", "done", "Project kick-off", "Completed", "#DCFCE7", "#15803D"),
        ("Planning", "May 5", "done", "Requirements", "Completed", "#DCFCE7", "#15803D"),
        ("Design", "May 20", "done", "Design sign-off", "Completed", "#DCFCE7", "#15803D"),
        ("Development", "May 31", "current", "Build and integrate", "In Progress", "#DBEAFE", "#1D4ED8"),
        ("Testing", "Jun 15", "todo", "UAT and testing", "Not Started", "#F1F5F9", "#475569"),
        ("Go-Live", "Jun 30", "todo", "Launch", "Not Started", "#F1F5F9", "#475569"),
    ]
    nodes_html = ""
    for idx, (st_title, st_date, state, box_title, box_sub, box_bg, box_fg) in enumerate(steps):
        if state == "done":
            circle = (
                '<div style="width:24px;height:24px;border-radius:50%;background:#00875A;'
                'color:white;display:flex;align-items:center;justify-content:center;'
                'font-size:13px;font-weight:800;z-index:2;position:relative">✓</div>'
            )
        elif state == "current":
            circle = (
                '<div style="width:24px;height:24px;border-radius:50%;background:#1473E6;'
                'border:3px solid #BFDBFE;color:white;display:flex;align-items:center;'
                'justify-content:center;font-size:10px;z-index:2;position:relative">●</div>'
            )
        else:
            circle = (
                '<div style="width:22px;height:22px;border-radius:50%;background:#94A3B8;'
                'color:white;display:flex;align-items:center;justify-content:center;'
                'font-size:10px;z-index:2;position:relative"></div>'
            )

        line_color = "#00875A" if idx < 2 else ("#1473E6" if idx == 2 else "#CBD5E1")
        line_right = (
            f'<div style="position:absolute;top:46px;left:50%;width:100%;height:3px;background:{line_color};z-index:1"></div>'
            if idx < len(steps) - 1
            else ""
        )
        nodes_html += clean_html(
            f"""
            <div style="flex:1;position:relative;display:flex;flex-direction:column;align-items:center;text-align:center;padding:0 4px">
                {line_right}
                <div style="font-size:11.5px;font-weight:700;color:#1E293B">{st_title}</div>
                <div style="font-size:10.5px;color:#64748B;margin-bottom:6px">{st_date}</div>
                {circle}
                <div style="margin-top:10px;background:{box_bg};border-radius:6px;padding:6px 6px;width:94%;min-height:44px">
                    <div style="font-size:10.5px;font-weight:600;color:{box_fg}">{box_title}</div>
                    <div style="font-size:10.5px;font-weight:700;color:{box_fg}">{box_sub}</div>
                </div>
            </div>
            """
        )
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


@st.dialog("View Project", width="large")
def view_project_modal(pid: int) -> None:
    row_df = df("SELECT * FROM projects WHERE id=?", (pid,))
    if row_df.empty:
        st.warning("Project not found.")
        return
    r = row_df.iloc[0]
    tags_html = "".join(tag_pill_html(t) for t in str(r.tags or "").split(",") if t.strip())
    full_desc = (
        "Redesign and launch the corporate website with improved user experience and modern design."
        if r["name"] == "Website Revamp"
        else str(r.description or r.short_desc or "")
    )
    pct = int(r.progress or 0)

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
                    <div><div style="font-size:11px;color:#64748B">Project Category</div><div style="font-size:13px;font-weight:700;color:#0F172A">{html.escape(str(r.category or 'Web Development'))}</div></div>
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
                    <div><div style="font-size:11px;color:#64748B">Days Left</div><div style="font-size:18px;font-weight:800;color:#00875A;line-height:1">{int(r.days_left or 45)}</div><div style="font-size:10px;color:#64748B">Days remaining</div></div>
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
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:12px;line-height:2.1;padding-top:14px">
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>14 (70%)</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>4 (20%)</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> On Hold</span><b>1 (5%)</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>1 (5%)</b></div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
    with c_right, st.container(border=True):
        section_title("⏳ Timeline", "View Full Timeline →")
        st.markdown(render_timeline_stepper_html(), unsafe_allow_html=True)

    b_left, b_right = st.columns([1.05, 1.95])
    with b_left, st.container(border=True):
        section_title("📄 Project Information")
        st.markdown(
            clean_html(
                f"""
                <div style="font-size:12.5px;line-height:1.95">
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Project Name</span><span style="font-weight:600">{html.escape(r['name'])}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px;line-height:1.4;margin:4px 0"><span style="color:#64748B">Description</span><span>{html.escape(full_desc)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Category</span><span>{html.escape(str(r.category or 'Web Development'))}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Status</span><span>{badge_html(r.status)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Priority</span><span>{badge_html(r.priority or 'High')}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Project Manager</span><span>{html.escape(str(r.owner))}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Start Date</span><span>{fmt_date(r.start_date)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Target End Date</span><span>{fmt_date(r.end_date)}</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Duration</span><span>67 days</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px"><span style="color:#64748B">Days Left</span><span style="color:#00875A;font-weight:700">{int(r.days_left or 45)} days</span></div>
                    <div style="display:grid;grid-template-columns:115px 1fr;gap:6px;margin-top:4px"><span style="color:#64748B">Tags</span><span>{tags_html}</span></div>
                </div>
                <hr style="border-color:#EEF2F6;margin:12px 0">
                <div style="font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:6px">📄 Project Summary</div>
                <div style="font-size:11.5px;color:#64748B;line-height:1.55">
                    {html.escape(str(r.description or r.short_desc or 'No project description provided.'))} Current calculated progress is {pct}% based on completed leaf tasks.
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    with b_right:
        with st.container(border=True):
            section_title("🎯 Key Milestones", "View All Milestones →")
            ms_rows = [
                (1, "#00875A", "Project Kick-off", "Apr 24, 2024", "Completed"),
                (2, "#00875A", "Requirements Finalized", "May 5, 2024", "Completed"),
                (3, "#00875A", "Design Sign-off", "Apr 30, 2024", "Completed"),
                (4, "#1473E6", "Development Complete", "May 31, 2024", "In Progress"),
                (5, "#64748B", "UAT Completion", "Jun 15, 2024", "Not Started"),
                (6, "#64748B", "Go-Live", "Jun 30, 2024", "Not Started"),
            ]
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
            _project_activity = df("SELECT created_at,event,details FROM audit_log WHERE item_type='Project' AND item_id=? ORDER BY id DESC LIMIT 5", (int(r["id"]),))
            act_rows = [(str(a.get("created_at") or ""), GREEN if str(a.get("event")) == "Created" else BLUE, str(a.get("details") or a.get("event") or "Updated"), "User") for _, a in _project_activity.iterrows()]
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
            default_desc = (
                "Redesign and launch the corporate website with improved user experience and modern design."
                if (r is not None and str(r.get("tags") or "").strip())
                else (str(r.description or "") if r is not None else "")
            )
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
            p_stat = st.selectbox("Status *", stats, index=stats.index(cur_st))
        prios = ["High", "Medium", "Low"]
        cur_pr = str(r.priority) if r is not None and r.priority in prios else "High"
        with c3:
            p_prio = st.selectbox("Priority *", prios, index=prios.index(cur_pr))

        c4, c5, c6 = st.columns(3)
        mgrs = R.name.tolist() if not R.empty else ["Alice Morgan"]
        cur_mgr = str(r.owner) if r is not None and r.owner in mgrs else mgrs[0]
        with c4:
            p_mgr = st.selectbox("Project Manager *", mgrs, index=mgrs.index(cur_mgr))
        with c5:
            s_dt = date.fromisoformat(r.start_date) if r is not None and r.start_date else date.today()
            p_start = st.date_input("Start Date *", s_dt)
        with c6:
            e_dt = date.fromisoformat(r.end_date) if r is not None and r.end_date else date(2024, 6, 30)
            p_end = st.date_input("Target End Date *", e_dt)

        all_tags_pool = ["Web", "Marketing", "Corporate", "IT", "Development", "Mobile", "Customer", "Data", "Cloud", "Infrastructure", "Security", "Compliance", "AI", "Innovation", "Automation", "Operations", "HR", "Training", "Procurement", "Sustainability", "ESG"]
        cur_tags = [t.strip() for t in str(r.tags or "Web,Marketing,Corporate").split(",") if t.strip() in all_tags_pool]
        p_tags = st.multiselect("Tags", all_tags_pool, default=cur_tags if cur_tags else ["Web", "Marketing", "Corporate"])

    with top_r, st.container(border=True):
        section_title("📄 Project Progress (auto-calculated)")
        pct_val = int(r.progress) if r is not None else 70
        st.plotly_chart(render_project_progress_donut(pct_val, 185), use_container_width=True, config={"displayModeBar": False})
        st.markdown(
            clean_html(
                """
                <div style="font-size:12px;line-height:2.1;padding:4px 8px">
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>14 (70%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>4 (20%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> On Hold</span><b>1 (5%)</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>1 (5%)</b></div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    with st.container(border=True):
        section_title("📋 Project Timeline & Milestones")
        st.markdown(render_timeline_stepper_html(), unsafe_allow_html=True)
        ms_edit_rows = [
            (1, "#00875A", "Project Kick-off", "Apr 24, 2024", "Completed"),
            (2, "#00875A", "Requirements Finalized", "May 5, 2024", "Completed"),
            (3, "#00875A", "Design Sign-off", "Apr 30, 2024", "Completed"),
            (4, "#1473E6", "Development Complete", "May 31, 2024", "In Progress"),
            (5, "#64748B", "UAT Completion", "Jun 15, 2024", "Not Started"),
            (6, "#64748B", "Go-Live", "Jun 30, 2024", "Not Started"),
        ]
        rows_mk = "".join(
            f"<tr><td>{i}</td><td><span style='color:{dot};margin-right:6px'>●</span>{nm}</td><td>{dt}</td><td>{badge_html(st_val)}</td><td>✎ &nbsp; <span style='color:#EF4444'>🗑</span></td></tr>"
            for i, dot, nm, dt, st_val in ms_edit_rows
        )
        st.markdown(
            f'<div class="vk-table-wrap"><table class="vk-table"><thead><tr><th>#</th><th>Milestone</th><th>Due Date</th><th>Status</th><th>Actions</th></tr></thead><tbody>{rows_mk}</tbody></table></div>',
            unsafe_allow_html=True,
        )

    b1, b2, b3 = st.columns([4, 1, 1.3])
    with b2:
        if st.button("Cancel", use_container_width=True, key="cancel_edit_modal"):
            st.session_state.pop("active_modal", None)
            st.rerun()
    with b3:
        if st.button("💾 Save Changes", type="primary", use_container_width=True, key="save_edit_modal"):
            tags_str = ",".join(p_tags)
            days_l = max(0, (p_end - date(2024, 5, 16)).days)
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
                        p_start.isoformat(), p_end.isoformat(), pct_val, days_l, 0, 20,
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
        ov_date_range = st.date_input("Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="ov_date_sel", label_visibility="collapsed")

    _ov_tasks = T.copy()
    if isinstance(ov_date_range, (tuple, list)) and len(ov_date_range) == 2:
        _ov_lo, _ov_hi = sorted(ov_date_range)
        _ov_task_start = pd.to_datetime(_ov_tasks.get("start_date", pd.Series(index=_ov_tasks.index,dtype=str)), errors="coerce").dt.date
        _ov_task_end = pd.to_datetime(_ov_tasks.get("due_date", pd.Series(index=_ov_tasks.index,dtype=str)), errors="coerce").dt.date
        _ov_tasks = _ov_tasks[(_ov_task_start.isna() | (_ov_task_start <= _ov_hi)) & (_ov_task_end.isna() | (_ov_task_end >= _ov_lo))].copy()
    if ov_proj_filter != "All Projects": _ov_tasks = _ov_tasks[_ov_tasks["project"] == ov_proj_filter].copy()
    st.write("")
    # 6 KPI Cards Row (live database metrics)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    _ov_total = len(_ov_tasks); _ov_completed = int((_ov_tasks["status"].astype(str).str.lower()=="completed").sum()) if _ov_total else 0
    _ov_progress = int((_ov_tasks["status"].astype(str).str.lower()=="in progress").sum()) if _ov_total else 0
    _ov_hold = int((_ov_tasks["status"].astype(str).str.lower()=="on hold").sum()) if _ov_total else 0
    _ov_due = pd.to_datetime(_ov_tasks.get("due_date", pd.Series(dtype=str)), errors="coerce") if _ov_total else pd.Series(dtype="datetime64[ns]")
    _ov_overdue = int(((_ov_due.dt.date < date.today()) & (_ov_tasks["status"].astype(str).str.lower()!="completed")).sum()) if _ov_total else 0
    _ov_pct = lambda n: f"({n/_ov_total:.0%})" if _ov_total else "(0%)"
    with k1: kpi_card_mockup("Total Projects", str(len(P)), sub_text="Loaded Projects", icon_html="❖", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k2: kpi_card_mockup("Total Tasks", str(_ov_total), sub_text="Across All Projects", icon_html="✓", icon_bg="#DCFCE7", icon_fg="#00875A")
    with k3: kpi_card_mockup("Completed Tasks", str(_ov_completed), _ov_pct(_ov_completed), bar_pct=round(_ov_completed/_ov_total*100) if _ov_total else 0, icon_html="◕", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k4: kpi_card_mockup("In Progress", str(_ov_progress), _ov_pct(_ov_progress), bar_pct=round(_ov_progress/_ov_total*100) if _ov_total else 0, icon_html="▶", icon_bg="#DBEAFE", icon_fg="#1473E6", bar_color="#1473E6")
    with k5: kpi_card_mockup("On Hold", str(_ov_hold), _ov_pct(_ov_hold), bar_pct=round(_ov_hold/_ov_total*100) if _ov_total else 0, icon_html="❚❚", icon_bg="#FEF3C7", icon_fg="#F28C00", bar_color="#F28C00")
    with k6: kpi_card_mockup("Overdue", str(_ov_overdue), _ov_pct(_ov_overdue), bar_pct=round(_ov_overdue/_ov_total*100) if _ov_total else 0, icon_html="✕", icon_bg="#FEE2E2", icon_fg="#EF4444", bar_color="#EF4444")

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
        if isinstance(ov_date_range, (tuple, list)) and len(ov_date_range) == 2:
            _ov_lo, _ov_hi = sorted(ov_date_range)
            _ov_st = pd.to_datetime(ov_projs["start_date"], errors="coerce").dt.date
            _ov_en = pd.to_datetime(ov_projs["end_date"], errors="coerce").dt.date
            ov_projs = ov_projs[(_ov_st.isna() | (_ov_st <= _ov_hi)) & (_ov_en.isna() | (_ov_en >= _ov_lo))].copy()
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
                f"<td style='text-align:center;font-weight:600'>{0 if pd.isna(pr.days_left) else int(pr.days_left)}</td>"
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
            ts_df = pd.DataFrame(
                {
                    "Status": ["Completed", "In Progress", "On Hold", "Overdue"],
                    "Count": [_ov_completed, _ov_progress, _ov_hold, _ov_overdue],
                }
            )
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
                        text=f"<b style='font-size:20px;color:#0F172A'>{_ov_total}</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_ts, use_container_width=True, config={"displayModeBar": False})
        with l_col:
            st.markdown(
                """
                <div style="font-size:12.5px;padding-top:34px;line-height:2.35">
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#00875A">●</span> &nbsp;Completed</span><b>78</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#1473E6">●</span> &nbsp;In Progress</span><b>46</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;On Hold</span><b>12</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> &nbsp;Overdue</span><b>6</b></div>
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
            _ms_live = M.copy()
            _ms_live["_due"] = pd.to_datetime(_ms_live["due_date"], errors="coerce")
            _ms_live = _ms_live.sort_values("_due", na_position="last").head(5)
            for _, _m in _ms_live.iterrows():
                _due_s = _as_iso_date(_m.get("due_date"))
                _days_left = (date.fromisoformat(_due_s) - date.today()).days if _due_s else 0
                ms_ov.append((GREEN, str(_m.get("name", "")), str(_m.get("project", "")), fmt_date(_due_s), _days_left, str(_m.get("status", "Not Started"))))
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
            tp_df = pd.DataFrame(
                {
                    "Priority": ["High", "Medium", "Low", "None"],
                    "Count": [int((_ov_tasks["priority"].astype(str).str.lower()=="high").sum()) if not _ov_tasks.empty else 0, int((_ov_tasks["priority"].astype(str).str.lower()=="medium").sum()) if not _ov_tasks.empty else 0, int((_ov_tasks["priority"].astype(str).str.lower()=="low").sum()) if not _ov_tasks.empty else 0, int(_ov_tasks["priority"].isna().sum()) if not _ov_tasks.empty else 0],
                }
            )
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
                        text=f"<b style='font-size:20px;color:#0F172A'>{_ov_total}</b><br><span style='font-size:11px;color:#64748B'>Total Tasks</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_tp, use_container_width=True, config={"displayModeBar": False})
        with l_col:
            st.markdown(
                """
                <div style="font-size:12.5px;padding-top:34px;line-height:2.35">
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#EF4444">●</span> &nbsp;High</span><b>{int((_ov_tasks['priority'].astype(str).str.lower()=='high').sum()) if not _ov_tasks.empty else 0}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;Medium</span><b>{int((_ov_tasks['priority'].astype(str).str.lower()=='medium').sum()) if not _ov_tasks.empty else 0}</b></div>
                    <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#1473E6">●</span> &nbsp;Low</span><b>{int((_ov_tasks['priority'].astype(str).str.lower()=='low').sum()) if not _ov_tasks.empty else 0}</b></div>
                    <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;None</span><b>{int(_ov_tasks['priority'].isna().sum()) if not _ov_tasks.empty else 0}</b></div>
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
        for _, _res in R.head(5).iterrows():
            _rn = str(_res.get("name") or "")
            _rt = T[T["assignee"].astype(str) == _rn] if not T.empty else pd.DataFrame()
            _assigned = len(_rt); _inprog = int((_rt["status"].astype(str).str.lower()=="in progress").sum()) if _assigned else 0
            _done = int((_rt["status"].astype(str).str.lower()=="completed").sum()) if _assigned else 0
            _due = pd.to_datetime(_rt.get("due_date", pd.Series(dtype=str)), errors="coerce") if _assigned else pd.Series(dtype="datetime64[ns]")
            _overdue = int(((_due.dt.date < date.today()) & (_rt["status"].astype(str).str.lower()!="completed")).sum()) if _assigned else 0
            tw_rows.append((_rn, _assigned, _inprog, _done, _overdue, round(_done/_assigned*100) if _assigned else 0))
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
        _audit_live = df("SELECT created_at,event,details,item_id FROM audit_log ORDER BY id DESC LIMIT 5")
        for _, _ar in _audit_live.iterrows():
            _tid = _ar.get("item_id")
            _task_name = ""; _project_name = ""
            if pd.notna(_tid):
                _tr_live = T[T["id"] == int(_tid)] if not T.empty else pd.DataFrame()
                if not _tr_live.empty:
                    _task_name = str(_tr_live.iloc[0].get("name") or "")
                    _project_name = str(_tr_live.iloc[0].get("project") or "")
            ra_rows.append((str(_ar.get("created_at") or ""), GREEN, str(_ar.get("event") or "Updated"), _project_name, _task_name or str(_ar.get("details") or "")))
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
        p_date_range = st.date_input("Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="p_date_box", label_visibility="collapsed")
    with h_b:
        if st.button("＋ New Project", type="primary", use_container_width=True, key="btn_new_proj_top"):
            st.session_state.active_modal = ("edit", None)
            st.rerun()

    st.write("")
    # 5 KPI Cards Row matching Projects screenshots
    total_p = len(P)
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        kpi_card_mockup("TOTAL PROJECTS", str(total_p), bar_pct=60, icon_html="❖", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k2:
        kpi_card_mockup("ON TRACK", "6", "(50%)", bar_pct=50, icon_html="✓", icon_bg="#DCFCE7", icon_fg="#00875A", bar_color="#00875A")
    with k3:
        kpi_card_mockup("IN PROGRESS", "3", "(25%)", bar_pct=25, icon_html="▶", icon_bg="#DBEAFE", icon_fg="#1473E6", bar_color="#1473E6")
    with k4:
        kpi_card_mockup("AT RISK", "2", "(17%)", bar_pct=17, icon_html="❚❚", icon_bg="#FEF3C7", icon_fg="#F28C00", bar_color="#F28C00")
    with k5:
        kpi_card_mockup("OVERDUE", "1", "(8%)", bar_pct=8, icon_html="✕", icon_bg="#FEE2E2", icon_fg="#EF4444", bar_color="#EF4444")

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
    if isinstance(p_date_range, (tuple, list)) and len(p_date_range) == 2:
        _plo, _phi = sorted(p_date_range)
        _pst = pd.to_datetime(vp["start_date"], errors="coerce").dt.date
        _pen = pd.to_datetime(vp["end_date"], errors="coerce").dt.date
        vp = vp[(_pst.isna() | (_pst <= _phi)) & (_pen.isna() | (_pen >= _plo))].copy()
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
                    ttot = int(item.get("tasks_total") or 20)

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
                ttot = int(item.get("tasks_total") or 20)
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
                    _days_left = pd.to_numeric(item.get("days_left"), errors="coerce")
                    st.markdown(f"<div style='padding-top:6px;font-size:12.5px;font-weight:700;color:#00875A'>{0 if pd.isna(_days_left) else int(_days_left)}</div>",unsafe_allow_html=True)
                with r_cols[9]:
                    st.markdown(
                        f"<div style='padding-top:2px'><div style='font-size:11.5px;font-weight:600'>{tdone}/{ttot}</div>"
                        f"<div style='height:5px;background:#E2E8F0;border-radius:99px;overflow:hidden;margin-top:2px'>"
                        f"<div style='width:{task_pct}%;height:100%;background:#00875A'></div></div></div>",
                        unsafe_allow_html=True, )
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

            # Pagination Footer
            st.markdown(
                f"""
                <div style="display:flex;align-items:center;justify-content:space-between;padding-top:8px;font-size:12.5px;color:#475569">
                    <div>Showing 1 to {len(vp)} of {len(P)} projects</div>
                    <div style="display:flex;align-items:center;gap:8px">
                        <span>Rows per page</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 10px;background:white;font-weight:600">10 ▾</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 9px;background:white">‹</span>
                        <span style="border-radius:6px;padding:3px 10px;background:#00875A;color:white;font-weight:700">1</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 9px;background:white">›</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

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
        gantt_date_range = st.date_input("Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="gantt_date_top", label_visibility="collapsed")

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

        # Build the Gantt hierarchy from imported project/task rows rather than embedded demo rows.
        _all_dates = []
        for _frame in (P, T):
            for _col in (["start_date", "end_date"] if _frame is P else ["start_date", "due_date"]):
                if _col in _frame.columns:
                    _all_dates.extend([d.date() for d in pd.to_datetime(_frame[_col], errors="coerce").dropna().tolist()])
        if isinstance(gantt_date_range, (tuple, list)) and len(gantt_date_range) == 2:
            _min_date, _max_date = sorted(gantt_date_range)
        else:
            _min_date = min(_all_dates) if _all_dates else date.today()
            _max_date = max(_all_dates) if _all_dates else date.today() + timedelta(days=30)
        _span_days = max(1, (_max_date - _min_date).days + 1)
        month_day_rows = []
        _project_prefix_map = {}
        for _pi, (_, _pr) in enumerate(P.iterrows(), 1):
            _prefix = str(_pi); _project_prefix_map[str(_pr["name"])] = _prefix
            _ps = _as_iso_date(_pr.get("start_date")); _pe = _as_iso_date(_pr.get("end_date"))
            try: _psd = date.fromisoformat(_ps) if _ps else _min_date
            except Exception: _psd = _min_date
            try: _ped = date.fromisoformat(_pe) if _pe else _psd
            except Exception: _ped = _psd
            _left = max(0.0, min(100.0, (_psd-_min_date).days/_span_days*100))
            _width = max(0.8, min(100.0-_left, ((_ped-_psd).days+1)/_span_days*100))
            _status = str(_pr.get("status") or "Not Started"); _pct = int(pd.to_numeric(pd.Series([_pr.get("progress",0)]),errors="coerce").fillna(0).iloc[0])
            _color = str(_pr.get("bar_color") or GREEN)
            month_day_rows.append((_prefix,0,"folder",_color,str(_pr["name"]),fmt_date(_ps),fmt_date(_pe),f"{max(1,(_ped-_psd).days+1)} days",_status,f"{_pct}%",_left,_width,_color,_pct,"",False,"#F8FAFC"))
            _project_tasks = T[T["project_id"] == _pr["id"]].copy() if not T.empty else pd.DataFrame()
            _source_to_wbs = {}
            for _ti, (_, _tr) in enumerate(_project_tasks.iterrows(), 1):
                _tid = int(_tr["id"]); _sid = str(_tr.get("source_id") or _tid)
                _wbs = f"{_prefix}.{_ti}"; _source_to_wbs[_tid] = _wbs
                _ts = _as_iso_date(_tr.get("start_date")); _te = _as_iso_date(_tr.get("due_date"))
                try: _tsd = date.fromisoformat(_ts) if _ts else _psd
                except Exception: _tsd = _psd
                try: _ted = date.fromisoformat(_te) if _te else _tsd
                except Exception: _ted = _tsd
                _tleft = max(0.0,min(100.0,(_tsd-_min_date).days/_span_days*100)); _twidth=max(0.8,min(100.0-_tleft,((_ted-_tsd).days+1)/_span_days*100))
                _tstatus=str(_tr.get("status") or "Not Started"); _tpct=int(pd.to_numeric(pd.Series([_tr.get("progress",0)]),errors="coerce").fillna(0).iloc[0]); _ttype=str(_tr.get("task_type") or "Task")
                _tcolor={"Completed":GREEN,"In Progress":BLUE,"At Risk":ORANGE,"Overdue":RED,"Not Started":GRAY,"On Hold":ORANGE}.get(_tstatus,BLUE)
                _parent=_tr.get("parent_task_id"); _indent=2 if pd.notna(_parent) else 1
                month_day_rows.append((_wbs,_indent,"play" if _ttype.lower()!="milestone" else "doc",_tcolor,str(_tr.get("name") or "Task"),fmt_date(_ts),fmt_date(_te),f"{max(1,(_ted-_tsd).days+1)} days",_tstatus,f"{_tpct}%",_tleft,_twidth,_tcolor,_tpct,"Milestone" if _ttype.lower()=="milestone" else "",bool(str(_tr.get("dependency") or "").strip()),"#FFFFFF"))
        week_rows = list(month_day_rows)
        day_rows = [tuple(r) + ("",) for r in month_day_rows]


        # Apply search, project, status, view-type, assignee and date-range filters.
        proj_prefix_map = _project_prefix_map
        if g_proj_sel != "All Projects" and g_proj_sel in proj_prefix_map:
            pfx = proj_prefix_map[g_proj_sel]
            active_rows = [r for r in active_rows if r[0] == pfx or r[0].startswith(pfx + ".")]
        if g_stat_sel != "All Statuses":
            active_rows = [r for r in active_rows if r[8] == g_stat_sel or r[1] == 0]
        if g_view_sel == "Milestones Only":
            active_rows = [r for r in active_rows if r[14] or r[1] == 0]
        elif g_view_sel == "Tasks Only":
            active_rows = [r for r in active_rows if r[1] != 0 and not r[14]]
        if gantt_q:
            q_low = gantt_q.lower()
            active_rows = [r for r in active_rows if q_low in r[4].lower() or q_low in r[0].lower()]
        if isinstance(gantt_date_range, (tuple, list)) and len(gantt_date_range) == 2:
            _glo, _ghi = sorted(gantt_date_range)
            _filtered_rows = []
            active_rows = []
            for _r in active_rows:
                _rs = pd.to_datetime(_r[5], errors="coerce")
                _re = pd.to_datetime(_r[6], errors="coerce")
                if (pd.isna(_rs) or _rs.date() <= _ghi) and (pd.isna(_re) or _re.date() >= _glo): _filtered_rows.append(_r)
            active_rows = _filtered_rows

        # Build Timeline Header HTML according to scale
        if scale == "Month":
            _mid_date = _min_date + timedelta(days=_span_days // 2)
            _month_labels = [(_min_date + timedelta(days=round(i * (_span_days-1) / 11))) for i in range(12)]
            _month_ticks = "".join(f'<div style="padding:5px 0;border-right:1px solid #EDF2F7;color:{"#1D4ED8" if d.month == date.today().month and d.year == date.today().year else "#475569"}">{d.strftime("%b %d")}</div>' for d in _month_labels)
            _today_pct = max(0,min(100,((date.today()-_min_date).days/_span_days)*100))
            tl_top_hdr = f"""
            <div style="display:grid;grid-template-columns:repeat(3,1fr);border-bottom:1px solid #E2E8F0;background:#F8FAFC;font-size:11.5px;font-weight:700;color:#1E293B;text-align:center">
                <div style="padding:6px 0;border-right:1px solid #E2E8F0">{_min_date.strftime('%b %Y')}</div>
                <div style="padding:6px 0;border-right:1px solid #E2E8F0">{_mid_date.strftime('%b %Y')}</div>
                <div style="padding:6px 0">{_max_date.strftime('%b %Y')}</div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(12,1fr);background:#F8FAFC;font-size:10.5px;font-weight:600;color:#475569;text-align:center">{_month_ticks}</div>
            """
            today_line_html = f'<div style="position:absolute;top:0;bottom:0;left:{_today_pct:.2f}%;width:1.5px;background:#10B981;z-index:4;pointer-events:none"></div>' if _min_date <= date.today() <= _max_date else ""
        elif scale == "Week":
            weeks_meta = []
            for _wi in range(10):
                _ws = _min_date + timedelta(days=round(_wi * max(0, _span_days-1) / 10))
                _we = min(_max_date, _ws + timedelta(days=6))
                weeks_meta.append((f"{_ws:%b %d} – {_we:%b %d}", f"Week {_ws.isocalendar().week}", _ws <= date.today() <= _we))
            w_cells = "".join(
                f'<div style="padding:4px 2px;border-right:1px solid #E2E8F0;background:{"#D1FAE5" if is_cur else "#F8FAFC"}">'
                f'<div style="font-size:10px;font-weight:700;color:#0F172A;line-height:1.2">{rng}</div>'
                f'<div style="font-size:9.5px;color:{"#065F46" if is_cur else "#64748B"};font-weight:{"700" if is_cur else "500"}">{wk}</div></div>'
                for rng, wk, is_cur in weeks_meta
            )
            tl_top_hdr = f'<div style="display:grid;grid-template-columns:repeat(10,1fr);text-align:center">{w_cells}</div>'
            today_line_html = '<div style="position:absolute;top:0;bottom:0;left:20%;width:10%;background:rgba(16,185,129,0.08);border-left:1.5px solid #10B981;border-right:1px solid #A7F3D0;z-index:1;pointer-events:none"></div>'
        else:
            hours = ["8 AM", "9 AM", "10 AM", "11 AM", "12 PM", "1 PM", "2 PM", "3 PM", "4 PM", "5 PM", "6 PM"]
            h_cells = "".join(f'<div style="padding:5px 0;border-right:1px solid #EDF2F7">{h}</div>' for h in hours)
            tl_top_hdr = f"""
            <div style="display:flex;align-items:center;justify-content:space-between;padding:5px 12px;border-bottom:1px solid #E2E8F0;background:#F8FAFC">
                <div></div>
                <div style="font-size:12px;font-weight:700;color:#0F172A">{date.today().strftime('%a, %b %d, %Y')}</div>
                <div style="display:flex;gap:6px;align-items:center">
                    <span style="border:1px solid #CBD5E1;background:white;color:#00875A;font-size:11px;font-weight:600;padding:2px 8px;border-radius:5px">Today ▾</span>
                    <span style="border:1px solid #CBD5E1;background:white;font-size:11px;padding:2px 6px;border-radius:5px">📅</span>
                </div>
            </div>
            <div style="display:grid;grid-template-columns:repeat(11,1fr);background:#F8FAFC;font-size:10.5px;font-weight:600;color:#475569;text-align:center;position:relative">
                {h_cells}
                <span style="position:absolute;left:50.5%;bottom:-8px;background:#EF4444;color:white;font-size:9.5px;font-weight:700;padding:1px 7px;border-radius:4px;z-index:6">Today 1:30 PM</span>
            </div>
            """
            today_line_html = '<div style="position:absolute;top:0;bottom:0;left:55%;border-left:1.5px dashed #EF4444;z-index:5;pointer-events:none"></div>'

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
            <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;background:white;margin-top:8px">
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
            _parent_options = ["None (Top Level Task)"] + (T["name"].dropna().astype(str).unique().tolist() if not T.empty else [])
            parent_t = st.selectbox("Parent Task", _parent_options, key="add_t_parent")

        r2_a, r2_b = st.columns(2)
        with r2_a:
            t_name = st.text_input("Task Name *", placeholder="Enter task name", key="add_t_name")
        with r2_b:
            st.text_input("Task ID (Auto)", value="Will be generated automatically", disabled=True, key="add_t_id_auto")

        r3_a, r3_b = st.columns([1.35, 1])
        with r3_a:
            t_desc = st.text_area("Description", placeholder="Enter task description, objectives, and details...", height=108, key="add_t_desc")
        with r3_b:
            t_prio = st.selectbox("Priority *", ["🔴 High", "🟠 Medium", "🔵 Low"], index=0, key="add_t_prio")
            t_stat = st.selectbox("Status *", ["🟦 Not Started", "🔵 In Progress", "🟢 Completed", "🟠 At Risk"], index=0, key="add_t_stat")

        r4_a, r4_b, r4_c = st.columns(3)
        with r4_a:
            t_start = st.date_input("Start Date *", date.today(), key="add_t_start")
        with r4_b:
            t_end = st.date_input("End Date *", date.today() + timedelta(days=7), key="add_t_end")
        with r4_c:
            st.number_input("Duration (Days)", min_value=1, max_value=365, value=7, key="add_t_dur")

        r5_a, r5_b, r5_c = st.columns([1.1, 1.1, 1])
        with r5_a:
            assignees_pool = sorted(R["name"].dropna().astype(str).unique().tolist()) if not R.empty else [""]
            t_ass = st.selectbox("Assignee *", assignees_pool, index=0, key="add_t_ass")
        with r5_b:
            t_pct = st.slider("% Complete *", 0, 100, 0, key="add_t_pct")
        with r5_c:
            st.selectbox("Task Type", ["Task", "Milestone", "Subtask"], key="add_t_type")

        r6_a, r6_b, r6_c = st.columns([1.1, 1.1, 1])
        with r6_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            st.toggle("◆ Mark as Milestone", value=False, key="add_t_ms_tog")
        with r6_b:
            t_dep = st.selectbox("Dependencies", ["Select dependent tasks", "1.1 Define project scope", "1.2 Stakeholder alignment"], key="add_t_dep")
        with r6_c:
            st.multiselect("Tags", ["Planning", "Design", "Development", "Testing", "Launch", "Research", "Strategy"], placeholder="Select or enter tags", key="add_t_tags")

        r7_a, r7_b, r7_c = st.columns([1, 1, 1.2])
        with r7_a:
            st.text_input("Estimated Effort (Hours)", placeholder="e.g. 16", key="add_t_est")
        with r7_b:
            st.text_input("Actual Effort (Hours)", placeholder="e.g. 8", key="add_t_act")
        with r7_c:
            t_notes = st.text_area("Progress Notes", placeholder="Enter progress notes...", height=68, key="add_t_notes")

        st.markdown(
            clean_html(
                """
                <div style="font-size:12.5px;font-weight:700;color:#0F172A;margin:6px 0 6px 0">Attachments</div>
                <div style="border:1.5px dashed #CBD5E1;border-radius:10px;padding:16px;text-align:center;background:#F8FAFC;margin-bottom:14px">
                    <div style="font-size:13px;font-weight:600;color:#334155">📎 &nbsp;Drag and drop files here or <span style="color:#1473E6;text-decoration:underline">browse</span></div>
                    <div style="font-size:11px;color:#64748B;margin-top:4px">Supports: PDF, DOCX, XLSX, PPTX, JPG, PNG (Max 25MB each)</div>
                </div>
                <div style="font-size:12.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Checklist / Subtasks</div>
                """
            ),
            unsafe_allow_html=True,
        )

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

        f_chk, f_can, f_sav = st.columns([3.2, 0.9, 1.1])
        with f_chk:
            create_another = st.checkbox("Create another task after saving", key="add_t_another")
        with f_can:
            if st.button("Cancel", key="add_t_cancel", use_container_width=True):
                st.session_state.pop("task_modal", None)
                st.rerun()
        with f_sav:
            if st.button("Save Task", type="primary", key="add_t_save", use_container_width=True):
                clean_prio = t_prio.split()[-1]
                clean_stat = " ".join(t_stat.split()[1:])
                _parent_id = None
                if parent_t != "None (Top Level Task)" and not T.empty:
                    _parent_match = T[T["name"].astype(str) == str(parent_t)]
                    if not _parent_match.empty: _parent_id = int(_parent_match.iloc[0]["id"])
                _task_type_clean = str(st.session_state.get("add_t_type", "Task")).replace("💼 ", "").replace("◆ ", "")
                _new_task_id = sql(
                    "INSERT INTO tasks(name,project_id,team,status,priority,start_date,due_date,assignee,progress,dependency,notes,parent_task_id,description,task_type,duration) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (t_name.strip() or "New Project Task", pmap.get(proj, 1), "", clean_stat, clean_prio, t_start.isoformat(), t_end.isoformat(), t_ass, int(t_pct),
                     "" if t_dep == "Select dependent tasks" else t_dep, t_notes or t_desc, _parent_id, t_desc, _task_type_clean, max(1,(t_end-t_start).days+1)),
                )
                for _subtask_name in st.session_state.get("add_t_sub_items", []):
                    sql("INSERT INTO task_subtasks(task_id,name,completed) VALUES(?,?,0)", (int(_new_task_id), str(_subtask_name)))
                if _new_task_id is not None:
                    sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_new_task_id), "User", "Task created", t_name.strip() or "New Project Task"))
                log_event("Created", "Task", _new_task_id, t_name.strip() or "New Project Task")
                st.session_state.pop("add_t_sub_items", None)
                if not create_another:
                    st.session_state.pop("task_modal", None)
                st.rerun()

    @st.dialog("View Task", width="large")
    def view_task_modal(task_rec: dict) -> None:
        tname = task_rec.get("name", "Task")
        tproj = task_rec.get("project", "Project")
        tstart = task_rec.get("start_date", "May 6, 2024")
        tend = task_rec.get("end_date", "May 12, 2024")
        tdur = task_rec.get("duration", "7 days")
        tpct = int(task_rec.get("pct", 100))
        tstat = task_rec.get("status", "Completed")
        tprio = task_rec.get("priority", "High")
        tass = task_rec.get("assignee", "")
        _task_id_for_modal = task_rec.get("id")
        _task_desc = str(task_rec.get("description") or task_rec.get("notes") or "No description provided.")
        _task_type = str(task_rec.get("task_type") or "Task")
        _parent_task_name = "None"
        _task_subtasks = pd.DataFrame()
        _task_attachments = pd.DataFrame()
        _task_activity = pd.DataFrame()
        if _task_id_for_modal is not None:
            _parent_id = task_rec.get("parent_task_id")
            if _parent_id is not None and pd.notna(_parent_id):
                _parent_df = df("SELECT name FROM tasks WHERE id=?", (int(_parent_id),))
                if not _parent_df.empty: _parent_task_name = str(_parent_df.iloc[0]["name"])
            _child_tasks = df("SELECT id,name,status,due_date,progress FROM tasks WHERE parent_task_id=? ORDER BY id", (int(_task_id_for_modal),))
            _child_tasks["completed"] = (_child_tasks["status"].astype(str).str.lower() == "completed").astype(int) if not _child_tasks.empty else pd.Series(dtype=int)
            if not _child_tasks.empty: _child_tasks["source_table"] = "tasks"
            _manual_subtasks = df("SELECT id,name,completed,due_date FROM task_subtasks WHERE task_id=? ORDER BY id", (int(_task_id_for_modal),))
            if not _manual_subtasks.empty: _manual_subtasks["source_table"] = "task_subtasks"
            _task_subtasks = pd.concat([_child_tasks[["id","name","completed","due_date","source_table"]] if not _child_tasks.empty else pd.DataFrame(columns=["id","name","completed","due_date","source_table"]), _manual_subtasks[["id","name","completed","due_date","source_table"]] if not _manual_subtasks.empty else pd.DataFrame(columns=["id","name","completed","due_date","source_table"])], ignore_index=True)
            _task_attachments = df("SELECT id,filename,mime_type,size_bytes,uploaded_at FROM task_attachments WHERE task_id=? ORDER BY id DESC", (int(_task_id_for_modal),))
            _task_activity = df("SELECT actor,action,details,created_at FROM task_activity WHERE task_id=? ORDER BY id DESC LIMIT 5", (int(_task_id_for_modal),))
        _subtasks_done = int((_task_subtasks["completed"] == 1).sum()) if not _task_subtasks.empty else 0
        _subtasks_count = len(_task_subtasks)
        _subtasks_html = "".join(f"<div style='display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9'><span>{'✅' if int(sr.get('completed',0) or 0) else '☐'} &nbsp;{html.escape(str(sr.get('name','')))}</span><span style='color:#64748B'>{fmt_date(sr.get('due_date')) if sr.get('due_date') else ''}</span></div>" for _, sr in _task_subtasks.iterrows()) or "<div style='color:#64748B'>No subtasks recorded.</div>"
        _attachments_html = "".join(f"<div style='border:1px solid #EDF2F7;border-radius:8px;padding:8px 10px;display:flex;align-items:center;justify-content:space-between;margin-bottom:6px'><span>📎 &nbsp;{html.escape(str(ar.get('filename','Attachment')))} <span style='color:#64748B;font-size:10.5px'>({int(ar.get('size_bytes') or 0):,} bytes)</span></span><span>Saved</span></div>" for _, ar in _task_attachments.iterrows()) or "<div style='color:#64748B'>No attachments uploaded.</div>"
        _activity_html = "".join(f"<div style='display:grid;grid-template-columns:110px 145px 1fr;gap:8px;border-bottom:1px solid #F1F5F9;align-items:center;padding:4px 0'><span>{avatar_html(str(ar.get('actor') or 'User'))}</span><span style='color:#64748B;font-size:11.5px'>{html.escape(str(ar.get('created_at') or ''))}</span><span>{html.escape(str(ar.get('details') or ar.get('action') or 'Updated'))}</span></div>" for _, ar in _task_activity.iterrows()) or "<div style='color:#64748B'>No activity recorded yet.</div>"

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
                            {html.escape(_task_desc)}
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">👤 &nbsp;Assignee</div>
                        <div style="background:#F8FAFC;border:1px solid #EDF2F7;border-radius:8px;padding:14px;display:flex;align-items:center;gap:12px">
                            <div style="width:44px;height:44px;border-radius:50%;background:#0284C7;color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px">JM</div>
                            <div>
                                <div style="font-size:14px;font-weight:800;color:#0F172A">{html.escape(tass)}</div>
                                <div style="font-size:12px;color:#64748B">{html.escape(str(task_rec.get('role') or 'Assignee'))}</div>
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
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600">☰ &nbsp;{html.escape(_parent_task_name)}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:10px 12px;background:white">
                        <div style="font-size:11.5px;font-weight:700;color:#0F172A;margin-bottom:6px">Milestone</div>
                        <div style="background:#F8FAFC;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:600"><span style="color:#7C3AED">◆</span> &nbsp;{'Milestone' if _task_type.lower() == 'milestone' else 'Not a milestone'}</div>
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
                        <div>{''.join(tag_pill_html(x.strip()) for x in str(task_rec.get('tags') or '').split(',') if x.strip()) or '<span style="color:#64748B">No tags</span>'}</div>
                    </div>
                </div>
                <div style="display:grid;grid-template-columns:1.35fr 1fr;gap:12px;margin-bottom:12px">
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">☑ &nbsp;Checklist / Subtasks ({_subtasks_count})</div>
                        <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
                            <div style="flex:1;height:8px;background:#E2E8F0;border-radius:99px;overflow:hidden"><div style="width:{round(_subtasks_done/_subtasks_count*100) if _subtasks_count else 0}%;height:100%;background:#00875A"></div></div>
                            <span style="font-size:11.5px;color:#475569;font-weight:600">{_subtasks_done} of {_subtasks_count} completed</span>
                        </div>
                        <div style="font-size:12.5px;line-height:2.1">{_subtasks_html}</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📎 &nbsp;Attachments ({len(_task_attachments)})</div>
                        {_attachments_html}
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:14px">
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">💬 &nbsp;Recent Activity / Notes</div>
                    <div style="font-size:12.5px;line-height:2.2">{_activity_html}</div>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        if _task_id_for_modal is not None:
            with st.expander("Manage checklist, attachments, and notes", expanded=False):
                _sub_col, _att_col = st.columns(2)
                with _sub_col:
                    _new_subtask = st.text_input("Add subtask", key=f"view_add_sub_{_task_id_for_modal}")
                    _sub_due = st.date_input("Subtask due date", value=date.today(), key=f"view_sub_due_{_task_id_for_modal}")
                    if st.button("Save Subtask", key=f"view_save_sub_{_task_id_for_modal}"):
                        if _new_subtask.strip():
                            sql("INSERT INTO task_subtasks(task_id,name,completed,due_date) VALUES(?,?,0,?)", (int(_task_id_for_modal), _new_subtask.strip(), _sub_due.isoformat()))
                            sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_task_id_for_modal), "User", "Subtask added", _new_subtask.strip()))
                            st.rerun()
                    if not _task_subtasks.empty:
                        for _, _sr in _task_subtasks.iterrows():
                            _is_done = st.checkbox(str(_sr["name"]), value=bool(_sr["completed"]), key=f"view_sub_done_{int(_sr['id'])}")
                            if _is_done != bool(_sr["completed"]):
                                if str(_sr.get("source_table")) == "tasks":
                                    sql("UPDATE tasks SET status=?,progress=? WHERE id=?", ("Completed" if _is_done else "Not Started", 100 if _is_done else 0, int(_sr["id"])))
                                    sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_task_id_for_modal), "User", "Subtask status changed", str(_sr["name"])))
                                else:
                                    sql("UPDATE task_subtasks SET completed=? WHERE id=?", (1 if _is_done else 0, int(_sr["id"])))
                                st.rerun()
                with _att_col:
                    _att_file = st.file_uploader("Upload attachment", key=f"view_attachment_{_task_id_for_modal}")
                    if _att_file is not None and st.button("Save Attachment", key=f"view_save_attachment_{_task_id_for_modal}"):
                        sql("INSERT INTO task_attachments(task_id,filename,mime_type,content,size_bytes) VALUES(?,?,?,?,?)", (int(_task_id_for_modal), _att_file.name, _att_file.type, _att_file.getvalue(), len(_att_file.getvalue())))
                        sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_task_id_for_modal), "User", "Attachment uploaded", _att_file.name))
                        st.rerun()
                    _attachment_files = df("SELECT id,filename,mime_type,content FROM task_attachments WHERE task_id=? ORDER BY id DESC", (int(_task_id_for_modal),))
                    for _, _af in _attachment_files.iterrows():
                        _dlc, _delc = st.columns([3, 1])
                        with _dlc:
                            st.download_button(f"Download {_af['filename']}", data=bytes(_af["content"] or b""), file_name=str(_af["filename"]), mime=str(_af.get("mime_type") or "application/octet-stream"), key=f"view_download_att_{int(_af['id'])}", use_container_width=True)
                        with _delc:
                            if st.button("Delete", key=f"view_delete_att_{int(_af['id'])}", use_container_width=True):
                                sql("DELETE FROM task_attachments WHERE id=?", (int(_af["id"]),))
                                sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_task_id_for_modal), "User", "Attachment deleted", str(_af["filename"])))
                                st.rerun()
                    _note_entry = st.text_area("Add note", key=f"view_note_{_task_id_for_modal}")
                    if st.button("Save Note", key=f"view_save_note_{_task_id_for_modal}"):
                        if _note_entry.strip():
                            sql("INSERT INTO task_activity(task_id,actor,action,details) VALUES(?,?,?,?)", (int(_task_id_for_modal), "User", "Note", _note_entry.strip()))
                            sql("UPDATE tasks SET notes=? WHERE id=?", (_note_entry.strip(), int(_task_id_for_modal)))
                            st.rerun()

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
        tname = task_rec.get("name", "Task")
        tproj = task_rec.get("project", "Project")
        twbs = task_rec.get("wbs", "1.3")

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
            ass_list = sorted(R["name"].dropna().astype(str).unique().tolist()) if not R.empty else [str(task_rec.get("assignee") or "")]
            _current_assignee = str(task_rec.get("assignee") or "")
            if _current_assignee and _current_assignee not in ass_list: ass_list = [_current_assignee] + ass_list
            ed_ass = st.selectbox("👤 Assignee *", ass_list, index=ass_list.index(_current_assignee) if _current_assignee in ass_list else 0, key="ed_t_ass")

        e3_a, e3_b, e3_c, e3_d = st.columns(4)
        with e3_a:
            ed_start = st.date_input("Start Date *", date.fromisoformat(str(task_rec.get("start_date"))[:10]) if task_rec.get("start_date") and _as_iso_date(task_rec.get("start_date")) else date.today(), key="ed_t_start")
        with e3_b:
            ed_end = st.date_input("End Date *", date.fromisoformat(str(task_rec.get("end_date") or task_rec.get("due_date"))[:10]) if (task_rec.get("end_date") or task_rec.get("due_date")) and _as_iso_date(task_rec.get("end_date") or task_rec.get("due_date")) else date.today() + timedelta(days=7), key="ed_t_end")
        with e3_c:
            st.number_input("Duration (Days) *", min_value=1, max_value=365, value=7, key="ed_t_dur")
        with e3_d:
            ed_pct = st.number_input("% Complete *", min_value=0, max_value=100, value=int(task_rec.get("pct", 100)), key="ed_t_pct")

        e4_a, e4_b, e4_c, e4_d = st.columns(4)
        with e4_a:
            _prio_opts = ["⚑ High", "⚑ Medium", "⚑ Low"]
            _prio_default = "⚑ " + str(task_rec.get("priority") or "Medium")
            ed_prio = st.selectbox("Priority *", _prio_opts, index=_prio_opts.index(_prio_default) if _prio_default in _prio_opts else 1, key="ed_t_prio")
        with e4_b:
            _status_opts = ["🟢 Completed", "🔵 In Progress", "⚪ Not Started", "🟠 At Risk", "🔴 Overdue", "🟠 On Hold"]
            _status_label = {"Completed":"🟢 Completed", "In Progress":"🔵 In Progress", "Not Started":"⚪ Not Started", "At Risk":"🟠 At Risk", "Overdue":"🔴 Overdue", "On Hold":"🟠 On Hold"}.get(str(task_rec.get("status") or "Not Started"), "⚪ Not Started")
            ed_stat = st.selectbox("Status *", _status_opts, index=_status_opts.index(_status_label), key="ed_t_stat")
        with e4_c:
            st.selectbox("Task Type *", ["💼 Task", "◆ Milestone", "Subtask"], index=0, key="ed_t_type")
        with e4_d:
            st.selectbox("Parent Task", ["1.2 Stakeholder alignment", "1.1 Define project scope", "None"], index=0, key="ed_t_parent")

        e5_a, e5_b, e5_c = st.columns([0.9, 0.9, 1.4])
        with e5_a:
            st.markdown("<div style='font-size:12.5px;font-weight:600;color:#0F172A;margin-bottom:6px'>Milestone</div>", unsafe_allow_html=True)
            st.toggle("◆ Mark as Milestone", value=False, key="ed_t_ms")
        with e5_b:
            st.markdown("<div style='font-size:12.5px;margin-bottom:6px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("Milestone Type", ["Not a milestone", "Project Kick-off", "Design Sign-off", "Go-Live"], key="ed_t_ms_sel", label_visibility="collapsed")
        with e5_c:
            st.multiselect("🏷 Tags", ["Design", "UI/UX", "Planning", "Development", "Testing", "Launch"], default=["Design", "UI/UX"], key="ed_t_tags")

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
                        st.session_state.setdefault("ed_sub_items", []).append(_subtask)
                        st.session_state["ed_sub_in"] = ""
                        st.rerun()
            if st.session_state.get("ed_sub_items"):
                for _idx, _subtask in enumerate(st.session_state["ed_sub_items"], 1):
                    st.caption(f"☐  {_subtask}")
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:12.5px;line-height:2.2;margin-top:6px">
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span>✅ &nbsp;Create homepage wireframe</span><span style="color:#64748B">May 8, 2024 &nbsp;🗑</span></div>
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span>✅ &nbsp;Create product pages wireframe</span><span style="color:#64748B">May 10, 2024 &nbsp;🗑</span></div>
                        <div style="display:flex;justify-content:space-between"><span>✅ &nbsp;Review with stakeholders</span><span style="color:#64748B">May 12, 2024 &nbsp;🗑</span></div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with e6_r, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px">📎 &nbsp;Attachments</div>
                    <div style="border:1.5px dashed #CBD5E1;border-radius:8px;padding:10px;text-align:center;background:#F8FAFC;margin-bottom:8px">
                        <div style="font-size:12px;font-weight:600;color:#334155">☁️ &nbsp;Drag and drop files here or <span style="color:#1473E6">browse</span></div>
                        <div style="font-size:10px;color:#64748B">Supports: PDF, DOCX, XLSX, PPTX, JPG, PNG (Max 25MB each)</div>
                    </div>
                    <div style="border:1px solid #EDF2F7;border-radius:6px;padding:6px 8px;display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;font-size:12px">
                        <span><b style="color:#EF4444">[PDF]</b> Wireframe_Overview.pdf <span style="color:#64748B;font-size:10.5px">(2.4 MB)</span></span>
                        <span>⬇ &nbsp;🗑</span>
                    </div>
                    <div style="border:1px solid #EDF2F7;border-radius:6px;padding:6px 8px;display:flex;justify-content:space-between;align-items:center;font-size:12px">
                        <span><b style="color:#1473E6">[IMG]</b> Homepage_Wireframe.png <span style="color:#64748B;font-size:10.5px">(1.1 MB)</span></span>
                        <span>⬇ &nbsp;🗑</span>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            st.markdown("<div style='font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px'>📄 &nbsp;Additional Information <span style='font-weight:400;color:#64748B'>(Optional)</span></div>", unsafe_allow_html=True)
            ai1, ai2, ai3, ai4 = st.columns([0.8, 0.8, 1.2, 1.4])
            with ai1:
                st.text_input("Estimated Effort (Hours)", value="40", key="ed_est_h")
            with ai2:
                st.text_input("Actual Effort (Hours)", value="38", key="ed_act_h")
            with ai3:
                st.selectbox("Dependencies", ["1.2 Stakeholder alignment", "1.1 Define project scope", "None"], index=0, key="ed_dep_s")
            with ai4:
                ed_notes = st.text_area("Progress Notes", value=str(task_rec.get("notes") or ""), height=68, key="ed_p_notes")

        f_meta, f_can, f_sav = st.columns([3.3, 0.85, 1.15])
        with f_meta:
            st.markdown(
                clean_html(
                    f"""
                    <div style="display:flex;gap:24px;align-items:center;font-size:11.5px;color:#475569;padding-top:4px">
                        <div><span style="color:#64748B">Created by</span><br>{avatar_html('James Miller')} <span style="color:#64748B">May 6, 2024 9:15 AM</span></div>
                        <div><span style="color:#64748B">Last updated by</span><br>{avatar_html('Emily Rodriguez')} <span style="color:#64748B">May 12, 2024 3:45 PM</span></div>
                    </div>
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
                _task_id = task_rec.get("id")
                if not str(ed_name).strip():
                    st.error("Task Name is required.")
                elif _task_id is None:
                    st.error("This task is missing a database ID and cannot be updated safely.")
                else:
                    _status_clean = str(ed_stat).split(" ", 1)[-1]
                    _priority_clean = str(ed_prio).split(" ", 1)[-1]
                    _type_clean = str(st.session_state.get("ed_t_type", "Task")).replace("💼 ", "").replace("◆ ", "")
                    _duration_clean = max(1, (ed_end - ed_start).days + 1)
                    sql("UPDATE tasks SET name=?,project_id=?,status=?,priority=?,start_date=?,due_date=?,assignee=?,progress=?,notes=?,description=?,task_type=?,estimated_effort=?,actual_effort=?,duration=? WHERE id=?",
                        (str(ed_name).strip(), int(pmap.get(ed_proj, task_rec.get("project_id") or 0)), _status_clean, _priority_clean, ed_start.isoformat(), ed_end.isoformat(), ed_ass, int(ed_pct), ed_notes, ed_desc, _type_clean, float(st.session_state.get("ed_est_h", 0) or 0), float(st.session_state.get("ed_act_h", 0) or 0), _duration_clean, int(_task_id)))
                    log_event("Updated", "Task", int(_task_id), ed_name)
                    st.session_state.pop("task_modal", None)
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
        tl_date_range = st.date_input("Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="tl_d_top", label_visibility="collapsed")
    with th_b:
        if st.button("＋ Add Task", type="primary", use_container_width=True, key="tl_btn_add_top"):
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
        tl_ass_f = st.selectbox("Assignee", ["All Assignees"] + sorted(T.assignee.dropna().astype(str).unique().tolist()), key="tl_f_ass", label_visibility="collapsed")
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
    _tl_kpi_df = T.copy()
    _tl_range = st.session_state.get("tl_d_top", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)))
    if isinstance(_tl_range, (tuple, list)) and len(_tl_range) == 2:
        _tl_lo, _tl_hi = sorted(_tl_range)
        _tl_sd = pd.to_datetime(_tl_kpi_df.get("start_date", pd.Series(index=_tl_kpi_df.index, dtype=str)), errors="coerce").dt.date
        _tl_ed = pd.to_datetime(_tl_kpi_df.get("due_date", pd.Series(index=_tl_kpi_df.index, dtype=str)), errors="coerce").dt.date
        _tl_kpi_df = _tl_kpi_df[((_tl_sd.isna() | (_tl_sd <= _tl_hi)) & (_tl_ed.isna() | (_tl_ed >= _tl_lo)))].copy()
    if tl_proj_f != "All Projects" and not _tl_kpi_df.empty: _tl_kpi_df = _tl_kpi_df[_tl_kpi_df["project"] == tl_proj_f]
    if tl_stat_f != "All Statuses" and not _tl_kpi_df.empty: _tl_kpi_df = _tl_kpi_df[_tl_kpi_df["status"] == tl_stat_f]
    if tl_prio_f != "All Priorities" and not _tl_kpi_df.empty: _tl_kpi_df = _tl_kpi_df[_tl_kpi_df["priority"] == tl_prio_f]
    if tl_ass_f != "All Assignees" and not _tl_kpi_df.empty: _tl_kpi_df = _tl_kpi_df[_tl_kpi_df["assignee"].fillna("") == tl_ass_f]
    _tl_query = (tl_top_q + " " + tl_sub_q).strip().lower()
    if _tl_query and not _tl_kpi_df.empty:
        _tl_kpi_df = _tl_kpi_df[_tl_kpi_df.apply(lambda _r: _tl_query in str(_r.get("name", "")).lower() or _tl_query in str(_r.get("project", "")).lower() or _tl_query in str(_r.get("assignee", "")).lower(), axis=1)]
    _tl_total = len(_tl_kpi_df)
    _tl_status = _tl_kpi_df["status"].astype(str).str.lower() if _tl_total else pd.Series(dtype=str)
    _tl_completed = int((_tl_status == "completed").sum()) if _tl_total else 0
    _tl_progress = int((_tl_status == "in progress").sum()) if _tl_total else 0
    _tl_not_started = int(_tl_status.isin(["not started", "planned"]).sum()) if _tl_total else 0
    _tl_due = pd.to_datetime(_tl_kpi_df.get("due_date", pd.Series(dtype=str)), errors="coerce") if _tl_total else pd.Series(dtype="datetime64[ns]")
    _tl_overdue = int(((_tl_due.dt.date < date.today()) & (_tl_status != "completed")).sum()) if _tl_total else 0
    _tl_risk = int(_tl_status.isin(["at risk", "delayed"]).sum()) if _tl_total else 0
    _pct = lambda n: f"({n / _tl_total:.0%})" if _tl_total else "(0%)"
    kpi_specs = [
        (tk1, "#1473E6", "#FFFFFF", "▤", "Total Tasks", "#475569", str(_tl_total), ""),
        (tk2, "#00875A", "#FFFFFF", "✓", "Completed", "#475569", str(_tl_completed), _pct(_tl_completed)),
        (tk3, "#1473E6", "#FFFFFF", "❂", "In Progress", "#475569", str(_tl_progress), _pct(_tl_progress)),
        (tk4, "#64748B", "#FFFFFF", "●", "Not Started", "#475569", str(_tl_not_started), _pct(_tl_not_started)),
        (tk5, "#F28C00", "#FFFFFF", "!", "At Risk", "#D97706", str(_tl_risk), _pct(_tl_risk)),
        (tk6, "#EF4444", "#FFFFFF", "!", "Overdue", "#DC2626", str(_tl_overdue), _pct(_tl_overdue)),
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
    _status_colors = {"Completed": GREEN, "In Progress": BLUE, "At Risk": ORANGE, "Delayed": RED, "Overdue": RED, "On Hold": ORANGE, "Not Started": "#CBD5E1"}
    _tag_colors = {"Planning": ("#DBEAFE", "#1D4ED8"), "Design": ("#F3E8FF", "#7C3AED"), "Development": ("#DCFCE7", "#15803D"), "Testing": ("#FEE2E2", "#B91C1C")}
    _today = date.today()
    _range = st.session_state.get("tl_d_top", (_today - timedelta(days=30), _today + timedelta(days=60)))
    if not isinstance(_range, (tuple, list)) or len(_range) != 2:
        _range = (_today - timedelta(days=30), _today + timedelta(days=60))
    _range_start, _range_end = _range
    if _range_start > _range_end:
        _range_start, _range_end = _range_end, _range_start
    for _, _pr in P.iterrows():
        _pr_tasks = T[T["project_id"] == _pr["id"]].copy() if not T.empty else pd.DataFrame()
        _rows = []
        for _, _tr in _pr_tasks.iterrows():
            _sd = _as_iso_date(_tr.get("start_date")); _ed = _as_iso_date(_tr.get("due_date"))
            try: _sd_obj = date.fromisoformat(_sd) if _sd else None
            except Exception: _sd_obj = None
            try: _ed_obj = date.fromisoformat(_ed) if _ed else None
            except Exception: _ed_obj = None
            if _sd_obj and _sd_obj > _range_end: continue
            if _ed_obj and _ed_obj < _range_start: continue
            _status = str(_tr.get("status") or "Not Started")
            _priority = str(_tr.get("priority") or "Medium")
            _assignee = str(_tr.get("assignee") or "")
            _name = str(_tr.get("name") or "")
            _tags = str(_tr.get("tags") or "")
            _pctv = int(pd.to_numeric(pd.Series([_tr.get("progress", 0)]), errors="coerce").fillna(0).iloc[0])
            _days = max(1, ((_ed_obj - _sd_obj).days + 1)) if _sd_obj and _ed_obj else 1
            _tag = next((x.strip() for x in _tags.split(",") if x.strip()), str(_tr.get("task_type") or "Task"))
            _tag_bg, _tag_fg = _tag_colors.get(_tag, ("#E2E8F0", "#334155"))
            _rows.append({"id": int(_tr["id"]), "wbs": str(_tr.get("source_id") or _tr["id"]), "ico": "doc" if str(_tr.get("task_type", "Task")).lower() == "milestone" else "play",
                "ic_col": _status_colors.get(_status, GRAY), "name": _name, "project": str(_pr["name"]),
                "start_date": fmt_date(_sd), "end_date": fmt_date(_ed), "duration": f"{_days} days", "status": _status,
                "priority": _priority, "assignee": _assignee, "pct": max(0,min(100,_pctv)), "bar_col": _status_colors.get(_status, BLUE),
                "tag": _tag, "tags": _tags, "description": str(_tr.get("description") or ""), "notes": str(_tr.get("notes") or ""),
                "task_type": str(_tr.get("task_type") or "Task"), "project_id": int(_pr["id"]), "parent_task_id": int(_tr.get("parent_task_id")) if pd.notna(_tr.get("parent_task_id")) else None,
                "tag_bg": _tag_bg, "tag_fg": _tag_fg})
        if not _rows: continue
        _grp_pct = round(sum(x["pct"] for x in _rows) / len(_rows))
        _pstart = _as_iso_date(_pr.get("start_date")); _pend = _as_iso_date(_pr.get("end_date"))
        _duration = 0
        try: _duration = (date.fromisoformat(_pend) - date.fromisoformat(_pstart)).days + 1 if _pstart and _pend else 0
        except Exception: pass
        task_groups.append({"group_title": f"{_pr['name']} ({len(_rows)})", "project": str(_pr["name"]), "icon_bg": str(_pr.get("bar_color") or GREEN),
            "start": fmt_date(_pstart), "end": fmt_date(_pend), "dur": f"{_duration} days" if _duration else "—", "pct": _grp_pct,
            "bar_col": str(_pr.get("bar_color") or GREEN), "row_bg": "#F8FAFC", "tasks": _rows})

    # Avatar initials and colors matching Image 12
    tl_avatars = {
        "Dana Scott": ("DS", "#0284C7"),
        "Emily Rodriguez": ("EC", "#0284C7"),
        "James Miller": ("JM", "#0284C7"),
        "Sarah Kim": ("SL", "#0284C7"),
        "Alice Morgan": ("AL", "#0284C7"),
        "Ryan Cruz": ("RJ", "#0284C7"),
        "Michael Tan": ("MT", "#0284C7"),
    }

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
                    st.markdown(f"<div style='font-size:12px;color:#475569;padding-top:5px'><span style='color:#94A3B8'>1</span> &nbsp; <b style='color:#1E293B'>{wbs_id}</b></div>", unsafe_allow_html=True)
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
                    ini, abg = tl_avatars.get(t_item["assignee"], ("JM", "#0284C7"))
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
                    pct_v = t_item["pct"]
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
                with rc[10]:
                    st.markdown(
                        f"<div style='padding-top:4px'><span class='vk-tag' style='background:{t_item['tag_bg']};color:{t_item['tag_fg']}'>{html.escape(t_item['tag'])}</span></div>",
                        unsafe_allow_html=True,
                    )
                with rc[11]:
                    b_v, b_e = st.columns(2)
                    with b_v:
                    if st.button("👁",key=f"tl_v_{project_id}_{wbs_id}",help="View Task"):pass
                        st.session_state.task_modal = ("view", t_item)
                        st.rerun()
                    with b_e:
                        if st.button("🗑", key=f"tl_e_{wbs_id}", help="Delete Task"):
                            st.session_state["confirm_delete_task"] = {"task_id": t_item.get("id"), "name": t_item.get("name", ""), "wbs": wbs_id}
                            st.rerun()
                st.markdown("<hr style='margin:4px 0;border-color:#F1F5F9'>", unsafe_allow_html=True)

        # Pagination Footer
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;justify-content:space-between;padding-top:10px;font-size:12.5px;color:#475569">
                    <div style="display:flex;align-items:center;gap:8px">
                        <span>Show</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 10px;background:white;font-weight:600;color:#0F172A">25 &nbsp;▾</span>
                        <span>entries</span>
                        <span style="margin-left:14px">Showing 1 to 17 of 17 tasks</span>
                    </div>
                    <div style="display:flex;align-items:center;gap:6px">
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white">‹</span>
                        <span style="border-radius:6px;padding:4px 11px;background:#00875A;color:white;font-weight:700">1</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white">›</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    # Trigger Task Modal if active
    if "task_modal" in st.session_state:
        tm_type, tm_payload = st.session_state.task_modal
        if tm_type == "add":
            add_task_modal()
        elif tm_type == "view":
            view_task_modal(tm_payload)
        elif tm_type == "edit":
            edit_task_modal(tm_payload)

    if st.session_state.get("confirm_delete_task"):
        _delete_task_info = st.session_state["confirm_delete_task"]
        @st.dialog("Delete Task")
        def _confirm_task_delete_modal():
            st.warning(f"Delete '{_delete_task_info.get('name', 'this task')}'? This action cannot be undone.")
            _del_a, _del_b = st.columns(2)
            with _del_a:
                if st.button("Cancel", key="tl_delete_cancel", use_container_width=True):
                    st.session_state.pop("confirm_delete_task", None)
                    st.rerun()
            with _del_b:
                if st.button("Delete Task", type="primary", key="tl_delete_confirm", use_container_width=True):
                    _task_id = _delete_task_info.get("task_id")
                    if _task_id is not None:
                        for _child_table in ("task_subtasks", "task_attachments", "task_activity", "calendar_events"):
                            sql(f"DELETE FROM {_child_table} WHERE task_id=?", (int(_task_id),))
                        sql("DELETE FROM tasks WHERE id=?", (int(_task_id),))
                    st.session_state.pop("confirm_delete_task", None)
                    st.rerun()
        _confirm_task_delete_modal()


# ===========================================================================
# PAGE 5: RESOURCES (Exact Replica of Main View + Calendar, Assignments, Add & 7 Action Modals)
# ===========================================================================
elif page == "Resources":

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
            st.file_uploader("Upload Photo", type=["jpg", "jpeg", "png"], key="ar_photo", label_visibility="collapsed")
            st.markdown(
                "<div style='text-align:center;font-size:11px;color:#64748B;line-height:1.4'>JPG, PNG (Max 2MB)<br>Recommended size: 200 x 200</div>",
                unsafe_allow_html=True,
            )

        mid_l, mid_r = st.columns([1.15, 1])
        with mid_l, st.container(border=True):
            st.markdown("<div style='font-size:14px;font-weight:700;color:#0F172A;margin-bottom:8px'>⚙ &nbsp;Skills</div>", unsafe_allow_html=True)
            skill_pool = ["UI/UX", "Figma", "React", "JavaScript", "Python", "API", "Testing", "Automation", "Content", "SEO", "Azure", "CI/CD", "Analysis", "Requirements", "Planning", "Stakeholder", "Infrastructure", "Support", "Power BI", "Excel"]
            st.multiselect("Select Skills *", skill_pool, default=[], placeholder="🔍 Search or select skills...", key="ar_skills")
            chips_preview = "".join(
                f'<span class="vk-tag" style="background:{"#DBEAFE" if i % 2 == 0 else "#F1F5F9"};color:{"#1D4ED8" if i % 2 == 0 else "#334155"};margin-bottom:6px">{s}</span>'
                for i, s in enumerate(["UI/UX", "Figma", "React", "JavaScript", "Python", "API", "Testing", "Automation", "Content", "SEO", "Azure", "CI/CD"])
            )
            st.markdown(
                clean_html(
                    f"""
                    <div style="margin-top:6px;margin-bottom:8px">{chips_preview}</div>
                    <div style="font-size:12px;font-weight:600;color:#1473E6;cursor:pointer">＋ Add Custom Skill</div>
                    """
                ),
                unsafe_allow_html=True,
            )

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
                st.text_input("Availability End Date", placeholder="📅 Optional", key="ar_av_end")
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
                if r_name.strip():
                    role_clean = "Team Member" if r_role == "Select role" else r_role
                    email_clean = r_email.strip() or f"{r_name.strip().lower().replace(' ', '.')}@hpe.com"
                    sql(
                        "INSERT INTO resources(name,email,team,role,status,capacity,notes) VALUES(?,?,?,?,?,?,?)",
                        (r_name.strip(), email_clean, "IT", role_clean, r_stat, int(r_cap), r_notes),
                    )
                    log_event("Created", "Resource", None, r_name.strip())
                st.session_state.pop("res_modal", None)
                st.rerun()

    @st.dialog("Resource Calendar", width="large")
    def resource_calendar_modal(res_rec: dict) -> None:
        rname = res_rec.get("name", "Resource")
        rrole = res_rec.get("role", "UX Designer")
        remail = res_rec.get("email", "james.miller@hpe.com")
        rini = res_rec.get("ini", "JM")
        rcolor = res_rec.get("av_col", "#0284C7")
        ralloc = int(res_rec.get("alloc", 100))
        rstat = res_rec.get("status", "Allocated")

        st.markdown(
            clean_html(
                f"""
                <div style="display:flex;align-items:center;gap:12px;padding-bottom:12px;border-bottom:1px solid #E2E8F0;margin-bottom:12px">
                    <div style="width:40px;height:40px;border-radius:50%;background:#00875A;color:white;display:flex;align-items:center;justify-content:center;font-size:19px">📅</div>
                    <div>
                        <div style="font-size:20px;font-weight:800;color:#0F172A;line-height:1.1">Resource Calendar</div>
                        <div style="font-size:12.5px;color:#64748B">View availability, leave, and project assignments for this resource.</div>
                    </div>
                </div>
                <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 18px;background:#F8FAFC;display:grid;grid-template-columns:1.8fr 1fr 1fr 1.2fr 1fr;align-items:center;gap:12px;margin-bottom:12px">
                    <div style="display:flex;align-items:center;gap:12px">
                        <div style="width:46px;height:46px;border-radius:50%;background:{rcolor};color:white;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:16px">{rini}</div>
                        <div>
                            <div style="font-size:15px;font-weight:800;color:#0F172A">{html.escape(rname)}</div>
                            <div style="font-size:12px;color:#475569">{html.escape(rrole)}</div>
                            <div style="font-size:11.5px;color:#64748B">{html.escape(remail)}</div>
                        </div>
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
                        <div style="font-size:11px;color:#64748B">Current Allocation</div>
                        <div style="font-size:15px;font-weight:800;color:#0F172A">{ralloc}%</div>
                        <div style="height:7px;background:#E2E8F0;border-radius:99px;overflow:hidden;margin-top:3px"><div style="width:{min(100, ralloc)}%;height:100%;background:#00875A"></div></div>
                    </div>
                    <div style="border-left:1px solid #E2E8F0;padding-left:14px">
                        <div style="font-size:11px;color:#64748B;margin-bottom:3px">Availability Status</div>
                        {badge_html(rstat)}
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        if "rc_scale" not in st.session_state:
            st.session_state.rc_scale = "Month"

        tb1, tb2, tb3, tb4, tb5, tb6 = st.columns([1.4, 0.65, 1.1, 0.55, 0.55, 0.95])
        if "rc_calendar_month" not in st.session_state:
            st.session_state["rc_calendar_month"] = date.today().replace(day=1)
        with tb1:
            _rc_l, _rc_label, _rc_r = st.columns([0.5, 2.2, 0.5])
            with _rc_l:
                if st.button("‹", key="rc_prev_month", use_container_width=True):
                    _rm = st.session_state["rc_calendar_month"]
                    st.session_state["rc_calendar_month"] = (_rm.replace(day=1)-timedelta(days=1)).replace(day=1)
                    st.rerun()
            with _rc_label:
                st.markdown(f"<div style='font-size:15px;font-weight:800;color:#0F172A;text-align:center;padding:7px 0'>{st.session_state['rc_calendar_month'].strftime('%B %Y')}</div>", unsafe_allow_html=True)
            with _rc_r:
                if st.button("›", key="rc_next_month", use_container_width=True):
                    _rm = st.session_state["rc_calendar_month"]
                    st.session_state["rc_calendar_month"] = (_rm.replace(day=28)+timedelta(days=4)).replace(day=1)
                    st.rerun()
        with tb2:
            if st.button("Today", key="rc_today_btn", use_container_width=True):
                st.session_state["rc_selected_date"] = date.today()
                st.session_state["rc_calendar_month"] = date.today().replace(day=1)
                st.rerun()
        with tb3:
            tile_toggle("", ["Month", "Week", "Day"], key="rc_scale", default="Month")
        with tb6:
            if st.button("＋ Add Event ▾", type="primary", key="rc_add_ev_btn", use_container_width=True):
                st.session_state["calendar_add_event"] = True
                st.session_state["page"] = "Calendar"
                st.rerun()

        cal_l, cal_r = st.columns([2.55, 1.05])
        with cal_l:
            _rc_month = st.session_state.get("rc_calendar_month", date.today().replace(day=1))
            _rc_first = _rc_month.replace(day=1)
            _rc_grid_start = _rc_first - timedelta(days=(_rc_first.weekday()+1)%7)
            _rc_events = {}
            _rc_tasks = T[T["assignee"].astype(str) == str(rname)].copy() if not T.empty else pd.DataFrame()
            for _, _rtask in _rc_tasks.iterrows():
                _rd = _as_iso_date(_rtask.get("start_date")) or _as_iso_date(_rtask.get("due_date"))
                if not _rd: continue
                try: _rday = date.fromisoformat(_rd)
                except Exception: continue
                _rproj = str(_rtask.get("project") or "")
                _rc_events.setdefault(_rday, []).append((_rproj or str(_rtask.get("name") or "Task"), str(_rtask.get("name") or "Task"), "g"))
            _rc_custom = df("SELECT ce.*,p.name AS project FROM calendar_events ce LEFT JOIN projects p ON p.id=ce.project_id WHERE ce.resources LIKE ? ORDER BY ce.start_date", (f"%{rname}%",))
            for _, _ev in _rc_custom.iterrows():
                try: _rday = date.fromisoformat(str(_ev.get("start_date"))[:10])
                except Exception: continue
                _rc_events.setdefault(_rday, []).append((str(_ev.get("title") or "Event"), str(_ev.get("start_time") or "Event"), "p"))
            rc_weeks = []
            for _rw in range(6):
                _cells=[]
                for _rdx in range(7):
                    _rday=_rc_grid_start+timedelta(days=_rw*7+_rdx)
                    _cells.append((str(_rday.day), _rday.month == _rc_month.month, _rc_events.get(_rday, [])))
                rc_weeks.append(_cells)
            pill_styles = {
                "g": ("#DCFCE7", "#15803D"),
                "p": ("#F3E8FF", "#7C3AED"),
                "r": ("#FEE2E2", "#B91C1C"),
            }
            grid_rows_html = ""
            for wk in rc_weeks:
                cells_html = ""
                for d_str, in_m, ev_list in wk:
                    is_sel = in_m and d_str == str(st.session_state.get("rc_selected_date", date.today()).day) and _rc_month.month == st.session_state.get("rc_selected_date", date.today()).month
                    bdr = "2px solid #3B82F6" if is_sel else "1px solid #EDF2F7"
                    cbg = "#EFF6FF" if is_sel else ("#FFFFFF" if in_m else "#F8FAFC")
                    dcol = "#1D4ED8" if is_sel else ("#0F172A" if in_m else "#94A3B8")
                    ev_pills = ""
                    for etitle, esub, etone in ev_list:
                        pbg, pfg = pill_styles[etone]
                        ev_pills += f'<div style="background:{pbg};color:{pfg};border-radius:5px;padding:3px 6px;font-size:10px;font-weight:600;margin-top:3px;line-height:1.25">{etitle}<br>{esub}</div>'
                    cells_html += f'<div style="min-height:84px;border:{bdr};background:{cbg};padding:5px 6px"><div style="font-size:11.5px;font-weight:700;color:{dcol}">{d_str}</div>{ev_pills}</div>'
                grid_rows_html += f'<div style="display:grid;grid-template-columns:repeat(7,1fr)">{cells_html}</div>'

            st.markdown(
                clean_html(
                    f"""
                    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;background:white">
                        <div style="display:grid;grid-template-columns:repeat(7,1fr);background:#F8FAFC;border-bottom:1px solid #E2E8F0;text-align:center;font-size:11.5px;font-weight:700;color:#1E293B;padding:7px 0">
                            <div>Sun</div><div>Mon</div><div>Tue</div><div>Wed</div><div>Thu</div><div>Fri</div><div>Sat</div>
                        </div>
                        {grid_rows_html}
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        with cal_r:
            st.markdown(
                clean_html(
                    """
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:10px">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">Legend</div>
                        <div style="font-size:12px;line-height:2.0;color:#334155">
                            <div><span style="color:#00875A;font-size:14px">●</span> &nbsp;Project Assignment</div>
                            <div><span style="color:#9333EA;font-size:14px">●</span> &nbsp;Meeting / Event</div>
                            <div><span style="color:#EF4444;font-size:14px">●</span> &nbsp;Leave</div>
                            <div><span style="color:#60A5FA;font-size:14px">■</span> &nbsp;Selected Date</div>
                            <div><span style="color:#CBD5E1;font-size:14px">●</span> &nbsp;Non-Working Day</div>
                        </div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white;margin-bottom:10px">
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">Selected Date</div>
                        <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">📅 &nbsp;Thu, May 9, 2024</div>
                        <div style="font-size:12.5px;color:#0F172A;font-weight:700"><span style="color:#00875A">●</span> &nbsp;Website Revamp</div>
                        <div style="font-size:11.5px;color:#64748B;padding-left:18px;margin-bottom:10px">9:00 AM - 5:00 PM<br>50% allocation</div>
                        <div style="background:#EFF6FF;border:1px solid #BFDBFE;color:#1D4ED8;border-radius:7px;padding:7px;text-align:center;font-size:12px;font-weight:700">＋ Add Event for this Date</div>
                    </div>
                    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;background:white">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
                            <span style="font-size:13px;font-weight:700;color:#0F172A">Upcoming Events</span>
                            <span style="font-size:11.5px;color:#1473E6;font-weight:600">View All</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #F1F5F9;font-size:12px">
                            <div><span style="color:#EF4444">📅</span> &nbsp;<b>Annual Leave</b><br><span style="font-size:11px;color:#64748B;padding-left:20px">May 20 – May 22, 2024</span></div>
                            <span>⋮</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #F1F5F9;font-size:12px">
                            <div><span style="color:#9333EA">📅</span> &nbsp;<b>Design Review</b><br><span style="font-size:11px;color:#64748B;padding-left:20px">May 1, 2024 2:00 PM – 4:00 PM</span></div>
                            <span>⋮</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;font-size:12px">
                            <div><span style="color:#9333EA">📅</span> &nbsp;<b>Team Meeting</b><br><span style="font-size:11px;color:#64748B;padding-left:20px">May 15, 2024 10:00 AM – 11:00 AM</span></div>
                            <span>⋮</span>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

        f_sp, f_cl = st.columns([4.5, 0.8])
        with f_cl:
            if st.button("Close", key="rc_close_btn", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()

    @st.dialog("Project Assignments", width="large")
    def project_assignments_modal(res_rec: dict) -> None:
        rname = res_rec.get("name", "Resource")
        rrole = res_rec.get("role", "UX Designer")
        remail = res_rec.get("email", "james.miller@hpe.com")
        rini = res_rec.get("ini", "JM")
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
            st.markdown("<div style='font-size:14.5px;font-weight:700;color:#0F172A;padding-top:6px'>💼 &nbsp;Current Project Assignments (3)</div>", unsafe_allow_html=True)
        with pa_h2:
            if st.button("＋ Assign to Project", type="primary", key="pa_assign_btn", use_container_width=True):
                st.session_state.res_modal = ("assign", res_rec)
                st.rerun()

        st.markdown(
            clean_html(
                f"""
                <div class="vk-table-wrap" style="margin-top:6px;margin-bottom:14px">
                    <table class="vk-table">
                        <thead>
                            <tr><th>Project</th><th>Role</th><th>Allocation</th><th>Start Date</th><th>End Date</th><th>Status</th><th>Primary Task(s)</th><th style="text-align:center">Actions</th></tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><span style="color:#1473E6;font-weight:700">Website Revamp</span></td>
                                <td>UX Designer</td>
                                <td>{progress_html(50, GREEN)}</td>
                                <td>Apr 24, 2024</td>
                                <td>Jun 30, 2024</td>
                                <td>{badge_html('In Progress')}</td>
                                <td>Wireframes, UI/UX design</td>
                                <td style="text-align:center;font-weight:700">⋮</td>
                            </tr>
                            <tr>
                                <td><span style="color:#1473E6;font-weight:700">Marketing Campaign</span></td>
                                <td>UI Designer</td>
                                <td>{progress_html(30, GREEN)}</td>
                                <td>May 6, 2024</td>
                                <td>Jun 25, 2024</td>
                                <td>{badge_html('In Progress')}</td>
                                <td>Landing pages, creatives</td>
                                <td style="text-align:center;font-weight:700">⋮</td>
                            </tr>
                            <tr>
                                <td><span style="color:#1473E6;font-weight:700">IT Infrastructure Upgrade</span></td>
                                <td>UX Support</td>
                                <td>{progress_html(20, GREEN)}</td>
                                <td>Jun 1, 2024</td>
                                <td>Jun 20, 2024</td>
                                <td>{badge_html('In Progress')}</td>
                                <td>UI assets, documentation</td>
                                <td style="text-align:center;font-weight:700">⋮</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            st.markdown("<div style='font-size:14.5px;font-weight:700;color:#0F172A;margin-bottom:6px'>🕒 &nbsp;Workload Overview</div>", unsafe_allow_html=True)
            wo_l, wo_m, wo_r = st.columns([0.9, 1.15, 1.85])
            with wo_l:
                fig_pa = go.Figure(
                    go.Pie(
                        values=[50, 30, 20],
                        labels=["Website Revamp", "Marketing Campaign", "IT Infrastructure Upgrade"],
                        hole=0.68,
                        sort=False,
                        marker=dict(colors=[GREEN, BLUE, PURPLE]),
                        textinfo="none",
                    )
                )
                fig_pa.update_layout(
                    height=165,
                    margin=dict(l=4, r=4, t=4, b=4),
                    showlegend=False,
                    paper_bgcolor="rgba(0,0,0,0)",
                    annotations=[
                        dict(
                            text="<b style='font-size:18px;color:#0F172A'>100%</b><br><span style='font-size:10px;color:#64748B'>Total Allocation</span>",
                            x=0.5,
                            y=0.5,
                            showarrow=False,
                        )
                    ],
                )
                st.plotly_chart(fig_pa, use_container_width=True, config={"displayModeBar": False})
            with wo_m:
                st.markdown(
                    clean_html(
                        """
                        <div style="font-size:12px;line-height:2.3;padding-top:26px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;Website Revamp</span><b>50%</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> &nbsp;Marketing Campaign</span><b>30%</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#7C3AED">●</span> &nbsp;IT Infrastructure Upgrade</span><b>20%</b></div>
                        </div>
                        """
                    ),
                    unsafe_allow_html=True,
                )
            with wo_r:
                st.markdown("<div style='font-size:12.5px;font-weight:700;color:#0F172A;margin-bottom:2px'>Weekly Allocation</div>", unsafe_allow_html=True)
                wk_df = pd.DataFrame(
                    {
                        "Day": ["Mon", "Tue", "Wed", "Thu", "Fri"],
                        "Website Revamp": [50, 50, 50, 50, 50],
                        "Marketing Campaign": [30, 30, 30, 30, 30],
                        "IT Infrastructure": [20, 20, 20, 20, 20],
                    }
                )
                fig_wk = px.bar(
                    wk_df,
                    x="Day",
                    y=["Website Revamp", "Marketing Campaign", "IT Infrastructure"],
                    color_discrete_map={"Website Revamp": GREEN, "Marketing Campaign": BLUE, "IT Infrastructure": PURPLE},
                )
                fig_wk.update_layout(
                    height=155,
                    margin=dict(l=0, r=0, t=4, b=0),
                    showlegend=False,
                    xaxis_title=None,
                    yaxis=dict(title=None, range=[0, 150], ticksuffix="%"),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                )
                st.plotly_chart(fig_wk, use_container_width=True, config={"displayModeBar": False})

        pa_tabs = st.tabs(["📅 Assigned Tasks (8)", "☆ Milestones (2)", "🕒 Time Allocation"])
        with pa_tabs[0]:
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-table-wrap">
                        <table class="vk-table">
                            <thead>
                                <tr><th>ID</th><th>Task Name</th><th>Project</th><th>Start Date</th><th>End Date</th><th>Status</th><th>Priority</th><th>% Complete</th><th>Due In</th></tr>
                            </thead>
                            <tbody>
                                <tr><td>1.1</td><td><b>Define project scope</b></td><td>Website Revamp</td><td>Apr 24, 2024</td><td>Apr 26, 2024</td><td>{badge_html('Completed')}</td><td>{badge_html('Medium')}</td><td>{progress_html(100, GREEN)}</td><td>-</td></tr>
                                <tr><td>1.3</td><td><b>Create wireframes</b></td><td>Website Revamp</td><td>May 6, 2024</td><td>May 12, 2024</td><td>{badge_html('Completed')}</td><td>{badge_html('High')}</td><td>{progress_html(100, GREEN)}</td><td>-</td></tr>
                                <tr><td>1.4</td><td><b>UI/UX design</b></td><td>Website Revamp</td><td>May 13, 2024</td><td>May 20, 2024</td><td>{badge_html('In Progress')}</td><td>{badge_html('High')}</td><td>{progress_html(70, BLUE)}</td><td><b>5 days</b></td></tr>
                                <tr><td>4.1</td><td><b>Content creation</b></td><td>Marketing Campaign</td><td>May 6, 2024</td><td>May 20, 2024</td><td>{badge_html('In Progress')}</td><td>{badge_html('Medium')}</td><td>{progress_html(50, BLUE)}</td><td><b>8 days</b></td></tr>
                                <tr><td>4.2</td><td><b>Campaign setup</b></td><td>Marketing Campaign</td><td>May 21, 2024</td><td>Jun 10, 2024</td><td>{badge_html('Not Started')}</td><td>{badge_html('Low')}</td><td>{progress_html(0, GRAY)}</td><td><b>21 days</b></td></tr>
                            </tbody>
                        </table>
                    </div>
                    <div style="font-size:12px;font-weight:700;color:#1473E6;margin-top:8px">☰ &nbsp;View All Tasks (8)</div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with pa_tabs[1]:
            st.info("2 Milestones linked to James Miller: Design Sign-off (Apr 30, 2024) and UAT Completion (Jun 15, 2024).")
        with pa_tabs[2]:
            st.info("Weekly allocation: 40 hrs/week (20 hrs Website Revamp, 12 hrs Marketing Campaign, 8 hrs IT Infrastructure Upgrade).")

        f_sp, f_cl = st.columns([4.5, 0.8])
        with f_cl:
            if st.button("Close", key="pa_close_btn", use_container_width=True):
                st.session_state.pop("res_modal", None)
                st.rerun()

    @st.dialog("Resource Action")
    def resource_action_submodal(act_mode: str, res_rec: dict) -> None:
        rname = res_rec.get("name", "Resource")
        rrole = res_rec.get("role", "UX Designer")
        remail = res_rec.get("email", "james.miller@hpe.com")
        rini = res_rec.get("ini", "JM")
        rcolor = res_rec.get("av_col", "#0284C7")
        rskills = res_rec.get("skills", ["UI/UX", "Figma"])
        rproj = res_rec.get("project", "—")
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
            roles_opt = ["UX Designer", "Business Analyst", "Frontend Developer", "Project Manager", "Backend Developer", "QA Engineer", "Content Specialist", "DevOps Engineer", "IT Support", "Data Analyst"]
            new_rl = st.selectbox("Role *", roles_opt, index=roles_opt.index(rrole) if rrole in roles_opt else 0, key="sub_ed_rl")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_ed_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Save Changes", type="primary", key="sub_ed_sav", use_container_width=True):
                    sql("UPDATE resources SET name=?, email=?, role=? WHERE name=?", (new_nm.strip(), new_em.strip(), new_rl, rname))
                    log_event("Updated", "Resource", None, new_nm.strip())
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "assign":
            st.markdown("<div style='font-size:12.5px;font-weight:700;color:#0F172A;margin-bottom:6px'>Select Project(s) *</div>", unsafe_allow_html=True)
            st.checkbox("Website Revamp", value=True, key="sub_as_p1")
            st.checkbox("Marketing Campaign", value=False, key="sub_as_p2")
            st.checkbox("Sustainability Initiative", value=False, key="sub_as_p3")
            st.checkbox("IT Infrastructure Upgrade", value=False, key="sub_as_p4")
            st.number_input("Allocation (%)", min_value=0, max_value=200, value=ralloc, step=10, key="sub_as_pct")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_as_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Assign", type="primary", key="sub_as_ok", use_container_width=True):
                    log_event("Assigned", "Resource", None, rname)
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
                    log_event("Updated Skills", "Resource", None, rname)
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "availability":
            st.date_input("Availability Start Date *", date.today(), key="sub_av_st")
            st.text_input("Availability End Date", placeholder="📅 Optional", key="sub_av_en")
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
                    sql("UPDATE resources SET capacity=? WHERE name=?", (int(new_cap), rname))
                    log_event("Updated Availability", "Resource", None, rname)
                    st.session_state.pop("res_modal", None)
                    st.rerun()

        elif act_mode == "leave":
            lv1, lv2 = st.columns(2)
            with lv1:
                st.date_input("Leave Start Date *", date(2024, 5, 1), key="sub_lv_st")
            with lv2:
                st.date_input("Leave End Date *", date(2024, 5, 10), key="sub_lv_en")
            st.selectbox("Reason (Optional)", ["Vacation", "Sick Leave", "Personal Leave", "Training"], key="sub_lv_rs")
            st.text_input("Notes (Optional)", placeholder="Enter notes...", key="sub_lv_nt")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Cancel", key="sub_lv_can", use_container_width=True):
                    st.session_state.pop("res_modal", None)
                    st.rerun()
            with c2:
                if st.button("Set on Leave", type="primary", key="sub_lv_ok", use_container_width=True):
                    sql("UPDATE resources SET status='On Leave', capacity=0 WHERE name=?", (rname,))
                    log_event("Set on Leave", "Resource", None, rname)
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
        res_date_range = st.date_input("Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="res_d_top", label_visibility="collapsed")
    with rh_b:
        if st.button("＋ Add Resource", type="primary", use_container_width=True, key="res_btn_add_top"):
            st.session_state.res_modal = ("add", None)
            st.rerun()

    st.write("")
    # --- 5 KPI CARDS ROW (Exact Match to Image 20) ---
    rk1, rk2, rk3, rk4, rk5 = st.columns(5)
    res_kpis = [
        (rk1, "#1473E6", "👥", "Total Resources", "#334155", "24", ""),
        (rk2, "#00875A", "✓", "Available", "#334155", "8", "(33%)"),
        (rk3, "#1473E6", "🕒", "Allocated", "#334155", "14", "(58%)"),
        (rk4, "#F28C00", "!", "Overallocated", "#DC2626", "2", "(8%)"),
        (rk5, "#64748B", "👤", "On Leave", "#334155", "0", "(0%)"),
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
                        <span><span style="color:#00875A">●</span> Billable</span>
                        <span><span style="color:#1473E6">●</span> Non-Billable</span>
                        <span><span style="color:#94A3B8">●</span> Leave</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )
        ra_df = pd.DataFrame(
            {
                "Month": ["Apr 2024", "May 2024", "Jun 2024"],
                "Billable": [6, 6.2, 6.8],
                "Non-Billable": [3.5, 4.2, 4.5],
                "Leave": [2, 2.2, 2.2],
            }
        )
        fig_ra = px.bar(
            ra_df,
            x="Month",
            y=["Billable", "Non-Billable", "Leave"],
            color_discrete_map={"Billable": GREEN, "Non-Billable": BLUE, "Leave": "#94A3B8"},
        )
        fig_ra.update_layout(
            height=195,
            margin=dict(l=0, r=0, t=6, b=0),
            showlegend=False,
            xaxis_title=None,
            yaxis=dict(title=None, range=[0, 16]),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            bargap=0.45,
        )
        st.plotly_chart(fig_ra, use_container_width=True, config={"displayModeBar": False})

    with rc_2, st.container(border=True):
        st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;margin-bottom:6px'>Utilization</div>", unsafe_allow_html=True)
        u_l, u_r = st.columns([1.15, 1])
        with u_l:
            fig_ut = go.Figure(
                go.Pie(
                    values=[58, 25, 17],
                    labels=["Billable", "Non-Billable", "Available"],
                    hole=0.66,
                    sort=False,
                    direction="clockwise",
                    marker=dict(colors=[GREEN, BLUE, "#94A3B8"]),
                    textinfo="none",
                )
            )
            fig_ut.update_layout(
                height=195,
                margin=dict(l=4, r=4, t=4, b=4),
                showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)",
                annotations=[
                    dict(
                        text="<b style='font-size:20px;color:#0F172A'>58%</b><br><span style='font-size:10.5px;color:#475569'>Team Utilization</span>",
                        x=0.5,
                        y=0.5,
                        showarrow=False,
                    )
                ],
            )
            st.plotly_chart(fig_ut, use_container_width=True, config={"displayModeBar": False})
        with u_r:
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:12.5px;line-height:2.4;padding-top:32px">
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;Billable</span><b>58%</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> &nbsp;Non-Billable</span><b>25%</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#94A3B8">●</span> &nbsp;Available</span><b>17%</b></div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

    with rc_3, st.container(border=True):
        st.markdown("<div style='font-size:15px;font-weight:700;color:#0F172A;margin-bottom:12px'>Resource Status</div>", unsafe_allow_html=True)
        rs_bars = [
            ("Available", 8, 42, "#6EE7B7"),
            ("Allocated", 14, 78, "#1473E6"),
            ("Overallocated", 2, 16, "#F87171"),
            ("On Leave", 0, 0, "#E2E8F0"),
        ]
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
        res_skill_f = st.selectbox("Skill", ["All Skills"] + sorted({x.strip() for v in R.notes.dropna().astype(str) for x in v.split(",") if x.strip()}), key="res_f_skill", label_visibility="collapsed")
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
    for _, _res in R.iterrows():
        _rid = int(_res["id"]); _name = str(_res.get("name") or ""); _notes = str(_res.get("notes") or "")
        _skills = [x.strip() for x in _notes.split(",") if x.strip()]
        _assigned_tasks = T[T["assignee"].astype(str) == _name].copy() if not T.empty else pd.DataFrame()
        if not _assigned_tasks.empty and isinstance(res_date_range, (tuple, list)) and len(res_date_range) == 2:
            _rlo, _rhi = sorted(res_date_range)
            _rsd = pd.to_datetime(_assigned_tasks["start_date"],errors="coerce").dt.date
            _red = pd.to_datetime(_assigned_tasks["due_date"],errors="coerce").dt.date
            _assigned_tasks = _assigned_tasks[(_rsd.isna() | (_rsd <= _rhi)) & (_red.isna() | (_red >= _rlo))].copy()
        _projects_assigned = list(dict.fromkeys(_assigned_tasks["project"].dropna().astype(str).tolist())) if not _assigned_tasks.empty else []
        _project_label = ", ".join(_projects_assigned) if _projects_assigned else "—"
        _alloc = 0
        _status = str(_res.get("status") or "Available")
        if _status.lower() not in {"available", "on leave", "inactive"}: _status = "Allocated" if not _assigned_tasks.empty else _status
        _avail = f"{max(0,100-_alloc)}%"
        _ini = "".join(part[0] for part in _name.split()[:2]).upper() or "U"
        res_directory.append({"id":_rid,"ini":_ini,"av_col":"#0284C7","name":_name,"email":str(_res.get("email") or ""),"role":str(_res.get("role") or "Team Member"),"skills":_skills,"sk_cols":[("#DBEAFE","#1D4ED8") if i%2==0 else ("#F3E8FF","#7C3AED") for i,_ in enumerate(_skills)],"project":_project_label,"alloc":_alloc,"avail":_avail,"status":_status})

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

        # Pagination Footer
        st.markdown(
            clean_html(
                """
                <div style="display:flex;align-items:center;justify-content:space-between;padding-top:10px;font-size:12.5px;color:#475569">
                    <div style="display:flex;align-items:center;gap:8px">
                        <span>Show</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 10px;background:white;font-weight:600;color:#0F172A">10 &nbsp;▾</span>
                        <span>entries</span>
                    </div>
                    <div style="display:flex;align-items:center;gap:6px">
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white">‹</span>
                        <span style="border-radius:6px;padding:4px 11px;background:#00875A;color:white;font-weight:700">1</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 11px;background:white;font-weight:600">2</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 11px;background:white;font-weight:600">3</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 10px;background:white">›</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

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
            ev_type = st.selectbox("Event Type *", ["🟢 Project Assignment", "🔵 Meeting / Event", "🔴 Leave", "🟣 Milestone"], key="ae_type")
        with e2:
            ev_title = st.text_input("Title *", placeholder="Enter event title", key="ae_title")

        e3, e4 = st.columns(2)
        with e3:
            ev_proj = st.selectbox("Project *", ["Select project"] + projects_list, key="ae_proj")
        with e4:
            st.selectbox("Task (Optional)", ["Select task (optional)"] + T.name.tolist(), key="ae_task")

        st.multiselect("Resource(s) *", sorted(R["name"].dropna().astype(str).unique().tolist()) if not R.empty else [], key="ae_res")

        dt1, dt2, dt3, dt4 = st.columns([1.2, 0.9, 1.2, 0.9])
        with dt1:
            ev_start_d = st.date_input("Start Date & Time *", date.today(), key="ae_sd")
        with dt2:
            st.markdown("<div style='font-size:12.5px;margin-bottom:4px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("Start Time", ["09:00 AM", "10:00 AM", "01:00 PM", "02:00 PM"], index=0, key="ae_st", label_visibility="collapsed")
        with dt3:
            ev_end_d = st.date_input("End Date & Time *", date.today(), key="ae_ed")
        with dt4:
            st.markdown("<div style='font-size:12.5px;margin-bottom:4px'>&nbsp;</div>", unsafe_allow_html=True)
            st.selectbox("End Time", ["05:00 PM", "11:00 AM", "03:00 PM", "04:00 PM"], index=0, key="ae_et", label_visibility="collapsed")

        al1, al2, rp_col = st.columns([1.5, 0.45, 1.95])
        with al1:
            st.text_input("Allocation (Optional)", value="50", key="ae_alloc")
            st.markdown("<div style='font-size:11px;color:#64748B;margin-top:-6px'>Enter allocation percentage for this assignment.</div>", unsafe_allow_html=True)
        with al2:
            st.markdown("<div style='margin-top:28px;border:1px solid #CBD5E1;border-radius:6px;padding:7px 8px;text-align:center;font-weight:600;background:#F8FAFC'>%</div>", unsafe_allow_html=True)
        with rp_col:
            st.selectbox("Repeat (Optional)", ["Does not repeat", "Daily", "Weekly", "Monthly"], key="ae_rep")

        ev_desc = st.text_area(
            "Description (Optional)",
            placeholder="Enter event description, meeting agenda, or additional details...",
            height=95,
            key="ae_desc",
        )
        st.markdown(f"<div style='text-align:right;font-size:11px;color:#64748B;margin-top:-6px;margin-bottom:10px'>{len(ev_desc)}/500</div>", unsafe_allow_html=True)

        f_sp, f_can, f_cr = st.columns([2.2, 1, 1.2])
        with f_can:
            if st.button("Cancel", key="ae_can_btn", use_container_width=True):
                st.session_state.pop("cal_modal", None)
                st.rerun()
        with f_cr:
            if st.button("Create Event", type="primary", key="ae_save_btn", use_container_width=True):
                if not str(ev_title).strip():
                    st.error("Event title is required.")
                elif ev_end_d < ev_start_d:
                    st.error("End date cannot be earlier than start date.")
                else:
                    _event_project = st.session_state.get("ae_proj")
                    _event_type_clean = str(ev_type).split(" ", 1)[-1]
                    _selected_task_name = st.session_state.get("ae_task", "")
                    _selected_task_rows = T[T["name"].astype(str) == str(_selected_task_name)] if not T.empty else pd.DataFrame()
                    _event_task_id = int(_selected_task_rows.iloc[0]["id"]) if not _selected_task_rows.empty and _selected_task_name != "Select task (optional)" else None
                    _event_team = st.session_state.get("ae_team", "")
                    _event_resources = ", ".join(st.session_state.get("ae_res", []))
                    _event_alloc = pd.to_numeric(pd.Series([st.session_state.get("ae_alloc", 0)]), errors="coerce").fillna(0).iloc[0]
                    _event_id = sql("INSERT INTO calendar_events(title,project_id,task_id,team,event_type,start_date,end_date,start_time,end_time,resources,allocation,repeat_rule,description) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (str(ev_title).strip(), int(pmap.get(_event_project, 0)) or None, _event_task_id, _event_team, _event_type_clean, ev_start_d.isoformat(), ev_end_d.isoformat(), st.session_state.get("ae_st", ""), st.session_state.get("ae_et", ""), _event_resources, float(_event_alloc), st.session_state.get("ae_rep", "Does not repeat"), ev_desc))
                    log_event("Created", "Calendar Event", _event_id, ev_title)
                    st.session_state.pop("cal_modal", None)
                    st.rerun()

    # --- CALENDAR TOP HEADER & FILTER BAR ---
    ch_l, ch_f1, ch_f2, ch_f3, ch_f4, ch_btn = st.columns([2.15, 0.72, 0.72, 0.72, 0.78, 0.68])
    with ch_l:
        st.markdown(
            clean_html(
                f"""
                <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Calendar</div>
                <div style='font-size:13.5px;color:#64748B;margin-top:2px'>View project timelines, task schedules, team availability, and resource assignments.</div>
                """
            ),
            unsafe_allow_html=True,
        )
    with ch_f1:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Project</div>", unsafe_allow_html=True)
        cal_proj_f = st.selectbox("Project", ["All Projects"] + projects_list, key="cal_f_proj", label_visibility="collapsed")
    with ch_f2:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Team</div>", unsafe_allow_html=True)
        cal_team_f = st.selectbox("Team", ["All Teams"] + teams_list, key="cal_f_team", label_visibility="collapsed")
    with ch_f3:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Resource</div>", unsafe_allow_html=True)
        cal_res_f = st.selectbox("Resource", ["All Resources"] + sorted(R.name.dropna().astype(str).unique().tolist()), key="cal_f_res", label_visibility="collapsed")
    with ch_f4:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Event Type</div>", unsafe_allow_html=True)
        cal_type_f = st.selectbox("Event Type", ["All Event Types", "Project Assignment", "Meeting / Event", "Leave", "Milestone"], key="cal_f_type", label_visibility="collapsed")
    with ch_btn:
        st.markdown("<div style='font-size:11px;margin-bottom:2px'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("＋ Add Event", type="primary", key="cal_top_add_btn", use_container_width=True):
            st.session_state.cal_modal = "add_event"
            st.rerun()

    st.write("")
    # --- SUB-TOOLBAR ROW: Month | Week | Day | Agenda + < May 2024 > + Today + Search + More Filters ---
    if "cal_view_mode" not in st.session_state:
        st.session_state.cal_view_mode = "Month"

    ct_v, ct_nav, ct_tod, ct_sp, ct_q, ct_mf = st.columns([1.55, 1.0, 0.48, 0.65, 1.55, 0.65])
    with ct_v:
        tile_toggle("", ["Month", "Week", "Day", "Agenda"], key="cal_view_mode", default="Month")
    if "cal_calendar_month" not in st.session_state:
        st.session_state["cal_calendar_month"] = date.today().replace(day=1)
    with ct_nav:
        _nav_l, _nav_label, _nav_r = st.columns([0.45, 2.2, 0.45])
        with _nav_l:
            if st.button("‹", key="cal_prev_month", use_container_width=True):
                _m = st.session_state["cal_calendar_month"]
                st.session_state["cal_calendar_month"] = (_m.replace(day=1) - timedelta(days=1)).replace(day=1)
                st.rerun()
        with _nav_label:
            st.markdown(f"<div style='text-align:center;font-size:14.5px;font-weight:800;color:#0F172A;padding:7px 0'>{st.session_state['cal_calendar_month'].strftime('%B %Y')}</div>", unsafe_allow_html=True)
        with _nav_r:
            if st.button("›", key="cal_next_month", use_container_width=True):
                _m = st.session_state["cal_calendar_month"]
                _next = (_m.replace(day=28) + timedelta(days=4)).replace(day=1)
                st.session_state["cal_calendar_month"] = _next
                st.rerun()
    with ct_tod:
        if st.button("Today", key="cal_tod_btn", use_container_width=True):
            st.session_state["cal_selected_date"] = date.today()
            st.session_state["cal_calendar_month"] = date.today().replace(day=1)
            st.rerun()
    with ct_q:
        cal_search_q = st.text_input("Search", placeholder="🔍 Search events, tasks, or team members...", key="cal_q_box", label_visibility="collapsed")
    with ct_mf:
        if st.button("▽ More Filters", key="cal_more_btn", use_container_width=True):
            st.session_state["cal_q_box"] = ""
            st.rerun()

    st.write("")
    # --- MAIN SPLIT: LEFT (Calendar Grid + 4 Bottom Cards) | RIGHT (Thu, May 9, 2024 Inspector) ---
    main_l, main_r = st.columns([2.75, 1.1])

    with main_l:
        _month = st.session_state.get("cal_calendar_month", date.today().replace(day=1))
        _month_start = _month.replace(day=1)
        _grid_start = _month_start - timedelta(days=(_month_start.weekday() + 1) % 7)
        _cal_event_map = {}
        # Task start/end dates and milestones are visible without requiring manually-created calendar entries.
        if not T.empty:
            for _, _tr in T.iterrows():
                _d = _as_iso_date(_tr.get("start_date")) or _as_iso_date(_tr.get("due_date"))
                if not _d: continue
                try: _day = date.fromisoformat(_d)
                except Exception: continue
                _title = str(_tr.get("name") or "Task")
                _proj = str(_tr.get("project") or "")
                _assignee = str(_tr.get("assignee") or "")
                if cal_proj_f != "All Projects" and _proj != cal_proj_f: continue
                if cal_team_f != "All Teams" and str(_tr.get("team") or "") != cal_team_f: continue
                if cal_res_f != "All Resources" and _assignee != cal_res_f: continue
                if cal_type_f not in ("All Event Types", "Project Assignment"): continue
                if cal_search_q and cal_search_q.lower() not in (_title + " " + _proj + " " + _assignee).lower(): continue
                _cal_event_map.setdefault(_day, []).append(("● " + _title, f"{_proj} · {_assignee}".strip(" ·"), "g"))
        if not M.empty and cal_type_f in ("All Event Types", "Milestone"):
            for _, _mr in M.iterrows():
                _d = _as_iso_date(_mr.get("due_date"))
                if not _d: continue
                try: _day = date.fromisoformat(_d)
                except Exception: continue
                _title = str(_mr.get("name") or "Milestone"); _proj = str(_mr.get("project") or "")
                if cal_proj_f != "All Projects" and _proj != cal_proj_f: continue
                if cal_search_q and cal_search_q.lower() not in (_title + " " + _proj).lower(): continue
                _cal_event_map.setdefault(_day, []).append(("◆ " + _title, _proj, "b"))
        _events_db = df("SELECT ce.*, p.name AS project FROM calendar_events ce LEFT JOIN projects p ON p.id=ce.project_id ORDER BY ce.start_date")
        if not _events_db.empty:
            for _, _er in _events_db.iterrows():
                try: _day = date.fromisoformat(str(_er.get("start_date"))[:10])
                except Exception: continue
                _proj = str(_er.get("project") or "")
                if cal_proj_f != "All Projects" and _proj != cal_proj_f: continue
                if cal_team_f != "All Teams" and str(_er.get("team") or "") != cal_team_f: continue
                if cal_res_f != "All Resources" and cal_res_f not in str(_er.get("resources") or ""): continue
                if cal_type_f != "All Event Types" and str(_er.get("event_type") or "") != cal_type_f: continue
                _title = str(_er.get("title") or "Event")
                if cal_search_q and cal_search_q.lower() not in (_title + " " + _proj + " " + str(_er.get("resources") or "")).lower(): continue
                _cal_event_map.setdefault(_day, []).append(("● " + _title, f"{_er.get('start_time') or ''} · {_proj}".strip(" ·"), "b"))
        cal_weeks = []
        for _week in range(6):
            _cells = []
            for _day_offset in range(7):
                _day = _grid_start + timedelta(days=_week*7 + _day_offset)
                _cells.append((str(_day.day), _day.month == _month.month, _cal_event_map.get(_day, [])))
            cal_weeks.append(_cells)

        ev_box_tones = {
            "g": ("#DCFCE7", "#065F46"),
            "b": ("#DBEAFE", "#1D4ED8"),
            "r": ("#FEE2E2", "#B91C1C"),
        }

        cal_rows_html = ""
        for wk in cal_weeks:
            row_cells = ""
            for d_str, in_m, evs in wk:
                is_sel = in_m and d_str == str(st.session_state.get("cal_selected_date", date.today()).day) and _month.month == st.session_state.get("cal_selected_date", date.today()).month
                bdr = "2px solid #2563EB" if is_sel else "1px solid #EDF2F7"
                bg = "#EFF6FF" if is_sel else ("#FFFFFF" if in_m else "#F8FAFC")
                d_col = "#0F172A" if in_m else "#94A3B8"
                pills_mk = ""
                for etitle, esub, etone in evs:
                    if cal_search_q and cal_search_q.lower() not in etitle.lower():
                        continue
                    pbg, pfg = ev_box_tones[etone]
                    pills_mk += f'<div style="background:{pbg};color:{pfg};border-radius:6px;padding:4px 7px;font-size:10.5px;margin-top:4px;line-height:1.3"><div style="font-weight:700">{etitle}</div><div style="font-size:9.8px;opacity:0.9">{esub}</div></div>'
                row_cells += f'<div style="min-height:104px;border:{bdr};background:{bg};padding:6px 7px"><div style="font-size:12px;font-weight:700;color:{d_col}">{d_str}</div>{pills_mk}</div>'
            cal_rows_html += f'<div style="display:grid;grid-template-columns:repeat(7,1fr)">{row_cells}</div>'

        st.markdown(
            clean_html(
                f"""
                <div style="border:1px solid #E2E8F0;border-radius:12px;overflow:hidden;background:white;margin-bottom:12px">
                    <div style="display:grid;grid-template-columns:repeat(7,1fr);background:#F8FAFC;border-bottom:1px solid #E2E8F0;text-align:center;font-size:12px;font-weight:700;color:#1E293B;padding:9px 0">
                        <div>Sun</div><div>Mon</div><div>Tue</div><div>Wed</div><div>Thu</div><div>Fri</div><div>Sat</div>
                    </div>
                    {cal_rows_html}
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        # Bottom 4 Filter / Legend Cards (Exact Match to Image 23)
        bc1, bc2, bc3, bc4 = st.columns(4)
        with bc1, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px">Resources</div>
                    <div style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 8px;font-size:11.5px;color:#64748B;margin-bottom:8px;background:#F8FAFC">🔍 Search team member...</div>
                    <div style="font-size:12px;line-height:2.0;color:#1E293B">
                        <div><span style="color:#1473E6">☑</span> &nbsp;Select All</div>
                        <div style="background:#EFF6FF;border-radius:4px;padding:0 4px"><span style="color:#1473E6">☑</span> &nbsp;James Miller</div>
                        <div style="background:#EFF6FF;border-radius:4px;padding:0 4px"><span style="color:#1473E6">☑</span> &nbsp;Sarah Kim</div>
                        <div style="background:#EFF6FF;border-radius:4px;padding:0 4px"><span style="color:#1473E6">☑</span> &nbsp;Alice Morgan</div>
                        <div><span style="color:#94A3B8">☐</span> &nbsp;Ryan Cruz</div>
                        <div style="font-size:11.5px;font-weight:600;color:#334155;margin-top:2px">Show More</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with bc2, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px">Event Types</div>
                    <div style="font-size:12px;line-height:2.15;color:#1E293B">
                        <div><span style="color:#1473E6">☑</span> &nbsp;Select All</div>
                        <div><span style="color:#1473E6">■</span> &nbsp;Project Assignment</div>
                        <div><span style="color:#EF4444">■</span> &nbsp;Meeting / Event</div>
                        <div><span style="color:#EF4444">■</span> &nbsp;Leave</div>
                        <div><span style="color:#9333EA">■</span> &nbsp;Milestone</div>
                        <div><span style="color:#94A3B8">■</span> &nbsp;Non-Working Day</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with bc3, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px">Projects</div>
                    <div style="font-size:12px;line-height:2.15;color:#1E293B">
                        <div><span style="color:#1473E6">☑</span> &nbsp;Select All</div>
                        <div><span style="color:#1473E6">☑</span> &nbsp;<span style="color:#00875A">●</span> Website Revamp</div>
                        <div><span style="color:#1473E6">☐</span> &nbsp;<span style="color:#1473E6">●</span> Marketing Campaign</div>
                        <div><span style="color:#9333EA">☐</span> &nbsp;<span style="color:#9333EA">●</span> IT Infrastructure Upgrade</div>
                        <div><span style="color:#9333EA">☐</span> &nbsp;<span style="color:#EF4444">●</span> Sustainability Initiative</div>
                        <div><span style="color:#EA580C">☐</span> &nbsp;<span style="color:#F28C00">●</span> Internal Operations</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with bc4, st.container(border=True):
            st.markdown(
                clean_html(
                    """
                    <div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:6px">Calendar Legend</div>
                    <div style="font-size:12px;line-height:2.15;color:#1E293B">
                        <div><span style="color:#00875A;font-size:14px">●</span> &nbsp;Project Assignment</div>
                        <div><span style="color:#1473E6;font-size:14px">●</span> &nbsp;Meeting / Event</div>
                        <div><span style="color:#EF4444;font-size:14px">●</span> &nbsp;Leave</div>
                        <div><span style="color:#9333EA;font-size:14px">●</span> &nbsp;Milestone</div>
                        <div><span style="color:#94A3B8;font-size:14px">●</span> &nbsp;Non-Working Day</div>
                        <div><span style="display:inline-block;width:11px;height:11px;border:2px solid #2563EB;border-radius:3px;vertical-align:middle"></span> &nbsp;Selected Date</div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )

    # --- RIGHT SIDE INSPECTOR PANEL (Thu, May 9, 2024) ---
    with main_r, st.container(border=True):
        st.markdown(
            clean_html(
                """
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
                    <span style="font-size:18px;font-weight:800;color:#0F172A">Thu, May 9, 2024</span>
                    <span style="color:#64748B;font-size:14px;cursor:pointer">✕</span>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

        insp_tabs = st.tabs(["Day View", "Team (5)", "Projects (3)"])
        with insp_tabs[0]:
            st.markdown(
                clean_html(
                    f"""
                    <div style="font-size:12px;padding-top:4px">
                        <div style="display:grid;grid-template-columns:58px 1fr;gap:10px;padding-bottom:12px;border-bottom:1px solid #F1F5F9;margin-bottom:10px">
                            <div style="color:#475569;font-weight:600;line-height:1.4">9:00 AM<br>5:00 PM</div>
                            <div>
                                <div style="display:flex;justify-content:space-between;align-items:center">
                                    <span style="font-weight:700;color:#0F172A;font-size:13px"><span style="color:#00875A">●</span> &nbsp;Website Revamp (50%)</span>
                                    {badge_html('On Track').replace('On Track', 'Project Task')}
                                </div>
                                <div style="color:#475569;padding-left:16px;margin-top:3px;line-height:1.5">
                                    UI/UX design and wireframes<br>Assignee: James Miller<br>Project: Website Revamp
                                </div>
                            </div>
                        </div>
                        <div style="display:grid;grid-template-columns:58px 1fr;gap:10px;padding-bottom:12px;border-bottom:1px solid #F1F5F9;margin-bottom:10px">
                            <div style="color:#475569;font-weight:600;line-height:1.4">10:00 AM<br>11:00 AM</div>
                            <div>
                                <div style="display:flex;justify-content:space-between;align-items:center">
                                    <span style="font-weight:700;color:#0F172A;font-size:13px"><span style="color:#1473E6">●</span> &nbsp;Team Meeting</span>
                                    {badge_html('In Progress').replace('In Progress', 'Meeting')}
                                </div>
                                <div style="color:#475569;padding-left:16px;margin-top:3px;line-height:1.5">
                                    Weekly project sync<br>Attendees: James Miller, Sarah Kim, Alice Morgan
                                </div>
                            </div>
                        </div>
                        <div style="display:grid;grid-template-columns:58px 1fr;gap:10px;padding-bottom:8px">
                            <div style="color:#475569;font-weight:600;line-height:1.4">2:00 PM<br>4:00 PM</div>
                            <div>
                                <div style="display:flex;justify-content:space-between;align-items:center">
                                    <span style="font-weight:700;color:#0F172A;font-size:13px"><span style="color:#9333EA">●</span> &nbsp;Design Review</span>
                                    {badge_html('In Progress').replace('In Progress', 'Meeting')}
                                </div>
                                <div style="color:#475569;padding-left:16px;margin-top:3px;line-height:1.5">
                                    Review UI concepts<br>Attendees: UX Team
                                </div>
                            </div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with insp_tabs[1]:
            st.caption("5 team members scheduled on Thu, May 9, 2024.")
        with insp_tabs[2]:
            st.caption("3 active projects scheduled on Thu, May 9, 2024.")

        st.markdown(
            clean_html(
                f"""
                <hr style="margin:8px 0 12px 0;border-color:#E2E8F0">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px">
                    <span style="font-size:13.5px;font-weight:700;color:#0F172A">Team Availability (May 9)</span>
                    <span style="font-size:11.5px;color:#1473E6;font-weight:600">View All</span>
                </div>
                <div style="font-size:12px;line-height:2.2">
                    <div style="display:grid;grid-template-columns:1.35fr 0.75fr 1fr;align-items:center;gap:6px;margin-bottom:4px">
                        <div><span class="vk-avatar" style="background:#0284C7">JM</span>James Miller</div>
                        <div>{badge_html('Available')}</div>
                        <div>{progress_html(50, GREEN)}</div>
                    </div>
                    <div style="display:grid;grid-template-columns:1.35fr 0.75fr 1fr;align-items:center;gap:6px;margin-bottom:4px">
                        <div><span class="vk-avatar" style="background:#0284C7">SK</span>Sarah Kim</div>
                        <div>{badge_html('Available')}</div>
                        <div>{progress_html(75, GREEN)}</div>
                    </div>
                    <div style="display:grid;grid-template-columns:1.35fr 0.75fr 1fr;align-items:center;gap:6px;margin-bottom:4px">
                        <div><span class="vk-avatar" style="background:#0284C7">AL</span>Alice Morgan</div>
                        <div>{badge_html('Available')}</div>
                        <div>{progress_html(50, GREEN)}</div>
                    </div>
                    <div style="display:grid;grid-template-columns:1.35fr 0.75fr 1fr;align-items:center;gap:6px;margin-bottom:4px">
                        <div><span class="vk-avatar" style="background:#0369A1">RJ</span>Ryan Cruz</div>
                        <div>{badge_html('Partial')}</div>
                        <div>{progress_html(80, BLUE)}</div>
                    </div>
                    <div style="display:grid;grid-template-columns:1.35fr 0.75fr 1fr;align-items:center;gap:6px">
                        <div><span class="vk-avatar" style="background:#0284C7">MT</span>Michael Tan</div>
                        <div>{badge_html('On Leave')}</div>
                        <div>{progress_html(0, GRAY)}</div>
                    </div>
                </div>
                <hr style="margin:12px 0;border-color:#E2E8F0">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
                    <span style="font-size:13.5px;font-weight:700;color:#0F172A">Projects on this Date (2)</span>
                    <span style="font-size:11.5px;color:#1473E6;font-weight:600">View All</span>
                </div>
                <div style="font-size:12px;line-height:2.1;color:#334155;margin-bottom:10px">
                    <div style="display:flex;justify-content:space-between">
                        <span><span style="color:#00875A">●</span> &nbsp;<b>Website Revamp</b></span>
                        <span style="color:#64748B">4 members &nbsp; 50% avg allocation</span>
                    </div>
                    <div style="display:flex;justify-content:space-between">
                        <span><span style="color:#9333EA">●</span> &nbsp;<b>IT Infrastructure Upgrade</b></span>
                        <span style="color:#64748B">3 members &nbsp; 20% avg allocation</span>
                    </div>
                </div>
                <hr style="margin:10px 0;border-color:#E2E8F0">
                <div style="font-size:13.5px;font-weight:700;color:#0F172A;margin-bottom:8px">Quick Actions</div>
                """
            ),
            unsafe_allow_html=True,
        )

        qa1, qa2, qa3, qa4 = st.columns(4)
        with qa1:
            if st.button("📅\nAdd Event", key="qa_add_ev", use_container_width=True):
                st.session_state.cal_modal = "add_event"
                st.rerun()
        with qa2:
            if st.button("👥\nView Team", key="qa_v_team", use_container_width=True):
                st.session_state.page = "Resources"
                st.rerun()
        with qa3:
            if st.button("🕒\nSet Leave", key="qa_set_lv", use_container_width=True):
                st.session_state.cal_modal = "add_event"
                st.rerun()
        with qa4:
            if st.button("⚙\nAvailability", key="qa_ed_av", use_container_width=True):
                st.session_state.page = "Resources"
                st.rerun()

    if st.session_state.get("cal_modal") == "add_event":
        add_event_modal()
        
# ===========================================================================
# PAGE 7: REPORTS (Exact Replica of All 7 Report Mockup Screens)
# ===========================================================================
elif page == "Reports":
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
        rep_date_range = st.date_input("Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="rep_top_dates", label_visibility="collapsed")
    with rh_f2:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Project</div>", unsafe_allow_html=True)
        rep_proj_f = st.selectbox("Project", ["All Projects"] + projects_list, key="rep_top_proj", label_visibility="collapsed")
    with rh_f3:
        st.markdown("<div style='font-size:11px;font-weight:600;color:#475569;margin-bottom:2px'>Team</div>", unsafe_allow_html=True)
        rep_team_f = st.selectbox("Team", ["All Teams"] + teams_list, key="rep_top_team", label_visibility="collapsed")
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

    if isinstance(rep_date_range, (tuple, list)) and len(rep_date_range) == 2:
        _rep_lo, _rep_hi = sorted(rep_date_range)
        _pstart = pd.to_datetime(P.get("start_date", pd.Series(index=P.index,dtype=str)), errors="coerce").dt.date
        _pend = pd.to_datetime(P.get("end_date", pd.Series(index=P.index,dtype=str)), errors="coerce").dt.date
        P = P[(_pstart.isna() | (_pstart <= _rep_hi)) & (_pend.isna() | (_pend >= _rep_lo))].copy()
        _tstart = pd.to_datetime(T.get("start_date", pd.Series(index=T.index,dtype=str)), errors="coerce").dt.date
        _tend = pd.to_datetime(T.get("due_date", pd.Series(index=T.index,dtype=str)), errors="coerce").dt.date
        T = T[(_tstart.isna() | (_tstart <= _rep_hi)) & (_tend.isna() | (_tend >= _rep_lo))].copy()
        _mstart = pd.to_datetime(M.get("start_date", pd.Series(index=M.index,dtype=str)), errors="coerce").dt.date
        _mend = pd.to_datetime(M.get("due_date", pd.Series(index=M.index,dtype=str)), errors="coerce").dt.date
        M = M[(_mstart.isna() | (_mstart <= _rep_hi)) & (_mend.isna() | (_mend >= _rep_lo))].copy()

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
        pg_btns = "".join(
            f"<span style='border-radius:6px;padding:4px 10px;background:{'#00875A' if p == active_p else 'white'};"
            f"color:{'white' if p == active_p else '#0F172A'};border:{'none' if p == active_p else '1px solid #CBD5E1'};"
            f"font-weight:700;font-size:12px'>{p}</span>"
            for p in pages
        )
        st.markdown(
            clean_html(
                f"""
                <div style="display:flex;align-items:center;justify-content:space-between;padding-top:10px;font-size:12px;color:#475569">
                    <div>{html.escape(left_text)}</div>
                    <div style="display:flex;align-items:center;gap:5px">
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 9px;background:white">‹</span>
                        {pg_btns}
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:4px 9px;background:white">›</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

    # =========================================================================
    # REPORT TAB 1: OVERVIEW (Exact Match to Images 30 & 31)
    # =========================================================================
    if active_rtab == "Overview":
        ov_k1, ov_k2, ov_k3, ov_k4 = st.columns(4)
        with ov_k1:
            st.markdown(
                clean_html(
                    """
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:22px">💼</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Projects</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">12</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> On Track</span><b>8</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> At Risk</span><b>3</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> Delayed</span><b>1</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k2:
            st.markdown(
                clean_html(
                    """
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:22px">📋</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Tasks</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">180</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:115px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Completed</span><b>120</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#1473E6">●</span> In Progress</span><b>40</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#64748B">●</span> Not Started</span><b>20</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k3:
            st.markdown(
                clean_html(
                    """
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#EDE9FE;color:#7C3AED;display:flex;align-items:center;justify-content:center;font-size:22px">👥</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Team Members</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">24</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> Available</span><b>18</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> Partial</span><b>4</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> On Leave</span><b>2</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with ov_k4:
            st.markdown(
                clean_html(
                    """
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:104px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:58px;height:58px;border-radius:50%;border:6px solid #00875A;border-left-color:#3B82F6;display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:800;color:#0F172A">68%</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B;margin-bottom:4px">Overall Progress</div>
                                <div style="font-size:11.5px;line-height:1.7">
                                    <div><span style="color:#00875A">●</span> <b>68%</b> Completed</div>
                                    <div><span style="color:#1473E6">●</span> <b>22%</b> In Progress</div>
                                    <div><span style="color:#64748B">●</span> <b>10%</b> Not Started</div>
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
                "12", "Projects",
                ["On Track", "At Risk", "Delayed"], [8, 3, 1], [GREEN, ORANGE, RED],
                [(GREEN, "On Track", "8 (67%)"), (ORANGE, "At Risk", "3 (25%)"), (RED, "Delayed", "1 (8%)")],
            )
        with r1_c2, st.container(border=True):
            section_title("Task Status")
            render_report_donut_card(
                "180", "Tasks",
                ["Completed", "In Progress", "Not Started"], [120, 40, 20], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", "120 (67%)"), (BLUE, "In Progress", "40 (22%)"), ("#94A3B8", "Not Started", "20 (11%)")],
            )
        with r1_c3, st.container(border=True):
            section_title("Team Availability")
            ta_df = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "Available": [18, 15, 18, 17, 16, 15],
                    "Partial": [4, 4, 4, 4, 5, 5],
                    "On Leave": [2, 3, 2, 3, 3, 4],
                }
            )
            fig_ta = px.bar(
                ta_df,
                x="Month",
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
            pp_items = [
                ("Website Revamp", 90),
                ("Marketing Campaign", 60),
                ("IT Infrastructure Upgrade", 40),
                ("Sustainability Initiative", 25),
                ("Internal Operations", 75),
            ]
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
            section_title("Task Completion Trend")
            tct_df = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "Completed": [25, 55, 100, 125, 155, 180],
                    "In Progress": [12, 30, 55, 72, 88, 102],
                    "Not Started": [8, 18, 28, 35, 42, 48],
                }
            )
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
            tw_bars = [
                ("JM", "#0284C7", "James Miller", 90, RED),
                ("SK", "#0284C7", "Sarah Kim", 75, ORANGE),
                ("AL", "#0F172A", "Alice Morgan", 60, BLUE),
                ("RJ", "#0F172A", "Ryan Cruz", 50, BLUE),
                ("MT", "#0284C7", "Michael Tan", 30, GREEN),
            ]
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
            um_rows = [
                ("UI/UX Design Complete", "Website Revamp", "May 10, 2024", "On Track"),
                ("Campaign Launch", "Marketing Campaign", "May 25, 2024", "At Risk"),
                ("Infrastructure Setup", "IT Infrastructure Upgrade", "Jun 15, 2024", "On Track"),
                ("Sustainability Plan", "Sustainability Initiative", "Jun 30, 2024", "Delayed"),
                ("Operations Review", "Internal Operations", "Jul 10, 2024", "On Track"),
            ]
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
            tp_rows = [
                ("Website Revamp", 90),
                ("Marketing Campaign", 60),
                ("IT Infrastructure Upgrade", 40),
                ("Sustainability Initiative", 25),
                ("Internal Operations", 75),
            ]
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
            ph_rows = [
                ("Website Revamp", 90, GREEN, "On Track"),
                ("Marketing Campaign", 60, BLUE, "At Risk"),
                ("IT Infrastructure Upgrade", 40, BLUE, "At Risk"),
                ("Sustainability Initiative", 25, RED, "Delayed"),
                ("Internal Operations", 75, GREEN, "On Track"),
            ]
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
                    """
                    <div class="vk-kpi-card" style="justify-content:space-between;padding:14px 18px;min-height:98px">
                        <div style="display:flex;align-items:center;gap:14px">
                            <div style="width:48px;height:48px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:22px">📁</div>
                            <div>
                                <div style="font-size:12px;font-weight:700;color:#1E293B">Total Projects</div>
                                <div style="font-size:28px;font-weight:800;color:#0F172A;line-height:1.15">12</div>
                            </div>
                        </div>
                        <div style="font-size:11.5px;line-height:1.85;min-width:105px">
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> On Track</span><b>8</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#F28C00">●</span> At Risk</span><b>3</b></div>
                            <div style="display:flex;justify-content:space-between"><span><span style="color:#EF4444">●</span> Delayed</span><b>1</b></div>
                        </div>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
        with pk2:
            kpi_card_mockup("Completed Projects", "5", sub_text="42%", icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with pk3:
            kpi_card_mockup("In Progress Projects", "5", sub_text="42%", icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with pk4:
            kpi_card_mockup("Not Started Projects", "2", sub_text="17%", icon_html="❚❚", icon_bg="#E2E8F0", icon_fg="#475569")

        st.write("")
        pr1_1, pr1_2, pr1_3 = st.columns(3)
        with pr1_1, st.container(border=True):
            section_title("Project Status")
            render_report_donut_card(
                "12", "Projects",
                ["On Track", "At Risk", "Delayed"], [8, 3, 1], [GREEN, ORANGE, RED],
                [(GREEN, "On Track", "8 (67%)"), (ORANGE, "At Risk", "3 (25%)"), (RED, "Delayed", "1 (8%)")],
            )
        with pr1_2, st.container(border=True):
            section_title("Project Progress")
            pp_items = [
                ("Website Revamp", 90),
                ("Marketing Campaign", 60),
                ("IT Infrastructure Upgrade", 40),
                ("Sustainability Initiative", 25),
                ("Internal Operations", 75),
            ]
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
            pst_df = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "On Track": [5, 6, 7, 7.5, 6.5, 5.5],
                    "At Risk": [2, 2, 2, 2.5, 2.5, 2.5],
                    "Delayed": [1, 2, 2, 2, 2, 2],
                }
            )
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
                "12", "Projects",
                ["Healthy", "At Risk", "Critical"], [8, 3, 1], [GREEN, ORANGE, RED],
                [(GREEN, "Healthy", "8 (67%)"), (ORANGE, "At Risk", "3 (25%)"), (RED, "Critical", "1 (8%)")],
            )
        with pr2_2, st.container(border=True):
            section_title("Milestone Status")
            render_report_donut_card(
                "25", "Milestones",
                ["Completed", "In Progress", "Not Started"], [15, 7, 3], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", "15 (60%)"), (BLUE, "In Progress", "7 (28%)"), ("#94A3B8", "Not Started", "3 (12%)")],
            )
        with pr2_3, st.container(border=True):
            section_title("Top Projects by Progress")
            tpp_items = [
                ("Website Revamp", 90),
                ("Internal Operations", 75),
                ("Marketing Campaign", 60),
                ("IT Infrastructure Upgrade", 40),
                ("Sustainability Initiative", 25),
            ]
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
            plist_rows = [
                ("Website Revamp", "On Track", "Apr 1, 2024", "Sep 30, 2024", 90, "4 / 5", "18 / 20", "Healthy"),
                ("Marketing Campaign", "At Risk", "Apr 15, 2024", "Aug 31, 2024", 60, "3 / 5", "12 / 20", "At Risk"),
                ("IT Infrastructure Upgrade", "At Risk", "May 1, 2024", "Oct 31, 2024", 40, "2 / 5", "8 / 20", "At Risk"),
                ("Sustainability Initiative", "Delayed", "May 15, 2024", "Dec 15, 2024", 25, "1 / 5", "5 / 20", "Critical"),
                ("Internal Operations", "On Track", "Jun 1, 2024", "Nov 30, 2024", 75, "3 / 5", "15 / 20", "Healthy"),
            ]
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
            render_pagination_footer("Showing 1 to 5 of 12 projects", [1, 2, 3], 1)

        with pr3_r, st.container(border=True):
            section_title("Upcoming Project Milestones", "View All")
            upm_rows = [
                ("UI/UX Design Complete", "Website Revamp", "May 10, 2024", "On Track"),
                ("Campaign Launch", "Marketing Campaign", "May 25, 2024", "At Risk"),
                ("Infrastructure Setup", "IT Infrastructure Upgrade", "Jun 15, 2024", "On Track"),
                ("Sustainability Plan", "Sustainability Initiative", "Jun 30, 2024", "Delayed"),
                ("Operations Review", "Internal Operations", "Jul 10, 2024", "On Track"),
            ]
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
        with rk1:
            kpi_card_mockup("Total Resources", str(len(R)), icon_html="👥", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with rk2:
            kpi_card_mockup("Available", str(int(R["status"].astype(str).str.lower().isin(["available", "active"]).sum()) if not R.empty else 0), sub_text=f"{(int(R["status"].astype(str).str.lower().isin(["available", "active"]).sum())/len(R)):.0%}" if len(R) else "0%", icon_html="👤", icon_bg="#D1FAE5", icon_fg="#00875A")
        with rk3:
            kpi_card_mockup("Partial Allocation", str(int((R["status"].astype(str).str.lower()=="partial").sum()) if not R.empty else 0), icon_html="◔", icon_bg="#FEF3C7", icon_fg="#F28C00")
        with rk4:
            kpi_card_mockup("On Leave", str(int((R["status"].astype(str).str.lower()=="on leave").sum()) if not R.empty else 0), icon_html="👤", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        rr1_1, rr1_2, rr1_3 = st.columns(3)
        with rr1_1, st.container(border=True):
            section_title("Resource Status")
            render_report_donut_card(
                "24", "Resources",
                ["Available", "Partial", "On Leave"], [18, 4, 2], [GREEN, ORANGE, RED],
                [(GREEN, "Available", "18 (75%)"), (ORANGE, "Partial", "4 (17%)"), (RED, "On Leave", "2 (8%)")],
            )
        with rr1_2, st.container(border=True):
            section_title("Team Distribution")
            td_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Members": [8, 6, 4, 4, 2],
                }
            )
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
                yaxis=dict(title=None, range=[0, 10.5]),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_td, use_container_width=True, config={"displayModeBar": False})

        with rr1_3, st.container(border=True):
            section_title("Assignments by Project")
            abp_items = [
                ("Website Revamp", 8, 88, GREEN),
                ("Internal Operations", 6, 68, "#0EA5E9"),
                ("Marketing Campaign", 5, 55, "#38BDF8"),
                ("IT Infrastructure Upgrade", 3, 36, PURPLE),
                ("Sustainability Initiative", 2, 24, ORANGE),
            ]
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
            rat_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Available": [5, 3.5, 3, 3.5, 2],
                    "Partial": [3.5, 3, 1, 1.5, 0.5],
                    "On Leave": [1.5, 1, 1, 1, 0.5],
                }
            )
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
                yaxis=dict(title=None, range=[0, 11]),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.48,
            )
            st.plotly_chart(fig_rat, use_container_width=True, config={"displayModeBar": False})

        with rr2_2, st.container(border=True):
            section_title("Resource Assignments (Count)")
            rac_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Primary Assignment": [4.5, 3.5, 2, 2, 1],
                    "Additional Assignment": [4, 3.5, 2.5, 2.5, 1],
                }
            )
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
                yaxis=dict(title=None, range=[0, 11]),
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

            rlist_rows = [
                ("JM", "#0284C7", "James Miller", "IT", "Available", 2, 8, "May 10, 2024", "—"),
                ("SK", "#0284C7", "Sarah Kim", "Design", "Available", 2, 6, "May 25, 2024", "—"),
                ("AL", "#0284C7", "Alice Morgan", "Marketing", "Partial", 1, 5, "Jun 15, 2024", "—"),
                ("RJ", "#0F172A", "Ryan Cruz", "Operations", "Available", 2, 7, "Jun 30, 2024", "—"),
                ("MT", "#0284C7", "Michael Tan", "IT", "On Leave", 0, 0, "—", "On leave until May 20"),
                ("DP", "#1D4ED8", "Daniel Perez", "Operations", "Available", 1, 4, "Jun 10, 2024", "—"),
                ("ER", "#EF4444", "Ella Ramirez", "Design", "Partial", 1, 3, "May 28, 2024", "—"),
                ("CP", "#0284C7", "Chris Padilla", "IT", "Available", 2, 6, "Jun 22, 2024", "—"),
            ]
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
            render_pagination_footer("Showing 1 to 8 of 24 resources", [1, 2, 3], 1)

        with rr3_r:
            with st.container(border=True):
                section_title("Upcoming Leave", "View All")
                ul_rows = [
                    ("MT", "#0284C7", "Michael Tan", "Vacation", "May 15, 2024", "May 20, 2024"),
                    ("AR", "#0F172A", "Anna Reyes", "Sick Leave", "May 22, 2024", "May 24, 2024"),
                    ("KS", "#0369A1", "Kevin Santos", "Vacation", "Jun 10, 2024", "Jun 14, 2024"),
                    ("LC", "#0284C7", "Lisa Chan", "Personal Leave", "Jun 25, 2024", "Jun 27, 2024"),
                ]
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
                ua_rows = [
                    ("SK", "#0284C7", "Sarah Kim", "Website Revamp", "May 1, 2024"),
                    ("RJ", "#1D4ED8", "Ryan Cruz", "IT Infrastructure Upgrade", "May 15, 2024"),
                    ("AL", "#0F172A", "Alice Morgan", "Marketing Campaign", "Jun 1, 2024"),
                    ("DP", "#0F172A", "Daniel Perez", "Internal Operations", "Jun 10, 2024"),
                ]
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
        with tk1:
            kpi_card_mockup("Total Tasks", "180", icon_html="📋", icon_bg="#D1FAE5", icon_fg="#00875A")
        with tk2:
            kpi_card_mockup("Completed Tasks", "120", sub_text="67%", icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with tk3:
            kpi_card_mockup("In Progress Tasks", "40", sub_text="22%", icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with tk4:
            kpi_card_mockup("Not Started Tasks", "20", sub_text="11%", icon_html="•••", icon_bg="#E2E8F0", icon_fg="#475569")

        st.write("")
        tr1_1, tr1_2, tr1_3 = st.columns(3)
        with tr1_1, st.container(border=True):
            section_title("Task Status")
            render_report_donut_card(
                "180", "Tasks",
                ["Completed", "In Progress", "Not Started"], [120, 40, 20], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", "120 (67%)"), (BLUE, "In Progress", "40 (22%)"), ("#94A3B8", "Not Started", "20 (11%)")],
            )
        with tr1_2, st.container(border=True):
            section_title("Tasks by Project")
            tbp_rows = [
                ("Website Revamp", 48, 25, 18, 36),
                ("Marketing Campaign", 38, 20, 15, 28),
                ("IT Infrastructure Upgrade", 28, 22, 14, 24),
                ("Sustainability Initiative", 20, 16, 14, 18),
                ("Internal Operations", 16, 12, 14, 16),
            ]
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
            render_report_donut_card(
                "180", "Tasks",
                ["High", "Medium", "Low"], [30, 90, 60], [RED, ORANGE, GREEN],
                [(RED, "High", "30 (17%)"), (ORANGE, "Medium", "90 (50%)"), (GREEN, "Low", "60 (33%)")],
            )

        tr2_1, tr2_2, tr2_3 = st.columns([1.2, 1.15, 0.85])
        with tr2_1, st.container(border=True):
            section_title("Task Trend")
            tt_df = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "Completed": [6, 15, 24, 31, 38, 45],
                    "In Progress": [3, 10, 14, 18, 21, 24],
                    "Not Started": [2, 5, 7, 8, 9, 10],
                }
            )
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
                yaxis=dict(title=None, range=[0, 52]),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_tt, use_container_width=True, config={"displayModeBar": False})

        with tr2_2, st.container(border=True):
            section_title("Tasks by Team")
            tbt_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Completed": [22, 14, 10, 11, 6],
                    "In Progress": [14, 10, 8, 8, 5],
                    "Not Started": [8, 12, 9, 8, 5],
                }
            )
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
                yaxis=dict(title=None, range=[0, 52]),
                legend_title_text="",
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.42,
            )
            st.plotly_chart(fig_tbt, use_container_width=True, config={"displayModeBar": False})

        with tr2_3, st.container(border=True):
            section_title("Overdue Tasks")
            st.markdown(
                clean_html(
                    """
                    <div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:14px;margin-bottom:12px">
                        <div style="width:40px;height:40px;border-radius:50%;background:#FEE2E2;color:#DC2626;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:800">⚠</div>
                        <div style="display:flex;align-items:baseline;gap:10px">
                            <span style="font-size:28px;font-weight:800;color:#DC2626">8</span>
                            <div style="font-size:11.5px;color:#B91C1C;line-height:1.3"><b>Overdue Tasks</b><br>4% of total</div>
                        </div>
                    </div>
                    <div style="font-size:12px;font-weight:700;color:#0F172A;margin-bottom:6px">By Priority</div>
                    <div style="font-size:12.5px;line-height:2.0">
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#EF4444">●</span> &nbsp;High</span><b>4</b></div>
                        <div style="display:flex;justify-content:space-between;border-bottom:1px solid #F1F5F9"><span><span style="color:#F28C00">●</span> &nbsp;Medium</span><b>3</b></div>
                        <div style="display:flex;justify-content:space-between"><span><span style="color:#00875A">●</span> &nbsp;Low</span><b>1</b></div>
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

            tlist_rows = [
                ("UI/UX Design Complete", "Website Revamp", "Design", "Completed", "High", "Apr 1, 2024", "Apr 30, 2024", "Alice Morgan (AL)"),
                ("Homepage Development", "Website Revamp", "IT", "In Progress", "High", "Apr 10, 2024", "May 15, 2024", "James Miller (JM)"),
                ("Content Review", "Marketing Campaign", "Marketing", "In Progress", "Medium", "Apr 15, 2024", "May 20, 2024", "Sarah Kim (SK)"),
                ("Infrastructure Setup", "IT Infrastructure Upgrade", "IT", "Not Started", "High", "May 1, 2024", "Jun 15, 2024", "Michael Tan (MT)"),
                ("Sustainability Plan", "Sustainability Initiative", "Operations", "Completed", "Medium", "May 5, 2024", "May 30, 2024", "Ryan Cruz (RJ)"),
                ("Process Documentation", "Internal Operations", "Operations", "In Progress", "Low", "May 10, 2024", "Jun 20, 2024", "Daniel Perez (DP)"),
                ("Training Materials", "Marketing Campaign", "Marketing", "Not Started", "Medium", "May 20, 2024", "Jun 30, 2024", "Ella Ramirez (ER)"),
                ("QA Testing", "Website Revamp", "IT", "In Progress", "High", "May 25, 2024", "Jun 25, 2024", "Chris Padilla (CP)"),
            ]
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
            render_pagination_footer("Showing 1 to 8 of 180 tasks", [1, 2, 3, 4, 5], 1)

        with tr3_r:
            with st.container(border=True):
                section_title("Upcoming Tasks (Next 14 Days)", "View All")
                ut_rows = [
                    ("Content Review", "Marketing Campaign", "May 20, 2024", "Medium"),
                    ("Infrastructure Setup", "IT Infrastructure Upgrade", "Jun 15, 2024", "High"),
                    ("Process Documentation", "Internal Operations", "Jun 20, 2024", "Low"),
                    ("QA Testing", "Website Revamp", "Jun 25, 2024", "High"),
                    ("Training Materials", "Marketing Campaign", "Jun 30, 2024", "Medium"),
                ]
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
                rc_rows = [
                    ("UI/UX Design Complete", "Website Revamp", "Apr 30, 2024"),
                    ("Sustainability Plan", "Sustainability Initiative", "May 30, 2024"),
                    ("Requirements Gathering", "Internal Operations", "Jun 2, 2024"),
                    ("Design Review", "Marketing Campaign", "Jun 5, 2024"),
                    ("System Configuration", "IT Infrastructure Upgrade", "Jun 8, 2024"),
                ]
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
        with mk1:
            kpi_card_mockup("Total Milestones", "25", icon_html="⚑", icon_bg="#D1FAE5", icon_fg="#00875A")
        with mk2:
            kpi_card_mockup("Completed", "15", sub_text="60%", icon_html="✓", icon_bg="#D1FAE5", icon_fg="#00875A")
        with mk3:
            kpi_card_mockup("In Progress", "7", sub_text="28%", icon_html="▷", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with mk4:
            kpi_card_mockup("Not Started", "3", sub_text="12%", icon_html="🕒", icon_bg="#E2E8F0", icon_fg="#475569")
        with mk5:
            kpi_card_mockup("Overdue", "4", sub_text="16%", icon_html="⚠", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        mr1_1, mr1_2, mr1_3 = st.columns(3)
        with mr1_1, st.container(border=True):
            section_title("Milestone Status")
            render_report_donut_card(
                "25", "Milestones",
                ["Completed", "In Progress", "Not Started"], [15, 7, 3], [GREEN, BLUE, "#94A3B8"],
                [(GREEN, "Completed", "15 (60%)"), (BLUE, "In Progress", "7 (28%)"), ("#94A3B8", "Not Started", "3 (12%)")],
            )
        with mr1_2, st.container(border=True):
            section_title("Milestones by Project")
            mbp_rows = [
                ("Website Revamp", 40, 25, 35, 5),
                ("Marketing Campaign", 40, 38, 22, 5),
                ("IT Infrastructure Upgrade", 40, 28, 32, 5),
                ("Sustainability Initiative", 48, 22, 30, 5),
                ("Internal Operations", 30, 28, 42, 5),
            ]
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
            mdd_df = pd.DataFrame(
                {
                    "Month": ["Apr", "May", "Jun", "Jul", "Aug", "Sep"],
                    "Completed": [1.5, 2.2, 3.5, 3, 2, 2.8],
                    "In Progress": [0.5, 1.5, 2.5, 1.5, 2.5, 2.5],
                    "Not Started": [1.5, 0, 1.5, 3, 2.8, 0],
                    "Overdue": [0, 1.8, 0, 0, 1.2, 2],
                }
            )
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
                yaxis=dict(title=None, range=[0, 10.5]),
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
            um30_rows = [
                ("Infrastructure Setup", "IT Infrastructure Upgrade", "May 15, 2024", "In Progress", "3", "#FEF3C7", "#D97706"),
                ("Campaign Launch", "Marketing Campaign", "May 20, 2024", "In Progress", "8", "#FEF3C7", "#D97706"),
                ("Sustainability Plan", "Sustainability Initiative", "May 25, 2024", "Not Started", "13", "#DCFCE7", "#15803D"),
                ("UI/UX Design Complete", "Website Revamp", "May 28, 2024", "In Progress", "16", "#DCFCE7", "#15803D"),
                ("Process Documentation", "Internal Operations", "May 30, 2024", "Not Started", "18", "#DCFCE7", "#15803D"),
            ]
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
            om_rows = [
                ("QA Testing", "Website Revamp", "Apr 20, 2024", 34),
                ("Content Review", "Marketing Campaign", "Apr 25, 2024", 29),
                ("Training Materials", "Marketing Campaign", "May 5, 2024", 19),
                ("Requirements Gathering", "Internal Operations", "May 10, 2024", 14),
            ]
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

            mlist_rows = []
            if not M.empty:
                for _, _mrow in M.iterrows():
                    _completion = _as_iso_date(_mrow.get("completion_date")) or "—"
                    mlist_rows.append((str(_mrow.get("name") or ""), str(_mrow.get("project") or ""), str(_mrow.get("team") or ""), str(_mrow.get("status") or "Not Started"), fmt_date(_as_iso_date(_mrow.get("start_date"))), fmt_date(_as_iso_date(_mrow.get("due_date"))), fmt_date(_completion) if _completion != "—" else "—", str(_mrow.get("dependency") or "—"), str(_mrow.get("notes") or "")))
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
            render_pagination_footer("Showing 1 to 8 of 25 milestones", [1, 2, 3, 4], 1)

    # =========================================================================
    # REPORT TAB 6: WORKLOAD (Exact Match to Image 27)
    # =========================================================================
    elif active_rtab == "Workload":
        wk1, wk2, wk3, wk4, wk5 = st.columns(5)
        with wk1:
            kpi_card_mockup("Total Team Members", "24", icon_html="👥", icon_bg="#DBEAFE", icon_fg="#1473E6")
        with wk2:
            kpi_card_mockup("Total Active Tasks", "180", icon_html="❖", icon_bg="#D1FAE5", icon_fg="#00875A")
        with wk3:
            kpi_card_mockup("Balanced Workload", "16", sub_text="67%", icon_html="👤", icon_bg="#D1FAE5", icon_fg="#00875A")
        with wk4:
            kpi_card_mockup("High Workload", "5", sub_text="21%", icon_html="△", icon_bg="#FEF3C7", icon_fg="#F28C00")
        with wk5:
            kpi_card_mockup("Overloaded", "3", sub_text="12%", icon_html="⚠", icon_bg="#FEE2E2", icon_fg="#EF4444")

        st.write("")
        wr1_1, wr1_2, wr1_3 = st.columns(3)
        with wr1_1, st.container(border=True):
            section_title("Workload Distribution")
            render_report_donut_card(
                "24", "Team Members",
                ["Balanced", "High Workload", "Overloaded"], [16, 5, 3], [GREEN, ORANGE, RED],
                [(GREEN, "Balanced", "16 (67%)"), (ORANGE, "High Workload", "5 (21%)"), (RED, "Overloaded", "3 (12%)")],
            )
        with wr1_2, st.container(border=True):
            section_title("Task Distribution by Team")
            tdt_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Completed": [18, 12, 13, 12, 6],
                    "In Progress": [10, 7, 9, 6, 6],
                    "Not Started": [8, 6, 8, 5, 3],
                    "High Priority": [9, 6, 7, 7, 3],
                }
            )
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
            wbp_rows = [
                ("Website Revamp", 36, 20, 18, 20, 42),
                ("Marketing Campaign", 30, 22, 16, 14, 36),
                ("IT Infrastructure Upgrade", 26, 20, 16, 12, 32),
                ("Sustainability Initiative", 24, 18, 15, 9, 28),
                ("Internal Operations", 20, 13, 16, 8, 22),
            ]
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
            twc_df = pd.DataFrame(
                {
                    "Team": ["IT", "Design", "Operations", "Marketing", "Others"],
                    "Available Capacity": [3.2, 2.3, 2.8, 2.0, 2.6],
                    "Assigned Tasks": [2.8, 1.8, 3.2, 1.5, 2.0],
                    "Remaining Capacity": [3.5, 3.0, 2.0, 2.5, 3.0],
                }
            )
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
                yaxis=dict(title=None, range=[0, 11]),
                legend_title_text="",
                legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=10.5)),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                bargap=0.45,
            )
            st.plotly_chart(fig_twc, use_container_width=True, config={"displayModeBar": False})

        with wr2_2, st.container(border=True):
            section_title("Individual Workload (Top 10)")
            iw_items = [
                ("James Miller", 18, 92, RED),
                ("Sarah Kim", 16, 82, RED),
                ("Alice Morgan", 14, 70, ORANGE),
                ("Ryan Cruz", 12, 60, ORANGE),
                ("Michael Tan", 10, 48, GREEN),
                ("Daniel Perez", 9, 43, GREEN),
                ("Ella Ramirez", 8, 38, GREEN),
                ("Chris Padilla", 7, 32, GREEN),
                ("Anna Reyes", 6, 26, GREEN),
                ("Kevin Santos", 6, 26, GREEN),
            ]
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
            wsp_rows = [
                ("Website Revamp", 50, 35, 15),
                ("Marketing Campaign", 42, 42, 16),
                ("IT Infrastructure Upgrade", 53, 27, 20),
                ("Sustainability Initiative", 50, 35, 15),
                ("Internal Operations", 44, 38, 18),
            ]
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

            wlist_rows = [
                ("JM", "#0284C7", "James Miller", "IT", 18, 6, 8, 4, 6, "Overloaded"),
                ("SK", "#0284C7", "Sarah Kim", "Design", 16, 5, 8, 3, 4, "Overloaded"),
                ("AL", "#0F172A", "Alice Morgan", "Marketing", 14, 5, 6, 3, 4, "High"),
                ("RJ", "#0F172A", "Ryan Cruz", "Operations", 12, 4, 6, 2, 3, "High"),
                ("MT", "#0284C7", "Michael Tan", "IT", 10, 5, 4, 1, 2, "Balanced"),
                ("DP", "#1D4ED8", "Daniel Perez", "Operations", 9, 4, 3, 2, 1, "Balanced"),
                ("ER", "#EF4444", "Ella Ramirez", "Design", 8, 3, 3, 2, 1, "Balanced"),
                ("CP", "#0284C7", "Chris Padilla", "IT", 7, 3, 2, 2, 1, "Balanced"),
            ]
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
                    cr_date_range = st.date_input("Filter Date Range", (date.today() - timedelta(days=30), date.today() + timedelta(days=60)), key="cr_flt_dates", label_visibility="collapsed")

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
        _custom_tasks = T.copy()
        if isinstance(cr_date_range, (tuple, list)) and len(cr_date_range) == 2:
            _cr_lo, _cr_hi = sorted(cr_date_range)
            _cr_sd = pd.to_datetime(_custom_tasks.get("start_date", pd.Series(index=_custom_tasks.index,dtype=str)), errors="coerce").dt.date
            _cr_ed = pd.to_datetime(_custom_tasks.get("due_date", pd.Series(index=_custom_tasks.index,dtype=str)), errors="coerce").dt.date
            _custom_tasks = _custom_tasks[(_cr_sd.isna() | (_cr_sd <= _cr_hi)) & (_cr_ed.isna() | (_cr_ed >= _cr_lo))].copy()
        custom_proj_df = pd.DataFrame([
            {"Project Name": pr.get("name", ""), "Project Status": pr.get("status", "Not Started"),
             "Total Tasks": int((_custom_tasks["project_id"] == pr.get("id")).sum()) if not _custom_tasks.empty else 0,
             "Completed": int(((_custom_tasks["project_id"] == pr.get("id")) & (_custom_tasks["status"].astype(str).str.lower() == "completed")).sum()) if not _custom_tasks.empty else 0,
             "In Progress": int(((_custom_tasks["project_id"] == pr.get("id")) & (_custom_tasks["status"].astype(str).str.lower() == "in progress")).sum()) if not _custom_tasks.empty else 0,
             "Not Started": int(((_custom_tasks["project_id"] == pr.get("id")) & (_custom_tasks["status"].astype(str).str.lower() == "not started")).sum()) if not _custom_tasks.empty else 0,
             "Overdue": int(((_custom_tasks["project_id"] == pr.get("id")) & (pd.to_datetime(_custom_tasks["due_date"], errors="coerce").dt.date < date.today()) & (_custom_tasks["status"].astype(str).str.lower() != "completed")).sum()) if not _custom_tasks.empty else 0,
             "Total Milestones": int((M["project_id"] == pr.get("id")).sum()) if not M.empty else 0,
             "Team Members": int((R["team"] == pr.get("team")).sum()) if not R.empty else 0}
            for _, pr in P.iterrows()])
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
                        """
                        <div style="font-size:14.5px;font-weight:800;color:#0F172A;margin-bottom:10px;padding-top:4px">Summary</div>
                        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:20px">📁</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Projects</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">5</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:20px">☰</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Tasks</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">180</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#D1FAE5;color:#00875A;display:flex;align-items:center;justify-content:center;font-size:20px">⚑</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Total Milestones</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">25</div>
                                </div>
                            </div>
                            <div style="border:1px solid #E2ECF5;border-radius:12px;padding:16px;background:white;display:flex;align-items:center;gap:12px">
                                <div style="width:44px;height:44px;min-width:44px;border-radius:50%;background:#DBEAFE;color:#1473E6;display:flex;align-items:center;justify-content:center;font-size:20px">👥</div>
                                <div>
                                    <div style="font-size:12px;font-weight:600;color:#334155">Team Members</div>
                                    <div style="font-size:26px;font-weight:800;color:#0F172A;line-height:1.15">24</div>
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
            render_pagination_footer(f"Showing 1 to {len(tbl_df)} of {len(tbl_df)} projects", [1], 1)

        with bot_r, st.container(border=True):
            sr_h1, sr_h2 = st.columns([1.4, 1.1])
            with sr_h1:
                st.markdown("<div style='font-size:15.5px;font-weight:800;color:#0F172A;padding-top:6px'>Saved Reports</div>", unsafe_allow_html=True)
            with sr_h2:
                if st.button("＋ Save Current Report", type="primary", key="cr_save_cur_btn", use_container_width=True):
                    sql(
                        "INSERT INTO saved_reports(name,report_type,config,last_generated,created_by) VALUES(?,?,?,?,?)",
                        ("Project Status Summary", "Project", "{}", "Jun 30, 2024", "A. Escabillas"),
                    )
                    log_event("Saved", "Custom Report", None, "Project Status Summary")
                    st.toast("Current report saved!")

            saved_rows = [
                ("Project Status Summary", "Project", "Jun 30, 2024", "A. Escabillas"),
                ("Team Workload Overview", "Resource", "Jun 28, 2024", "A. Escabillas"),
                ("Overdue Tasks Report", "Task", "Jun 25, 2024", "A. Escabillas"),
                ("Upcoming Milestones", "Milestone", "Jun 20, 2024", "A. Escabillas"),
                ("Marketing Campaign Details", "Project", "Jun 18, 2024", "A. Escabillas"),
            ]
            sr_tr = "".join(
                f"<tr>"
                f"<td><b style='color:#0F172A'>{r_nm}</b></td>"
                f"<td>{r_tp}</td>"
                f"<td>{r_dt}</td>"
                f"<td>{r_by}</td>"
                f"<td style='white-space:nowrap'>"
                f"<span style='border:1px solid #CBD5E1;border-radius:4px;padding:2px 5px;margin-right:3px;font-size:11px'>👁</span>"
                f"<span style='border:1px solid #CBD5E1;border-radius:4px;padding:2px 5px;margin-right:3px;font-size:11px'>✎</span>"
                f"<span style='border:1px solid #FECACA;background:#FEF2F2;color:#DC2626;border-radius:4px;padding:2px 5px;font-size:11px'>🗑</span>"
                f"</td></tr>"
                for r_nm, r_tp, r_dt, r_by in saved_rows
            )
            st.markdown(
                clean_html(
                    f"""
                    <div class="vk-table-wrap">
                        <table class="vk-table">
                            <thead>
                                <tr>
                                    <th>Report Name</th>
                                    <th>Type</th>
                                    <th>Last Generated</th>
                                    <th>Created By</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>{sr_tr}</tbody>
                        </table>
                    </div>
                    """
                ),
                unsafe_allow_html=True,
            )
            render_pagination_footer("Showing 1 to 5 of 5 reports", [1], 1)


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
            u_tm = st.selectbox("Team", teams_list, key="set_add_u_tm")
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

    st.markdown(
        clean_html(
            f"""
            <div style='font-size:28px;font-weight:800;color:{TEXT_COLOR};letter-spacing:-0.03em;line-height:1.15'>Settings</div>
            <div style='font-size:13.5px;color:#64748B;margin-top:2px;margin-bottom:14px'>Manage workspace, users, defaults, notifications, integrations, and system preferences.</div>
            """
        ),
        unsafe_allow_html=True,
    )

    # Workbook is the sole operational source; visual styling and existing settings layout remain unchanged.
    with st.container(border=True):
        st.markdown("**Portfolio Excel Source**")
        st.caption("Import or refresh all project, task, subtask, owner, date, status, priority, and detail fields from one Portfolio sheet.")
        _src_col, _clear_col = st.columns([2, 1])
        with _src_col:
            _source_xlsx = st.file_uploader("Upload Portfolio workbook", type=["xlsx", "xlsm"], key="settings_portfolio_excel_source")
            if _source_xlsx is not None and st.button("Import / Refresh Data", type="primary", key="settings_import_excel_source"):
                try:
                    _import_summary = import_portfolio_excel(_source_xlsx)
                    st.success(f"Imported {_import_summary['projects']} projects, {_import_summary['tasks']} task/subtask rows, and {_import_summary['owners']} owners.")
                    st.rerun()
                except Exception as _import_exc:
                    st.error(f"Import failed: {_import_exc}")
        with _clear_col:
            if st.button("🗑 Clear Database Data", key="settings_open_clear_data", use_container_width=True):
                st.session_state["show_clear_data_modal"] = True
            if st.session_state.get("show_clear_data_modal"):
                @st.dialog("Clear Database Data", width="large")
                def _clear_data_selection_modal():
                    st.markdown("<div style='display:flex;align-items:center;gap:10px;margin-bottom:8px'><span style='display:inline-flex;width:34px;height:34px;border-radius:50%;align-items:center;justify-content:center;background:#FEE2E2;color:#B91C1C;font-size:18px'>🗑</span><span style='font-size:19px;font-weight:800;color:#0F172A'>Clear Database Data</span></div>", unsafe_allow_html=True)
                    st.warning("Select the data categories to permanently clear from the local database and shared snapshot. App preferences are retained.")
                    _clear_projects = st.checkbox("Projects", value=True, key="clear_sel_projects")
                    _clear_tasks = st.checkbox("Tasks and subtasks", value=True, key="clear_sel_tasks")
                    _clear_resources = st.checkbox("Resources / owners", value=True, key="clear_sel_resources")
                    _clear_milestones = st.checkbox("Milestones", value=True, key="clear_sel_milestones")
                    _clear_events = st.checkbox("Calendar events", value=True, key="clear_sel_events")
                    _clear_reports = st.checkbox("Saved reports", value=True, key="clear_sel_reports")
                    _clear_audit = st.checkbox("Audit history", value=False, key="clear_sel_audit")
                    _ca, _cb = st.columns(2)
                    with _ca:
                        if st.button("Cancel", key="clear_data_cancel", use_container_width=True):
                            st.session_state.pop("show_clear_data_modal", None)
                            st.rerun()
                    with _cb:
                        if st.button("Confirm Clear", type="primary", key="clear_data_confirm", use_container_width=True):
                            _chosen = []
                            if _clear_projects: _chosen.append("projects")
                            if _clear_tasks: _chosen.append("tasks")
                            if _clear_resources: _chosen.append("resources")
                            if _clear_milestones: _chosen.append("milestones")
                            if _clear_events: _chosen.append("calendar_events")
                            if _clear_reports: _chosen.append("saved_reports")
                            if _clear_audit: _chosen.append("audit_log")
                            clear_operational_data(_chosen)
                            st.session_state.pop("show_clear_data_modal", None)
                            st.success("Selected database data cleared.")
                            st.rerun()
                _clear_data_selection_modal()

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

        with ws_c2:
            theme_opts = ["🟩 HPE (Green)", "🟦 Blue", "⬛ Dark"]
            cur_theme_val = get_setting("theme", "HPE Green")
            t_idx = 1 if cur_theme_val == "Blue" else (2 if cur_theme_val == "Dark" else 0)
            sel_theme = st.selectbox("Theme", theme_opts, index=t_idx, key="set_ws_theme")
            mapped_theme = "Blue" if "Blue" in sel_theme else ("Dark" if "Dark" in sel_theme else "HPE Green")
            if mapped_theme != cur_theme_val:
                put_setting("theme", mapped_theme)
                st.rerun()

            st.selectbox("Language", ["English", "Filipino"], index=0, key="set_ws_lang")
            st.selectbox("Date Format", ["Apr 24, 2024", "24/04/2024", "04/24/2024"], index=0, key="set_ws_dfmt")
            land_opts = ["Overview", "Projects", "Gantt Chart", "Task List", "Resources", "Calendar", "Reports"]
            cur_land = get_setting("default_landing", "Overview")
            sel_land = st.selectbox("Default Landing Page", land_opts, index=land_opts.index(cur_land) if cur_land in land_opts else 0, key="set_ws_land")
            if sel_land != cur_land:
                put_setting("default_landing", sel_land)
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
        if not R.empty:
            for _, _ur in R.iterrows():
                _nm = str(_ur.get("name", "")); _email = str(_ur.get("email", ""))
                _role = str(_ur.get("role", "Team Member")); _team = str(_ur.get("team", ""))
                _status = str(_ur.get("status", "Available"))
                _initials = "".join(part[0] for part in _nm.split()[:2]).upper() or "U"
                um_users.append((_initials, "#0284C7", _nm, _email, _role, _team, _status))

        if um_q:
            um_users = [u for u in um_users if um_q.lower() in u[2].lower() or um_q.lower() in u[3].lower() or um_q.lower() in u[4].lower()]

        um_tr = "".join(
            f"<tr>"
            f"<td style='padding:7px 10px'><span class='vk-avatar' style='background:{abg}'>{ini}</span><b>{nm}</b></td>"
            f"<td style='padding:7px 10px;color:#334155'>{em}</td>"
            f"<td style='padding:7px 10px'>{rl}</td>"
            f"<td style='padding:7px 10px'>{tm}</td>"
            f"<td style='padding:7px 10px'><span class='vk-badge {'vk-badge-green' if st_s == 'Active' else 'vk-badge-red'}' style='min-width:64px'>{st_s}</span></td>"
            f"<td style='padding:7px 10px;white-space:nowrap'>"
            f"<span style='color:#1E293B;margin-right:10px;cursor:pointer'>✎</span>"
            f"<span style='color:#EF4444;cursor:pointer'>🗑</span>"
            f"</td></tr>"
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
                <div style="display:flex;align-items:center;justify-content:space-between;padding-top:10px;font-size:12px;color:#475569">
                    <div>Showing 1 to {len(um_users)} of 24 users</div>
                    <div style="display:flex;align-items:center;gap:4px">
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 8px;background:white">‹</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 8px;background:white">‹</span>
                        <span style="border-radius:6px;padding:3px 10px;background:#00875A;color:white;font-weight:700">1</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 9px;background:white;font-weight:600">2</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 9px;background:white;font-weight:600">3</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 9px;background:white;font-weight:600">3</span>
                        <span style="border:1px solid #CBD5E1;border-radius:6px;padding:3px 8px;background:white">›</span>
                    </div>
                </div>
                """
            ),
            unsafe_allow_html=True,
        )

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
        st.selectbox("Default Project Status", ["Not Started", "On Track", "At Risk", "Delayed"], index=0, key="set_ps_stat")
        st.selectbox("Default Project Template", ["Standard Project", "Software Delivery", "Marketing Campaign", "Infrastructure"], index=0, key="set_ps_tmpl")
        st.toggle("Enable project key/ID auto-generation", value=True, key="set_ps_t1")
        st.toggle("Require project description", value=True, key="set_ps_t2")
        st.toggle("Allow project attachments", value=True, key="set_ps_t3")
        st.toggle("Enable project archiving", value=False, key="set_ps_t4")

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
        st.selectbox("Default Task Status", ["Not Started", "In Progress", "Completed"], index=0, key="set_ts_stat")
        st.selectbox("Default Task Priority", ["Medium", "High", "Low"], index=0, key="set_ts_prio")
        st.toggle("Enable task dependencies", value=True, key="set_ts_t1")
        st.toggle("Allow subtasks", value=True, key="set_ts_t2")
        st.toggle("Enable recurring tasks", value=False, key="set_ts_t3")
        st.toggle("Require assignee", value=True, key="set_ts_t4")
        st.toggle("Allow task attachments", value=True, key="set_ts_t5")

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
        st.selectbox("Default Milestone Status", ["Not Started", "In Progress", "Completed"], index=0, key="set_ms_stat")
        st.write("")
        st.toggle("Enable milestones", value=True, key="set_ms_t1")
        st.toggle("Require due date", value=True, key="set_ms_t2")
        st.toggle("Allow milestone dependencies", value=True, key="set_ms_t3")
        st.toggle("Auto-notify on milestone updates", value=True, key="set_ms_t4")

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
                st.toggle(n_lbl, value=n_def, key=f"set_nf_{idx}")
        with nf_r:
            st.markdown("<div style='padding-top:26px'></div>", unsafe_allow_html=True)
            st.selectbox("Notification Frequency", ["Real-time", "Hourly", "Daily digest", "Weekly digest"], index=0, key="set_nf_freq")

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
            ("Microsoft 365", "Connect to sync calendar and teams.", "⊞", "#EA580C"),
            ("Slack", "Receive notifications in Slack.", "❖", "#E11D48"),
            ("Teams", "Connect for team collaboration.", "👥", "#4F46E5"),
            ("OneDrive", "Access project files from OneDrive.", "☁", "#0284C7"),
        ]
        for s_nm, s_ds, s_ic, s_cl in int_services:
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
                is_c = get_setting(f"int_{s_nm}", "false") == "true"
                if st.button("Connected" if is_c else "Connect", key=f"set_int_{s_nm}", type="primary" if is_c else "secondary", use_container_width=True):
                    put_setting(f"int_{s_nm}", "false" if is_c else "true")
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
        st.toggle("Enable audit logs", value=True, key="set_sys_t1")
        st.toggle("Allow data export", value=True, key="set_sys_t2")
        st.toggle("Enable automatic backups", value=True, key="set_sys_t3")
        st.toggle("Enable maintenance mode", value=False, key="set_sys_t4")

        st_l, st_r = st.columns([1.1, 1.3])
        with st_l:
            st.markdown("<div style='font-size:12.5px;font-weight:700;color:#0F172A;padding-top:10px'>Session Timeout</div>", unsafe_allow_html=True)
        with st_r:
            st.selectbox("Session Timeout", ["8 hours", "4 hours", "2 hours", "1 hour", "30 minutes"], index=0, key="set_sys_to", label_visibility="collapsed")

        st.write("")
        if st.button("🗑 Clear Cache", key="set_clear_cache_btn"):
            st.cache_data.clear()
            st.toast("Application cache cleared!")

    if st.session_state.get("settings_modal") == "add_user":
        add_user_settings_modal()


# ---------------------------------------------------------------------------
# Deployment notes and lightweight runtime integrity checks
# ---------------------------------------------------------------------------
# Configure Streamlit Secrets (or environment variable) without committing credentials:
# .streamlit/secrets.toml
# MONGODB_URI = "mongodb+srv://<db-user>:<password>@<cluster-host>/?retryWrites=true&w=majority"
# The URI is intentionally never embedded in this source file.
